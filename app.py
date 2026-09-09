import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import os
import io
import requests
import hashlib
import html
from decimal import Decimal, ROUND_HALF_UP
from datetime import date, datetime
from urllib.parse import quote

# =========================================================
# 0. ROUNDING & CLOUD PERSISTENCE UTILITIES
# =========================================================
LOCAL_CACHE_FILE = "persistent_scm_data.xlsx"
DEFAULT_SUPABASE_BUCKET = "scm-dashboard"
DEFAULT_SUPABASE_OBJECT = "persistent_scm_data.xlsx"

XLSX_SIGNATURES = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")
XLS_SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"

def detect_excel_engine(file_path_or_buffer):
    signature = b""
    if isinstance(file_path_or_buffer, (str, os.PathLike)):
        path = os.fspath(file_path_or_buffer)
        try:
            with open(path, "rb") as handle:
                signature = handle.read(8)
        except OSError as exc:
            raise ValueError(f"Cannot open workbook file: {exc}") from exc
    elif isinstance(file_path_or_buffer, (bytes, bytearray)):
        signature = bytes(file_path_or_buffer[:8])
    elif hasattr(file_path_or_buffer, "read"):
        original_position = getattr(file_path_or_buffer, "tell", lambda: None)()
        if hasattr(file_path_or_buffer, "seek"):
            file_path_or_buffer.seek(0)
        signature = file_path_or_buffer.read(8)
        if hasattr(file_path_or_buffer, "seek"):
            file_path_or_buffer.seek(original_position if original_position is not None else 0)
    else:
        raise ValueError("Unsupported workbook input. Please upload a valid .xlsx or .xls file.")

    if any(signature.startswith(sig) for sig in XLSX_SIGNATURES):
        return "openpyxl"
    if signature.startswith(XLS_SIGNATURE):
        return "xlrd"

    raise ValueError("The uploaded/saved file is not a valid Excel workbook.")

def validate_excel_bytes(file_bytes):
    if not file_bytes:
        return False
    detect_excel_engine(io.BytesIO(file_bytes))
    return True

def round_half_up(value, ndigits=0):
    if pd.isna(value):
        return np.nan
    quantizer = Decimal("1").scaleb(-ndigits)
    rounded = Decimal(str(float(value))).quantize(quantizer, rounding=ROUND_HALF_UP)
    return int(rounded) if ndigits == 0 else float(rounded)

def round_series_half_up(series):
    """Optimized vectorized rounding using numpy instead of Decimal apply loops."""
    numeric = pd.to_numeric(series, errors="coerce")
    # np.floor(x + 0.5) perfectly replicates ROUND_HALF_UP for positive numbers (SCM inventory context)
    return np.floor(numeric + 0.5)

def _secret(name, default=None):
    try:
        return st.secrets[name]
    except Exception:
        return default

def get_cloud_storage_config():
    url = str(_secret("SUPABASE_URL", "") or "").strip().rstrip("/")
    key = str(_secret("SUPABASE_SECRET_KEY", "") or _secret("SUPABASE_KEY", "") or "").strip()
    bucket = str(_secret("SUPABASE_BUCKET", DEFAULT_SUPABASE_BUCKET) or DEFAULT_SUPABASE_BUCKET).strip()
    object_name = str(_secret("SUPABASE_OBJECT", DEFAULT_SUPABASE_OBJECT) or DEFAULT_SUPABASE_OBJECT).strip().lstrip("/")
    return {
        "configured": bool(url and key and bucket and object_name),
        "url": url,
        "key": key,
        "bucket": bucket,
        "object_name": object_name,
    }

def _supabase_object_endpoint(config):
    encoded_bucket = quote(config["bucket"], safe="")
    encoded_object = quote(config["object_name"], safe="/")
    return f'{config["url"]}/storage/v1/object/{encoded_bucket}/{encoded_object}'

def _supabase_auth_headers(config):
    headers = {"apikey": config["key"]}
    if not config["key"].startswith("sb_"):
        headers["Authorization"] = f'Bearer {config["key"]}'
    return headers

def download_cloud_workbook():
    config = get_cloud_storage_config()
    if not config["configured"]:
        return None
    headers = _supabase_auth_headers(config)
    response = requests.get(_supabase_object_endpoint(config), headers=headers, timeout=30)
    if response.status_code == 404:
        return None
    response.raise_for_status()
    return response.content

def upload_cloud_workbook(file_bytes):
    config = get_cloud_storage_config()
    if not config["configured"]:
        return False
    headers = _supabase_auth_headers(config)
    headers.update({
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "x-upsert": "true",
    })
    response = requests.post(_supabase_object_endpoint(config), headers=headers, data=file_bytes, timeout=60)
    if response.status_code not in (200, 201):
        raise RuntimeError(f"Cloud persistence upload failed ({response.status_code})")
    return True

def save_local_cache(file_bytes):
    with open(LOCAL_CACHE_FILE, "wb") as file_handle:
        file_handle.write(file_bytes)

def load_local_cache():
    if not os.path.exists(LOCAL_CACHE_FILE):
        return None
    with open(LOCAL_CACHE_FILE, "rb") as file_handle:
        return file_handle.read()

def initialize_persistent_workbook():
    if "scm_workbook_bytes" in st.session_state:
        return st.session_state["scm_workbook_bytes"], st.session_state.get("scm_storage_source", "Session cache")

    cloud_config = get_cloud_storage_config()
    if cloud_config["configured"]:
        try:
            cloud_bytes = download_cloud_workbook()
            if cloud_bytes:
                validate_excel_bytes(cloud_bytes)
                st.session_state["scm_workbook_bytes"] = cloud_bytes
                st.session_state["scm_storage_source"] = "Cloud • Supabase"
                try: save_local_cache(cloud_bytes)
                except Exception: pass
                return cloud_bytes, "Cloud • Supabase"
        except Exception as exc:
            st.session_state["scm_cloud_warning"] = str(exc)

    local_bytes = load_local_cache()
    if local_bytes:
        try:
            validate_excel_bytes(local_bytes)
        except Exception as exc:
            st.session_state["scm_local_warning"] = "Local cache invalid: " + str(exc)
        else:
            st.session_state["scm_workbook_bytes"] = local_bytes
            st.session_state["scm_storage_source"] = "Local cache"
            return local_bytes, "Local cache"

    return None, "No saved workbook"

# =========================================================
# 1. PAGE CONFIGURATION
# =========================================================
st.set_page_config(
    page_title="SCM Executive Control Tower",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

if "scm_theme" not in st.session_state:
    st.session_state["scm_theme"] = "dark"

def toggle_scm_theme():
    current_theme = st.session_state.get("scm_theme", "dark")
    st.session_state["scm_theme"] = "light" if current_theme == "dark" else "dark"

SCM_THEME = st.session_state["scm_theme"]
SCM_IS_DARK = SCM_THEME == "dark"

PLOTLY_TEMPLATE = "plotly_dark" if SCM_IS_DARK else "plotly_white"
PLOTLY_HOVER_BG = "#0f172a" if SCM_IS_DARK else "#ffffff"
PLOTLY_HOVER_TEXT = "#f8fafc" if SCM_IS_DARK else "#0f172a"
PLOTLY_HOVER_BORDER = "rgba(148,163,184,0.28)" if SCM_IS_DARK else "rgba(15,23,42,0.16)"

PLOTLY_CHART_TEXT = "#f8fafc" if SCM_IS_DARK else "#0f172a"
PLOTLY_CHART_MUTED = "#94a3b8" if SCM_IS_DARK else "#64748b"
PLOTLY_CHART_GRID = "rgba(148,163,184,0.12)" if SCM_IS_DARK else "rgba(15,23,42,0.08)"
PLOTLY_CHART_AXIS = "rgba(148,163,184,0.20)" if SCM_IS_DARK else "rgba(15,23,42,0.14)"
PLOTLY_BAR_CONFIG = {"displayModeBar": False, "displaylogo": False, "scrollZoom": False, "responsive": True}

# =========================================================
# 2. EXECUTIVE UI / THEME
# =========================================================
st.markdown(
    """
    <style>
        :root {
            --scm-navy: #0b1220;
            --scm-indigo: #6366f1;
            --scm-blue: #0ea5e9;
            --scm-green: #10b981;
            --scm-amber: #f59e0b;
            --scm-red: #f43f5e;
            --scm-border: rgba(148, 163, 184, 0.18);
            --scm-muted: #94a3b8;
        }

        html, body, [data-testid="stAppViewContainer"] { overflow-x: hidden !important; }

        .block-container, [data-testid="stMainBlockContainer"] {
            width: 100% !important; max-width: 100% !important; min-width: 0 !important;
            padding-top: 0.55rem !important; padding-bottom: 2.25rem !important;
            padding-left: clamp(0.5rem, 2vw, 1.5rem) !important;
            padding-right: clamp(0.5rem, 2vw, 1.5rem) !important; margin: 0 !important;
        }

        section[data-testid="stMain"] { min-width: 0 !important; max-width: none !important; margin-left: 0 !important; padding-left: 0 !important; overflow-x: clip !important; }
        
        div[data-testid="stTabs"] { width: 100% !important; min-width: 0 !important; }
        div[data-baseweb="tab-panel"] { width: 100% !important; padding: 1rem 0 0 0 !important; }
        div[data-baseweb="tab-panel"] > div, div[data-baseweb="tab-panel"] > div > div[data-testid="stVerticalBlock"] { padding: 0 !important; }

        [data-testid="stSidebar"], [data-testid="collapsedControl"], [data-testid="stToolbar"], #MainMenu, header[data-testid="stHeader"], [data-testid="stDecoration"], [data-testid="viewerBadge"] {
            display: none !important; visibility: hidden !important;
        }
        header[data-testid="stHeader"] { height: 0 !important; min-height: 0 !important; }

        .data-sync-caption { display: block; margin: 0 0 0.28rem 0; color: var(--scm-muted); font-size: 0.62rem; font-weight: 850; line-height: 1.2; letter-spacing: 0.09em; text-transform: uppercase; text-align: right; white-space: nowrap; }
        .data-sync-shell { width: 100%; min-width: 0; padding-top: 0.02rem; }

        div[data-testid="stButton"] > button { min-height: 42px; border-radius: 11px; font-weight: 800; letter-spacing: 0.01em; border: 1px solid rgba(99, 102, 241, 0.34); box-shadow: 0 5px 14px rgba(15, 23, 42, 0.06); }
        div[data-testid="stHorizontalBlock"] > div:last-child div[data-testid="stButton"] > button { min-height: 44px; border-radius: 12px; border: 1px solid rgba(99, 102, 241, 0.50); background: linear-gradient(135deg, rgba(79,70,229,0.98), rgba(37,99,235,0.96)); color: #ffffff; font-weight: 850; box-shadow: 0 8px 20px rgba(37, 99, 235, 0.18); transition: transform 150ms ease, box-shadow 150ms ease, filter 150ms ease; }
        div[data-testid="stHorizontalBlock"] > div:last-child div[data-testid="stButton"] > button:hover { transform: translateY(-1px); box-shadow: 0 11px 24px rgba(37, 99, 235, 0.24); filter: brightness(1.03); }

        div[data-testid="stDialog"] [data-testid="stFileUploader"] section { min-height: 118px !important; border-radius: 14px !important; border: 1px dashed rgba(99, 102, 241, 0.48) !important; background: rgba(99, 102, 241, 0.035) !important; }
        .import-dialog-note { padding: 0.72rem 0.85rem; border: 1px solid rgba(148, 163, 184, 0.18); border-radius: 12px; background: rgba(148, 163, 184, 0.035); color: var(--scm-muted); font-size: 0.78rem; line-height: 1.5; margin-bottom: 0.85rem; }

        /* RESPONSIVE PARETO HTML TABLES */
        .pareto-html-shell {
            width: 100%; max-width: 100%; min-width: 0;
            overflow-x: auto; /* Optimized for mobile */
            -webkit-overflow-scrolling: touch;
            border: 1px solid var(--scm-border); border-radius: 12px;
        }
        .pareto-html-table {
            width: 100%; min-width: 550px; /* Forces scroll on small screens */
            table-layout: fixed; border-collapse: collapse; font-size: 0.78rem;
        }
        .pareto-html-table thead th { padding: 0.62rem 0.56rem; text-align: left; font-size: 0.67rem; font-weight: 850; letter-spacing: 0.05em; text-transform: uppercase; color: var(--scm-muted); background: rgba(148, 163, 184, 0.055); border-bottom: 1px solid var(--scm-border); }
        .pareto-html-table tbody td { padding: 0.60rem 0.56rem; border-bottom: 1px solid rgba(148, 163, 184, 0.11); vertical-align: middle; overflow-wrap: anywhere; word-break: normal; }
        .pareto-html-table tbody tr:last-child td { border-bottom: none; }
        .pareto-html-table tbody tr:hover { background: rgba(99, 102, 241, 0.035); }
        .pareto-html-table .num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }

        .pareto-status { display: inline-flex; align-items: center; gap: 6px; max-width: 100%; min-height: 25px; padding: 4px 9px 4px 7px; border-radius: 999px; border: 1px solid rgba(148, 163, 184, 0.20); background: rgba(148, 163, 184, 0.055); color: var(--scm-muted); box-shadow: inset 0 1px 0 rgba(255,255,255,0.035); font-size: 0.64rem; font-weight: 900; line-height: 1.1; letter-spacing: 0.045em; text-transform: uppercase; white-space: nowrap; }
        .pareto-status::before { content: ""; width: 7px; height: 7px; flex: 0 0 7px; border-radius: 50%; background: #94a3b8; box-shadow: 0 0 0 3px rgba(148,163,184,0.10); }
        .pareto-status.stockout { color: #fb7185; background: rgba(244, 63, 94, 0.10); border-color: rgba(244, 63, 94, 0.30); }
        .pareto-status.stockout::before { background: #f43f5e; box-shadow: 0 0 0 3px rgba(244,63,94,0.14), 0 0 10px rgba(244,63,94,0.30); }
        .pareto-status.critical { color: #fb923c; background: rgba(249, 115, 22, 0.10); border-color: rgba(249, 115, 22, 0.30); }
        .pareto-status.critical::before { background: #f97316; box-shadow: 0 0 0 3px rgba(249,115,22,0.14); }
        .pareto-status.low { color: #fbbf24; background: rgba(245, 158, 11, 0.10); border-color: rgba(245, 158, 11, 0.30); }
        .pareto-status.low::before { background: #f59e0b; box-shadow: 0 0 0 3px rgba(245,158,11,0.14); }
        .pareto-status.ok { color: #34d399; background: rgba(16, 185, 129, 0.10); border-color: rgba(16, 185, 129, 0.30); }
        .pareto-status.ok::before { background: #10b981; box-shadow: 0 0 0 3px rgba(16,185,129,0.14); }
        .pareto-status.overstock { color: #38bdf8; background: rgba(14, 165, 233, 0.10); border-color: rgba(14, 165, 233, 0.30); }
        .pareto-status.overstock::before { background: #0ea5e9; box-shadow: 0 0 0 3px rgba(14,165,233,0.14); }

        .stock-status-legend { display: flex; align-items: center; flex-wrap: wrap; gap: 7px; margin: 0.15rem 0 0.78rem 0; padding: 9px 11px; border: 1px solid var(--scm-border); border-radius: 12px; background: rgba(148, 163, 184, 0.022); }
        .stock-status-legend-label { margin-right: 3px; color: var(--scm-muted); font-size: 0.62rem; font-weight: 900; letter-spacing: 0.075em; text-transform: uppercase; }

        .hero-shell { box-sizing: border-box; width: 100%; min-width: 0; margin: 0.08rem 0 0.92rem 0; padding: clamp(20px, 1.65vw, 28px) clamp(22px, 2.15vw, 34px); border: 1px solid rgba(99, 102, 241, 0.34); border-left: 4px solid var(--scm-indigo); border-radius: 18px; background: #0b1220; box-shadow: 0 10px 28px rgba(2, 6, 23, 0.14); overflow: visible; }
        .hero-copy { display: block; width: 100%; min-width: 0; max-width: 100%; }
        .hero-kicker { color: #a5b4fc; font-size: clamp(0.62rem, 0.66vw, 0.72rem); font-weight: 850; line-height: 1.40 !important; letter-spacing: 0.12em; text-transform: uppercase; margin: 0 0 10px 0 !important; }
        .hero-title { color: #f8fafc; font-size: clamp(1.65rem, 2.10vw, 2.45rem); font-weight: 900; line-height: 1.13 !important; letter-spacing: -0.025em; margin: 0 0 12px 0 !important; }
        .hero-subtitle { max-width: 1180px; color: #cbd5e1; font-size: clamp(0.80rem, 0.86vw, 0.94rem); font-weight: 500; line-height: 1.55 !important; margin: 0 !important; }
        
        .drp-action-btn { display: inline-flex; align-items: center; margin-top: 18px; padding: 8px 18px; background: rgba(99, 102, 241, 0.12); color: #a5b4fc; border: 1px solid rgba(99, 102, 241, 0.4); border-radius: 8px; font-size: 0.72rem; font-weight: 850; text-decoration: none; letter-spacing: 0.05em; text-transform: uppercase; transition: all 0.2s ease; }
        .drp-action-btn:hover { background: rgba(99, 102, 241, 0.25); color: #ffffff; border-color: rgba(99, 102, 241, 0.8); transform: translateY(-1px); }

        .section-heading { display: grid; grid-template-columns: auto minmax(0, 1fr) minmax(0, auto); align-items: center; gap: 9px; min-width: 0; margin: 0.44rem 0 0.72rem 0; padding: 0.18rem 0.12rem; }
        .section-heading .dot { width: 9px; height: 9px; border-radius: 50%; background: var(--scm-indigo); box-shadow: 0 0 0 4px rgba(99, 102, 241, 0.11); }
        .section-heading .title { font-size: 1.12rem; font-weight: 900; letter-spacing: -0.012em; }
        .section-heading .subtitle { color: var(--scm-muted); font-size: 0.75rem; font-weight: 550; text-align: right; white-space: normal; overflow-wrap: anywhere; }
        .trend-heading-inline { min-height: 62px; margin: 0.20rem 0 0.18rem 0 !important; align-content: center; }

        .metric-card, .metric-card-base { position: relative; box-sizing: border-box; width: 100%; min-width: 0; min-height: 120px; height: 100%; padding: clamp(10px, 1.5vw, 15px) clamp(12px, 1.5vw, 17px); border: 1px solid var(--scm-border); border-radius: 15px; background: rgba(148, 163, 184, 0.025); box-shadow: 0 7px 20px rgba(15, 23, 42, 0.045); overflow: hidden; }
        .metric-card::before, .metric-card-base::before { content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 3px; background: var(--scm-indigo); opacity: 0.90; }
        .metric-card-base { display: flex; flex-direction: column; justify-content: space-between; }
        .metric-title, .metric-header { color: var(--scm-muted); font-size: 0.67rem; font-weight: 850; text-transform: uppercase; letter-spacing: 0.075em; }
        .metric-header { display: flex; justify-content: space-between; align-items: center; gap: 8px; }
        .metric-value, .metric-value-sm { font-size: clamp(1.4rem, 1.85vw, 2.05rem); font-weight: 900; line-height: 1.05; margin: 9px 0 6px 0; letter-spacing: -0.025em; }
        .metric-footnote { color: var(--scm-muted); font-size: 0.69rem; line-height: 1.35; }

        .badge { display: inline-block; padding: 4px 8px; border-radius: 999px; font-size: 0.60rem; font-weight: 850; text-transform: uppercase; letter-spacing: 0.045em; }
        .badge-red { background: rgba(244, 63, 94, 0.12); color: #fb7185; }
        .badge-yellow { background: rgba(245, 158, 11, 0.12); color: #fbbf24; }
        .badge-green { background: rgba(16, 185, 129, 0.12); color: #34d399; }
        .badge-blue { background: rgba(14, 165, 233, 0.12); color: #38bdf8; }
        .icon-box { display: flex; align-items: center; justify-content: center; width: 28px; height: 28px; flex: 0 0 28px; border-radius: 8px; background: rgba(99, 102, 241, 0.09); }

        .status-strip { display: flex; flex-wrap: wrap; align-items: center; gap: 6px; width: 100%; min-width: 0; margin: 0.45rem 0 0.12rem 0; padding: 8px 10px; border: 1px solid var(--scm-border); border-radius: 12px; background: rgba(148, 163, 184, 0.018); }
        .status-pill, .info-chip { display: inline-flex; align-items: center; gap: 5px; min-width: 0; padding: 4px 8px; border: 1px solid var(--scm-border); border-radius: 999px; font-size: 0.67rem; font-weight: 700; color: var(--scm-muted); line-height: 1.25; }
        .info-chip { margin-right: 5px; margin-bottom: 4px; }
        
        div[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 15px !important; border-color: var(--scm-border) !important; box-shadow: 0 6px 18px rgba(15, 23, 42, 0.035); background: rgba(148, 163, 184, 0.012); }
        div[data-testid="stVerticalBlockBorderWrapper"] > div { min-width: 0 !important; }
        div[data-testid="stPlotlyChart"] { width: 100% !important; min-width: 0 !important; border-radius: 13px; overflow: hidden; }

        .pareto-panel-spacer { height: 0.28rem; }
        .pareto-header { display: flex; align-items: center; justify-content: space-between; gap: 10px; width: 100%; min-width: 0; padding: 0.58rem 0 0.54rem 0; border-bottom: 2px solid; margin-bottom: 0.62rem; }
        .pareto-count { flex: 0 0 auto; white-space: nowrap; border: 1px solid rgba(128,128,128,0.22); padding: 4px 8px; border-radius: 999px; font-size: 0.68rem; font-weight: 850; }
        hr { border: none; border-top: 1px solid var(--scm-border); margin: 1.10rem 0; }
    </style>
    """,
    unsafe_allow_html=True,
)

if SCM_IS_DARK:
    theme_tokens = { "page_bg": "#070b14", "surface": "#0b1220", "surface_2": "#0f172a", "card_bg": "rgba(15, 23, 42, 0.72)", "input_bg": "#0f172a", "text": "#f8fafc", "muted": "#94a3b8", "border": "rgba(148, 163, 184, 0.18)", "soft_border": "rgba(148, 163, 184, 0.11)", "table_head": "rgba(148, 163, 184, 0.055)", "hover": "rgba(99, 102, 241, 0.07)", "hero_bg": "#0b1220", "hero_title": "#f8fafc", "hero_subtitle": "#cbd5e1", "hero_kicker": "#a5b4fc", "shadow": "0 10px 28px rgba(2, 6, 23, 0.22)", }
else:
    theme_tokens = { "page_bg": "#f6f8fc", "surface": "#ffffff", "surface_2": "#f8fafc", "card_bg": "#ffffff", "input_bg": "#ffffff", "text": "#0f172a", "muted": "#64748b", "border": "rgba(15, 23, 42, 0.12)", "soft_border": "rgba(15, 23, 42, 0.08)", "table_head": "#f8fafc", "hover": "rgba(79, 70, 229, 0.055)", "hero_bg": "#ffffff", "hero_title": "#0f172a", "hero_subtitle": "#475569", "hero_kicker": "#4f46e5", "shadow": "0 10px 28px rgba(15, 23, 42, 0.08)", }

st.markdown(
    f"""
    <style>
        :root {{
            --scm-page-bg: {theme_tokens["page_bg"]}; --scm-surface: {theme_tokens["surface"]}; --scm-surface-2: {theme_tokens["surface_2"]};
            --scm-card-bg: {theme_tokens["card_bg"]}; --scm-input-bg: {theme_tokens["input_bg"]}; --scm-text: {theme_tokens["text"]};
            --scm-muted: {theme_tokens["muted"]}; --scm-border: {theme_tokens["border"]}; --scm-soft-border: {theme_tokens["soft_border"]};
            --scm-table-head: {theme_tokens["table_head"]}; --scm-hover: {theme_tokens["hover"]}; --scm-hero-bg: {theme_tokens["hero_bg"]};
            --scm-hero-title: {theme_tokens["hero_title"]}; --scm-hero-subtitle: {theme_tokens["hero_subtitle"]};
            --scm-hero-kicker: {theme_tokens["hero_kicker"]}; --scm-theme-shadow: {theme_tokens["shadow"]};
        }}
        html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {{ background: var(--scm-page-bg) !important; color: var(--scm-text) !important; }}
        [data-testid="stMainBlockContainer"], [data-testid="stMarkdownContainer"] *, [data-testid="stWidgetLabel"] * {{ color: var(--scm-text); }}
        [data-testid="stCaptionContainer"] * {{ color: var(--scm-muted) !important; }}
        .hero-shell {{ background: var(--scm-hero-bg) !important; border-color: var(--scm-border) !important; border-left-color: var(--scm-indigo) !important; box-shadow: var(--scm-theme-shadow) !important; }}
        .hero-title {{ color: var(--scm-hero-title) !important; }} .hero-subtitle {{ color: var(--scm-hero-subtitle) !important; }} .hero-kicker {{ color: var(--scm-hero-kicker) !important; }}
        .metric-card, .metric-card-base, div[data-testid="stVerticalBlockBorderWrapper"] {{ background: var(--scm-card-bg) !important; border-color: var(--scm-border) !important; box-shadow: var(--scm-theme-shadow) !important; color: var(--scm-text) !important; }}
        .metric-value, .metric-value-sm, .section-heading .title {{ color: var(--scm-text) !important; }}
        .metric-title, .metric-header, .metric-footnote, .section-heading .subtitle {{ color: var(--scm-muted) !important; }}
        div[data-baseweb="select"] > div, div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div {{ background: var(--scm-input-bg) !important; border-color: var(--scm-border) !important; color: var(--scm-text) !important; }}
        div[data-baseweb="select"] *, div[data-baseweb="input"] *, div[data-baseweb="textarea"] * {{ color: var(--scm-text) !important; }}
        [data-baseweb="popover"] [role="listbox"], [data-baseweb="popover"] ul {{ background: var(--scm-surface) !important; color: var(--scm-text) !important; }}
        [data-baseweb="popover"] [role="option"] {{ color: var(--scm-text) !important; }} [data-baseweb="popover"] [role="option"]:hover {{ background: var(--scm-hover) !important; }}
        button[data-baseweb="tab"] {{ color: var(--scm-muted) !important; }} button[data-baseweb="tab"][aria-selected="true"] {{ color: var(--scm-text) !important; }}
        [data-testid="stExpander"] {{ background: var(--scm-card-bg) !important; border-color: var(--scm-border) !important; color: var(--scm-text) !important; }}
        .pareto-html-shell {{ background: var(--scm-card-bg) !important; border-color: var(--scm-border) !important; }}
        .pareto-html-table {{ color: var(--scm-text) !important; }}
        .pareto-html-table thead th {{ color: var(--scm-muted) !important; background: var(--scm-table-head) !important; border-bottom-color: var(--scm-border) !important; }}
        .pareto-html-table tbody td {{ color: var(--scm-text) !important; border-bottom-color: var(--scm-soft-border) !important; }}
        .pareto-html-table tbody tr:hover {{ background: var(--scm-hover) !important; }}
        .status-strip {{ background: var(--scm-card-bg) !important; border-color: var(--scm-border) !important; }}
        .status-pill, .info-chip, .pareto-count, hr {{ border-color: var(--scm-border) !important; }}
        [data-testid="stDialog"] [role="dialog"] {{ background: var(--scm-surface) !important; color: var(--scm-text) !important; }}
        div[data-testid="stDialog"] [data-testid="stFileUploader"] section, .import-dialog-note {{ background: var(--scm-surface-2) !important; border-color: var(--scm-border) !important; color: var(--scm-muted) !important; }}
        
        .st-key-scm_executive_header {{ box-sizing: border-box; width: 100%; margin: 0.08rem 0 0.92rem 0; padding: clamp(20px, 1.65vw, 28px) clamp(22px, 2.15vw, 34px); border: 1px solid var(--scm-border); border-left: 4px solid var(--scm-indigo); border-radius: 18px; background: var(--scm-hero-bg); box-shadow: var(--scm-theme-shadow); }}
        .st-key-scm_header_theme_toggle {{ display: flex; justify-content: flex-end; align-items: flex-start; width: 100%; padding-top: 0.05rem; }}
        .st-key-scm_header_theme_toggle [data-testid="stButton"] {{ width: auto !important; margin-left: auto !important; }}
        .st-key-scm_theme_toggle button, .st-key-scm_header_theme_toggle button {{ width: 42px !important; height: 42px !important; border-radius: 12px !important; border: 1px solid var(--scm-border) !important; background: var(--scm-surface) !important; color: var(--scm-text) !important; box-shadow: 0 5px 14px rgba(15, 23, 42, 0.10) !important; transition: transform 150ms ease, box-shadow 150ms ease, border-color 150ms ease, background 150ms ease !important; }}
        .st-key-scm_theme_toggle button:hover {{ transform: translateY(-1px) !important; border-color: rgba(99, 102, 241, 0.62) !important; background: var(--scm-hover) !important; box-shadow: 0 8px 18px rgba(15, 23, 42, 0.14) !important; }}
        .st-key-scm_header_logo {{ display: flex; align-items: center; justify-content: flex-start; min-height: 84px; }}
        .st-key-scm_header_logo img {{ max-width: 92px !important; max-height: 76px !important; object-fit: contain !important; filter: drop-shadow(0 6px 14px rgba(2, 6, 23, 0.18)); }}
        
        @media (max-width: 760px) {{
            .st-key-scm_executive_header {{ padding: 18px 18px 20px 20px; }}
            .st-key-scm_header_logo img {{ max-width: 72px !important; max-height: 58px !important; }}
            .st-key-scm_theme_toggle button {{ width: 38px !important; height: 38px !important; border-radius: 10px !important; }}
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# EXECUTIVE HEADER WITH OPTIONAL COMPANY LOGO + THEME TOGGLE
# ---------------------------------------------------------
COMPANY_LOGO_SOURCE = str(_secret("COMPANY_LOGO", "") or os.getenv("SCM_COMPANY_LOGO", "assets/company_logo.png") or "").strip()
company_logo_available = bool(COMPANY_LOGO_SOURCE and (COMPANY_LOGO_SOURCE.lower().startswith(("http://", "https://")) or os.path.exists(COMPANY_LOGO_SOURCE)))

theme_toggle_icon = "🌙" if SCM_IS_DARK else "☀️"
theme_toggle_help = "Dark mode is active • Click to switch to Light mode" if SCM_IS_DARK else "Light mode is active • Click to switch to Dark mode"

with st.container(key="scm_executive_header"):
    if company_logo_available:
        header_logo_col, header_copy_col, header_theme_col = st.columns([1.15, 10.35, 0.5], gap="small", vertical_alignment="top")
        with header_logo_col:
            st.image(COMPANY_LOGO_SOURCE, width=92)
    else:
        header_copy_col, header_theme_col = st.columns([11.5, 0.5], gap="small", vertical_alignment="top")

    with header_copy_col:
        st.markdown(
            """
            <div class="hero-copy">
                <div class="hero-kicker">Supply Chain Management • Executive Analytics</div>
                <div class="hero-title">MUTI MC SCM Executive Control Tower</div>
                <div class="hero-subtitle">Inventory visibility, Pareto risk prioritization, stockout trends, Days of Inventory, and branch-level action monitoring.</div>
                <a href="https://scmdrp.streamlit.app/" target="_blank" class="drp-action-btn">Launch Delivery Requirements Plan (DRP) ↗</a>
            </div>
            """, unsafe_allow_html=True
        )

    with header_theme_col:
        st.button(theme_toggle_icon, key="scm_theme_toggle", help=theme_toggle_help, use_container_width=False, on_click=toggle_scm_theme)


# =========================================================
# 3. DATA IMPORT & PROCESSING
# =========================================================
@st.cache_data(show_spinner=False)
def process_excel_file(file_path_or_buffer):
    excel_engine = detect_excel_engine(file_path_or_buffer)
    try:
        xls = pd.ExcelFile(file_path_or_buffer, engine=excel_engine)
    except Exception as exc:
        raise ValueError(f"Unable to open the Excel workbook using {excel_engine}: {exc}") from exc

    sheet_map = {str(sheet).lower().strip().replace(" ", "_"): sheet for sheet in xls.sheet_names}
    def get_actual_sheet_name(target):
        if target in sheet_map: return sheet_map[target]
        raise ValueError(f"Missing sheet '{target}'")

    raw_df = pd.read_excel(xls, sheet_name=get_actual_sheet_name("raw_data"))
    raw_df.columns = raw_df.columns.astype(str).str.lower().str.strip().str.replace(" ", "_", regex=False)
    if "class" in raw_df.columns: raw_df.rename(columns={"class": "pareto_class"}, inplace=True)

    required_raw_cols = ["area", "branch", "pareto_class", "stock_status", "remaining_inventory", "suggested_transfer", "doi", "model"]
    missing_raw_cols = [c for c in required_raw_cols if c not in raw_df.columns]
    if missing_raw_cols: raise ValueError("Raw_Data is missing: " + ", ".join(missing_raw_cols))

    for col in ["remaining_inventory", "suggested_transfer", "doi"]:
        raw_df[col] = round_series_half_up(pd.to_numeric(raw_df[col], errors="coerce").fillna(0)).astype(int)

    raw_df["pareto_class"] = raw_df["pareto_class"].astype(str).str.strip()
    raw_df["stock_status"] = raw_df["stock_status"].fillna("").astype(str).str.strip()
    raw_df["area"] = raw_df["area"].fillna("").astype(str).str.strip()
    raw_df["branch"] = raw_df["branch"].fillna("").astype(str).str.strip()

    def parse_kpi_sheet(sheet_name):
        df_raw = pd.read_excel(xls, sheet_name=get_actual_sheet_name(sheet_name), header=None)
        empty_kpi = pd.DataFrame(columns=["period", "after_po", "per_branch", "class_a_out", "before_po", "overall_doi", "class_a_doi"])
        if df_raw.empty or df_raw.shape[1] < 2: return empty_kpi

        kpis_to_extract = {
            "MUTI MC : Stock Outrate - Overall after PO Balance": "after_po",
            "MUTI MC : Stock Outrate - Per Branch": "per_branch",
            "Overall Class A Stock Out Rate": "class_a_out",
            "MUTI MC : Stock Outrate - Overall (Before PO Balance)": "before_po",
            "MUTI MC : DoI": "overall_doi",
            "MC Class A Doi": "class_a_doi",
        }

        def parse_header_date(value):
            if pd.isna(value): return pd.NaT
            if isinstance(value, (pd.Timestamp, datetime, date, np.datetime64)): parsed = pd.to_datetime(value, errors="coerce")
            elif isinstance(value, (int, float, np.integer, np.floating)):
                num = float(value)
                if 20000 <= num <= 60000: parsed = pd.Timestamp("1899-12-30") + pd.to_timedelta(num, unit="D")
                else: return pd.NaT
            else:
                val_text = str(value).strip()
                parsed = pd.to_datetime(val_text, errors="coerce") if val_text else pd.NaT
            if pd.isna(parsed) or parsed.year < 2000 or parsed.year > 2100: return pd.NaT
            return pd.Timestamp(parsed).normalize()

        best_date_columns, best_dates = [], []
        for row_idx in range(min(10, len(df_raw))):
            r_cols, r_dates = [], []
            for col_idx in range(df_raw.shape[1]):
                p_date = parse_header_date(df_raw.iat[row_idx, col_idx])
                if pd.notna(p_date): r_cols.append(col_idx); r_dates.append(p_date)
            if len(r_cols) > len(best_date_columns): best_date_columns = r_cols; best_dates = r_dates

        if not best_date_columns: raise ValueError(f"No valid dates in '{sheet_name}'.")

        data = {"period": best_dates}
        first_col = df_raw[0].fillna("").astype(str)
        for kpi, col in kpis_to_extract.items():
            idx = df_raw.index[first_col.str.contains(kpi, regex=False, na=False)].tolist()
            data[col] = pd.to_numeric(df_raw.loc[idx[0], best_date_columns].values, errors="coerce") if idx else [np.nan] * len(best_dates)

        clean_df = pd.DataFrame(data).dropna(subset=["period"])
        metric_cols = list(kpis_to_extract.values())
        clean_df = clean_df.dropna(subset=metric_cols, how="all")

        def last_valid(series):
            nn = series.dropna()
            return nn.iloc[-1] if not nn.empty else np.nan

        clean_df = clean_df.sort_values("period").groupby("period", as_index=False).agg({c: last_valid for c in metric_cols})

        if "ytd" in sheet_name.lower() and not clean_df.empty:
            clean_df["month_year"] = clean_df["period"].dt.to_period("M")
            clean_df = clean_df.drop_duplicates(subset=["month_year"], keep="last").drop(columns=["month_year"])
            
        return clean_df.sort_values("period").reset_index(drop=True)

    kpi_ytd = parse_kpi_sheet("kpi_ytd_input")
    kpi_weekly = parse_kpi_sheet("kpi_weekly_input")

    return raw_df, kpi_ytd, kpi_weekly


# =========================================================
# 4. DATA SYNCHRONIZATION & PERSISTENCE
# =========================================================
cloud_config = get_cloud_storage_config()
saved_workbook_bytes, storage_source = initialize_persistent_workbook()

def persist_uploaded_workbook(uploaded_bytes):
    process_excel_file(io.BytesIO(uploaded_bytes))
    persistence_messages = []
    try:
        save_local_cache(uploaded_bytes)
        persistence_messages.append("local cache")
    except Exception as exc: st.warning(f"Local cache fail: {exc}")

    new_src = "Local cache"
    if cloud_config["configured"]:
        upload_cloud_workbook(uploaded_bytes)
        new_src = "Cloud • Supabase"
        persistence_messages.append("Supabase cloud")

    st.session_state["scm_workbook_bytes"] = uploaded_bytes
    st.session_state["scm_storage_source"] = new_src
    dest = " + ".join(persistence_messages) if persistence_messages else "active session"
    return f"Validated and saved to {dest}."

@st.dialog("Data Sync", width="large")
def data_sync_dialog():
    st.markdown("<div class='import-dialog-note'>Sync latest SCM Excel. Requires Sheets: Raw_Data, KPI_YTD_Input, KPI_Weekly_Input.</div>", unsafe_allow_html=True)
    d_file = st.file_uploader("Select Excel workbook", type=["xlsx", "xls"], key="scm_u")
    if not d_file: return
    f_bytes = d_file.getvalue()
    st.caption(f"Selected: {d_file.name} • {len(f_bytes)/(1024*1024):.2f} MB")
    col1, col2 = st.columns([1.25, 2.75])
    with col1: do_import = st.button("Validate & Sync", type="primary", use_container_width=True)
    with col2: st.caption("Previous workbook retained if validation fails.")

    if do_import:
        u_hash = hashlib.sha256(f_bytes).hexdigest()
        if st.session_state.get("scm_last_hash") == u_hash: st.info("Already active dataset."); return
        try:
            with st.spinner("Validating..."): success_msg = persist_uploaded_workbook(f_bytes)
        except Exception as exc: st.error(f"Sync failed: {exc}"); return
        st.session_state["scm_last_hash"] = u_hash; st.session_state["scm_import_success"] = success_msg
        st.cache_data.clear(); st.rerun()

# =========================================================
# 5. HELPERS
# =========================================================
def calculate_stockout_rate(df, pareto_class=None):
    sub = df[df["pareto_class"] == pareto_class] if pareto_class else df
    if sub.empty: return 0
    return int(round_half_up((sub["stock_status"].str.lower().str.strip() == "stockout").mean() * 100))

def section_heading(title, subtitle=""):
    st.markdown(f"<div class='section-heading'><span class='dot'></span><span class='title'>{title}</span><span class='subtitle'>{subtitle}</span></div>", unsafe_allow_html=True)

def stock_status_style_class(status_val):
    val = str(status_val or "").strip().casefold().replace("_", " ")
    if val in {"stockout", "stock out", "out of stock", "oos"}: return "stockout"
    if "critical" in val: return "critical"
    if val in {"low", "low stock", "below min", "below minimum"}: return "low"
    if val in {"ok", "normal", "healthy", "available", "in stock", "instock"}: return "ok"
    if val in {"overstock", "over stock", "excess", "excess stock"}: return "overstock"
    return "neutral"

def render_stock_status_legend(df):
    if df.empty or "stock_status" not in df.columns: return
    observed = [s for s in df["stock_status"].fillna("").astype(str).unique() if s.strip()]
    if not observed: return
    chips = [f"<span class='pareto-status {stock_status_style_class(l)}'>{html.escape(l.strip())}</span>" for l in observed]
    st.markdown(f"<div class='stock-status-legend'><span class='stock-status-legend-label'>Stock Status</span>{''.join(chips)}</div>", unsafe_allow_html=True)

def apply_executive_bar_style(fig, *, accent_color, value_axis_title, value_max, category_order, orientation="v", percent=True, height=425, right_margin=28):
    fig.update_traces(marker=dict(color=accent_color, line=dict(width=0)), opacity=0.96, textfont=dict(size=12, color=PLOTLY_CHART_TEXT), cliponaxis=False)
    v_axis = dict(title=dict(text=value_axis_title, font=dict(size=11, color=PLOTLY_CHART_MUTED)), range=[0, value_max], gridcolor=PLOTLY_CHART_GRID, zeroline=False, showline=False, tickfont=dict(size=10, color=PLOTLY_CHART_MUTED), automargin=True)
    if percent: v_axis["ticksuffix"] = "%"
    c_axis = dict(title="", type="category", categoryorder="array", categoryarray=category_order, showgrid=False, showline=True, linecolor=PLOTLY_CHART_AXIS, tickfont=dict(size=10, color=PLOTLY_CHART_MUTED), automargin=True)
    
    if orientation == "h": c_axis["autorange"] = "reversed"; xaxis, yaxis = v_axis, c_axis
    else: xaxis, yaxis = c_axis, v_axis

    fig.update_layout(template=PLOTLY_TEMPLATE, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=height, margin=dict(t=76, b=52, l=42 if orientation=="v" else 20, r=right_margin), showlegend=False, bargap=0.34 if orientation=="v" else 0.3, font=dict(color=PLOTLY_CHART_TEXT), title=dict(x=0.02, y=0.97, font=dict(size=16, color=PLOTLY_CHART_TEXT)), hoverlabel=dict(bgcolor=PLOTLY_HOVER_BG, bordercolor=PLOTLY_HOVER_BORDER, font=dict(color=PLOTLY_HOVER_TEXT)), xaxis=xaxis, yaxis=yaxis)
    return fig

def prepare_chart_series(df, y_col):
    if df.empty or y_col not in df.columns: return pd.DataFrame(columns=["period", y_col])
    c_df = df[["period", y_col]].copy()
    c_df["period"] = pd.to_datetime(c_df["period"], errors="coerce")
    c_df[y_col] = pd.to_numeric(c_df[y_col], errors="coerce")
    return c_df.dropna().sort_values("period").drop_duplicates(subset=["period"], keep="last").reset_index(drop=True)

def create_styled_line_chart(df, y_col, title, subtitle, line_color, is_weekly, is_percentage=True, fill=False):
    chart_df = prepare_chart_series(df, y_col)
    fig = go.Figure()
    if chart_df.empty:
        fig.update_layout(template=PLOTLY_TEMPLATE, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=350, margin=dict(t=72, b=34, l=42, r=20), title=dict(text=f"{title}<br><span style='font-size:10px'>{subtitle}</span>", x=0.02), xaxis=dict(visible=False), yaxis=dict(visible=False))
        fig.add_annotation(text="No data", x=0.5, y=0.5, showarrow=False, font=dict(color="#94a3b8"))
        return fig

    plot_y = chart_df[y_col].copy()
    if is_percentage:
        scale = 0.01 if plot_y.abs().max() > 1.5 else 1.0
        plot_y = np.floor((plot_y * scale) * 100.0 + 0.5) / 100.0
        text_labels, tick_fmt, y_max_base = [f"{v*100:.0f}%" for v in plot_y], ".0%", 0.10
    else:
        plot_y = np.floor(plot_y + 0.5).astype(int)
        text_labels, tick_fmt, y_max_base = [f"{v:,.0f}" for v in plot_y], ",.0f", 10

    if is_weekly:
        chart_x = chart_df["period"].dt.strftime("%d %b %Y")
        tick_text = chart_df["period"].dt.strftime("%d %b")
        cat_arr = chart_x.tolist()
    else:
        chart_x = chart_df["period"].dt.strftime("%b")
        tick_text = chart_x.tolist()
        cat_arr = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    hover_tpl = f"<b>%{{customdata}}</b><br>{title}: <b>%{{y:{tick_fmt}}}</b><extra></extra>"
    text_pos = ["top center" if i % 2 == 0 else "bottom center" for i in range(len(chart_df))]
    
    fig.add_trace(go.Scatter(x=chart_x, y=plot_y, customdata=chart_df["period"].dt.strftime("%d %b %Y"), name="Actual", mode="lines+markers+text", text=text_labels, textposition=text_pos, textfont=dict(size=10), line=dict(shape="linear" if is_weekly else "spline", width=3.2, color=line_color), marker=dict(size=8, color=line_color, line=dict(width=2, color="#fff")), fill="tozeroy" if fill else "none", fillcolor=f"rgba({int(line_color[1:3],16)},{int(line_color[3:5],16)},{int(line_color[5:7],16)},0.08)" if fill else None, hovertemplate=hover_tpl, cliponaxis=False))

    y_max = plot_y.max() * 1.16 if not plot_y.empty else y_max_base
    y_max = max(y_max, 0.05 if is_percentage else 5)

    x_conf = dict(type="category", showgrid=False, categoryorder="array", categoryarray=cat_arr, linecolor="rgba(148,163,184,0.18)")
    if is_weekly: x_conf.update(tickmode="array", tickvals=chart_x.tolist(), ticktext=tick_text.tolist())
    else: x_conf.update(tickmode="array", tickvals=chart_x.tolist(), ticktext=tick_text)

    latest = chart_df["period"].max()
    latest_txt = latest.strftime("%d %b %Y") if pd.notna(latest) else "N/A"

    fig.update_layout(template=PLOTLY_TEMPLATE, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=365, margin=dict(t=82, b=44, l=46, r=20), showlegend=False, hoverlabel=dict(bgcolor=PLOTLY_HOVER_BG, bordercolor=PLOTLY_HOVER_BORDER, font=dict(color=PLOTLY_HOVER_TEXT)), title=dict(text=f"{title}<br><span style='font-size:10px; color:#94a3b8;'>{subtitle} • LATEST {latest_txt.upper()}</span>", x=0.02, y=0.96), xaxis=x_conf, yaxis=dict(showgrid=True, gridwidth=1, gridcolor="rgba(148,163,184,0.14)", zeroline=False, tickformat=tick_fmt, range=[0, y_max]))
    return fig

# =========================================================
# INITIALIZE PRIMARY TABS
# =========================================================
tab_inventory, tab_procurements = st.tabs(["📊 Inventory Control Tower", "📦 Procurements"])

with tab_inventory:
    trend_title_col, trend_sync_col = st.columns([5.45, 0.55], gap="small", vertical_alignment="center")
    with trend_title_col: st.markdown("<div class='section-heading trend-heading-inline'><span class='dot'></span><span class='title'>MUTI MC Trends</span></div>", unsafe_allow_html=True)
    with trend_sync_col:
        st.markdown('<div class="data-sync-shell"><span class="data-sync-caption">Latest workbook</span></div>', unsafe_allow_html=True)
        if st.button("Data Sync", type="primary", use_container_width=True, key="sync"): data_sync_dialog()

    if "scm_import_success" in st.session_state: st.success(st.session_state.pop("scm_import_success"))
    if st.session_state.get("scm_cloud_warning"): st.caption(st.session_state["scm_cloud_warning"])

    if not saved_workbook_bytes:
        st.info("Click Data Sync to upload your workbook.")
        st.stop()

    try:
        raw_data, kpi_ytd, kpi_weekly = process_excel_file(io.BytesIO(saved_workbook_bytes))
    except Exception as e:
        st.error(f"Import Failed: {e}"); st.stop()

    control_col1, control_col2 = st.columns([1.35, 4.65], gap="small")
    with control_col1: timeframe = st.selectbox("TIMEFRAME", ["Year-to-Date (YTD)", "Weekly View"], index=0)
    kpi_data = kpi_weekly if timeframe == "Weekly View" else kpi_ytd
    is_weekly = timeframe == "Weekly View"

    with control_col2:
        l_date = kpi_data["period"].max().strftime("%d %b %Y") if not kpi_data.empty and kpi_data["period"].notna().any() else "No KPI date"
        st.markdown(f"<div style='padding-top:28px;'><span class='info-chip'>View: {timeframe}</span><span class='info-chip'>Latest KPI: {l_date}</span></div>", unsafe_allow_html=True)

    r1, r2 = st.columns(2, gap="small")
    with r1: st.plotly_chart(create_styled_line_chart(kpi_data, "class_a_doi", "MC Class A DoI", "CLASS A DAYS OF INVENTORY", "#7c3aed", is_weekly, False, True), use_container_width=True)
    with r2: st.plotly_chart(create_styled_line_chart(kpi_data, "overall_doi", "Days of Inventory", "OVERALL INVENTORY COVERAGE", "#2563eb", is_weekly, False, True), use_container_width=True)

    r3, r4 = st.columns(2, gap="small")
    with r3: st.plotly_chart(create_styled_line_chart(kpi_data, "per_branch", "Per Branch OOS", "STOCKOUT RATE", "#0ea5e9", is_weekly, True), use_container_width=True)
    with r4: st.plotly_chart(create_styled_line_chart(kpi_data, "class_a_out", "Overall Class A Rate", "CLASS A STOCKOUT RATE", "#f43f5e", is_weekly, True), use_container_width=True)

    r5, r6 = st.columns(2, gap="small")
    with r5: st.plotly_chart(create_styled_line_chart(kpi_data, "before_po", "Overall Before PO", "STOCKOUT RATE BEFORE PO", "#f59e0b", is_weekly, True), use_container_width=True)
    with r6: st.plotly_chart(create_styled_line_chart(kpi_data, "after_po", "Overall After PO", "STOCKOUT RATE AFTER PO", "#10b981", is_weekly, True, True), use_container_width=True)

    st.markdown("---")
    
    # ------------------------------------
    # Network Scope Filter & Performance
    # ------------------------------------
    section_heading("Network Scope", "Filter the operational view")
    s1, s2 = st.columns([1.5, 3.5])
    areas = ["All Areas"] + sorted([a for a in raw_data["area"].dropna().unique() if str(a).strip()])
    with s1: selected_area = st.selectbox("NETWORK SCOPE", areas)
    
    area_data = raw_data[raw_data["area"] == selected_area] if selected_area != "All Areas" else raw_data.copy()
    
    with s2:
        st.markdown(f"<div style='padding-top:28px;'><span class='info-chip'>Scope: {selected_area}</span><span class='info-chip'>{area_data['branch'].nunique()} Branch(es)</span><span class='info-chip'>{len(area_data):,} Records</span></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    m1, m2, m3, m4 = st.columns(4, gap="small")
    m1.markdown(f"<div class='metric-card'><div class='metric-title'>Class A Rate</div><div class='metric-value-sm'>{calculate_stockout_rate(area_data, 'Class A')}%</div></div>", unsafe_allow_html=True)
    m2.markdown(f"<div class='metric-card'><div class='metric-title'>Class B Rate</div><div class='metric-value-sm'>{calculate_stockout_rate(area_data, 'Class B')}%</div></div>", unsafe_allow_html=True)
    m3.markdown(f"<div class='metric-card'><div class='metric-title'>Class C Rate</div><div class='metric-value-sm'>{calculate_stockout_rate(area_data, 'Class C')}%</div></div>", unsafe_allow_html=True)
    m4.markdown(f"<div class='metric-card'><div class='metric-title'>Average Rate</div><div class='metric-value-sm'>{calculate_stockout_rate(area_data)}%</div></div>", unsafe_allow_html=True)

    # ------------------------------------
    # Branch Level Pareto Tables
    # ------------------------------------
    st.markdown("---")
    section_heading("Branch-Level Stockout Performance", f"Operational drill-down • {selected_area}")
    
    h1, h2 = st.columns([3.4, 1.6])
    with h2: selected_branch = st.selectbox("BRANCH", ["All Branches"] + sorted(area_data["branch"].dropna().unique().tolist()))
    
    branch_data = area_data[area_data["branch"] == selected_branch] if selected_branch != "All Branches" else area_data.copy()
    
    render_stock_status_legend(branch_data)

    def render_pareto_table(df, p_class, hex_col):
        c_df = df[df["pareto_class"] == p_class].copy()
        count = len(c_df)
        st.markdown(f"<div class='pareto-header' style='border-color:{hex_col};'><span style='font-size:1.05rem; font-weight:900; color:{hex_col};'>{p_class.upper()}</span><span class='pareto-count' style='color:{hex_col};'>● {count} Items</span></div>", unsafe_allow_html=True)
        if count == 0: st.info(f"No {p_class} active."); return

        c_df = c_df.sort_values(["suggested_transfer", "doi"], ascending=[False, True]).reset_index(drop=True)
        c_df.index += 1
        
        # Super-fast HTML generation using itertuples
        rows_html = [
            f"<tr><td class='num'>{row.Index}</td><td>{html.escape(str(row.model))}</td><td><span class='pareto-status {stock_status_style_class(row.stock_status)}'>{html.escape(str(row.stock_status))}</span></td><td class='num'>{int(row.remaining_inventory):,}</td><td class='num'>{int(row.suggested_transfer):,}</td><td class='num'>{int(row.doi):,}</td></tr>"
            for row in c_df.itertuples()
        ]
        
        st.markdown(f"<div class='pareto-html-shell'><table class='pareto-html-table'><colgroup><col style='width:7%'><col style='width:35%'><col style='width:16%'><col style='width:14%'><col style='width:14%'><col style='width:14%'></colgroup><thead><tr><th class='num'>Rank</th><th>Model</th><th>Status</th><th class='num'>Inventory</th><th class='num'>Transfer</th><th class='num'>DOI</th></tr></thead><tbody>{''.join(rows_html)}</tbody></table></div>", unsafe_allow_html=True)

    for p_class, p_color in [("Class A", "#f87171"), ("Class B", "#fbbf24"), ("Class C", "#4ade80")]:
        with st.container(border=True): render_pareto_table(branch_data, p_class, p_color)
        st.markdown("<div class='pareto-panel-spacer'></div>", unsafe_allow_html=True)

with tab_procurements:
    st.markdown("<br>", unsafe_allow_html=True)
    section_heading("Procurements Control", "Manage allocations and lead times")
    pc1, pc2 = st.columns(2, gap="small")
    with pc1: st.markdown("<div class='metric-card' style='min-height: 250px;'><div class='metric-header'><span style='color:#6366f1; font-weight:900;'>🏍️ MOTORCYCLE UNITS</span></div><hr style='border-color: rgba(148,163,184,0.1);'><div style='color:var(--scm-muted); font-size:0.85rem;'><b>Pipeline Visibility</b><br>• Supplier Lead Times<br>• Incoming Allocations<br><br><i>(Pending integration)</i></div></div>", unsafe_allow_html=True)
    with pc2: st.markdown("<div class='metric-card' style='min-height: 250px;'><div class='metric-header'><span style='color:#10b981; font-weight:900;'>⚙️ SPARE PARTS</span></div><hr style='border-color: rgba(148,163,184,0.1);'><div style='color:var(--scm-muted); font-size:0.85rem;'><b>Replenishment Status</b><br>• Active POs<br>• Critical Shortages<br><br><i>(Pending integration)</i></div></div>", unsafe_allow_html=True)
