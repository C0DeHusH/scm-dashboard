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


# Excel file signatures. Detect from the actual bytes instead of trusting
# a temporary/cloud filename, which may not retain an extension online.
XLSX_SIGNATURES = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")
XLS_SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"


def detect_excel_engine(file_path_or_buffer):
    """
    Return the explicit pandas Excel engine based on the workbook bytes.

    .xlsx/.xlsm/.xltx files are ZIP containers -> openpyxl
    legacy .xls files are OLE Compound Documents -> xlrd

    This prevents the deployment error:
      "Excel file format cannot be determined, you must specify an engine manually."
    """
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
        # Preserve the caller's current stream position.
        original_position = None
        try:
            if hasattr(file_path_or_buffer, "tell"):
                original_position = file_path_or_buffer.tell()
            if hasattr(file_path_or_buffer, "seek"):
                file_path_or_buffer.seek(0)
            signature = file_path_or_buffer.read(8)
        finally:
            if hasattr(file_path_or_buffer, "seek"):
                file_path_or_buffer.seek(
                    original_position if original_position is not None else 0
                )
    else:
        raise ValueError(
            "Unsupported workbook input. Please upload a valid .xlsx or .xls file."
        )

    if any(signature.startswith(sig) for sig in XLSX_SIGNATURES):
        return "openpyxl"

    if signature.startswith(XLS_SIGNATURE):
        return "xlrd"

    raise ValueError(
        "The uploaded/saved file is not a valid Excel workbook. "
        "Please open it in Microsoft Excel and save/export it as .xlsx, "
        "then upload it again."
    )


def validate_excel_bytes(file_bytes):
    """Validate that persisted bytes are a supported Excel container."""
    if not file_bytes:
        return False
    detect_excel_engine(io.BytesIO(file_bytes))
    return True


def round_half_up(value, ndigits=0):
    """
    Standard business rounding (ROUND_HALF_UP):
      8.4 -> 8
      8.5 -> 9
      8.6 -> 9
    Returns NaN unchanged.
    """
    if pd.isna(value):
        return np.nan

    quantizer = Decimal("1").scaleb(-ndigits)
    rounded = Decimal(str(float(value))).quantize(
        quantizer,
        rounding=ROUND_HALF_UP,
    )

    if ndigits == 0:
        return int(rounded)
    return float(rounded)


def round_series_half_up(series):
    """Round a numeric pandas Series to whole numbers using ROUND_HALF_UP."""
    numeric = pd.to_numeric(series, errors="coerce")
    return numeric.map(
        lambda value: round_half_up(value)
        if pd.notna(value)
        else np.nan
    )


def _secret(name, default=None):
    """Safely read a Streamlit secret without failing during local development."""
    try:
        return st.secrets[name]
    except Exception:
        return default


def get_cloud_storage_config():
    """
    Supabase Storage configuration.

    Required Streamlit secrets:
        SUPABASE_URL
        SUPABASE_SECRET_KEY

    Backward-compatible alias:
        SUPABASE_KEY

    Optional:
        SUPABASE_BUCKET  (default: scm-dashboard)
        SUPABASE_OBJECT  (default: persistent_scm_data.xlsx)
    """
    url = str(_secret("SUPABASE_URL", "") or "").strip().rstrip("/")
    key = str(
        _secret("SUPABASE_SECRET_KEY", "")
        or _secret("SUPABASE_KEY", "")
        or ""
    ).strip()
    bucket = str(
        _secret("SUPABASE_BUCKET", DEFAULT_SUPABASE_BUCKET)
        or DEFAULT_SUPABASE_BUCKET
    ).strip()
    object_name = str(
        _secret("SUPABASE_OBJECT", DEFAULT_SUPABASE_OBJECT)
        or DEFAULT_SUPABASE_OBJECT
    ).strip().lstrip("/")

    configured = bool(url and key and bucket and object_name)

    return {
        "configured": configured,
        "url": url,
        "key": key,
        "bucket": bucket,
        "object_name": object_name,
    }


def _supabase_object_endpoint(config):
    encoded_bucket = quote(config["bucket"], safe="")
    encoded_object = quote(config["object_name"], safe="/")
    return (
        f'{config["url"]}/storage/v1/object/'
        f"{encoded_bucket}/{encoded_object}"
    )


def _supabase_auth_headers(config):
    """
    Support both current Supabase secret keys (sb_secret_...) and the
    legacy JWT service_role key.

    Current secret keys belong in the apikey header only.
    Legacy service_role JWTs also use Authorization: Bearer.
    """
    headers = {"apikey": config["key"]}

    if not config["key"].startswith("sb_"):
        headers["Authorization"] = f'Bearer {config["key"]}'

    return headers


def download_cloud_workbook():
    """
    Download the persisted workbook from Supabase Storage.
    Returns bytes when available, None when the object does not exist.
    """
    config = get_cloud_storage_config()
    if not config["configured"]:
        return None

    headers = _supabase_auth_headers(config)

    response = requests.get(
        _supabase_object_endpoint(config),
        headers=headers,
        timeout=30,
    )

    if response.status_code == 404:
        return None

    response.raise_for_status()
    return response.content


def upload_cloud_workbook(file_bytes):
    """
    Persist the workbook in Supabase Storage.
    x-upsert=true replaces the previous dashboard workbook atomically.
    """
    config = get_cloud_storage_config()
    if not config["configured"]:
        return False

    headers = _supabase_auth_headers(config)
    headers.update(
        {
            "Content-Type": (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            "x-upsert": "true",
        }
    )

    response = requests.post(
        _supabase_object_endpoint(config),
        headers=headers,
        data=file_bytes,
        timeout=60,
    )

    # Some Supabase configurations return 200, others 201.
    if response.status_code not in (200, 201):
        raise RuntimeError(
            "Cloud persistence upload failed "
            f"({response.status_code}): {response.text[:300]}"
        )

    return True


def save_local_cache(file_bytes):
    """Local fallback/cache. Cloud storage remains the durable source online."""
    with open(LOCAL_CACHE_FILE, "wb") as file_handle:
        file_handle.write(file_bytes)


def load_local_cache():
    if not os.path.exists(LOCAL_CACHE_FILE):
        return None

    with open(LOCAL_CACHE_FILE, "rb") as file_handle:
        return file_handle.read()


def initialize_persistent_workbook():
    """
    Load the workbook once per browser session.

    Priority:
      1. Supabase Storage (durable across refresh/restart/redeploy)
      2. Local server cache (development / fallback)
    """
    if "scm_workbook_bytes" in st.session_state:
        return (
            st.session_state["scm_workbook_bytes"],
            st.session_state.get("scm_storage_source", "Session cache"),
        )

    cloud_config = get_cloud_storage_config()

    if cloud_config["configured"]:
        try:
            cloud_bytes = download_cloud_workbook()
            if cloud_bytes:
                # Never activate a cloud object unless it is really an Excel file.
                # This also protects the app if an old/incorrect object was uploaded
                # to the same Supabase Storage path.
                validate_excel_bytes(cloud_bytes)
                st.session_state["scm_workbook_bytes"] = cloud_bytes
                st.session_state["scm_storage_source"] = "Cloud • Supabase"
                # Refresh the local cache for faster fallback/debugging.
                try:
                    save_local_cache(cloud_bytes)
                except Exception:
                    pass
                return cloud_bytes, "Cloud • Supabase"
        except Exception as exc:
            st.session_state["scm_cloud_warning"] = str(exc)

    local_bytes = load_local_cache()
    if local_bytes:
        try:
            validate_excel_bytes(local_bytes)
        except Exception as exc:
            st.session_state["scm_local_warning"] = (
                "The saved local cache is not a valid Excel workbook: " + str(exc)
            )
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

# =========================================================
# 1A. CUSTOM LIGHT / DARK THEME CONTROL
#     UI-only enhancement. Dashboard calculations and data logic are unchanged.
# =========================================================
if "scm_theme" not in st.session_state:
    # Keep the current executive appearance as the default.
    st.session_state["scm_theme"] = "dark"


def set_scm_theme(theme_name):
    """Switch the custom dashboard presentation theme."""
    if theme_name in {"light", "dark"}:
        st.session_state["scm_theme"] = theme_name


SCM_THEME = st.session_state["scm_theme"]
SCM_IS_DARK = SCM_THEME == "dark"

# Plotly presentation follows the dashboard theme.
PLOTLY_TEMPLATE = "plotly_dark" if SCM_IS_DARK else "plotly_white"
PLOTLY_HOVER_BG = "#0f172a" if SCM_IS_DARK else "#ffffff"
PLOTLY_HOVER_TEXT = "#f8fafc" if SCM_IS_DARK else "#0f172a"
PLOTLY_HOVER_BORDER = (
    "rgba(148,163,184,0.28)"
    if SCM_IS_DARK
    else "rgba(15,23,42,0.16)"
)

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

        /* FULL-WIDTH EXECUTIVE CANVAS */
        html, body, [data-testid="stAppViewContainer"] {
            overflow-x: hidden !important;
        }

        .block-container,
        [data-testid="stMainBlockContainer"] {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            padding-top: 0.55rem !important;
            padding-bottom: 2.25rem !important;
            padding-left: clamp(0.5rem, 2vw, 1.5rem) !important;
            padding-right: clamp(0.5rem, 2vw, 1.5rem) !important;
            margin: 0 !important;
        }

        section[data-testid="stMain"] {
            min-width: 0 !important;
            max-width: none !important;
            margin-left: 0 !important;
            padding-left: 0 !important;
            overflow-x: clip !important;
        }

        /* TABS MAXIMIZATION - Aggressively remove native inner padding */
        div[data-testid="stTabs"] {
            width: 100% !important;
            min-width: 0 !important;
        }
        div[data-baseweb="tab-panel"] {
            width: 100% !important;
            padding-left: 0 !important;
            padding-right: 0 !important;
            padding-bottom: 0 !important;
            padding-top: 1rem !important;
        }
        div[data-baseweb="tab-panel"] > div {
            padding-left: 0 !important;
            padding-right: 0 !important;
        }
        div[data-baseweb="tab-panel"] > div > div[data-testid="stVerticalBlock"] {
            padding-left: 0 !important;
            padding-right: 0 !important;
        }

        /* FULL-SCREEN MODE — sidebar removed entirely */
        [data-testid="stSidebar"],
        [data-testid="collapsedControl"] {
            display: none !important;
        }

        /* STREAMLIT NATIVE HEADER / TOOLBAR — hidden */
        [data-testid="stToolbar"],
        #MainMenu,
        header[data-testid="stHeader"],
        [data-testid="stDecoration"],
        [data-testid="viewerBadge"] {
            display: none !important;
            visibility: hidden !important;
        }

        header[data-testid="stHeader"] {
            height: 0 !important;
            min-height: 0 !important;
        }

        /* Data Sync action beside MUTI MC Trends */
        .data-sync-caption {
            display: block;
            margin: 0 0 0.28rem 0;
            color: var(--scm-muted);
            font-size: 0.62rem;
            font-weight: 850;
            line-height: 1.2;
            letter-spacing: 0.09em;
            text-transform: uppercase;
            text-align: right;
            white-space: nowrap;
        }

        .data-sync-shell {
            width: 100%;
            min-width: 0;
            padding-top: 0.02rem;
        }

        /* Streamlit buttons — executive treatment */
        div[data-testid="stButton"] > button {
            min-height: 42px;
            border-radius: 11px;
            font-weight: 800;
            letter-spacing: 0.01em;
            border: 1px solid rgba(99, 102, 241, 0.34);
            box-shadow: 0 5px 14px rgba(15, 23, 42, 0.06);
        }

        /* The compact header action is intentionally stronger than ordinary buttons. */
        div[data-testid="stHorizontalBlock"] > div:last-child div[data-testid="stButton"] > button {
            min-height: 44px;
            border-radius: 12px;
            border: 1px solid rgba(99, 102, 241, 0.50);
            background: linear-gradient(135deg, rgba(79,70,229,0.98), rgba(37,99,235,0.96));
            color: #ffffff;
            font-weight: 850;
            box-shadow: 0 8px 20px rgba(37, 99, 235, 0.18);
            transition: transform 150ms ease, box-shadow 150ms ease, filter 150ms ease;
        }

        div[data-testid="stHorizontalBlock"] > div:last-child div[data-testid="stButton"] > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 11px 24px rgba(37, 99, 235, 0.24);
            filter: brightness(1.03);
        }

        /* Upload dialog */
        div[data-testid="stDialog"] [data-testid="stFileUploader"] section {
            min-height: 118px !important;
            border-radius: 14px !important;
            border: 1px dashed rgba(99, 102, 241, 0.48) !important;
            background: rgba(99, 102, 241, 0.035) !important;
        }

        .import-dialog-note {
            padding: 0.72rem 0.85rem;
            border: 1px solid rgba(148, 163, 184, 0.18);
            border-radius: 12px;
            background: rgba(148, 163, 184, 0.035);
            color: var(--scm-muted);
            font-size: 0.78rem;
            line-height: 1.5;
            margin-bottom: 0.85rem;
        }

        /* Responsive Pareto HTML tables */
        .pareto-html-shell {
            width: 100%;
            max-width: 100%;
            min-width: 0;
            overflow: hidden;
            border: 1px solid var(--scm-border);
            border-radius: 12px;
        }

        .pareto-html-table {
            width: 100%;
            max-width: 100%;
            table-layout: fixed;
            border-collapse: collapse;
            font-size: 0.78rem;
        }

        .pareto-html-table thead th {
            padding: 0.62rem 0.56rem;
            text-align: left;
            font-size: 0.67rem;
            font-weight: 850;
            letter-spacing: 0.05em;
            text-transform: uppercase;
            color: var(--scm-muted);
            background: rgba(148, 163, 184, 0.055);
            border-bottom: 1px solid var(--scm-border);
        }

        .pareto-html-table tbody td {
            padding: 0.60rem 0.56rem;
            border-bottom: 1px solid rgba(148, 163, 184, 0.11);
            vertical-align: middle;
            overflow-wrap: anywhere;
            word-break: normal;
        }

        .pareto-html-table tbody tr:last-child td {
            border-bottom: none;
        }

        .pareto-html-table tbody tr:hover {
            background: rgba(99, 102, 241, 0.035);
        }

        .pareto-html-table .num {
            text-align: right;
            font-variant-numeric: tabular-nums;
            white-space: nowrap;
        }

        .pareto-status {
            display: inline-flex;
            align-items: center;
            max-width: 100%;
            padding: 3px 7px;
            border-radius: 999px;
            border: 1px solid rgba(148, 163, 184, 0.20);
            font-size: 0.67rem;
            font-weight: 800;
            line-height: 1.2;
            white-space: normal;
        }

        .pareto-status.stockout {
            color: #f87171;
            background: rgba(248, 113, 113, 0.08);
            border-color: rgba(248, 113, 113, 0.22);
        }

        /* WORLD-CLASS EXECUTIVE HERO — ONE FLAT SURFACE */
        .hero-shell {
            box-sizing: border-box;
            width: 100%;
            min-width: 0;
            margin: 0.08rem 0 0.92rem 0;
            padding: clamp(20px, 1.65vw, 28px) clamp(22px, 2.15vw, 34px);
            border: 1px solid rgba(99, 102, 241, 0.34);
            border-left: 4px solid var(--scm-indigo);
            border-radius: 18px;
            background: #0b1220;
            box-shadow: 0 10px 28px rgba(2, 6, 23, 0.14);
            overflow: visible;
        }

        .hero-grid,
        .hero-copy {
            display: block;
            width: 100%;
            min-width: 0;
            max-width: 100%;
            position: static !important;
        }

        .hero-kicker,
        .hero-title,
        .hero-subtitle {
            display: block;
            position: static !important;
            float: none !important;
            clear: both;
            width: 100%;
            height: auto !important;
            min-height: 0 !important;
            padding: 0 !important;
            transform: none !important;
            white-space: normal !important;
            overflow: visible !important;
            text-overflow: clip !important;
            overflow-wrap: break-word;
            word-break: normal;
        }

        .hero-kicker {
            color: #a5b4fc;
            font-size: clamp(0.62rem, 0.66vw, 0.72rem);
            font-weight: 850;
            line-height: 1.40 !important;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            margin: 0 0 10px 0 !important;
        }

        .hero-title {
            color: #f8fafc;
            font-size: clamp(1.65rem, 2.10vw, 2.45rem);
            font-weight: 900;
            line-height: 1.13 !important;
            letter-spacing: -0.025em;
            margin: 0 0 12px 0 !important;
        }

        .hero-subtitle {
            max-width: 1180px;
            color: #cbd5e1;
            font-size: clamp(0.80rem, 0.86vw, 0.94rem);
            font-weight: 500;
            line-height: 1.55 !important;
            margin: 0 !important;
        }
        
        /* DRP MODULE LINK BUTTON */
        .drp-action-btn {
            display: inline-flex;
            align-items: center;
            margin-top: 18px;
            padding: 8px 18px;
            background: rgba(99, 102, 241, 0.12);
            color: #a5b4fc;
            border: 1px solid rgba(99, 102, 241, 0.4);
            border-radius: 8px;
            font-size: 0.72rem;
            font-weight: 850;
            text-decoration: none;
            letter-spacing: 0.05em;
            text-transform: uppercase;
            transition: all 0.2s ease;
        }
        .drp-action-btn:hover {
            background: rgba(99, 102, 241, 0.25);
            color: #ffffff;
            border-color: rgba(99, 102, 241, 0.8);
            transform: translateY(-1px);
        }

        /* SECTION HEADINGS */
        .section-heading {
            display: grid;
            grid-template-columns: auto minmax(0, 1fr) minmax(0, auto);
            align-items: center;
            gap: 9px;
            min-width: 0;
            margin: 0.44rem 0 0.72rem 0;
            padding: 0.18rem 0.12rem;
        }

        .section-heading .dot {
            width: 9px;
            height: 9px;
            border-radius: 50%;
            background: var(--scm-indigo);
            box-shadow: 0 0 0 4px rgba(99, 102, 241, 0.11);
        }

        .section-heading .title {
            font-size: 1.12rem;
            font-weight: 900;
            letter-spacing: -0.012em;
        }

        .section-heading .subtitle {
            color: var(--scm-muted);
            font-size: 0.75rem;
            font-weight: 550;
            text-align: right;
            white-space: normal;
            overflow-wrap: anywhere;
        }

        .trend-heading-inline {
            min-height: 62px;
            margin: 0.20rem 0 0.18rem 0 !important;
            align-content: center;
        }

        /* EXECUTIVE KPI CARDS */
        .metric-card,
        .metric-card-base {
            position: relative;
            box-sizing: border-box;
            width: 100%;
            min-width: 0;
            min-height: 120px;
            height: 100%;
            padding: 15px 17px;
            border: 1px solid var(--scm-border);
            border-radius: 15px;
            background: rgba(148, 163, 184, 0.025);
            box-shadow: 0 7px 20px rgba(15, 23, 42, 0.045);
            overflow: hidden;
        }

        .metric-card::before,
        .metric-card-base::before {
            content: "";
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 3px;
            background: var(--scm-indigo);
            opacity: 0.90;
        }

        .metric-card-base {
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }

        .metric-title,
        .metric-header {
            color: var(--scm-muted);
            font-size: 0.67rem;
            font-weight: 850;
            text-transform: uppercase;
            letter-spacing: 0.075em;
        }

        .metric-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 8px;
        }

        .metric-value,
        .metric-value-sm {
            font-size: clamp(1.72rem, 1.85vw, 2.05rem);
            font-weight: 900;
            line-height: 1.05;
            margin: 9px 0 6px 0;
            letter-spacing: -0.025em;
        }

        .metric-footnote {
            color: var(--scm-muted);
            font-size: 0.69rem;
            line-height: 1.35;
        }

        .badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 999px;
            font-size: 0.60rem;
            font-weight: 850;
            text-transform: uppercase;
            letter-spacing: 0.045em;
        }

        .badge-red { background: rgba(244, 63, 94, 0.12); color: #fb7185; }
        .badge-yellow { background: rgba(245, 158, 11, 0.12); color: #fbbf24; }
        .badge-green { background: rgba(16, 185, 129, 0.12); color: #34d399; }
        .badge-blue { background: rgba(14, 165, 233, 0.12); color: #38bdf8; }

        .icon-box {
            display: flex;
            align-items: center;
            justify-content: center;
            width: 28px;
            height: 28px;
            flex: 0 0 28px;
            border-radius: 8px;
            background: rgba(99, 102, 241, 0.09);
        }

        /* STATUS / FILTER CHIPS */
        .status-strip {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 6px;
            width: 100%;
            min-width: 0;
            margin: 0.45rem 0 0.12rem 0;
            padding: 8px 10px;
            border: 1px solid var(--scm-border);
            border-radius: 12px;
            background: rgba(148, 163, 184, 0.018);
        }

        .status-pill,
        .info-chip {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            min-width: 0;
            padding: 4px 8px;
            border: 1px solid var(--scm-border);
            border-radius: 999px;
            font-size: 0.67rem;
            font-weight: 700;
            color: var(--scm-muted);
            line-height: 1.25;
        }

        .info-chip {
            margin-right: 5px;
            margin-bottom: 4px;
        }

        .status-pill strong {
            color: inherit;
            font-weight: 850;
        }

        .status-dot {
            width: 6px;
            height: 6px;
            flex: 0 0 6px;
            border-radius: 50%;
            background: var(--scm-green);
            box-shadow: 0 0 0 3px rgba(16,185,129,0.10);
        }

        /* STREAMLIT COMPONENTS / CHART CARDS */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 15px !important;
            border-color: var(--scm-border) !important;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.035);
            background: rgba(148, 163, 184, 0.012);
        }

        div[data-testid="stVerticalBlockBorderWrapper"] > div {
            min-width: 0 !important;
        }

        div[data-testid="stPlotlyChart"] {
            width: 100% !important;
            min-width: 0 !important;
            border-radius: 13px;
            overflow: hidden;
        }

        div[data-baseweb="select"] > div,
        div[data-testid="stFileUploader"] section {
            border-radius: 10px;
        }

        /* PARETO ACTION PANELS */
        .pareto-panel-spacer {
            height: 0.28rem;
        }

        .pareto-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
            width: 100%;
            min-width: 0;
            padding: 0.58rem 0 0.54rem 0;
            border-bottom: 2px solid;
            margin-bottom: 0.62rem;
        }

        .pareto-header > span:first-child {
            min-width: 0;
            overflow-wrap: anywhere;
        }

        .pareto-count {
            flex: 0 0 auto;
            white-space: nowrap;
            border: 1px solid rgba(128,128,128,0.22);
            padding: 4px 8px;
            border-radius: 999px;
            font-size: 0.68rem;
            font-weight: 850;
        }

        .pareto-table-shell {
            width: 100%;
            min-width: 0;
            overflow: hidden;
        }

        hr {
            border: none;
            border-top: 1px solid var(--scm-border);
            margin: 1.10rem 0;
        }
    </style><br>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# CUSTOM THEME PALETTE
# ---------------------------------------------------------
if SCM_IS_DARK:
    theme_tokens = {
        "page_bg": "#070b14",
        "surface": "#0b1220",
        "surface_2": "#0f172a",
        "card_bg": "rgba(15, 23, 42, 0.72)",
        "input_bg": "#0f172a",
        "text": "#f8fafc",
        "muted": "#94a3b8",
        "border": "rgba(148, 163, 184, 0.18)",
        "soft_border": "rgba(148, 163, 184, 0.11)",
        "table_head": "rgba(148, 163, 184, 0.055)",
        "hover": "rgba(99, 102, 241, 0.07)",
        "hero_bg": "#0b1220",
        "hero_title": "#f8fafc",
        "hero_subtitle": "#cbd5e1",
        "hero_kicker": "#a5b4fc",
        "shadow": "0 10px 28px rgba(2, 6, 23, 0.22)",
    }
else:
    theme_tokens = {
        "page_bg": "#f6f8fc",
        "surface": "#ffffff",
        "surface_2": "#f8fafc",
        "card_bg": "#ffffff",
        "input_bg": "#ffffff",
        "text": "#0f172a",
        "muted": "#64748b",
        "border": "rgba(15, 23, 42, 0.12)",
        "soft_border": "rgba(15, 23, 42, 0.08)",
        "table_head": "#f8fafc",
        "hover": "rgba(79, 70, 229, 0.055)",
        "hero_bg": "#ffffff",
        "hero_title": "#0f172a",
        "hero_subtitle": "#475569",
        "hero_kicker": "#4f46e5",
        "shadow": "0 10px 28px rgba(15, 23, 42, 0.08)",
    }

st.markdown(
    f"""
    <style>
        :root {{
            --scm-page-bg: {theme_tokens["page_bg"]};
            --scm-surface: {theme_tokens["surface"]};
            --scm-surface-2: {theme_tokens["surface_2"]};
            --scm-card-bg: {theme_tokens["card_bg"]};
            --scm-input-bg: {theme_tokens["input_bg"]};
            --scm-text: {theme_tokens["text"]};
            --scm-muted: {theme_tokens["muted"]};
            --scm-border: {theme_tokens["border"]};
            --scm-soft-border: {theme_tokens["soft_border"]};
            --scm-table-head: {theme_tokens["table_head"]};
            --scm-hover: {theme_tokens["hover"]};
            --scm-hero-bg: {theme_tokens["hero_bg"]};
            --scm-hero-title: {theme_tokens["hero_title"]};
            --scm-hero-subtitle: {theme_tokens["hero_subtitle"]};
            --scm-hero-kicker: {theme_tokens["hero_kicker"]};
            --scm-theme-shadow: {theme_tokens["shadow"]};
        }}

        html,
        body,
        .stApp,
        [data-testid="stAppViewContainer"],
        [data-testid="stMain"] {{
            background: var(--scm-page-bg) !important;
            color: var(--scm-text) !important;
        }}

        [data-testid="stMainBlockContainer"] {{
            color: var(--scm-text) !important;
        }}

        [data-testid="stMarkdownContainer"],
        [data-testid="stMarkdownContainer"] p,
        [data-testid="stMarkdownContainer"] li,
        [data-testid="stMarkdownContainer"] h1,
        [data-testid="stMarkdownContainer"] h2,
        [data-testid="stMarkdownContainer"] h3,
        [data-testid="stMarkdownContainer"] h4,
        [data-testid="stWidgetLabel"],
        [data-testid="stWidgetLabel"] p {{
            color: var(--scm-text);
        }}

        [data-testid="stCaptionContainer"],
        [data-testid="stCaptionContainer"] p {{
            color: var(--scm-muted) !important;
        }}

        .hero-shell {{
            background: var(--scm-hero-bg) !important;
            border-color: var(--scm-border) !important;
            border-left-color: var(--scm-indigo) !important;
            box-shadow: var(--scm-theme-shadow) !important;
        }}

        .hero-title {{
            color: var(--scm-hero-title) !important;
        }}

        .hero-subtitle {{
            color: var(--scm-hero-subtitle) !important;
        }}

        .hero-kicker {{
            color: var(--scm-hero-kicker) !important;
        }}

        .metric-card,
        .metric-card-base,
        div[data-testid="stVerticalBlockBorderWrapper"] {{
            background: var(--scm-card-bg) !important;
            border-color: var(--scm-border) !important;
            box-shadow: var(--scm-theme-shadow) !important;
            color: var(--scm-text) !important;
        }}

        .metric-value,
        .metric-value-sm,
        .section-heading .title {{
            color: var(--scm-text) !important;
        }}

        .metric-title,
        .metric-header,
        .metric-footnote,
        .section-heading .subtitle {{
            color: var(--scm-muted) !important;
        }}

        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        div[data-baseweb="textarea"] > div {{
            background: var(--scm-input-bg) !important;
            border-color: var(--scm-border) !important;
            color: var(--scm-text) !important;
        }}

        div[data-baseweb="select"] *,
        div[data-baseweb="input"] *,
        div[data-baseweb="textarea"] * {{
            color: var(--scm-text) !important;
        }}

        [data-baseweb="popover"] [role="listbox"],
        [data-baseweb="popover"] ul {{
            background: var(--scm-surface) !important;
            color: var(--scm-text) !important;
        }}

        [data-baseweb="popover"] [role="option"] {{
            color: var(--scm-text) !important;
        }}

        [data-baseweb="popover"] [role="option"]:hover {{
            background: var(--scm-hover) !important;
        }}

        button[data-baseweb="tab"] {{
            color: var(--scm-muted) !important;
        }}

        button[data-baseweb="tab"][aria-selected="true"] {{
            color: var(--scm-text) !important;
        }}

        [data-testid="stExpander"] {{
            background: var(--scm-card-bg) !important;
            border-color: var(--scm-border) !important;
            color: var(--scm-text) !important;
        }}

        .pareto-html-shell {{
            background: var(--scm-card-bg) !important;
            border-color: var(--scm-border) !important;
        }}

        .pareto-html-table {{
            color: var(--scm-text) !important;
        }}

        .pareto-html-table thead th {{
            color: var(--scm-muted) !important;
            background: var(--scm-table-head) !important;
            border-bottom-color: var(--scm-border) !important;
        }}

        .pareto-html-table tbody td {{
            color: var(--scm-text) !important;
            border-bottom-color: var(--scm-soft-border) !important;
        }}

        .pareto-html-table tbody tr:hover {{
            background: var(--scm-hover) !important;
        }}

        .status-strip {{
            background: var(--scm-card-bg) !important;
            border-color: var(--scm-border) !important;
        }}

        .status-pill,
        .info-chip,
        .pareto-count,
        .pareto-status {{
            border-color: var(--scm-border) !important;
        }}

        hr {{
            border-top-color: var(--scm-border) !important;
        }}

        [data-testid="stDialog"] [role="dialog"] {{
            background: var(--scm-surface) !important;
            color: var(--scm-text) !important;
        }}

        div[data-testid="stDialog"] [data-testid="stFileUploader"] section {{
            background: var(--scm-surface-2) !important;
        }}

        .import-dialog-note {{
            background: var(--scm-surface-2) !important;
            border-color: var(--scm-border) !important;
            color: var(--scm-muted) !important;
        }}

        .scm-theme-control-label {{
            margin: 0 0 0.28rem 0;
            color: var(--scm-muted);
            font-size: 0.60rem;
            font-weight: 850;
            line-height: 1.1;
            letter-spacing: 0.10em;
            text-transform: uppercase;
            text-align: right;
        }}

        .st-key-scm_theme_light button,
        .st-key-scm_theme_dark button {{
            min-height: 36px !important;
            height: 36px !important;
            padding: 0.25rem 0.65rem !important;
            border-radius: 10px !important;
            border: 1px solid var(--scm-border) !important;
            background: var(--scm-surface) !important;
            color: var(--scm-text) !important;
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.06) !important;
            font-size: 0.72rem !important;
            font-weight: 850 !important;
            white-space: nowrap !important;
        }}

        .st-key-scm_theme_light button:hover,
        .st-key-scm_theme_dark button:hover {{
            border-color: rgba(99, 102, 241, 0.55) !important;
            background: var(--scm-hover) !important;
            transform: translateY(-1px);
        }}

        .st-key-scm_theme_light button[kind="primary"],
        .st-key-scm_theme_dark button[kind="primary"] {{
            background: linear-gradient(
                135deg,
                rgba(79,70,229,0.98),
                rgba(37,99,235,0.96)
            ) !important;
            border-color: rgba(99, 102, 241, 0.72) !important;
            color: #ffffff !important;
            box-shadow: 0 7px 18px rgba(37, 99, 235, 0.18) !important;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

theme_space, theme_light_col, theme_dark_col = st.columns(
    [10.0, 1.0, 1.0],
    gap="small",
)

with theme_space:
    st.markdown(
        "<div class='scm-theme-control-label'>Display Theme</div>",
        unsafe_allow_html=True,
    )

with theme_light_col:
    st.button(
        "☀️ Light",
        key="scm_theme_light",
        type="primary" if SCM_THEME == "light" else "secondary",
        use_container_width=True,
        on_click=set_scm_theme,
        args=("light",),
    )

with theme_dark_col:
    st.button(
        "🌙 Dark",
        key="scm_theme_dark",
        type="primary" if SCM_THEME == "dark" else "secondary",
        use_container_width=True,
        on_click=set_scm_theme,
        args=("dark",),
    )


st.markdown(
    """
    <div class="hero-shell">
        <div class="hero-grid">
            <div class="hero-copy">
                <div class="hero-kicker">Supply Chain Management • Executive Analytics</div>
                <div class="hero-title">MUTI MC SCM Executive Control Tower</div>
                <div class="hero-subtitle">
                    Inventory visibility, Pareto risk prioritization, stockout trends, Days of Inventory, and branch-level action monitoring.
                </div>
                <a href="https://scmdrp.streamlit.app/" target="_blank" class="drp-action-btn">
                    Launch Delivery Requirements Plan (DRP) ↗
                </a>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# 3. DATA IMPORT & PROCESSING
# =========================================================
@st.cache_data(show_spinner=False)
def process_excel_file(file_path_or_buffer):
    # Explicit engine selection is required for reliable online deployment,
    # especially when a workbook is passed as BytesIO or downloaded from
    # object storage without a filename/extension.
    excel_engine = detect_excel_engine(file_path_or_buffer)

    try:
        xls = pd.ExcelFile(file_path_or_buffer, engine=excel_engine)
    except ImportError as exc:
        required_package = "openpyxl" if excel_engine == "openpyxl" else "xlrd"
        raise ValueError(
            f"Excel engine '{excel_engine}' is unavailable. "
            f"Add '{required_package}' to requirements.txt and redeploy."
        ) from exc
    except Exception as exc:
        raise ValueError(
            f"Unable to open the Excel workbook using {excel_engine}: {exc}"
        ) from exc
    sheet_map = {
        str(sheet).lower().strip().replace(" ", "_"): sheet
        for sheet in xls.sheet_names
    }

    def get_actual_sheet_name(target):
        if target in sheet_map:
            return sheet_map[target]
        raise ValueError(f"Missing sheet '{target}'. Found: {xls.sheet_names}")

    # ---------------------------
    # Raw data
    # ---------------------------
    raw_df = pd.read_excel(xls, sheet_name=get_actual_sheet_name("raw_data"))
    raw_df.columns = (
        raw_df.columns.astype(str)
        .str.lower()
        .str.strip()
        .str.replace(" ", "_", regex=False)
    )

    if "class" in raw_df.columns:
        raw_df.rename(columns={"class": "pareto_class"}, inplace=True)

    required_raw_cols = [
        "area",
        "branch",
        "pareto_class",
        "stock_status",
        "remaining_inventory",
        "suggested_transfer",
        "doi",
        "model",
    ]
    missing_raw_cols = [col for col in required_raw_cols if col not in raw_df.columns]
    if missing_raw_cols:
        raise ValueError(
            "Raw_Data is missing required column(s): " + ", ".join(missing_raw_cols)
        )

    for col in ["remaining_inventory", "suggested_transfer", "doi"]:
        raw_df[col] = (
            round_series_half_up(
                pd.to_numeric(raw_df[col], errors="coerce").fillna(0)
            )
            .fillna(0)
            .astype(int)
        )

    raw_df["pareto_class"] = raw_df["pareto_class"].astype(str).str.strip()
    raw_df["stock_status"] = raw_df["stock_status"].fillna("").astype(str).str.strip()
    raw_df["area"] = raw_df["area"].fillna("").astype(str).str.strip()
    raw_df["branch"] = raw_df["branch"].fillna("").astype(str).str.strip()

    # ---------------------------
    # KPI sheets
    # ---------------------------
    def parse_kpi_sheet(sheet_name):
        df_raw = pd.read_excel(
            xls,
            sheet_name=get_actual_sheet_name(sheet_name),
            header=None,
        )

        empty_kpi_df = pd.DataFrame(
            columns=[
                "period",
                "after_po",
                "per_branch",
                "class_a_out",
                "before_po",
                "overall_doi",
                "class_a_doi",
            ]
        )

        if df_raw.empty or df_raw.shape[1] < 2:
            return empty_kpi_df

        kpis_to_extract = {
            "MUTI MC : Stock Outrate - Overall after PO Balance": "after_po",
            "MUTI MC : Stock Outrate - Per Branch": "per_branch",
            "Overall Class A Stock Out Rate": "class_a_out",
            "MUTI MC : Stock Outrate - Overall (Before PO Balance)": "before_po",
            "MUTI MC : DoI": "overall_doi",
            "MC Class A Doi": "class_a_doi",
        }

        # -------------------------------------------------
        # FIX: Detect the KPI date header dynamically.
        #
        # The previous version assumed dates always started
        # at Excel column index 2. If January is in the
        # preceding column, that assumption makes the graphs
        # begin in February. We now scan the top rows and use
        # every column that contains a plausible date.
        # -------------------------------------------------
        def parse_header_date(value):
            if pd.isna(value):
                return pd.NaT

            if isinstance(value, (pd.Timestamp, datetime, date, np.datetime64)):
                parsed = pd.to_datetime(value, errors="coerce")
            elif isinstance(value, (int, float, np.integer, np.floating)):
                # Typical Excel serial date range.
                numeric_value = float(value)
                if 20000 <= numeric_value <= 60000:
                    parsed = pd.Timestamp("1899-12-30") + pd.to_timedelta(
                        numeric_value, unit="D"
                    )
                else:
                    return pd.NaT
            else:
                value_text = str(value).strip()
                if not value_text:
                    return pd.NaT
                parsed = pd.to_datetime(value_text, errors="coerce")

            if pd.isna(parsed):
                return pd.NaT

            parsed = pd.Timestamp(parsed)
            # Reject accidental conversions that are not realistic KPI dates.
            if parsed.year < 2000 or parsed.year > 2100:
                return pd.NaT

            return parsed.normalize()

        best_date_row = None
        best_date_columns = []
        best_dates = []

        # The KPI input template normally keeps its date row near the top.
        for row_idx in range(min(10, len(df_raw))):
            row_date_columns = []
            row_dates = []

            for col_idx in range(df_raw.shape[1]):
                parsed_date = parse_header_date(df_raw.iat[row_idx, col_idx])
                if pd.notna(parsed_date):
                    row_date_columns.append(col_idx)
                    row_dates.append(parsed_date)

            if len(row_date_columns) > len(best_date_columns):
                best_date_row = row_idx
                best_date_columns = row_date_columns
                best_dates = row_dates

        if best_date_row is None or not best_date_columns:
            raise ValueError(
                f"No valid KPI date headers found in sheet '{sheet_name}'."
            )

        # Keep date columns in their actual Excel left-to-right order.
        date_col_indexes = best_date_columns
        dates = best_dates

        data = {"period": dates}
        first_col_as_text = df_raw[0].fillna("").astype(str)

        for kpi_name, col_name in kpis_to_extract.items():
            idx = df_raw.index[
                first_col_as_text.str.contains(kpi_name, regex=False, na=False)
            ].tolist()

            if idx:
                values = df_raw.loc[idx[0], date_col_indexes].values
                data[col_name] = pd.to_numeric(values, errors="coerce")
            else:
                data[col_name] = [np.nan] * len(dates)

        clean_df = pd.DataFrame(data)
        clean_df["period"] = pd.to_datetime(clean_df["period"], errors="coerce")

        metric_cols = [
            "after_po",
            "per_branch",
            "class_a_out",
            "before_po",
            "overall_doi",
            "class_a_doi",
        ]

        clean_df = clean_df.dropna(subset=["period"])
        clean_df = clean_df.dropna(subset=metric_cols, how="all")

        # One exact date = one source point.
        # If the workbook repeats the same date, keep the last non-null KPI value.
        def last_valid(series):
            non_null = series.dropna()
            return non_null.iloc[-1] if not non_null.empty else np.nan

        clean_df = (
            clean_df.sort_values("period")
            .groupby("period", as_index=False, sort=True)
            .agg({col: last_valid for col in metric_cols})
        )

        if "ytd" in sheet_name.lower() and not clean_df.empty:
            # YTD = one point per calendar month.
            # The January point is now preserved whenever January exists
            # anywhere in the detected KPI date columns.
            clean_df["month_year"] = clean_df["period"].dt.to_period("M")
            clean_df = (
                clean_df.sort_values("period")
                .drop_duplicates(subset=["month_year"], keep="last")
                .drop(columns=["month_year"])
            )
        else:
            # Weekly = actual dated observations only.
            # Never manufacture/interpolate calendar dates.
            clean_df = clean_df.sort_values("period")

        clean_df = (
            clean_df.drop_duplicates(subset=["period"], keep="last")
            .sort_values("period")
            .reset_index(drop=True)
        )

        return clean_df

    kpi_ytd = parse_kpi_sheet("kpi_ytd_input")
    kpi_weekly = parse_kpi_sheet("kpi_weekly_input")

    return raw_df, kpi_ytd, kpi_weekly


# =========================================================
# 4. DATA SYNCHRONIZATION & PERSISTENCE
#    Sidebar removed. Import opens in a modal dialog.
# =========================================================
cloud_config = get_cloud_storage_config()
saved_workbook_bytes, storage_source = initialize_persistent_workbook()


def persist_uploaded_workbook(uploaded_bytes):
    """Validate and persist one uploaded workbook. Returns a success message."""
    # Full workbook validation happens before any persistent copy is replaced.
    process_excel_file(io.BytesIO(uploaded_bytes))

    persistence_messages = []

    try:
        save_local_cache(uploaded_bytes)
        persistence_messages.append("local cache")
    except Exception as exc:
        # Local cache is operational fallback only; cloud may still succeed.
        st.warning(f"Local cache could not be updated: {exc}")

    new_storage_source = "Local cache"
    if cloud_config["configured"]:
        upload_cloud_workbook(uploaded_bytes)
        new_storage_source = "Cloud • Supabase"
        persistence_messages.append("Supabase cloud storage")

    st.session_state["scm_workbook_bytes"] = uploaded_bytes
    st.session_state["scm_storage_source"] = new_storage_source

    destination_text = (
        " + ".join(persistence_messages)
        if persistence_messages
        else "active dashboard session"
    )
    return f"SCM workbook validated and saved to {destination_text}."


@st.dialog("Data Sync", width="large")
def data_sync_dialog():
    st.markdown(
        """
        <div class="import-dialog-note">
            Synchronize the latest SCM Excel workbook. The file is validated first and
            only then replaces the active persistent dataset. Expected sheets:
            <b>Raw_Data</b>, <b>KPI_YTD_Input</b>, and <b>KPI_Weekly_Input</b>.
        </div>
        """,
        unsafe_allow_html=True,
    )

    dialog_file = st.file_uploader(
        "Select SCM Excel workbook",
        type=["xlsx", "xls"],
        help="Use a genuine Microsoft Excel workbook.",
        key="scm_dialog_uploader",
    )

    if dialog_file is None:
        st.caption("Choose a file, then click Validate & Sync.")
        return

    file_bytes = dialog_file.getvalue()
    file_size_mb = len(file_bytes) / (1024 * 1024)
    st.caption(f"Selected: {dialog_file.name} • {file_size_mb:.2f} MB")

    action_col, info_col = st.columns([1.25, 2.75], gap="small")
    with action_col:
        do_import = st.button(
            "Validate & Sync",
            type="primary",
            use_container_width=True,
            key="scm_data_sync_confirm_button",
        )
    with info_col:
        st.caption(
            "The previous persisted workbook is retained if validation or cloud upload fails."
        )

    if do_import:
        upload_hash = hashlib.sha256(file_bytes).hexdigest()
        if st.session_state.get("scm_last_successful_upload_hash") == upload_hash:
            st.info("This exact workbook is already the active dataset.")
            return

        try:
            with st.spinner("Validating workbook and updating dashboard data..."):
                success_message = persist_uploaded_workbook(file_bytes)
        except Exception as exc:
            st.error(f"Data Sync failed: {exc}")
            return

        st.session_state["scm_last_successful_upload_hash"] = upload_hash
        st.session_state["scm_import_success"] = success_message
        st.cache_data.clear()
        st.rerun()

# =========================================================
# 5. HELPERS
# =========================================================
def calculate_stockout_rate(df, pareto_class=None):
    """Preserves the original dashboard stockout-rate logic."""
    subset = df[df["pareto_class"] == pareto_class] if pareto_class else df

    if len(subset) == 0:
        return 0

    stock_status = subset["stock_status"].fillna("").astype(str).str.lower().str.strip()
    stockouts = (stock_status == "stockout").sum()
    return round_half_up((stockouts / len(subset)) * 100)


def section_heading(title, subtitle=""):
    st.markdown(
        f"""
        <div class="section-heading">
            <span class="dot"></span>
            <span class="title">{title}</span>
            <span class="subtitle">{subtitle}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def prepare_chart_series(df, y_col):
    """
    Build a clean chart-specific series.
    This prevents duplicate x-axis dates and prevents blank rows for other KPIs
    from creating misleading points on the current chart.
    """
    if df.empty or y_col not in df.columns:
        return pd.DataFrame(columns=["period", y_col])

    chart_df = df[["period", y_col]].copy()
    chart_df["period"] = pd.to_datetime(chart_df["period"], errors="coerce")
    chart_df[y_col] = pd.to_numeric(chart_df[y_col], errors="coerce")
    chart_df = chart_df.dropna(subset=["period", y_col])
    chart_df = chart_df.sort_values("period")

    # Hard safety rule: only one visible point per date.
    chart_df = chart_df.drop_duplicates(subset=["period"], keep="last")

    return chart_df.reset_index(drop=True)


def infer_percentage_scale(series):
    """
    Supports KPI files stored either as Excel percentages (0.08 = 8%)
    or as percentage points (8 = 8%).
    """
    numeric = pd.to_numeric(series, errors="coerce").dropna()
    if numeric.empty:
        return 1.0

    # If values are generally above 1, treat them as percentage points.
    return 0.01 if numeric.abs().max() > 1.5 else 1.0


def create_styled_line_chart(
    df,
    y_col,
    title,
    subtitle,
    line_color,
    is_weekly,
    is_percentage=True,
    fill=False,
):
    chart_df = prepare_chart_series(df, y_col)
    fig = go.Figure()

    if chart_df.empty:
        fig.add_annotation(
            text="No valid data available for this KPI",
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            showarrow=False,
            font=dict(size=14, color="#94a3b8"),
        )
        fig.update_layout(
            template=PLOTLY_TEMPLATE,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            height=350,
            margin=dict(t=72, b=34, l=42, r=20),
            title=dict(
                text=f"{title}<br><span style='font-size:10px'>{subtitle}</span>",
                x=0.02,
                xanchor="left",
            ),
            xaxis=dict(visible=False),
            yaxis=dict(visible=False),
        )
        return fig

    plot_y = chart_df[y_col].copy()

    if is_percentage:
        scale = infer_percentage_scale(plot_y)
        # Convert to percentage points, apply standard half-up rounding,
        # then convert back to a Plotly fraction.
        # Example: 8.4% -> 8%, 8.5% -> 9%.
        percentage_points = (plot_y * scale) * 100.0
        rounded_percentage_points = round_series_half_up(percentage_points)
        plot_y = rounded_percentage_points / 100.0
        text_labels = [f"{v * 100:.0f}%" for v in plot_y]
        tick_format = ".0%"
        default_ceiling = 0.10
    else:
        # DOI and other numeric KPIs use standard half-up rounding.
        plot_y = round_series_half_up(plot_y).astype(int)
        text_labels = [f"{v:,.0f}" for v in plot_y]
        tick_format = ",.0f"
        default_ceiling = 10

    # -----------------------------------------------------
    # X-AXIS BEHAVIOR
    # -----------------------------------------------------
    # Weekly View:
    #   Use a categorical axis based only on dates that have data.
    #   Plotly therefore cannot display dates between observations.
    #
    # YTD View:
    #   Use month categories ordered Jan -> Dec.
    #   This guarantees January is the first visible month whenever
    #   a January KPI point exists in the source workbook.
    # -----------------------------------------------------
    if is_weekly:
        chart_x = chart_df["period"].dt.strftime("%d %b %Y")
        tick_text = chart_df["period"].dt.strftime("%d %b")
        category_array = chart_x.tolist()
    else:
        chart_x = chart_df["period"].dt.strftime("%b")
        tick_text = chart_x.tolist()
        category_array = [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        ]

    hover_dates = chart_df["period"].dt.strftime("%d %b %Y")

    if is_percentage:
        hover_template = (
            "<b>%{customdata}</b><br>"
            f"{title}: <b>%{{y:.0%}}</b>"
            "<extra></extra>"
        )
    else:
        hover_template = (
            "<b>%{customdata}</b><br>"
            f"{title}: <b>%{{y:,.0f}}</b>"
            "<extra></extra>"
        )

    line_shape = "linear" if is_weekly else "spline"

    fillcolor = None
    if fill:
        fillcolor = (
            f"rgba({int(line_color[1:3], 16)}, "
            f"{int(line_color[3:5], 16)}, "
            f"{int(line_color[5:7], 16)}, 0.08)"
        )

    # Alternate data-label positions to reduce collisions when neighboring values are close.
    text_positions = [
        "top center" if index % 2 == 0 else "bottom center"
        for index in range(len(chart_df))
    ]

    fig.add_trace(
        go.Scatter(
            x=chart_x,
            y=plot_y,
            customdata=hover_dates,
            name="Actual",
            mode="lines+markers+text",
            text=text_labels,
            textposition=text_positions,
            textfont=dict(size=10, family="Arial"),
            line=dict(
                shape=line_shape,
                width=3.2,
                dash="solid",
                color=line_color,
            ),
            marker=dict(
                size=8,
                color=line_color,
                line=dict(width=2, color="#ffffff"),
            ),
            fill="tozeroy" if fill else "none",
            fillcolor=fillcolor,
            hovertemplate=hover_template,
            cliponaxis=False,
        )
    )

    # -----------------------------------------------------
    # DIRECTION / TREND GUIDE — POSITIONED ABOVE ACTUAL
    # -----------------------------------------------------
    # The linear fit determines direction only. For presentation clarity, the
    # dashed guide is vertically repositioned into a dedicated band above the
    # highest Actual observation. This guarantees that it never crosses or
    # obscures the measured KPI line and avoids implying a second KPI value.
    trend_y = None
    trend_direction = "Stable"
    if len(chart_df) >= 2:
        trend_x = (
            (chart_df["period"] - chart_df["period"].min())
            .dt.total_seconds()
            .to_numpy(dtype=float)
            / 86400.0
        )
        trend_source_y = pd.to_numeric(plot_y, errors="coerce").to_numpy(dtype=float)
        valid_trend = np.isfinite(trend_x) & np.isfinite(trend_source_y)

        if valid_trend.sum() >= 2 and np.ptp(trend_x[valid_trend]) > 0:
            slope, intercept = np.polyfit(
                trend_x[valid_trend],
                trend_source_y[valid_trend],
                1,
            )
            fitted_y = slope * trend_x + intercept

            actual_valid = trend_source_y[np.isfinite(trend_source_y)]
            actual_max = float(np.nanmax(actual_valid))
            actual_min = float(np.nanmin(actual_valid))
            actual_span = max(actual_max - actual_min, 0.0)

            # Minimum visual separation is unit-aware.
            minimum_gap = 0.010 if is_percentage else 1.0
            minimum_amplitude = 0.006 if is_percentage else 0.65
            gap = max(actual_span * 0.18, minimum_gap)
            amplitude = max(actual_span * 0.10, minimum_amplitude)

            fitted_range = np.ptp(fitted_y[valid_trend])
            if fitted_range > 0:
                normalized_fit = (
                    fitted_y - np.nanmin(fitted_y[valid_trend])
                ) / fitted_range
            else:
                normalized_fit = np.full_like(fitted_y, 0.5, dtype=float)

            # Entire dashed guide sits above the highest Actual point.
            trend_y = actual_max + gap + (normalized_fit * amplitude)

            fitted_delta = float(fitted_y[-1] - fitted_y[0])
            flat_threshold = max(actual_span * 0.03, 0.001 if is_percentage else 0.10)
            if fitted_delta > flat_threshold:
                trend_direction = "Upward"
            elif fitted_delta < -flat_threshold:
                trend_direction = "Downward"
            else:
                trend_direction = "Stable"

            direction_symbol = {
                "Upward": "↑",
                "Downward": "↓",
                "Stable": "→",
            }[trend_direction]

            fig.add_trace(
                go.Scatter(
                    x=chart_x,
                    y=trend_y,
                    name=f"Trend Direction {direction_symbol}",
                    mode="lines",
                    line=dict(
                        width=2.4,
                        dash="dash",
                        color="rgba(100,116,139,0.95)",
                    ),
                    hovertemplate=(
                        f"Trend direction: <b>{trend_direction} {direction_symbol}</b>"
                        "<br><span style='font-size:10px'>Guide positioned above Actual for readability</span>"
                        "<extra></extra>"
                    ),
                    connectgaps=False,
                )
            )

    max_observed = pd.to_numeric(plot_y, errors="coerce").max()
    if trend_y is not None and len(trend_y):
        trend_max = np.nanmax(trend_y)
        if np.isfinite(trend_max):
            max_observed = max(max_observed, trend_max)
    if pd.isna(max_observed) or max_observed <= 0:
        y_max = default_ceiling
    else:
        # Keep enough breathing room above the separated trend guide without
        # compressing the Actual series excessively.
        y_max = max_observed * 1.16

    if is_percentage:
        y_max = max(y_max, 0.05)
    else:
        y_max = max(y_max, 5)

    latest_period = chart_df["period"].max()
    latest_text = (
        latest_period.strftime("%d %b %Y")
        if pd.notna(latest_period)
        else "N/A"
    )

    xaxis_config = dict(
        type="category",
        showgrid=False,
        categoryorder="array",
        categoryarray=category_array,
        tickangle=0,
        linecolor="rgba(148,163,184,0.18)",
    )

    if is_weekly:
        # Exact observations only: no in-between dates.
        xaxis_config.update(
            tickmode="array",
            tickvals=chart_x.tolist(),
            ticktext=tick_text.tolist(),
        )
    else:
        # Jan-Dec month ordering; only months present in the data render.
        xaxis_config.update(
            tickmode="array",
            tickvals=chart_x.tolist(),
            ticktext=tick_text,
        )

    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=365,
        margin=dict(t=82, b=44, l=46, r=20),
        hovermode="closest",
        hoverlabel=dict(
            bgcolor=PLOTLY_HOVER_BG,
            bordercolor=PLOTLY_HOVER_BORDER,
            font=dict(color=PLOTLY_HOVER_TEXT, size=11),
        ),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.015,
            xanchor="right",
            x=0.99,
            bgcolor="rgba(0,0,0,0)",
            font=dict(size=10, color="#94a3b8"),
        ),
        title=dict(
            text=(
                f"{title}<br>"
                f"<span style='font-size:10px; color:#94a3b8;'>"
                f"{subtitle} • {len(chart_df)} DATA PERIOD(S) • LATEST {latest_text.upper()}"
                f"</span>"
            ),
            x=0.02,
            y=0.96,
            xanchor="left",
            yanchor="top",
            font=dict(size=16, family="Arial"),
        ),
        xaxis=xaxis_config,
        yaxis=dict(
            showgrid=True,
            gridwidth=1,
            gridcolor="rgba(148,163,184,0.14)",
            zeroline=False,
            tickformat=tick_format,
            range=[0, y_max],
        ),
    )

    return fig


def performance_card(title, value, note):
    return f"""
    <div class='metric-card'>
        <div class='metric-title'>{title}</div>
        <div class='metric-value-sm'>{value}%</div>
        <div class='metric-footnote'>{note}</div>
    </div>
    """


def card_html(title, value, badge_text, color_theme, icon):
    return f"""
    <div class='metric-card-base'>
        <div class='metric-header'>
            {title}
            <div class='icon-box'>{icon}</div>
        </div>
        <div class='metric-value'>{value}%</div>
        <div><span class='badge badge-{color_theme}'>{badge_text}</span></div>
    </div>
    """

# =========================================================
# INITIALIZE PRIMARY TABS
# =========================================================
tab_inventory, tab_procurements = st.tabs(["📊 Inventory Control Tower", "📦 Procurements"])

with tab_inventory:
    # =========================================================
    # 4A. MUTI MC TRENDS HEADER + DATA SYNC ACTION
    # =========================================================
    trend_title_col, trend_sync_col = st.columns([5.45, 0.55], gap="small", vertical_alignment="center")

    with trend_title_col:
        st.markdown(
            """
            <div class="section-heading trend-heading-inline">
                <span class="dot"></span>
                <span class="title">MUTI MC Trends</span>
                
            </div>
            """,
            unsafe_allow_html=True,
        )

    with trend_sync_col:
        st.markdown(
            '<div class="data-sync-shell"><span class="data-sync-caption">Latest workbook</span></div>',
            unsafe_allow_html=True,
        )
        if st.button(
            "Data Sync",
            type="primary",
            use_container_width=True,
            key="open_scm_data_sync_dialog",
            help="Validate and synchronize the latest SCM Excel workbook.",
        ):
            data_sync_dialog()

    import_success = st.session_state.pop("scm_import_success", None)
    if import_success:
        st.success(import_success)

    if st.session_state.get("scm_cloud_warning"):
        with st.expander("Cloud storage connection notice"):
            st.caption(st.session_state["scm_cloud_warning"])

    if st.session_state.get("scm_local_warning"):
        with st.expander("Local workbook notice"):
            st.caption(st.session_state["scm_local_warning"])


    if saved_workbook_bytes is None:
        st.info(
            "Click Data Sync beside MUTI MC Trends to upload your workbook. "
            "For online deployment, configure Supabase secrets so the last uploaded workbook "
            "survives server restarts and redeployments."
        )
        st.stop()

    try:
        raw_data, kpi_ytd, kpi_weekly = process_excel_file(
            io.BytesIO(saved_workbook_bytes)
        )
    except Exception as e:
        st.error(f"Import Failed: {e}")
        st.info(
            "Use a genuine Microsoft Excel .xlsx or .xls workbook. "
            "For .xlsx, requirements.txt must include openpyxl; for legacy .xls, "
            "it must include xlrd. If an older invalid workbook is stored in the cloud, "
            "upload a valid workbook once to replace it."
        )
        st.stop()

    # Executive operating-status strip
    all_kpi_dates = pd.concat(
        [
            kpi_ytd["period"] if "period" in kpi_ytd else pd.Series(dtype="datetime64[ns]"),
            kpi_weekly["period"] if "period" in kpi_weekly else pd.Series(dtype="datetime64[ns]"),
        ],
        ignore_index=True,
    )
    all_kpi_dates = pd.to_datetime(all_kpi_dates, errors="coerce").dropna()
    dashboard_latest_date = (
        all_kpi_dates.max().strftime("%d %b %Y")
        if not all_kpi_dates.empty
        else "No KPI date"
    )

    persistence_label = (
        "Cloud persistent"
        if storage_source == "Cloud • Supabase"
        else "Local fallback"
    )

    st.markdown(
        f"""

        """,
        unsafe_allow_html=True,
    )

    # =========================================================
    # 6. MUTI MC TRENDS — CONTROLS & CHARTS
    # =========================================================
    st.markdown("<div style='height:0.15rem'></div>", unsafe_allow_html=True)
    control_col1, control_col2 = st.columns([1.35, 4.65], gap="small")

    with control_col1:
        timeframe = st.selectbox(
            "TIMEFRAME",
            ["Year-to-Date (YTD)", "Weekly View"],
            index=0,
        )

    kpi_data = kpi_weekly if timeframe == "Weekly View" else kpi_ytd
    is_weekly = timeframe == "Weekly View"

    with control_col2:
        if not kpi_data.empty and kpi_data["period"].notna().any():
            latest_kpi_date = pd.to_datetime(kpi_data["period"], errors="coerce").max()
            latest_label = latest_kpi_date.strftime("%d %b %Y")
        else:
            latest_label = "No KPI date"

        st.markdown(
            f"""
            <div style='padding-top:28px;'>
                <span class='info-chip'>View: {timeframe}</span>
                <span class='info-chip'>Latest KPI: {latest_label}</span>
                <span class='info-chip'>Actual data dates only</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # Requested graph order:
    # 1. MC Class A DoI
    # 2. Days of Inventory
    # 3. Per Branch OOS
    # 4. Overall Class A Rate
    # 5. Overall Before PO Balance
    # 6. Overall After PO Balance

    row1_left, row1_right = st.columns(2, gap="small")
    with row1_left:
        with st.container(border=True):
            st.plotly_chart(
                create_styled_line_chart(
                kpi_data,
                "class_a_doi",
                "MC Class A DoI",
                "CLASS A DAYS OF INVENTORY",
                "#7c3aed",
                is_weekly=is_weekly,
                is_percentage=False,
                fill=True,
                ),
                use_container_width=True,
            )

    with row1_right:
        with st.container(border=True):
            st.plotly_chart(
                create_styled_line_chart(
                kpi_data,
                "overall_doi",
                "Days of Inventory",
                "OVERALL INVENTORY COVERAGE",
                "#2563eb",
                is_weekly=is_weekly,
                is_percentage=False,
                fill=True,
                ),
                use_container_width=True,
            )

    row2_left, row2_right = st.columns(2, gap="small")
    with row2_left:
        with st.container(border=True):
            st.plotly_chart(
                create_styled_line_chart(
                kpi_data,
                "per_branch",
                "Per Branch OOS",
                "STOCKOUT RATE",
                "#0ea5e9",
                is_weekly=is_weekly,
                is_percentage=True,
                ),
                use_container_width=True,
            )

    with row2_right:
        with st.container(border=True):
            st.plotly_chart(
                create_styled_line_chart(
                kpi_data,
                "class_a_out",
                "Overall Class A Rate",
                "CLASS A STOCKOUT RATE",
                "#f43f5e",
                is_weekly=is_weekly,
                is_percentage=True,
                ),
                use_container_width=True,
            )

    row3_left, row3_right = st.columns(2, gap="small")
    with row3_left:
        with st.container(border=True):
            st.plotly_chart(
                create_styled_line_chart(
                kpi_data,
                "before_po",
                "Overall Before PO Balance",
                "STOCKOUT RATE BEFORE PO BALANCE",
                "#f59e0b",
                is_weekly=is_weekly,
                is_percentage=True,
                ),
                use_container_width=True,
            )

    with row3_right:
        with st.container(border=True):
            st.plotly_chart(
                create_styled_line_chart(
                kpi_data,
                "after_po",
                "Overall After PO Balance",
                "STOCKOUT RATE AFTER PO BALANCE",
                "#10b981",
                is_weekly=is_weekly,
                is_percentage=True,
                fill=True,
                ),
                use_container_width=True,
            )

    st.markdown("---")


    # =========================================================
    # 7. NETWORK SCOPE — MOVED BELOW MUTI MC TRENDS
    # =========================================================
    section_heading(
        "Network Scope",
        "Filter the operational view without changing the network-level trend history",
    )

    scope_col1, scope_col2 = st.columns([1.5, 3.5])

    with scope_col1:
        areas = ["All Areas"] + sorted(
            [area for area in raw_data["area"].dropna().unique() if str(area).strip()]
        )
        selected_area = st.selectbox("NETWORK SCOPE", areas)

    area_data = raw_data.copy()
    if selected_area != "All Areas":
        area_data = area_data[area_data["area"] == selected_area]

    with scope_col2:
        active_branches = area_data["branch"].replace("", np.nan).dropna().nunique()
        active_models = area_data["model"].replace("", np.nan).dropna().nunique()
        stock_records = len(area_data)

        st.markdown(
            f"""
            <div style='padding-top:28px;'>
                <span class='info-chip'>Scope: {selected_area}</span>
                <span class='info-chip'>{active_branches} Branch(es)</span>
                <span class='info-chip'>{active_models} Model(s)</span>
                <span class='info-chip'>{stock_records:,} Stock Record(s)</span>
            </div>
            """,
            unsafe_allow_html=True,
        )


    # =========================================================
    # 8. PERFORMANCE OVERVIEW — MOVED BELOW TRENDS
    # =========================================================
    rate_a = calculate_stockout_rate(area_data, "Class A")
    rate_b = calculate_stockout_rate(area_data, "Class B")
    rate_c = calculate_stockout_rate(area_data, "Class C")
    avg_rate = round_half_up((rate_a + rate_b + rate_c) / 3)

    st.markdown("<br>", unsafe_allow_html=True)
    section_heading(
        "Performance Overview",
        f"Current raw-data stockout profile • {selected_area}",
    )

    m1, m2, m3, m4 = st.columns(4, gap="small")
    m1.markdown(
        performance_card("Class A Rate", rate_a, "Highest-priority Pareto inventory"),
        unsafe_allow_html=True,
    )
    m2.markdown(
        performance_card("Class B Rate", rate_b, "Medium-priority Pareto inventory"),
        unsafe_allow_html=True,
    )
    m3.markdown(
        performance_card("Class C Rate", rate_c, "Lower-priority Pareto inventory"),
        unsafe_allow_html=True,
    )
    m4.markdown(
        performance_card("Average Rate", avg_rate, "Average of Class A, B and C rates"),
        unsafe_allow_html=True,
    )


    # =========================================================
    # 9. STOCK OUT RATE PER AREA — AVERAGE + CLASS A
    # =========================================================
    st.markdown("<br>", unsafe_allow_html=True)
    section_heading(
        "Stock Out Rate per Area",
        "Average = mean of Class A/B/C rates • Class A = Class A Stock Out Count ÷ Class A Total Stock Status Count",
    )

    area_rates = []

    # IMPORTANT: calculate every KPI independently PER AREA first.
    # Example: AREA I uses only AREA I records; AREA II uses only AREA II records.
    for area, a_df in raw_data.groupby("area", sort=True, dropna=True):
        if not str(area).strip():
            continue

        a_df = a_df.copy()

        # Existing Average Stock Out Rate logic is intentionally preserved,
        # but it is calculated only from the current area's records.
        avg_area_rate = round_half_up(
            (
                calculate_stockout_rate(a_df, "Class A")
                + calculate_stockout_rate(a_df, "Class B")
                + calculate_stockout_rate(a_df, "Class C")
            )
            / 3
        )

        # CLASS A STOCK OUT RATE — PER AREA
        # Numerator   = Class A Stock Out Count in THIS AREA
        # Denominator = Total Class A Stock Status Count in THIS AREA
        # Formula     = Numerator / Denominator * 100
        normalized_class = (
            a_df["pareto_class"].fillna("").astype(str).str.strip().str.casefold()
        )
        class_a_df = a_df[normalized_class.eq("class a")].copy()

        class_a_status = (
            class_a_df["stock_status"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.casefold()
        )

        class_a_stockout_count = int(class_a_status.eq("stockout").sum())
        # Total Stock Status Count follows the dashboard logic: one Class A row = one
        # Class A stock-status record for the current area.
        class_a_total_stock_status_count = int(len(class_a_df))

        class_a_area_rate = (
            round_half_up(
                (class_a_stockout_count / class_a_total_stock_status_count) * 100
            )
            if class_a_total_stock_status_count > 0
            else 0
        )

        area_rates.append(
            {
                "Area": area,
                "Average Stock Out Rate": avg_area_rate,
                "Class A Stock Out Rate": class_a_area_rate,
                "Class A Stock Out Count": class_a_stockout_count,
                "Class A Total Stock Status Count": class_a_total_stock_status_count,
            }
        )

    area_rates_df = pd.DataFrame(area_rates)

    if area_rates_df.empty:
        st.info("No area-level stockout data available.")
    else:
        # Keep one common area order so the two side-by-side charts are easy to compare.
        area_rates_df = area_rates_df.sort_values(
            "Average Stock Out Rate", ascending=False
        ).reset_index(drop=True)
        area_order = area_rates_df["Area"].tolist()

        avg_col, class_a_col = st.columns(2, gap="small")

        with avg_col:
            fig_bar = px.bar(
                area_rates_df,
                x="Area",
                y="Average Stock Out Rate",
                text="Average Stock Out Rate",
                template=PLOTLY_TEMPLATE,
                title="Average Stock Out Rate per Area",
            )

            fig_bar.update_traces(
                marker_color="#6366f1",
                marker_line=dict(width=0),
                opacity=0.92,
                texttemplate="%{text:.0f}%",
                textposition="outside",
                cliponaxis=False,
                hovertemplate=(
                    "<b>%{x}</b><br>"
                    "Average Stock Out Rate: <b>%{y:.0f}%</b>"
                    "<extra></extra>"
                ),
            )

            avg_bar_max = area_rates_df["Average Stock Out Rate"].max()
            fig_bar.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                height=425,
                margin=dict(t=62, b=58, l=44, r=20),
                showlegend=False,
                bargap=0.28,
                title=dict(x=0.02, xanchor="left", font=dict(size=16)),
                hoverlabel=dict(
                    bgcolor=PLOTLY_HOVER_BG,
                    bordercolor=PLOTLY_HOVER_BORDER,
                    font=dict(color=PLOTLY_HOVER_TEXT, size=11),
                ),
                xaxis=dict(
                    title="",
                    type="category",
                    categoryorder="array",
                    categoryarray=area_order,
                    showgrid=False,
                    tickangle=0,
                    automargin=True,
                    linecolor="rgba(148,163,184,0.18)",
                ),
                yaxis=dict(
                    title="Stockout Rate",
                    ticksuffix="%",
                    range=[0, max(10, avg_bar_max * 1.24)],
                    gridcolor="rgba(148,163,184,0.14)",
                    zeroline=False,
                    automargin=True,
                ),
            )

            st.plotly_chart(fig_bar, use_container_width=True)

        with class_a_col:
            fig_class_a = px.bar(
                area_rates_df,
                x="Area",
                y="Class A Stock Out Rate",
                text="Class A Stock Out Rate",
                template=PLOTLY_TEMPLATE,
                title="Class A Stock Out Rate per Area",
                custom_data=["Class A Stock Out Count", "Class A Total Stock Status Count"],
            )

            fig_class_a.update_traces(
                marker_color="#f43f5e",
                marker_line=dict(width=0),
                opacity=0.92,
                texttemplate="%{text:.0f}%",
                textposition="outside",
                cliponaxis=False,
                hovertemplate=(
                    "<b>%{x}</b><br>"
                    "Class A Stock Out Rate: <b>%{y:.0f}%</b><br>"
                    "Class A Stock Out Count: <b>%{customdata[0]}</b><br>"
                    "Class A Total Stock Status Count: <b>%{customdata[1]}</b>"
                    "<extra></extra>"
                ),
            )

            class_a_bar_max = area_rates_df["Class A Stock Out Rate"].max()
            fig_class_a.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                height=425,
                margin=dict(t=62, b=58, l=44, r=20),
                showlegend=False,
                bargap=0.28,
                title=dict(x=0.02, xanchor="left", font=dict(size=16)),
                hoverlabel=dict(
                    bgcolor=PLOTLY_HOVER_BG,
                    bordercolor=PLOTLY_HOVER_BORDER,
                    font=dict(color=PLOTLY_HOVER_TEXT, size=11),
                ),
                xaxis=dict(
                    title="",
                    type="category",
                    categoryorder="array",
                    categoryarray=area_order,
                    showgrid=False,
                    tickangle=0,
                    automargin=True,
                    linecolor="rgba(148,163,184,0.18)",
                ),
                yaxis=dict(
                    title="Class A Stockout Rate",
                    ticksuffix="%",
                    range=[0, max(10, class_a_bar_max * 1.24)],
                    gridcolor="rgba(148,163,184,0.14)",
                    zeroline=False,
                    automargin=True,
                ),
            )

            st.plotly_chart(fig_class_a, use_container_width=True)


    # =========================================================
    # 10. CLASS A BRANCH RANKING — HIGH RISK + ZERO STOCKOUT
    # =========================================================
    st.markdown("<br>", unsafe_allow_html=True)
    section_heading(
        "Class A Branch Stockout Ranking",
        f"Top branch risks and zero-stockout leaders • {selected_area}",
    )

    TOP_BRANCH_LIMIT = 10

    # Use the active Network Scope. When "All Areas" is selected this ranks the
    # entire branch network; when one area is selected it ranks only that area's branches.
    branch_rank_source = area_data.copy()
    branch_rank_source["_normalized_class"] = (
        branch_rank_source["pareto_class"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
    )
    branch_rank_source["_normalized_status"] = (
        branch_rank_source["stock_status"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
    )
    branch_rank_source["branch"] = (
        branch_rank_source["branch"].fillna("").astype(str).str.strip()
    )
    branch_rank_source["area"] = (
        branch_rank_source["area"].fillna("").astype(str).str.strip()
    )

    branch_class_a = branch_rank_source[
        branch_rank_source["_normalized_class"].eq("class a")
        & branch_rank_source["branch"].ne("")
    ].copy()

    if branch_class_a.empty:
        st.info("No Class A branch records are available for the selected network scope.")
    else:
        # One Class A row = one Class A stock-status record, matching the area KPI logic.
        branch_class_a["_is_stockout"] = branch_class_a["_normalized_status"].eq(
            "stockout"
        ).astype(int)

        branch_class_a_summary = (
            branch_class_a.groupby(["area", "branch"], as_index=False, dropna=False)
            .agg(
                **{
                    "Class A Stock Out Count": ("_is_stockout", "sum"),
                    "Class A Total Stock Status Count": ("_is_stockout", "size"),
                }
            )
        )

        branch_class_a_summary["Class A Stock Out Rate"] = branch_class_a_summary.apply(
            lambda row: round_half_up(
                (
                    row["Class A Stock Out Count"]
                    / row["Class A Total Stock Status Count"]
                )
                * 100
            )
            if row["Class A Total Stock Status Count"] > 0
            else 0,
            axis=1,
        )

        # Add the area to labels only when the dashboard is showing the full network.
        if selected_area == "All Areas":
            branch_class_a_summary["Branch Display"] = (
                branch_class_a_summary["branch"]
                + "  •  "
                + branch_class_a_summary["area"]
            )
        else:
            branch_class_a_summary["Branch Display"] = branch_class_a_summary["branch"]

        # Highest-risk branches: positive Class A OOS only, ranked descending.
        high_class_a_branches = (
            branch_class_a_summary[
                branch_class_a_summary["Class A Stock Out Rate"] > 0
            ]
            .sort_values(
                [
                    "Class A Stock Out Rate",
                    "Class A Stock Out Count",
                    "Class A Total Stock Status Count",
                    "branch",
                ],
                ascending=[False, False, False, True],
            )
            .head(TOP_BRANCH_LIMIT)
            .reset_index(drop=True)
        )

        # Zero-stockout leaders: exactly 0% Class A OOS, ranked by Class A coverage count.
        # The bar length uses coverage count because every qualifying OOS rate is 0%.
        zero_class_a_branches = (
            branch_class_a_summary[
                branch_class_a_summary["Class A Stock Out Rate"] == 0
            ]
            .sort_values(
                ["Class A Total Stock Status Count", "branch"],
                ascending=[False, True],
            )
            .head(TOP_BRANCH_LIMIT)
            .reset_index(drop=True)
        )

        high_rank_col, zero_rank_col = st.columns(2, gap="small")

        with high_rank_col:
            if high_class_a_branches.empty:
                st.success("No branch has a Class A Stock Out Rate above 0% in this scope.")
            else:
                high_order = high_class_a_branches["Branch Display"].tolist()
                high_rate_max = float(
                    high_class_a_branches["Class A Stock Out Rate"].max()
                )

                fig_high_class_a = px.bar(
                    high_class_a_branches,
                    x="Class A Stock Out Rate",
                    y="Branch Display",
                    orientation="h",
                    text="Class A Stock Out Rate",
                    template=PLOTLY_TEMPLATE,
                    title=f"Top {len(high_class_a_branches)} Highest Class A Stock Out Rate",
                    custom_data=[
                        "area",
                        "branch",
                        "Class A Stock Out Count",
                        "Class A Total Stock Status Count",
                    ],
                )

                fig_high_class_a.update_traces(
                    marker_color="#f43f5e",
                    marker_line=dict(width=0),
                    opacity=0.94,
                    texttemplate="%{text:.0f}%",
                    textposition="outside",
                    cliponaxis=False,
                    hovertemplate=(
                        "<b>%{customdata[1]}</b><br>"
                        "Area: <b>%{customdata[0]}</b><br>"
                        "Class A Stock Out Rate: <b>%{x:.0f}%</b><br>"
                        "Class A Stock Out Count: <b>%{customdata[2]}</b><br>"
                        "Class A Total Stock Status Count: <b>%{customdata[3]}</b>"
                        "<extra></extra>"
                    ),
                )

                fig_high_class_a.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    height=max(410, 42 * len(high_class_a_branches) + 120),
                    margin=dict(t=68, b=46, l=24, r=46),
                    showlegend=False,
                    bargap=0.28,
                    title=dict(x=0.02, xanchor="left", font=dict(size=16)),
                    hoverlabel=dict(
                        bgcolor=PLOTLY_HOVER_BG,
                        bordercolor=PLOTLY_HOVER_BORDER,
                        font=dict(color=PLOTLY_HOVER_TEXT, size=11),
                    ),
                    xaxis=dict(
                        title="Class A Stockout Rate",
                        ticksuffix="%",
                        range=[0, min(105, max(10, high_rate_max * 1.18))],
                        gridcolor="rgba(148,163,184,0.14)",
                        zeroline=False,
                        automargin=True,
                    ),
                    yaxis=dict(
                        title="",
                        type="category",
                        categoryorder="array",
                        categoryarray=high_order,
                        autorange="reversed",
                        showgrid=False,
                        automargin=True,
                    ),
                )

                st.plotly_chart(fig_high_class_a, use_container_width=True)

        with zero_rank_col:
            if zero_class_a_branches.empty:
                st.warning("No branch currently has a 0% Class A Stock Out Rate in this scope.")
            else:
                zero_class_a_branches = zero_class_a_branches.copy()
                zero_class_a_branches["Zero Rate Label"] = "0% OOS"
                zero_order = zero_class_a_branches["Branch Display"].tolist()
                zero_coverage_max = float(
                    zero_class_a_branches["Class A Total Stock Status Count"].max()
                )

                fig_zero_class_a = px.bar(
                    zero_class_a_branches,
                    x="Class A Total Stock Status Count",
                    y="Branch Display",
                    orientation="h",
                    text="Zero Rate Label",
                    template=PLOTLY_TEMPLATE,
                    title=f"Top {len(zero_class_a_branches)} Branches with 0% Class A Stock Out Rate",
                    custom_data=[
                        "area",
                        "branch",
                        "Class A Stock Out Rate",
                        "Class A Stock Out Count",
                        "Class A Total Stock Status Count",
                    ],
                )

                fig_zero_class_a.update_traces(
                    marker_color="#10b981",
                    marker_line=dict(width=0),
                    opacity=0.92,
                    textposition="outside",
                    cliponaxis=False,
                    hovertemplate=(
                        "<b>%{customdata[1]}</b><br>"
                        "Area: <b>%{customdata[0]}</b><br>"
                        "Class A Stock Out Rate: <b>%{customdata[2]:.0f}%</b><br>"
                        "Class A Stock Out Count: <b>%{customdata[3]}</b><br>"
                        "Class A Total Stock Status Count: <b>%{customdata[4]}</b>"
                        "<extra></extra>"
                    ),
                )

                fig_zero_class_a.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    height=max(410, 42 * len(zero_class_a_branches) + 120),
                    margin=dict(t=68, b=46, l=24, r=58),
                    showlegend=False,
                    bargap=0.28,
                    title=dict(x=0.02, xanchor="left", font=dict(size=16)),
                    hoverlabel=dict(
                        bgcolor=PLOTLY_HOVER_BG,
                        bordercolor=PLOTLY_HOVER_BORDER,
                        font=dict(color=PLOTLY_HOVER_TEXT, size=11),
                    ),
                    xaxis=dict(
                        title="Class A Stock Status Coverage Count",
                        range=[0, max(1, zero_coverage_max * 1.22)],
                        gridcolor="rgba(148,163,184,0.14)",
                        zeroline=False,
                        automargin=True,
                    ),
                    yaxis=dict(
                        title="",
                        type="category",
                        categoryorder="array",
                        categoryarray=zero_order,
                        autorange="reversed",
                        showgrid=False,
                        automargin=True,
                    ),
                )

                st.plotly_chart(fig_zero_class_a, use_container_width=True)
                st.caption(
                    "Zero-stockout leaders are ranked by Class A stock-status coverage count; "
                    "every branch shown has exactly 0% Class A Stock Out Rate."
                )


    st.markdown("---")


    # =========================================================
    # 11. BRANCH-LEVEL SECTION & PARETO TABLES
    # =========================================================
    section_heading(
        "Branch-Level Stockout Performance",
        f"Operational drill-down • {selected_area}",
    )

    head_col1, head_col2 = st.columns([3.4, 1.6])

    with head_col1:
        st.markdown(
            "<div style='padding-top:12px; color:#94a3b8; font-size:0.82rem;'>"
            "Select a branch to review stockout risk and Pareto action models."
            "</div>",
            unsafe_allow_html=True,
        )

    with head_col2:
        branch_options = ["All Branches"] + sorted(
            [
                branch
                for branch in area_data["branch"].dropna().unique()
                if str(branch).strip()
            ]
        )
        selected_branch = st.selectbox("BRANCH", branch_options)

    branch_data = area_data.copy()
    if selected_branch != "All Branches":
        branch_data = branch_data[branch_data["branch"] == selected_branch]

    br_rate_a = calculate_stockout_rate(branch_data, "Class A")
    br_rate_b = calculate_stockout_rate(branch_data, "Class B")
    br_rate_c = calculate_stockout_rate(branch_data, "Class C")
    br_avg = calculate_stockout_rate(branch_data)

    c1, c2, c3, c4 = st.columns(4, gap="small")
    c1.markdown(
        card_html(
            "BRANCH CLASS A RATE",
            br_rate_a,
            "HIGH PRIORITY RISK",
            "red",
            "A",
        ),
        unsafe_allow_html=True,
    )
    c2.markdown(
        card_html(
            "BRANCH CLASS B RATE",
            br_rate_b,
            "MEDIUM PRIORITY RISK",
            "yellow",
            "B",
        ),
        unsafe_allow_html=True,
    )
    c3.markdown(
        card_html(
            "BRANCH CLASS C RATE",
            br_rate_c,
            "LOW PRIORITY RISK",
            "green",
            "C",
        ),
        unsafe_allow_html=True,
    )
    c4.markdown(
        card_html(
            "BRANCH AVERAGE",
            br_avg,
            "PERFORMANCE INDEX",
            "blue",
            "Σ",
        ),
        unsafe_allow_html=True,
    )

    st.markdown("<br><br>", unsafe_allow_html=True)
    section_heading(
        "Pareto Action Models",
        f"Transfer priorities for {selected_branch}",
    )

    def render_pareto_table(df, pareto_class, hex_color):
        class_df = df[df["pareto_class"] == pareto_class].copy()
        item_count = len(class_df)

        st.markdown(
            f"""
            <div class='pareto-header' style='border-color: {hex_color};'>
                <span style='font-size:1.05rem; font-weight:900; color:{hex_color};'>
                    {pareto_class.upper()}
                </span>
                <span class='pareto-count' style='color:{hex_color};'>
                    ● {item_count} Items
                </span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if item_count == 0:
            st.info(f"No {pareto_class} items active.")
            return

        class_df["remaining_inventory"] = (
            round_series_half_up(class_df["remaining_inventory"]).astype(int)
        )
        class_df["suggested_transfer"] = (
            round_series_half_up(class_df["suggested_transfer"]).astype(int)
        )
        class_df["doi"] = round_series_half_up(class_df["doi"]).astype(int)

        class_df = class_df.sort_values(
            by=["suggested_transfer", "doi"],
            ascending=[False, True],
        ).reset_index(drop=True)

        class_df.index += 1
        class_df = class_df.reset_index().rename(columns={"index": "Rank"})

        display_cols = [
            "Rank",
            "model",
            "stock_status",
            "remaining_inventory",
            "suggested_transfer",
            "doi",
        ]

        final_df = class_df[display_cols].copy()
        final_df.columns = ["Rank", "Model", "Status", "Inventory", "Transfer", "DOI"]

        rows_html = []
        for _, row in final_df.iterrows():
            status_text = str(row["Status"])
            status_class = " stockout" if status_text.strip().lower() == "stockout" else ""
            rows_html.append(
                "<tr>"
                f"<td class='num'>{int(row['Rank'])}</td>"
                f"<td>{html.escape(str(row['Model']))}</td>"
                f"<td><span class='pareto-status{status_class}'>{html.escape(status_text)}</span></td>"
                f"<td class='num'>{int(row['Inventory']):,}</td>"
                f"<td class='num'>{int(row['Transfer']):,}</td>"
                f"<td class='num'>{int(row['DOI']):,}</td>"
                "</tr>"
            )

        table_html = (
            "<div class='pareto-html-shell'>"
            "<table class='pareto-html-table'>"
            "<colgroup>"
            "<col style='width:7%'>"
            "<col style='width:35%'>"
            "<col style='width:16%'>"
            "<col style='width:14%'>"
            "<col style='width:14%'>"
            "<col style='width:14%'>"
            "</colgroup>"
            "<thead><tr>"
            "<th class='num'>Rank</th><th>Model</th><th>Status</th>"
            "<th class='num'>Inventory</th><th class='num'>Transfer</th><th class='num'>DOI</th>"
            "</tr></thead>"
            "<tbody>" + "".join(rows_html) + "</tbody>"
            "</table></div>"
        )
        st.markdown(table_html, unsafe_allow_html=True)


    # Full-width stacked Pareto panels eliminate horizontal scrolling and keep
    # all six operational columns readable at standard laptop/browser widths.
    for pareto_class, pareto_color in [
        ("Class A", "#f87171"),
        ("Class B", "#fbbf24"),
        ("Class C", "#4ade80"),
    ]:
        with st.container(border=True):
            render_pareto_table(branch_data, pareto_class, pareto_color)
        st.markdown("<div class='pareto-panel-spacer'></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.caption(
        "SCM Executive Control Tower  • Modal data import. • [Access DRP Module](https://scmdrp.streamlit.app/)"
    )


# =========================================================
# SECONDARY TAB: PROCUREMENTS
# =========================================================
with tab_procurements:
    st.markdown("<br>", unsafe_allow_html=True)
    section_heading(
        "Procurements Control", 
        "Manage purchase orders, incoming stock allocations, and supplier lead times"
    )

    proc_col1, proc_col2 = st.columns(2, gap="small")

    with proc_col1:
        st.markdown(
            """
            <div class='metric-card' style='min-height: 250px; width: 100%;'>
                <div class='metric-header'>
                    <span style='color: #6366f1; font-weight: 900; font-size: 1.1rem;'>🏍️ MOTORCYCLE UNITS</span>
                </div>
                <hr style='margin: 10px 0; border-color: rgba(148, 163, 184, 0.1);'>
                <div style='color: var(--scm-muted); font-size: 0.85rem; line-height: 1.6;'>
                    <b>Pipeline Visibility</b><br>
                    • Supplier Lead Times<br>
                    • Incoming Allocations<br>
                    • Backorder Tracking<br><br>
                    <i>(Procurement data integration pending)</i>
                </div>
            </div>
            """, 
            unsafe_allow_html=True
        )

    with proc_col2:
        st.markdown(
            """
            <div class='metric-card' style='min-height: 250px; width: 100%;'>
                <div class='metric-header'>
                    <span style='color: #10b981; font-weight: 900; font-size: 1.1rem;'>⚙️ SPARE PARTS</span>
                </div>
                <hr style='margin: 10px 0; border-color: rgba(148, 163, 184, 0.1);'>
                <div style='color: var(--scm-muted); font-size: 0.85rem; line-height: 1.6;'>
                    <b>Replenishment Status</b><br>
                    • Active Purchase Orders<br>
                    • Critical Shortages<br>
                    • Parts Delivery Schedule<br><br>
                    <i>(Procurement data integration pending)</i>
                </div>
            </div>
            """, 
            unsafe_allow_html=True
        )
