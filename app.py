import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import os
import io
import requests
import hashlib
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
    initial_sidebar_state="expanded",
)

# =========================================================
# 2. EXECUTIVE UI / THEME
# =========================================================
st.markdown(
    """
    <style>
        /* ---------- Global / maximum-width layout ---------- */
        html, body, [data-testid="stAppViewContainer"] {
            overflow-x: hidden !important;
        }

        /* Let Streamlit calculate the space beside the sidebar, then use all of it. */
        .block-container,
        [data-testid="stMainBlockContainer"] {
            box-sizing: border-box !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            padding-top: 0.55rem !important;
            padding-bottom: 2.50rem !important;
            padding-left: 0.45rem !important;
            padding-right: 0.45rem !important;
            margin: 0 !important;
            overflow-x: hidden !important;
        }

        section[data-testid="stMain"],
        section[data-testid="stMain"] > div,
        [data-testid="stMainBlockContainer"] > div {
            min-width: 0 !important;
            max-width: 100% !important;
            overflow-x: hidden !important;
        }

        @media (min-width: 1600px) {
            .block-container,
            [data-testid="stMainBlockContainer"] {
                padding-left: 0.55rem !important;
                padding-right: 0.55rem !important;
            }
        }

        @media (max-width: 900px) {
            .block-container,
            [data-testid="stMainBlockContainer"] {
                padding-left: 0.35rem !important;
                padding-right: 0.35rem !important;
                padding-top: 0.45rem !important;
            }
        }

        /* Calm page background: never layer behind the executive banner. */
        [data-testid="stAppViewContainer"] {
            background: var(--background-color) !important;
            background-image: none !important;
        }

        /* ---------- Sidebar control center ---------- */
        [data-testid="stSidebar"] {
            width: 286px !important;
            min-width: 286px !important;
            max-width: 286px !important;
            border-right: 1px solid rgba(148,163,184,0.16);
            box-shadow: none !important;
            overflow-x: hidden !important;
        }

        [data-testid="stSidebar"] > div:first-child {
            width: 286px !important;
            min-width: 286px !important;
            max-width: 286px !important;
        }

        [data-testid="stSidebar"] * {
            box-sizing: border-box;
            max-width: 100%;
        }

        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] li,
        [data-testid="stSidebar"] span {
            overflow-wrap: anywhere;
            word-break: normal;
        }

        [data-testid="stSidebar"] h2 {
            line-height: 1.18 !important;
            margin-bottom: 0.55rem !important;
        }

        [data-testid="stSidebar"] [data-testid="stFileUploader"] {
            margin-top: 0.65rem;
            margin-bottom: 0.80rem;
        }

        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
            min-height: 92px;
            border-radius: 12px !important;
        }

        @media (max-width: 900px) {
            [data-testid="stSidebar"],
            [data-testid="stSidebar"] > div:first-child {
                width: min(286px, 86vw) !important;
                min-width: min(286px, 86vw) !important;
                max-width: min(286px, 86vw) !important;
            }
        }

        /* ---------- Executive hero: flat, single-layer, overlap-proof ---------- */
        .hero-shell {
            box-sizing: border-box !important;
            display: block !important;
            position: relative !important;
            isolation: isolate;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            height: auto !important;
            min-height: 0 !important;
            margin: 0 0 1.10rem 0 !important;
            padding: clamp(20px, 1.7vw, 28px) clamp(20px, 2.0vw, 32px) !important;
            border: 1px solid #263247 !important;
            border-left: 5px solid #6366f1 !important;
            border-radius: 15px !important;
            background: #0f172a !important;
            background-color: #0f172a !important;
            background-image: none !important;
            box-shadow: 0 8px 20px rgba(2, 6, 23, 0.16) !important;
            backdrop-filter: none !important;
            -webkit-backdrop-filter: none !important;
            overflow: hidden !important;
            transform: none !important;
            z-index: 0 !important;
        }

        /* Explicitly suppress decorative layers from older cached CSS versions. */
        .hero-shell::before,
        .hero-shell::after,
        .hero-grid::before,
        .hero-grid::after,
        .hero-copy::before,
        .hero-copy::after {
            content: none !important;
            display: none !important;
            background: none !important;
        }

        .hero-grid,
        .hero-copy {
            box-sizing: border-box !important;
            display: block !important;
            position: static !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            height: auto !important;
            margin: 0 !important;
            padding: 0 !important;
            background: transparent !important;
            box-shadow: none !important;
            transform: none !important;
        }

        .hero-kicker,
        .hero-title,
        .hero-subtitle {
            box-sizing: border-box !important;
            display: block !important;
            position: static !important;
            float: none !important;
            clear: both !important;
            width: 100% !important;
            max-width: 100% !important;
            height: auto !important;
            min-height: 0 !important;
            padding: 0 !important;
            margin-left: 0 !important;
            margin-right: 0 !important;
            white-space: normal !important;
            overflow-wrap: break-word !important;
            word-break: normal !important;
            overflow: visible !important;
            text-overflow: clip !important;
            transform: none !important;
            background: transparent !important;
        }

        .hero-kicker {
            color: #a5b4fc !important;
            font-size: clamp(0.64rem, 0.67vw, 0.74rem) !important;
            font-weight: 800 !important;
            line-height: 1.35 !important;
            letter-spacing: 0.10em !important;
            text-transform: uppercase !important;
            margin-top: 0 !important;
            margin-bottom: 9px !important;
        }

        .hero-title {
            color: #f8fafc !important;
            font-size: clamp(1.55rem, 2.05vw, 2.35rem) !important;
            font-weight: 900 !important;
            line-height: 1.14 !important;
            letter-spacing: -0.02em !important;
            margin-top: 0 !important;
            margin-bottom: 12px !important;
        }

        .hero-subtitle {
            color: #cbd5e1 !important;
            font-size: clamp(0.82rem, 0.88vw, 0.96rem) !important;
            font-weight: 500 !important;
            line-height: 1.50 !important;
            margin-top: 0 !important;
            margin-bottom: 0 !important;
        }

        @media (max-width: 760px) {
            .hero-shell {
                padding: 18px 16px 19px 16px !important;
                border-left-width: 4px !important;
                border-radius: 13px !important;
                margin-bottom: 0.95rem !important;
            }

            .hero-kicker {
                font-size: 0.60rem !important;
                margin-bottom: 8px !important;
            }

            .hero-title {
                font-size: clamp(1.30rem, 6vw, 1.80rem) !important;
                line-height: 1.16 !important;
                margin-bottom: 10px !important;
            }

            .hero-subtitle {
                font-size: 0.78rem !important;
                line-height: 1.48 !important;
            }
        }

        /* ---------- Section labels ---------- */
        .section-heading {
            display: grid;
            grid-template-columns: auto minmax(0, 1fr) minmax(0, auto);
            align-items: center;
            gap: 10px;
            margin: 0.45rem 0 0.85rem 0;
            min-width: 0;
        }

        .section-heading .dot {
            width: 10px;
            height: 10px;
            border-radius: 50%;
            background: #6366f1;
            box-shadow: 0 0 0 5px rgba(99, 102, 241, 0.12);
        }

        .section-heading .title {
            font-size: 1.16rem;
            font-weight: 850;
        }

        .section-heading .subtitle {
            color: #94a3b8;
            font-size: 0.78rem;
            margin-left: 0;
            text-align: right;
            white-space: normal;
            overflow-wrap: anywhere;
        }

        @media (max-width: 760px) {
            .section-heading {
                grid-template-columns: auto minmax(0, 1fr);
                align-items: start;
            }

            .section-heading .subtitle {
                grid-column: 2;
                text-align: left;
                font-size: 0.72rem;
            }
        }

        /* ---------- KPI cards ---------- */
        .metric-card,
        .metric-card-base {
            position: relative;
            border: 1px solid rgba(148, 163, 184, 0.20);
            border-radius: 16px;
            padding: 17px 19px;
            min-height: 126px;
            background:
                linear-gradient(145deg, rgba(255,255,255,0.045), rgba(99,102,241,0.035));
            box-shadow:
                0 10px 28px rgba(15, 23, 42, 0.065),
                inset 0 1px 0 rgba(255,255,255,0.04);
            transition: transform 160ms ease, box-shadow 160ms ease;
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
            background: linear-gradient(#6366f1, #0ea5e9);
            opacity: 0.85;
        }

        .metric-card:hover,
        .metric-card-base:hover {
            transform: translateY(-1px);
            box-shadow: 0 14px 34px rgba(15, 23, 42, 0.09);
        }

        .metric-card-base {
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            height: 100%;
        }

        .metric-title,
        .metric-header {
            color: #94a3b8;
            font-size: 0.70rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.08em;
        }

        .metric-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .metric-value,
        .metric-value-sm {
            font-size: 2rem;
            font-weight: 900;
            line-height: 1.1;
            margin: 10px 0 6px 0;
        }

        .metric-footnote {
            color: #94a3b8;
            font-size: 0.72rem;
        }

        .badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 999px;
            font-size: 0.62rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        .badge-red { background: rgba(244, 63, 94, 0.14); color: #fb7185; }
        .badge-yellow { background: rgba(234, 179, 8, 0.14); color: #facc15; }
        .badge-green { background: rgba(16, 185, 129, 0.14); color: #34d399; }
        .badge-blue { background: rgba(59, 130, 246, 0.14); color: #60a5fa; }

        .icon-box {
            display: flex;
            align-items: center;
            justify-content: center;
            width: 30px;
            height: 30px;
            border-radius: 9px;
            background: rgba(99, 102, 241, 0.10);
        }

        /* ---------- Info chips ---------- */
        .info-chip {
            display: inline-block;
            padding: 5px 9px;
            border-radius: 999px;
            border: 1px solid rgba(148, 163, 184, 0.25);
            font-size: 0.70rem;
            color: #94a3b8;
            margin-right: 6px;
            margin-bottom: 4px;
        }

        /* ---------- Dashboard utility cards ---------- */
        .status-strip {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 8px;
            margin: 0.65rem 0 0.25rem 0;
            padding: 10px 12px;
            border: 1px solid rgba(148,163,184,0.18);
            border-radius: 13px;
            background: rgba(15,23,42,0.025);
        }

        .status-pill {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 5px 9px;
            border: 1px solid rgba(148,163,184,0.20);
            border-radius: 999px;
            font-size: 0.70rem;
            font-weight: 700;
            color: #94a3b8;
            white-space: nowrap;
        }

        .status-pill strong {
            color: inherit;
            font-weight: 850;
        }

        .status-dot {
            width: 7px;
            height: 7px;
            border-radius: 50%;
            background: #10b981;
            box-shadow: 0 0 0 3px rgba(16,185,129,0.12);
        }

        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 16px !important;
            border-color: rgba(148,163,184,0.18) !important;
            box-shadow: 0 8px 22px rgba(15,23,42,0.045);
            background: rgba(255,255,255,0.012);
        }

        div[data-testid="stPlotlyChart"] {
            border-radius: 14px;
            overflow: hidden;
        }

        /* ---------- Pareto ---------- */
        .pareto-panel-spacer {
            height: 0.35rem;
        }

        .pareto-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            width: 100%;
            min-width: 0;
            padding: 0.70rem 0 0.62rem 0;
            border-bottom: 2px solid;
            margin-bottom: 0.75rem;
        }

        .pareto-header > span:first-child {
            min-width: 0;
            overflow-wrap: anywhere;
        }

        .pareto-count {
            flex: 0 0 auto;
            white-space: nowrap;
            border: 1px solid rgba(128,128,128,0.25);
            padding: 4px 9px;
            border-radius: 999px;
            font-size: 0.72rem;
            font-weight: 800;
        }

        /* Full-width Pareto tables: no browser-level horizontal scrolling. */
        .pareto-table-shell {
            width: 100%;
            min-width: 0;
            overflow: hidden;
        }

        /* ---------- Streamlit controls ---------- */
        div[data-baseweb="select"] > div,
        div[data-testid="stFileUploader"] section {
            border-radius: 10px;
        }

        div[data-testid="stDataFrame"] {
            width: 100% !important;
            min-width: 0 !important;
            max-width: 100% !important;
            border: 1px solid rgba(148, 163, 184, 0.18);
            border-radius: 12px;
            overflow: hidden !important;
        }

        div[data-testid="stDataFrame"] > div {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
        }

        hr {
            border: none;
            border-top: 1px solid rgba(148, 163, 184, 0.20);
            margin: 1.4rem 0;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero-shell">
        <div class="hero-grid">
            <div class="hero-copy">
                <div class="hero-kicker">Supply Chain Management • Executive Analytics</div>
                <div class="hero-title">MUTI MC SCM Executive Control Tower</div>
                <div class="hero-subtitle">Inventory visibility, Pareto risk prioritization, stockout trends, Days of Inventory, and branch-level action monitoring.</div>
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
# 4. SIDEBAR / DATA SYNCHRONIZATION & PERSISTENCE
# =========================================================
cloud_config = get_cloud_storage_config()
saved_workbook_bytes, storage_source = initialize_persistent_workbook()

with st.sidebar:
    st.markdown("## SCM Control Center")
    st.caption(
        "Upload the latest SCM workbook. The newest successful upload "
        "becomes the dashboard's active persistent dataset."
    )

    if cloud_config["configured"]:
        st.success("Cloud persistence: Connected")
        st.caption(
            f"Supabase bucket: {cloud_config['bucket']}  \n"
            f"Object: {cloud_config['object_name']}"
        )
    else:
        st.warning("Cloud persistence: Not configured")
        st.caption(
            "Local cache works for development, but a hosted app may lose "
            "local files after a restart or redeploy."
        )

    uploaded_file = st.file_uploader(
        "Upload SCM Excel",
        type=["xlsx", "xls"],
        help="Expected sheets: Raw_Data, KPI_YTD_Input, KPI_Weekly_Input",
    )

    if uploaded_file is not None:
        uploaded_bytes = uploaded_file.getvalue()
        upload_hash = hashlib.sha256(uploaded_bytes).hexdigest()
        already_processed = (
            st.session_state.get("scm_last_upload_hash") == upload_hash
        )

        # Streamlit keeps the uploaded file in widget state across reruns.
        # Process each unique upload once so filters do not repeatedly write
        # the same workbook to cloud storage.
        if not already_processed:
            # Validate the workbook before replacing the persisted copy.
            try:
                process_excel_file(io.BytesIO(uploaded_bytes))
            except Exception as exc:
                st.error(f"Upload rejected: {exc}")
                st.session_state["scm_last_upload_hash"] = upload_hash
            else:
                persistence_messages = []

                # Always keep a local cache for local use / operational fallback.
                try:
                    save_local_cache(uploaded_bytes)
                    persistence_messages.append("local cache")
                except Exception as exc:
                    st.warning(f"Local cache could not be updated: {exc}")

                if cloud_config["configured"]:
                    try:
                        upload_cloud_workbook(uploaded_bytes)
                        storage_source = "Cloud • Supabase"
                        persistence_messages.append("Supabase cloud storage")
                    except Exception as exc:
                        st.error(
                            "Workbook was validated, but cloud persistence failed. "
                            f"The previous cloud workbook remains unchanged. Details: {exc}"
                        )
                        storage_source = "Local cache"
                else:
                    storage_source = "Local cache"

                st.session_state["scm_workbook_bytes"] = uploaded_bytes
                st.session_state["scm_storage_source"] = storage_source
                st.session_state["scm_last_upload_hash"] = upload_hash
                saved_workbook_bytes = uploaded_bytes
                st.cache_data.clear()

                if persistence_messages:
                    st.success(
                        "SCM workbook updated and saved to "
                        + " + ".join(persistence_messages)
                        + "."
                    )
    else:
        # Clearing the uploader allows the same file to be intentionally
        # uploaded again later in the same browser session.
        st.session_state.pop("scm_last_upload_hash", None)

    if st.session_state.get("scm_cloud_warning"):
        with st.expander("Cloud storage connection notice"):
            st.caption(st.session_state["scm_cloud_warning"])

    if st.session_state.get("scm_local_warning"):
        with st.expander("Local workbook notice"):
            st.caption(st.session_state["scm_local_warning"])

    st.markdown("---")
    st.caption("Dashboard navigation")
    st.markdown(
        "1. **MUTI MC Trends**  \n"
        "2. **Network Scope**  \n"
        "3. **Performance Overview**  \n"
        "4. **Average OOS per Area**  \n"
        "5. **Branch-Level Actions**"
    )

    st.markdown("---")
    st.caption("Rounding standard")
    st.markdown("**0.0–0.4 ↓** • **0.5–0.9 ↑**")
    st.caption("Example: 8.4 → 8 • 8.5 → 9")


if saved_workbook_bytes is None:
    st.info(
        "👈 Upload your SCM Excel workbook to begin. "
        "For online deployment, configure Supabase secrets so the last "
        "uploaded workbook survives server restarts and redeployments."
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
            template="plotly",
            height=350,
            margin=dict(t=75, b=35, l=45, r=25),
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
    # DIRECTION / TREND LINE
    # -----------------------------------------------------
    # A linear best-fit trend is drawn as a broken (dashed) line.
    # It is calculated only from the dates that contain real KPI observations;
    # no additional dates or data points are created.
    trend_y = None
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
            trend_y = slope * trend_x + intercept
            # Stockout percentages and DOI cannot be negative in this dashboard.
            trend_y = np.maximum(trend_y, 0)

            trend_hover = (
                "Trend direction: <b>%{y:.0%}</b><extra></extra>"
                if is_percentage
                else "Trend direction: <b>%{y:,.0f}</b><extra></extra>"
            )

            fig.add_trace(
                go.Scatter(
                    x=chart_x,
                    y=trend_y,
                    name="Trend",
                    mode="lines",
                    line=dict(
                        width=2.2,
                        dash="dash",
                        color="rgba(148,163,184,0.90)",
                    ),
                    hovertemplate=trend_hover,
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
        y_max = max_observed * 1.35

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
        template="plotly",
        height=365,
        margin=dict(t=82, b=48, l=50, r=25),
        hovermode="closest",
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
    <div class="status-strip">
        <span class="status-pill">
            <span class="status-dot"></span>
            <strong>Dashboard online</strong>
        </span>
        <span class="status-pill">
            Data source: <strong>{storage_source}</strong>
        </span>
        <span class="status-pill">
            Persistence: <strong>{persistence_label}</strong>
        </span>
        <span class="status-pill">
            Latest KPI: <strong>{dashboard_latest_date}</strong>
        </span>
        <span class="status-pill">
            Rounding: <strong>Half-up</strong>
        </span>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# 6. MUTI MC TRENDS — FIRST SECTION
# =========================================================
st.markdown("<br>", unsafe_allow_html=True)
section_heading(
    "MUTI MC Trends",
    "YTD begins with January when January data exists • Weekly shows actual data dates only",
)

control_col1, control_col2, control_col3 = st.columns([1.6, 1.2, 3.2])

with control_col1:
    timeframe = st.selectbox(
        "TIMEFRAME",
        ["Year-to-Date (YTD)", "Weekly View"],
        index=0,
    )

kpi_data = kpi_weekly if timeframe == "Weekly View" else kpi_ytd
is_weekly = timeframe == "Weekly View"



with control_col3:
    if not kpi_data.empty and kpi_data["period"].notna().any():
        latest_kpi_date = pd.to_datetime(kpi_data["period"], errors="coerce").max()
        latest_label = latest_kpi_date.strftime("%d %b %Y")
    else:
        latest_label = "No KPI date"

    st.markdown(
        f"""
        <div style='padding-top:30px;'>
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

row1_left, row1_right = st.columns(2, gap="large")
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

row2_left, row2_right = st.columns(2, gap="large")
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

row3_left, row3_right = st.columns(2, gap="large")
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
        <div style='padding-top:30px;'>
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

m1, m2, m3, m4 = st.columns(4, gap="medium")
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
# 9. AVERAGE STOCK OUT RATE PER AREA — MOVED BELOW TRENDS
# =========================================================
st.markdown("<br>", unsafe_allow_html=True)
section_heading(
    "Average Stock Out Rate per Area",
    "Raw Data • average of Class A, B and C stockout rates",
)

area_rates = []
for area in sorted([a for a in raw_data["area"].dropna().unique() if str(a).strip()]):
    a_df = raw_data[raw_data["area"] == area]
    avg_area_rate = round_half_up(
        (
            calculate_stockout_rate(a_df, "Class A")
            + calculate_stockout_rate(a_df, "Class B")
            + calculate_stockout_rate(a_df, "Class C")
        )
        / 3
    )
    area_rates.append(
        {"Area": area, "Average Stock Out Rate": avg_area_rate}
    )

area_rates_df = pd.DataFrame(area_rates)

if area_rates_df.empty:
    st.info("No area-level stockout data available.")
else:
    # Vertical executive column chart: one bar per area.
    # Sort highest risk first while keeping each area label fully visible.
    area_rates_df = area_rates_df.sort_values(
        "Average Stock Out Rate", ascending=False
    ).reset_index(drop=True)

    fig_bar = px.bar(
        area_rates_df,
        x="Area",
        y="Average Stock Out Rate",
        text="Average Stock Out Rate",
        template="plotly",
    )

    fig_bar.update_traces(
        marker_color="#6366f1",
        marker_line=dict(width=0),
        texttemplate="%{text:.0f}%",
        textposition="outside",
        cliponaxis=False,
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Average Stock Out Rate: <b>%{y:.0f}%</b>"
            "<extra></extra>"
        ),
    )

    bar_max = area_rates_df["Average Stock Out Rate"].max()
    fig_bar.update_layout(
        height=430,
        margin=dict(t=34, b=70, l=48, r=32),
        showlegend=False,
        bargap=0.34,
        xaxis=dict(
            title="",
            type="category",
            categoryorder="array",
            categoryarray=area_rates_df["Area"].tolist(),
            showgrid=False,
            tickangle=0,
            automargin=True,
            linecolor="rgba(148,163,184,0.18)",
        ),
        yaxis=dict(
            title="Stockout Rate",
            ticksuffix="%",
            range=[0, max(10, bar_max * 1.24)],
            gridcolor="rgba(148,163,184,0.14)",
            zeroline=False,
            automargin=True,
        ),
    )

    st.plotly_chart(fig_bar, use_container_width=True)

st.markdown("---")


# =========================================================
# 10. BRANCH-LEVEL SECTION & PARETO TABLES
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

c1, c2, c3, c4 = st.columns(4, gap="medium")
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

    st.markdown("<div class='pareto-table-shell'>", unsafe_allow_html=True)
    st.dataframe(
        final_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Rank": st.column_config.NumberColumn("Rank", format="%d", width="small"),
            "Model": st.column_config.TextColumn("Model", width="medium"),
            "Status": st.column_config.TextColumn("Status", width="medium"),
            "Inventory": st.column_config.NumberColumn("Inventory", format="%.0f", width="small"),
            "Transfer": st.column_config.NumberColumn("Transfer", format="%.0f", width="small"),
            "DOI": st.column_config.NumberColumn("DOI", format="%.0f", width="small"),
        },
    )
    st.markdown("</div>", unsafe_allow_html=True)


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
    "SCM Executive Control Tower • Executive dashboard UI • Standard half-up rounding • January-aware YTD • Weekly actual-data dates only • Cloud persistence ready."
)
