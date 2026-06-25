import streamlit as st
import pandas as pd
import io
import re
import sys
import tempfile
from copy import copy
from pathlib import Path
from datetime import datetime

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Import validate_invoice if present
sys.path.insert(0, str(Path(__file__).parent))
try:
    from validate_invoice import run as validate_run
    FREEZPAK_ENABLED = True
except ImportError:
    FREEZPAK_ENABLED = False

# Page config
st.set_page_config(
    page_title="Invoice Automation Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
    background-color: #0f1117;
    color: #e0e0e0;
}
.dashboard-header {
    background: linear-gradient(135deg, #1a1f2e 0%, #16213e 100%);
    border: 1px solid #2a3a5c;
    border-radius: 12px;
    padding: 28px 36px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    gap: 20px;
}
.dashboard-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 26px;
    font-weight: 600;
    color: #4fc3f7;
    letter-spacing: -0.5px;
    margin: 0;
}
.dashboard-subtitle {
    font-size: 13px;
    color: #7a8aaa;
    margin: 4px 0 0 0;
    font-weight: 300;
}
.status-dot {
    width: 10px; height: 10px;
    border-radius: 50%;
    background: #4caf50;
    box-shadow: 0 0 8px #4caf50;
    display: inline-block;
    margin-right: 8px;
}
.step-badge {
    display: inline-block;
    background: #4fc3f7;
    color: #0f1117;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    font-weight: 600;
    padding: 2px 10px;
    border-radius: 20px;
    margin-right: 8px;
}
.info-box {
    background: #1a2744;
    border-left: 3px solid #4fc3f7;
    border-radius: 0 8px 8px 0;
    padding: 12px 16px;
    margin: 10px 0;
    font-size: 13px;
    color: #b0c4de;
}
.success-box {
    background: #1a2e1a;
    border-left: 3px solid #4caf50;
    border-radius: 0 8px 8px 0;
    padding: 12px 16px;
    margin: 10px 0;
    font-size: 13px;
    color: #a5d6a7;
}
.warning-box {
    background: #2a2010;
    border-left: 3px solid #ff9800;
    border-radius: 0 8px 8px 0;
    padding: 12px 16px;
    margin: 10px 0;
    font-size: 13px;
    color: #ffcc80;
}
.error-box {
    background: #2a1010;
    border-left: 3px solid #f44336;
    border-radius: 0 8px 8px 0;
    padding: 12px 16px;
    margin: 10px 0;
    font-size: 13px;
    color: #ef9a9a;
}
.placeholder-box {
    background: #1a1a2e;
    border-left: 3px solid #546e8a;
    border-radius: 0 8px 8px 0;
    padding: 12px 16px;
    margin: 10px 0;
    font-size: 13px;
    color: #7a8aaa;
}
.metric-row { display: flex; gap: 12px; margin: 16px 0; }
.metric-card {
    background: #1a1f2e;
    border: 1px solid #2a3a5c;
    border-radius: 10px;
    padding: 16px 20px;
    flex: 1;
    text-align: center;
}
.metric-value { font-family: 'IBM Plex Mono', monospace; font-size: 22px; font-weight: 600; color: #4fc3f7; }
.metric-label { font-size: 11px; color: #7a8aaa; margin-top: 2px; text-transform: uppercase; letter-spacing: 0.5px; }
.sub-type-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    color: #7a8aaa;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 4px;
}
.stTabs [data-baseweb="tab-list"] {
    background: #1a1f2e;
    border-radius: 10px;
    padding: 4px;
    gap: 4px;
    border: 1px solid #2a3a5c;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    color: #7a8aaa;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 13px;
    padding: 8px 20px;
}
.stTabs [aria-selected="true"] {
    background: #4fc3f7 !important;
    color: #0f1117 !important;
    font-weight: 600;
}
.stButton > button {
    background: #4fc3f7;
    color: #0f1117;
    border: none;
    border-radius: 8px;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    font-size: 13px;
    padding: 10px 28px;
    width: 100%;
    transition: opacity 0.2s;
}
.stButton > button:hover { opacity: 0.85; }
.stDataFrame { border-radius: 10px; border: 1px solid #2a3a5c; overflow: hidden; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 24px 32px; }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown(f"""
<div class="dashboard-header">
    <div style="font-size:36px">📦</div>
    <div>
        <p class="dashboard-title">XDOCK INVOICE DASHBOARD</p>
        <p class="dashboard-subtitle">
            <span class="status-dot"></span>
            Jetro / Restaurant Depot &nbsp;·&nbsp; Logistics Operations &nbsp;·&nbsp;
            {datetime.now().strftime("%B %d, %Y")}
        </p>
    </div>
</div>
""", unsafe_allow_html=True)

# XDock tab config
XDOCKS = {
    "🏭  Halls":    {"key": "halls",    "desc": "Halls XDock invoices",    "color": "#4fc3f7"},
    "❄️  Freezpak": {"key": "freezpak", "desc": "Freezpak XDock invoices", "color": "#81d4fa"},
    "📦  EXP":      {"key": "exp",      "desc": "EXP XDock invoices",      "color": "#4db6ac"},
    "🚚  EFH":      {"key": "efh",      "desc": "EFH XDock invoices",      "color": "#80cbc4"},
}

# Invoice type registry
INVOICE_TYPES = {

    "halls": [
        {"name": "IBT",            "key": "halls_ibt",          "placeholder": True},
        {"name": "Trucking & FSC", "key": "halls_trucking_fsc",  "placeholder": True},
        {
            "name": "Warehousing",
            "key":  "halls_warehousing",
            "placeholder": False,
            "sub_types": [
                {"name": "Ancillary",        "key": "halls_warehousing_ancillary",      "placeholder": True},
                {"name": "Inbound",          "key": "halls_warehousing_inbound",        "placeholder": True},
                {"name": "Renewal",          "key": "halls_warehousing_renewal",        "placeholder": True},
                {"name": "Sort & Selection", "key": "halls_warehousing_sort_selection", "placeholder": True},
            ],
        },
    ],

    "freezpak": [
        {"name": "Ancillary",           "key": "freezpak_ancillary",         "placeholder": True},
        {"name": "Inbound",             "key": "freezpak_inbound",           "placeholder": True},
        {"name": "Recurring / Storage", "key": "freezpak_recurring_storage", "placeholder": True},
        {"name": "XDock",               "key": "freezpak_xdock",             "placeholder": False},
    ],

    "exp": [
        {"name": "Inout Billing",        "key": "exp_inout_billing",      "placeholder": False},
        {"name": "Rejected & Repacking", "key": "exp_rejected_repacking", "placeholder": True},
    ],

    "efh": [
        {"name": "CGA",    "key": "efh_cga",    "placeholder": False},
        {"name": "Haines", "key": "efh_haines", "placeholder": False},
        {"name": "Streets","key": "efh_streets","placeholder": False},
        {"name": "Stults", "key": "efh_stults", "placeholder": False},
    ],
}

# Invoice type keys that use validate_invoice.py (Detail + AP tabs, rate check)
VALIDATED_INVOICE_KEYS = {"freezpak_xdock", "exp_inout_billing"}

# Invoice type keys that use the EFH freight-validation workflow
EFH_INVOICE_KEYS = {"efh_cga", "efh_haines", "efh_streets", "efh_stults"}


# =============================================================================
# HELPERS
# =============================================================================

def _norm(value) -> str:
    """Normalise a cell header for case/punctuation-insensitive comparison."""
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def _find_column(ws, keyword: str) -> int | None:
    """Return 1-based column index whose normalised header contains keyword."""
    for col in range(1, ws.max_column + 1):
        raw = ws.cell(row=1, column=col).value
        if raw is not None and keyword in _norm(raw):
            return col
    return None


def _po_to_int(value):
    """Convert a PO# cell value to int, returning None if not convertible."""
    if value is None or str(value).strip() == "":
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    try:
        return int(float(str(value).strip().replace(",", "")))
    except (ValueError, TypeError):
        return None


def _sum_column(ws, col: int, data_start: int = 2) -> float:
    """Sum numeric values in a column (skipping header row)."""
    total = 0.0
    for row in range(data_start, ws.max_row + 1):
        v = ws.cell(row=row, column=col).value
        if v is not None and v != "":
            try:
                total += float(v)
            except (TypeError, ValueError):
                pass
    return total


def _last_data_row(ws, data_start: int = 2) -> int:
    """Return the last row that has at least one non-None value."""
    last = data_start - 1
    for row in range(data_start, ws.max_row + 1):
        if any(ws.cell(row=row, column=c).value is not None
               for c in range(1, ws.max_column + 1)):
            last = row
    return last


# =============================================================================
# PROCESSING FUNCTIONS
# =============================================================================

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Serialise a DataFrame to an in-memory .xlsx blob."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Processed")
    return buf.getvalue()


def process_validated_invoice(uploaded_file) -> tuple:
    """
    Run validate_invoice.run() on any workbook with Detail + AP tabs.
    Used by: Freezpak XDock, EXP Inout Billing.

    Steps: paste-as-values -> rate check (<=0.475/lb) -> suffix groups -> AP export.
    Returns (output_bytes, output_filename, error_message).
    """
    if not FREEZPAK_ENABLED:
        return None, None, (
            "validate_invoice.py not found. "
            "Place it next to app.py and restart."
        )
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path   = Path(tmpdir)
        input_path = tmp_path / uploaded_file.name
        input_path.write_bytes(uploaded_file.getbuffer())
        result_code = validate_run(input_path, tmp_path)

        if result_code == 2:
            exc_files = list(tmp_path.glob("*Exception Report*.xlsx"))
            if exc_files:
                return (
                    exc_files[0].read_bytes(), exc_files[0].name,
                    "Rate validation failed -- Exception Report generated. "
                    "Fix the flagged rows and resubmit.",
                )
            return None, None, "Rate validation failed but no exception report was generated."

        if result_code == 1:
            return None, None, (
                "Processing failed -- check that your file has both a Detail tab and an AP tab."
            )

        ap_files = list(tmp_path.glob("* - AP.xlsx"))
        if ap_files:
            return ap_files[0].read_bytes(), ap_files[0].name, None

        return None, None, "Processing completed but output file could not be found."


def process_efh_invoice(uploaded_file) -> tuple:
    """
    EFH invoice processor -- identical logic for CGA, Haines, Streets, Stults.

    Steps:
      1. Locate AP tab and optional Details tab (case-insensitive sheet names).
      2. Find the 'P.O. #' and 'Freight Amount' columns by header keyword.
      3. Convert all PO values to integers (in memory only).
      4. If a Details tab is present:
           - Sum Freight Amount on AP tab.
           - Sum Freight Amount on Details tab.
           - If totals do not match (tolerance 0.01) -> return Exception Report.
      5. Format the AP tab:
           - Bold header row.
           - Center-align all cells.
           - Auto-filter on header row.
           - Set every column to the same width (max content width, capped at 25).
      6. Remove the Details tab (if present).
      7. Save AP-only workbook as '<original stem>-AP.xlsx'.

    Returns (output_bytes, output_filename, error_message).
    error_message is None on clean success.
    """
    # ── Load workbook ─────────────────────────────────────────────────────────
    file_bytes = uploaded_file.getbuffer()
    try:
        wb = load_workbook(io.BytesIO(bytes(file_bytes)), data_only=True)
    except Exception as exc:
        return None, None, f"Could not open workbook: {exc}"

    # ── Locate AP and Details tabs (case-insensitive) ─────────────────────────
    sheet_map = {s.lower(): s for s in wb.sheetnames}

    ap_name = None
    for candidate in ("ap",):
        if candidate in sheet_map:
            ap_name = sheet_map[candidate]
            break
    if ap_name is None:
        # Fall back: first sheet
        ap_name = wb.sheetnames[0]

    details_name = None
    for candidate in ("details", "detail"):
        if candidate in sheet_map:
            details_name = sheet_map[candidate]
            break

    ws_ap = wb[ap_name]

    # ── Locate required columns on AP ─────────────────────────────────────────
    ap_po_col      = _find_column(ws_ap, "po")
    ap_freight_col = _find_column(ws_ap, "freight")

    if ap_freight_col is None:
        return None, None, (
            "Could not find a 'Freight Amount' column on the AP tab. "
            "Check that your header row uses 'Freight Amount'."
        )

    # ── Step 3: Convert PO values to integers (in memory, AP tab) ────────────
    if ap_po_col is not None:
        for row in range(2, ws_ap.max_row + 1):
            cell = ws_ap.cell(row=row, column=ap_po_col)
            if cell.value is not None:
                cell.value = _po_to_int(cell.value)

    # ── Step 4: Freight validation against Details tab ────────────────────────
    if details_name is not None:
        ws_det = wb[details_name]
        det_freight_col = _find_column(ws_det, "freight")

        if det_freight_col is None:
            return None, None, (
                "Could not find a 'Freight Amount' column on the Details tab. "
                "Check the header names match 'Freight Amount'."
            )

        # Convert PO on Details tab as well
        det_po_col = _find_column(ws_det, "po")
        if det_po_col is not None:
            for row in range(2, ws_det.max_row + 1):
                cell = ws_det.cell(row=row, column=det_po_col)
                if cell.value is not None:
                    cell.value = _po_to_int(cell.value)

        ap_total  = round(_sum_column(ws_ap,  ap_freight_col),  2)
        det_total = round(_sum_column(ws_det, det_freight_col), 2)

        if abs(ap_total - det_total) > 0.01:
            # Build exception report workbook
            exc_wb = Workbook()
            exc_ws = exc_wb.active
            exc_ws.title = "Freight Mismatch"
            exc_ws.append(["Check", "Amount"])
            exc_ws.append(["AP Freight Total",      ap_total])
            exc_ws.append(["Details Freight Total", det_total])
            exc_ws.append(["Difference",            round(ap_total - det_total, 2)])
            buf = io.BytesIO()
            exc_wb.save(buf)
            stem = Path(uploaded_file.name).stem
            return (
                buf.getvalue(),
                f"{stem} - Exception Report.xlsx",
                f"Freight totals do not match -- AP: ${ap_total:,.2f} vs "
                f"Details: ${det_total:,.2f} (difference: ${abs(ap_total - det_total):,.2f}). "
                "Exception Report generated.",
            )

    # ── Step 5: Format the AP tab ─────────────────────────────────────────────
    last_row = _last_data_row(ws_ap)
    last_col = ws_ap.max_column

    center  = Alignment(horizontal="center", vertical="center", wrap_text=False)
    bold    = Font(bold=True)
    normal  = Font(bold=False)

    # Compute max content width per column, then pick the largest for uniform width
    col_widths = []
    for col in range(1, last_col + 1):
        max_len = 0
        for row in range(1, last_row + 1):
            v = ws_ap.cell(row=row, column=col).value
            if v is not None:
                # Date objects format to roughly 10 chars
                text_len = len(str(v).split(" ")[0]) if hasattr(v, "strftime") else len(str(v))
                max_len = max(max_len, text_len)
        col_widths.append(max_len)

    uniform_width = max(min(max(col_widths, default=10) + 2, 25), 10)

    for col in range(1, last_col + 1):
        col_letter = get_column_letter(col)
        ws_ap.column_dimensions[col_letter].width = uniform_width

        for row in range(1, last_row + 1):
            cell = ws_ap.cell(row=row, column=col)
            cell.alignment = center
            cell.font = bold if row == 1 else normal

    # Auto-filter on header row
    ws_ap.auto_filter.ref = (
        f"A1:{get_column_letter(last_col)}{last_row}"
    )

    # ── Step 6: Remove Details tab ────────────────────────────────────────────
    if details_name is not None:
        del wb[details_name]

    # ── Step 7: Save AP-only output ───────────────────────────────────────────
    # Remove any sheet that is not the AP tab
    for sheet in wb.sheetnames:
        if sheet != ap_name:
            del wb[sheet]

    buf = io.BytesIO()
    wb.save(buf)
    stem     = Path(uploaded_file.name).stem
    filename = f"{stem}-AP.xlsx"
    return buf.getvalue(), filename, None


def process_generic(uploaded_file, xdock_label: str) -> tuple:
    """Generic passthrough: normalise columns, tag source, return .xlsx."""
    try:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") \
             else pd.read_excel(uploaded_file)
    except Exception as exc:
        return None, None, f"Could not read file: {exc}"

    df.columns = [str(c).strip().upper().replace(" ", "_") for c in df.columns]
    df = df.dropna(how="all")
    df["XDOCK_SOURCE"]   = xdock_label
    df["PROCESSED_DATE"] = datetime.now().strftime("%Y-%m-%d")
    stem     = Path(uploaded_file.name).stem
    filename = f"{stem} - AP.xlsx"
    return to_excel_bytes(df), filename, None


# =============================================================================
# INVOICE ROUTER
# =============================================================================

def route_invoice(uploaded_file, xdock_key: str, invoice_type_key: str) -> tuple:
    """
    Dispatch to the correct processor.

    Returns (output_bytes | None, output_filename | None, error_message | None).
      None error     -> clean success
      'PLACEHOLDER'  -> logic not yet implemented

    HOW TO ADD A NEW TYPE
    1. Write a process_<name>() function above.
    2. Add an if/elif below.
    3. Set placeholder: False in INVOICE_TYPES.
    4. For validate_invoice.py workflow: add key to VALIDATED_INVOICE_KEYS.
    5. For EFH workflow: add key to EFH_INVOICE_KEYS.
    """

    # Freezpak XDock + EXP Inout Billing: validate_invoice.py pipeline
    if invoice_type_key in VALIDATED_INVOICE_KEYS:
        return process_validated_invoice(uploaded_file)

    # EFH (CGA, Haines, Streets, Stults): freight validation + AP formatting
    if invoice_type_key in EFH_INVOICE_KEYS:
        return process_efh_invoice(uploaded_file)

    # -- ADD NEW ROUTES ABOVE THIS LINE ---------------------------------------
    # Pending: halls_ibt, halls_trucking_fsc, halls_warehousing_*, freezpak_ancillary,
    # freezpak_inbound, freezpak_recurring_storage, exp_rejected_repacking
    # -------------------------------------------------------------------------

    return None, None, "PLACEHOLDER"


# =============================================================================
# SHARED UPLOAD / PREVIEW / PROCESS SECTION
# =============================================================================

def render_invoice_section(
    xdock_key: str,
    invoice_type_cfg: dict,
    xdock_color: str,
    xdock_display: str,
) -> None:
    inv_key      = invoice_type_cfg["key"]
    inv_name     = invoice_type_cfg["name"]
    is_ph        = invoice_type_cfg["placeholder"]
    is_validated = (inv_key in VALIDATED_INVOICE_KEYS)
    is_efh       = (inv_key in EFH_INVOICE_KEYS)

    # Instructions banner
    if is_ph:
        st.markdown("""
        <div class="placeholder-box">
            <b>Logic not configured yet.</b>
            Upload and processing workflow will be added later.
            You can still upload a file to preview its contents.
        </div>
        """, unsafe_allow_html=True)

    elif is_validated:
        if not FREEZPAK_ENABLED:
            st.markdown("""
            <div class="warning-box">
                <b>validate_invoice.py not detected.</b>
                Place it next to app.py and restart.
            </div>
            """, unsafe_allow_html=True)
        st.markdown(f"""
        <div class="info-box">
            <b>How to use:</b><br>
            <span class="step-badge">1</span>
                Upload your {xdock_display} &mdash; {inv_name} workbook (.xlsx) &mdash; must have
                <b>Detail</b> and <b>AP</b> tabs<br>
            <span class="step-badge">2</span> Review the raw data preview<br>
            <span class="step-badge">3</span>
                Click <b>Process Invoice</b> &mdash; rate check runs automatically (max $0.475/lb)<br>
            <span class="step-badge">4</span>
                Download the <b>AP-only output</b>, or the <b>Exception Report</b> if rates fail
        </div>
        """, unsafe_allow_html=True)

    elif is_efh:
        st.markdown(f"""
        <div class="info-box">
            <b>How to use:</b><br>
            <span class="step-badge">1</span>
                Upload your {inv_name} invoice workbook (.xlsx) &mdash; should have an
                <b>AP</b> tab and optionally a <b>Details</b> tab<br>
            <span class="step-badge">2</span> Review the raw data preview<br>
            <span class="step-badge">3</span>
                Click <b>Process Invoice</b> &mdash; PO numbers normalised, freight totals
                validated between AP and Details tabs<br>
            <span class="step-badge">4</span>
                Download the formatted <b>AP-only output</b>, or the
                <b>Exception Report</b> if freight totals do not match
        </div>
        """, unsafe_allow_html=True)

    else:
        st.markdown(f"""
        <div class="info-box">
            <b>How to use:</b><br>
            <span class="step-badge">1</span>
                Upload your {xdock_display} &mdash; {inv_name} invoice file (Excel or CSV)<br>
            <span class="step-badge">2</span> Review the raw data preview<br>
            <span class="step-badge">3</span> Click <b>Process Invoice</b> to run the automation<br>
            <span class="step-badge">4</span> Download the output file
        </div>
        """, unsafe_allow_html=True)

    # File uploader
    accepted_types = ["xlsx", "xls"] if (is_validated or is_efh) else ["xlsx", "xls", "csv"]
    file_hint      = (
        "Excel (.xlsx) with AP tab (+ optional Details tab)"
        if is_efh
        else ("Excel (.xlsx) with Detail and AP tabs" if is_validated else "Excel (.xlsx, .xls) or CSV")
    )

    uploaded = st.file_uploader(
        label=f"Upload {xdock_display} -- {inv_name}",
        type=accepted_types,
        key=f"upload_{inv_key}",
        label_visibility="collapsed",
    )

    if uploaded:
        try:
            raw_df = (
                pd.read_csv(uploaded)
                if uploaded.name.endswith(".csv")
                else pd.read_excel(uploaded, sheet_name=0)
            )
            uploaded.seek(0)
        except Exception as exc:
            st.markdown(
                f'<div class="warning-box">Could not preview file: {exc}</div>',
                unsafe_allow_html=True,
            )
            return

        sheet_label = "Detail tab rows" if is_validated else "AP tab rows" if is_efh else "Total Rows"

        # Detect a freight/amount column for the total card
        freight_total_html = ""
        freight_col_label  = ""
        for col in raw_df.columns:
            col_norm = re.sub(r"[^a-z0-9]", "", str(col).lower())
            if "freight" in col_norm or ("amount" in col_norm and "freight" in col_norm):
                try:
                    total = pd.to_numeric(raw_df[col], errors="coerce").sum()
                    freight_col_label = str(col)
                    freight_total_html = f"""
            <div class="metric-card" style="border-color:#4fc3f7; background:#132030">
                <div class="metric-value" style="color:#4fc3f7">${total:,.2f}</div>
                <div class="metric-label">{freight_col_label} Total</div>
            </div>"""
                    break
                except Exception:
                    pass
        # Fallback: any column with "amount" if no freight column found
        if not freight_total_html:
            for col in raw_df.columns:
                col_norm = re.sub(r"[^a-z0-9]", "", str(col).lower())
                if "amount" in col_norm:
                    try:
                        total = pd.to_numeric(raw_df[col], errors="coerce").sum()
                        freight_col_label = str(col)
                        freight_total_html = f"""
            <div class="metric-card" style="border-color:#4fc3f7; background:#132030">
                <div class="metric-value" style="color:#4fc3f7">${total:,.2f}</div>
                <div class="metric-label">{freight_col_label} Total</div>
            </div>"""
                        break
                    except Exception:
                        pass

        st.markdown(f"""
        <div class="metric-row">
            <div class="metric-card">
                <div class="metric-value">{len(raw_df):,}</div>
                <div class="metric-label">{sheet_label}</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">{len(raw_df.columns)}</div>
                <div class="metric-label">Columns</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">{raw_df.isnull().sum().sum()}</div>
                <div class="metric-label">Blank Cells</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">{uploaded.name.split(".")[-1].upper()}</div>
                <div class="metric-label">File Type</div>
            </div>{freight_total_html}
        </div>
        """, unsafe_allow_html=True)

        with st.expander("📋 Raw Data Preview (first sheet)", expanded=True):
            st.dataframe(raw_df.head(20), use_container_width=True, height=280)

        st.markdown("<br>", unsafe_allow_html=True)

        if not is_ph:
            sess_processed = f"processed_{inv_key}"
            sess_output    = f"output_{inv_key}"

            col1, _col2, _col3 = st.columns([1, 1, 2])
            with col1:
                process_clicked = st.button("Process Invoice", key=f"process_{inv_key}")

            if process_clicked:
                st.session_state.pop(sess_output, None)
                st.session_state[sess_processed] = True

            if st.session_state.get(sess_processed):
                if sess_output not in st.session_state:
                    with st.spinner("Running validation and processing..."):
                        uploaded.seek(0)
                        st.session_state[sess_output] = route_invoice(
                            uploaded, xdock_key, inv_key
                        )

                output_bytes, output_filename, error_msg = st.session_state[sess_output]

                if error_msg and error_msg != "PLACEHOLDER":
                    is_warning = any(w in error_msg.lower()
                                     for w in ("failed", "exception", "mismatch", "do not match"))
                    box = "warning-box" if is_warning else "error-box"
                    st.markdown(f'<div class="{box}">{error_msg}</div>', unsafe_allow_html=True)

                if output_bytes:
                    if not error_msg:
                        st.markdown(
                            '<div class="success-box">Validation passed -- output ready for download</div>',
                            unsafe_allow_html=True,
                        )
                    col_a, _col_b, _col_c = st.columns([1, 1, 2])
                    with col_a:
                        st.download_button(
                            label="Download Output",
                            data=output_bytes,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_{inv_key}",
                        )

    else:
        st.markdown(f"""
        <div style="text-align:center; padding:48px 0; color:#3a4a6a;
                    border: 1.5px dashed #2a3a5c; border-radius:12px; margin-top:12px;">
            <div style="font-size:40px; margin-bottom:12px">&#x2191;</div>
            <div style="font-family:'IBM Plex Mono',monospace; font-size:14px">
                Drop your {inv_name} invoice file here
            </div>
            <div style="font-size:12px; margin-top:6px">{file_hint}</div>
        </div>
        """, unsafe_allow_html=True)


# =============================================================================
# MAIN -- TAB RENDERING
# =============================================================================

tab_labels = list(XDOCKS.keys())
tabs       = st.tabs(tab_labels)

for tab, (label, cfg) in zip(tabs, XDOCKS.items()):
    with tab:
        xdock_key     = cfg["key"]
        xdock_display = xdock_key.capitalize()
        xdock_color   = cfg["color"]
        inv_type_list = INVOICE_TYPES[xdock_key]

        st.markdown(f"""
        <div style="margin-bottom:20px">
            <span style="font-family:'IBM Plex Mono',monospace; font-size:18px;
                         font-weight:600; color:{xdock_color}">{label.strip()}</span>
            <span style="font-size:12px; color:#7a8aaa; margin-left:12px">{cfg['desc']}</span>
        </div>
        """, unsafe_allow_html=True)

        type_names = [t["name"] for t in inv_type_list]
        sel_idx = st.radio(
            "Invoice Type",
            options=range(len(type_names)),
            format_func=lambda i, tn=type_names: tn[i],
            horizontal=True,
            key=f"radio_{xdock_key}",
            label_visibility="collapsed",
        )
        selected_type = inv_type_list[sel_idx]

        st.markdown("<hr style='border:1px solid #2a3a5c; margin:16px 0'>", unsafe_allow_html=True)

        if "sub_types" in selected_type:
            sub_list  = selected_type["sub_types"]
            sub_names = [s["name"] for s in sub_list]
            st.markdown('<div class="sub-type-label">Warehousing Sub-Type</div>', unsafe_allow_html=True)
            sub_idx = st.radio(
                "Warehousing Sub-Type",
                options=range(len(sub_names)),
                format_func=lambda i, sn=sub_names: sn[i],
                horizontal=True,
                key=f"radio_{xdock_key}_warehousing_sub",
                label_visibility="collapsed",
            )
            selected_type = sub_list[sub_idx]
            st.markdown("<hr style='border:1px solid #2a3a5c; margin:16px 0'>", unsafe_allow_html=True)

        render_invoice_section(xdock_key, selected_type, xdock_color, xdock_display)
