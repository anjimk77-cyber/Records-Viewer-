import re

import pandas as pd
import streamlit as st
from datetime import date

import gspread
from google.oauth2.service_account import Credentials

# =========================================================================
# CONFIG
#
# This is a SEPARATE app meant for the manager. It shares the same Google
# Sheet and the same "Customer List.xlsx" as the data-entry app (app.py).
# It has no Pond Details editor and no Harvest Details form. The only
# writes it performs are the recycle-bin row-delete actions below: in
# Sales Details that writes Settle = 'Yes' (Sales Details sheet); in All
# Harvest Details that writes Harvest Status = 'H' (main WaterQualityData
# sheet). Everything else here is read-only:
#   1) "📋 Enter Water Quality Data" — Customer / Farm selection (used only
#      to choose which farm's records to view)
#   2) "📊 All Saved Records" — a live, read-only view of that farm's data
#      straight from the Google Sheet.
# Deploy this file as its own Streamlit app (its own URL/link) so the
# manager gets a separate link from the data-entry app.
# =========================================================================
st.set_page_config(page_title="Shrimp FarmFlow - KMN ", layout="wide", page_icon="📊")

CUSTOMER_FILE = "Customer List.xlsx"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Second Google Sheet — Sales Details. Separate spreadsheet from the
# water-quality WaterQualityData sheet above; the same service account
# must also be shared (as Viewer or Editor) on this sheet.
SALES_SHEET_ID = "1S3csAE-E_hN8vstuHR0KkeAN7yCVQTFe4AkEVlw4vQw"

# Must match app.py's COLUMN_ORDER exactly, since both apps read/write the
# same Google Sheet. Includes the second Harvest slot (Harvest Date 2 /
# Harvest Type 2 / Harvest KG 2 / Harvest ABW 2), and the "Expect Harvest
# (KG)" / "Survival QTY" auto-calculated columns, both added to app.py.
COLUMN_ORDER = [
    "Timestamp", "Customer", "Farm Name with Code", "Zone", "Area",
    "Pond Number", "Date", "Species Culture", "Cycle Type",
    "DOC", "Density", "Feed Per Day", "ABW",
    "Expect Harvest (KG)", "Survival QTY",
    "Issues", "Water Color", "Grade", "Remark", "Technician",
    "Harvest Date", "Harvest Type", "Harvest KG", "Harvest ABW",
    "Harvest Date 2", "Harvest Type 2", "Harvest KG 2", "Harvest ABW 2",
    "Deleted",
]

# Expected columns in the Sales Details Google Sheet. "Settle" holds the
# persisted checkbox state from the "Financial Status" tick below (blank
# or "Yes") — it's created automatically in the sheet the first time a row
# is ticked if it doesn't already exist there.
SALES_COLUMN_ORDER = [
    "Date", "Item No.", "Item Description", "Customer Code",
    "Customer Name", "Quantity", "Sales Amt", "Settle",
]

# =========================================================================
# STYLE (kept visually consistent with the data-entry app)
# =========================================================================
st.markdown("""
<style>
ul[role="listbox"], div[role="listbox"] {
    width: max-content !important;
    min-width: 220px !important;
    max-width: 92vw !important;
}
[role="option"] {
    width: auto !important;
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    word-break: break-word !important;
}
[role="option"] * {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
}

@media (max-width: 700px) {
    div[data-testid="stHorizontalBlock"] {
        flex-direction: column !important;
    }
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        width: 100% !important;
        min-width: 100% !important;
    }
}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center;'>Shrimp FarmFlow - KMN</h1>",
            unsafe_allow_html=True)
st.subheader("KMN Aqua Services")
st.markdown("---")

# =========================================================================
# GOOGLE SHEETS BACKEND. Read-only except for the two recycle-bin delete
# flags: get_or_create_column() finds/creates the 'Settle' header (Sales
# Details sheet) or 'Harvest Status' header (main sheet), and the write
# calls that use it sit next to the Sales Details / All Harvest Details
# tables further down.
# =========================================================================
def _gsheet_configured():
    return "gcp_service_account" in st.secrets and "gsheet" in st.secrets and "sheet_id" in st.secrets["gsheet"]

@st.cache_resource(show_spinner=False)
def get_worksheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet_id = st.secrets["gsheet"]["sheet_id"]
    worksheet_name = st.secrets["gsheet"].get("worksheet_name", "WaterQualityData")
    sh = client.open_by_key(sheet_id)
    return sh.worksheet(worksheet_name)

@st.cache_resource(show_spinner=False)
def get_sales_worksheet():
    """Separate spreadsheet (Sales Details) — same service account creds,
    different spreadsheet key. Worksheet/tab name can be overridden via
    st.secrets["gsheet"]["sales_worksheet_name"] (defaults to the first
    sheet/tab in the spreadsheet if not set)."""
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SALES_SHEET_ID)
    worksheet_name = st.secrets.get("gsheet", {}).get("sales_worksheet_name", "")
    if worksheet_name:
        return sh.worksheet(worksheet_name)
    return sh.sheet1

def get_or_create_column(ws, header_name):
    """Returns the 1-based column number of the given header in ws,
    creating that header (in the first empty column) if it isn't there
    yet. Used for 'Settle' (Sales Details sheet) and 'Harvest Status'
    (main WaterQualityData sheet). Expands the sheet's column count first if
    the new header would land past the sheet's current grid width —
    writing past the grid is what raises gspread's APIError, since the
    underlying Sheets API rejects it."""
    headers = ws.row_values(1)
    if header_name in headers:
        return headers.index(header_name) + 1
    new_col_idx = len(headers) + 1
    if new_col_idx > ws.col_count:
        ws.add_cols(new_col_idx - ws.col_count)
    ws.update_cell(1, new_col_idx, header_name)
    return new_col_idx

def load_data():
    """Always reads fresh from the Google Sheet (no caching), so the
    manager always sees the latest saved records — including anything
    just added by the data-entry app or edited directly in the Sheet.
    Soft-deleted rows (Deleted = Yes) are filtered out, same as the
    data-entry app. Rows marked Harvest Status = 'H' (via the recycle-bin
    control on the All Harvest Details table below) are filtered out the
    same way — that flag is written straight to the Sheet, so the removal
    persists across refreshes instead of resetting."""
    ws = get_worksheet()
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for c in COLUMN_ORDER:
        if c not in df.columns:
            df[c] = ""
    if "Harvest Status" not in df.columns:
        df["Harvest Status"] = ""
    # "Harvest Submitted Date" already exists as its own column in the
    # Sheet (outside COLUMN_ORDER, same as "Harvest Status") — kept here
    # so the All Harvest Details table below can offer it as a Sort by
    # option instead of it getting dropped like any other unlisted column.
    if "Harvest Submitted Date" not in df.columns:
        df["Harvest Submitted Date"] = ""
    if len(df) > 0:
        df = df[COLUMN_ORDER + ["Harvest Status", "Harvest Submitted Date"]]
    df = df.astype(str).replace("nan", "")
    if "Deleted" in df.columns:
        is_deleted = df["Deleted"].astype(str).str.strip().str.lower().isin(["yes", "true", "1"])
        df = df[~is_deleted].reset_index(drop=True)
    if "Harvest Status" in df.columns:
        is_harvest_hidden = df["Harvest Status"].astype(str).str.strip().str.upper() == "H"
        df = df[~is_harvest_hidden].reset_index(drop=True)
    return df

def load_sales_data():
    """Always reads fresh from the Sales Details Google Sheet (no caching),
    same pattern as load_data() above. Also attaches each row's real sheet
    row number (_RowNumber) so a tick in the UI can be written back to the
    correct cell in the 'Settle' column."""
    ws = get_sales_worksheet()
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for c in SALES_COLUMN_ORDER:
        if c not in df.columns:
            df[c] = ""
    df["_RowNumber"] = range(2, len(df) + 2)
    if len(df) > 0:
        df = df[["_RowNumber"] + SALES_COLUMN_ORDER]
    return df

if not _gsheet_configured():
    st.error("❌ Google Sheets is not configured yet. This app needs the same "
             "`.streamlit/secrets.toml` (the `[gcp_service_account]` and `[gsheet]` sections) "
             "used by the data-entry app.")
    st.stop()

try:
    get_worksheet()
except Exception as e:
    st.error(f"❌ Could not connect to the Google Sheet. Check your secrets and sharing settings.\n\n{e}")
    st.stop()

# =========================================================================
# LOAD CUSTOMER LIST (for the Customer / Farm selectors)
# =========================================================================
@st.cache_data
def load_customer_data():
    return pd.read_excel(CUSTOMER_FILE)

try:
    customer_df = load_customer_data()
except Exception as e:
    st.error(f"❌ Could not load '{CUSTOMER_FILE}'. Make sure it's in the app folder. ({e})")
    st.stop()

REQUIRED_COLS = ["Customer Name", "Farm Name with Code", "Zone", "Area"]
missing_cols = [c for c in REQUIRED_COLS if c not in customer_df.columns]
if missing_cols:
    st.error(f"❌ 'Customer List.xlsx' is missing required column(s): {', '.join(missing_cols)}")
    st.stop()

for _col in REQUIRED_COLS:
    customer_df[_col] = customer_df[_col].apply(
        lambda v: "" if pd.isna(v) else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))
    )

all_customers = sorted(customer_df["Customer Name"].replace("", pd.NA).dropna().unique().tolist())

# =========================================================================
# STEP 1: "Enter Water Quality Data" — Customer / Farm selection.
# For the manager this is just a filter to choose whose records to view —
# there is no data entry after this, only the read-only table below.
# =========================================================================
st.subheader("📋 Enter Customer Details")

col1, col2 = st.columns(2)
with col1:
    customer = st.selectbox("Customer Name *", all_customers, key="customer_select")

farm_options = sorted(
    customer_df.loc[customer_df["Customer Name"] == customer, "Farm Name with Code"]
    .dropna().unique().tolist()
)
if not farm_options:
    farm_options = ["-- No farms found for this customer --"]

with col2:
    farm = st.selectbox("Farm Name with Code *", farm_options, key=f"farm_select_{customer}")

farm_row_match = customer_df[
    (customer_df["Customer Name"] == customer) & (customer_df["Farm Name with Code"] == farm)
]
if len(farm_row_match) > 0 and "Marketing Manager" in customer_df.columns:
    mm = farm_row_match.iloc[0].get("Marketing Manager", "")
    if str(mm).strip():
        st.caption(f"Marketing Manager: {mm}")

# The Customer Code / Customer ID for the selected farm (used below to
# filter the Sales Details sheet). "Customer List.xlsx" may store this
# under any of a few likely column names — the first one that exists and
# has a non-blank value for this row wins.
_CUSTOMER_CODE_COLUMN_CANDIDATES = [
    "Customer Code", "Customer ID", "Customer Code with Code", "Code", "Cust Code",
]
selected_customer_code = ""
if len(farm_row_match) > 0:
    for _cand in _CUSTOMER_CODE_COLUMN_CANDIDATES:
        if _cand in customer_df.columns:
            _val = str(farm_row_match.iloc[0].get(_cand, "")).strip()
            if _val and _val.lower() != "nan":
                selected_customer_code = _val
                break

# =========================================================================
# ALL SAVED RECORDS — read-only, straight from the Google Sheet, for the
# selected Customer + Farm across every pond.
# =========================================================================
st.markdown("---")
st.markdown(f"#### 📊 All Saved Records — {farm}")

if st.button("🔄 Refresh"):
    st.rerun()

# Hidden by default — tick this to show the full "All Saved Records"
# table below. The totals / Pond Layout sections further down are
# unaffected and still show either way.
show_all_saved_records = st.checkbox(
    "Show All Saved Records table", value=False, key="show_all_saved_records"
)

df_farm_summary = load_data()
_farm_required = {"Customer", "Farm Name with Code"}
if len(df_farm_summary) > 0 and _farm_required.issubset(df_farm_summary.columns):
    df_farm_summary = df_farm_summary[
        (df_farm_summary["Customer"] == customer) & (df_farm_summary["Farm Name with Code"] == farm)
    ].copy()
else:
    df_farm_summary = pd.DataFrame(columns=COLUMN_ORDER)

# Initialized here (before the branch below) so both are always defined —
# even when this farm has no saved records yet — since the Sales Details
# section further down the page reads total_density_by_species to compute
# the NANAMI feed "Do not exceed" limits.
total_expect_harvest_kg = None
total_density_by_species = {}

if len(df_farm_summary) > 0:
    if "Date" in df_farm_summary.columns:
        df_farm_summary["_ParsedDate"] = pd.to_datetime(df_farm_summary["Date"], errors="coerce")

        # Latest saved record per pond — shared basis for both totals below.
        _latest_per_pond = None
        if "Pond Number" in df_farm_summary.columns:
            _latest_per_pond = (
                df_farm_summary.dropna(subset=["_ParsedDate"])
                .sort_values("_ParsedDate")
                .groupby("Pond Number", as_index=False)
                .last()
            )

        # Total Expect Harvest (KG) for the farm = each pond's MOST RECENT
        # saved record's "Expect Harvest (KG)" value, summed across every
        # pond on this farm (not every historical row, which would double
        # count a pond's earlier daily estimates). Same logic as app.py.
        if _latest_per_pond is not None and "Expect Harvest (KG)" in df_farm_summary.columns:
            _harvest_vals = pd.to_numeric(_latest_per_pond["Expect Harvest (KG)"], errors="coerce").dropna()
            if len(_harvest_vals) > 0:
                total_expect_harvest_kg = float(_harvest_vals.sum())

        # Total Density for the farm = each pond's MOST RECENT saved
        # record's "Density" value, summed across every pond on this farm
        # EXCEPT ponds whose latest record shows a Full Harvest (checking
        # Harvest Type 2 first, then Harvest Type — same "2nd slot wins"
        # rule used by the Pond Layout section below). Split out per
        # Species Culture (e.g. Vannamei ponds get their own total,
        # separate from Monodon ponds) instead of one combined number,
        # since a farm can be running more than one species at once.
        if _latest_per_pond is not None and "Density" in df_farm_summary.columns:
            def _is_pond_full_h_for_density(prow):
                _t = str(prow.get("Harvest Type 2", "")).strip() or str(prow.get("Harvest Type", "")).strip()
                return "full" in _t.lower()

            _not_full_h_mask = ~_latest_per_pond.apply(_is_pond_full_h_for_density, axis=1)
            _density_pool = _latest_per_pond.loc[_not_full_h_mask].copy()
            _density_pool["Density"] = pd.to_numeric(_density_pool["Density"], errors="coerce")
            _density_pool = _density_pool.dropna(subset=["Density"])
            if len(_density_pool) > 0:
                if "Species Culture" in _density_pool.columns:
                    _density_species_label = (
                        _density_pool["Species Culture"].astype(str).str.strip().replace("", "Unspecified")
                    )
                else:
                    _density_species_label = pd.Series("Unspecified", index=_density_pool.index)
                total_density_by_species = (
                    _density_pool.groupby(_density_species_label)["Density"].sum().to_dict()
                )

        sort_cols = [c for c in ["Pond Number"] if c in df_farm_summary.columns] + ["_ParsedDate"]
        df_farm_summary = df_farm_summary.sort_values(by=sort_cols).drop(columns=["_ParsedDate"])

    # "DOC Today" = this row's saved DOC + however many days have passed
    # between its Date and today (i.e. what the DOC would be right now).
    # It's a live, always-changing number rather than something actually
    # saved in the Sheet, so it's shown in red/bold to stand out. Rows
    # whose Cycle Type is "Soon to be" haven't actually started yet, so
    # DOC Today just stays 0 for them instead of counting elapsed days.
    # A row whose Harvest Type (checking the more recent Harvest Type 2
    # first, then Harvest Type) says "Full" instead STOPS ADVANCING at
    # that row's Full Harvest date — the pond was fully harvested there,
    # so DOC Today shouldn't keep counting up to today. Same rule used
    # for "DOC" in the All Harvest Details table below.
    def _compute_doc_today(row):
        if str(row.get("Cycle Type") or "").strip() == "Soon to be":
            return "0"
        parsed = pd.to_datetime(row.get("Date"), errors="coerce")
        if pd.isna(parsed):
            return ""
        try:
            doc_num = int(float(row.get("DOC")))
        except (TypeError, ValueError):
            return ""
        _t2 = str(row.get("Harvest Type 2", "")).strip().lower()
        _t1 = str(row.get("Harvest Type", "")).strip().lower()
        _full_harvest_date_str = ""
        if "full" in _t2:
            _full_harvest_date_str = str(row.get("Harvest Date 2", "")).strip()
        elif "full" in _t1:
            _full_harvest_date_str = str(row.get("Harvest Date", "")).strip()
        if _full_harvest_date_str:
            _full_harvest_date = pd.to_datetime(_full_harvest_date_str, errors="coerce")
            if pd.notna(_full_harvest_date):
                return str(doc_num + (_full_harvest_date - parsed).days)
        days_passed = (pd.Timestamp(date.today()) - parsed).days
        return str(doc_num + days_passed)

    df_farm_summary["DOC Today"] = df_farm_summary.apply(_compute_doc_today, axis=1)

    _farm_display_cols = ["Pond Number", "Date", "Species Culture", "Cycle Type", "DOC", "DOC Today", "Density",
                           "Feed Per Day", "ABW", "Expect Harvest (KG)", "Survival QTY",
                           "Issues", "Water Color", "Grade", "Remark", "Technician",
                           "Harvest Date", "Harvest Type", "Harvest KG", "Harvest ABW",
                           "Harvest Date 2", "Harvest Type 2", "Harvest KG 2", "Harvest ABW 2"]
    _farm_display_cols = [c for c in _farm_display_cols if c in df_farm_summary.columns]

    # st.dataframe has no way to color/bold an individual column's text, so
    # this one table is rendered as a plain HTML table instead ("All Harvest
    # Details" below keeps using st.dataframe as before).
    def _escape_html(v):
        return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def _render_highlighted_table(df, cols, highlight_col, highlight_style="color:red;font-weight:bold;"):
        header_html = "".join(
            f"<th style='padding:6px 10px;border-bottom:2px solid #ccc;text-align:left;white-space:nowrap;'>{_escape_html(c)}</th>"
            for c in cols
        )
        rows_html = ""
        for _, r in df.iterrows():
            cells_html = ""
            for c in cols:
                cell_style = "padding:6px 10px;border-bottom:1px solid #eee;white-space:nowrap;"
                if c == highlight_col:
                    cell_style += highlight_style
                cells_html += f"<td style='{cell_style}'>{_escape_html(r.get(c, ''))}</td>"
            rows_html += f"<tr>{cells_html}</tr>"
        return (
            "<div style='overflow-x:auto; width:100%;'>"
            "<table style='width:100%; border-collapse:collapse; font-size:0.9rem;'>"
            f"<thead><tr>{header_html}</tr></thead>"
            f"<tbody>{rows_html}</tbody>"
            "</table></div>"
        )

    if show_all_saved_records:
        st.markdown(
            _render_highlighted_table(df_farm_summary, _farm_display_cols, "DOC Today"),
            unsafe_allow_html=True,
        )
        st.caption(f"{len(df_farm_summary)} saved record(s) across all ponds for {farm}.")

    if total_density_by_species:
        for _density_species, _density_val in total_density_by_species.items():
            st.markdown(
                f"**🧮 {_density_species} Ponds Total Density — {farm}: {_density_val:,.2f}**"
            )
        st.caption(
            "(sum of each pond's latest Density, excluding ponds at Full Harvest, split by Species Culture)"
        )

    if total_expect_harvest_kg is not None:
        st.markdown(
            f"**🌾 Total Expect Harvest (KG) — {farm}: {total_expect_harvest_kg:,.2f} kg** "
            "(sum of each pond's latest Expect Harvest (KG) estimate)"
        )

    # =========================================================================
    # POND LAYOUT — one rectangle per Pond Number (using that pond's most
    # recent saved record). Running / Partial H ponds show DOC Today
    # centered inside the box; Full H ponds show "Full H" and its Harvest
    # Date instead.
    # =========================================================================
    if "Pond Number" in df_farm_summary.columns and "DOC Today" in df_farm_summary.columns:
    st.markdown("---")
    st.markdown(f"#### 🟦 Pond Layout — {farm}")

        _pond_latest = (
            df_farm_summary.assign(_PondSortDate=pd.to_datetime(df_farm_summary["Date"], errors="coerce"))
            .dropna(subset=["_PondSortDate"])
            .sort_values("_PondSortDate")
            .groupby("Pond Number", as_index=False)
            .last()
            .sort_values("Pond Number")
        )

        def _escape_html_pond(v):
            return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        def _pond_harvest_type(prow):
            return str(prow.get("Harvest Type 2", "")).strip() or str(prow.get("Harvest Type", "")).strip()

        # A pond keeps showing Partial H if ANY of its saved records ever
        # had a Partial harvest — not just its most recent row (daily
        # entries often leave Harvest Type blank after the harvest day).
        # Full H intentionally does NOT carry forward this way — it only
        # reflects the pond's latest record, same as before.
        _partial_history_by_pond = (
            df_farm_summary.assign(
                _HasPartial=(
                    df_farm_summary.get("Harvest Type", pd.Series("", index=df_farm_summary.index))
                    .astype(str).str.lower().str.contains("partial")
                    | df_farm_summary.get("Harvest Type 2", pd.Series("", index=df_farm_summary.index))
                    .astype(str).str.lower().str.contains("partial")
                )
            )
            .groupby("Pond Number")["_HasPartial"]
            .any()
        )

        def _pond_status(prow):
            _pond_no_status = prow.get("Pond Number", "")
            _h_type_lower = _pond_harvest_type(prow).lower()
            _has_partial_history = bool(_partial_history_by_pond.get(_pond_no_status, False))
            if "full" in _h_type_lower:
                return "Full H"
            elif "partial" in _h_type_lower or _has_partial_history:
                return "Partial H"
            elif str(prow.get("Cycle Type", "")).strip() == "Soon to be":
                return "Soon to be"
            else:
                return "Running"

        def _pond_box_color(prow):
            _status = _pond_status(prow)
            if _status == "Partial H":
                return "#fff3cd"  # yellow — Partial Harvest
            elif _status == "Full H":
                return "#d4edda"  # green — Full Harvest
            elif _status == "Soon to be":
                return "#e2e2e2"  # gray — cycle hasn't started yet
            else:
                return "#eaf4ff"  # default blue — no harvest yet

        def _species_letter(prow):
            _species = str(prow.get("Species Culture", "")).strip().lower()
            if "vannamei" in _species:
                return "V"
            elif "monodon" in _species:
                return "M"
            else:
                return ""

        _pond_boxes_html = ""
        for _, _prow in _pond_latest.iterrows():
            _pond_no = _escape_html_pond(_prow.get("Pond Number", ""))
            _box_color = _pond_box_color(_prow)
            _status_box = _pond_status(_prow)

            if _status_box == "Full H":
                # Full H ponds: show "Full H" + its Harvest Date instead of DOC Today.
                _h_date = str(_prow.get("Harvest Date 2", "")).strip() or str(_prow.get("Harvest Date", "")).strip()
                _h_date = _escape_html_pond(_h_date or "-")
                _box_middle_html = (
                    "<div style='font-size:1.2rem;font-weight:bold;color:red;'>Full H</div>"
                    f"<div style='font-size:0.75rem;color:#333;'>{_h_date}</div>"
                )
            elif _status_box == "Soon to be":
                # Soon to be ponds: show "Soon to be" instead of DOC Today,
                # matching the gray box color already assigned by
                # _pond_box_color() for this status.
                _box_middle_html = (
                    "<div style='font-size:1.1rem;font-weight:bold;color:#555;'>Soon to be</div>"
                )
            else:
                # Running / Partial H ponds: keep showing the DOC Today
                # number, but the label under it now shows the pond's
                # Started Date (today's date minus DOC Today days) instead
                # of the literal "DOC Today" text.
                _doc_today_raw = _prow.get("DOC Today", "")
                _doc_today_val = _escape_html_pond(_doc_today_raw or "-")
                try:
                    _started_date = (
                        pd.Timestamp(date.today()) - pd.Timedelta(days=int(float(_doc_today_raw)))
                    ).strftime("%Y-%m-%d")
                    _started_label = f"Started on {_started_date}"
                except (TypeError, ValueError):
                    _started_label = "Started on ---"
                _box_middle_html = (
                    f"<div style='font-size:1.4rem;font-weight:bold;color:red;'>{_doc_today_val}</div>"
                    f"<div style='font-size:0.7rem;color:#777;'>{_escape_html_pond(_started_label)}</div>"
                )

            _species_label = _species_letter(_prow)
            _species_html = (
                f"<div style='font-size:0.75rem;font-weight:bold;color:#444;margin-top:2px;'>{_species_label}</div>"
                if _species_label else ""
            )

            _pond_boxes_html += (
                "<div style='display:flex;flex-direction:column;align-items:center;margin:6px;'>"
                f"<div style='width:140px;height:90px;border:2px solid #333;border-radius:6px;"
                "display:flex;flex-direction:column;align-items:center;justify-content:center;"
                f"background:{_box_color};'>"
                f"<div style='font-size:0.8rem;color:#555;'>Pond {_pond_no}</div>"
                f"{_box_middle_html}"
                "</div>"
                f"{_species_html}"
                "</div>"
            )

        st.markdown(
            "<div style='display:flex;gap:18px;justify-content:center;margin-bottom:8px;font-size:0.85rem;'>"
            "<div><span style='display:inline-block;width:14px;height:14px;background:#eaf4ff;"
            "border:1px solid #333;border-radius:3px;vertical-align:middle;margin-right:6px;'></span>Running</div>"
            "<div><span style='display:inline-block;width:14px;height:14px;background:#fff3cd;"
            "border:1px solid #333;border-radius:3px;vertical-align:middle;margin-right:6px;'></span>Partial H</div>"
            "<div><span style='display:inline-block;width:14px;height:14px;background:#d4edda;"
            "border:1px solid #333;border-radius:3px;vertical-align:middle;margin-right:6px;'></span>Full H</div>"
            "</div>",
            unsafe_allow_html=True,
        )

        if "Pond Number" in df_farm_summary.columns and "DOC Today" in df_farm_summary.columns:
        st.markdown("---")
        st.markdown(f"#### 🟦 Pond Layout — {farm}")

        _pond_latest = (
            df_farm_summary.assign(_PondSortDate=pd.to_datetime(df_farm_summary["Date"], errors="coerce"))
            .dropna(subset=["_PondSortDate"])
            .sort_values("_PondSortDate")
            .groupby("Pond Number", as_index=False)
            .last()
            .sort_values("Pond Number")
        )

        def _escape_html_pond(v):
            return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        def _pond_harvest_type(prow):
            return str(prow.get("Harvest Type 2", "")).strip() or str(prow.get("Harvest Type", "")).strip()

        _partial_history_by_pond = (
            df_farm_summary.assign(
                _HasPartial=(
                    df_farm_summary.get("Harvest Type", pd.Series("", index=df_farm_summary.index))
                    .astype(str).str.lower().str.contains("partial")
                    | df_farm_summary.get("Harvest Type 2", pd.Series("", index=df_farm_summary.index))
                    .astype(str).str.lower().str.contains("partial")
                )
            )
            .groupby("Pond Number")["_HasPartial"]
            .any()
        )

        def _pond_status(prow):
            _pond_no_status = prow.get("Pond Number", "")
            _h_type_lower = _pond_harvest_type(prow).lower()
            _has_partial_history = bool(_partial_history_by_pond.get(_pond_no_status, False))
            if "full" in _h_type_lower:
                return "Full H"
            elif "partial" in _h_type_lower or _has_partial_history:
                return "Partial H"
            elif str(prow.get("Cycle Type", "")).strip() == "Soon to be":
                return "Soon to be"
            else:
                return "Running"

        def _pond_box_color(prow):
            _status = _pond_status(prow)
            if _status == "Partial H":
                return "#fff3cd"  # yellow — Partial Harvest
            elif _status == "Full H":
                return "#d4edda"  # green — Full Harvest
            elif _status == "Soon to be":
                return "#e2e2e2"  # gray — cycle hasn't started yet
            else:
                return "#eaf4ff"  # default blue — no harvest yet

        def _species_letter(prow):
            _species = str(prow.get("Species Culture", "")).strip().lower()
            if "vannamei" in _species:
                return "V"
            elif "monodon" in _species:
                return "M"
            else:
                return ""

        # Legend (unchanged)
        st.markdown(
            "<div style='display:flex;gap:18px;justify-content:center;margin-bottom:8px;font-size:0.85rem;'>"
            "<div><span style='display:inline-block;width:14px;height:14px;background:#eaf4ff;"
            "border:1px solid #333;border-radius:3px;vertical-align:middle;margin-right:6px;'></span>Running</div>"
            "<div><span style='display:inline-block;width:14px;height:14px;background:#fff3cd;"
            "border:1px solid #333;border-radius:3px;vertical-align:middle;margin-right:6px;'></span>Partial H</div>"
            "<div><span style='display:inline-block;width:14px;height:14px;background:#d4edda;"
            "border:1px solid #333;border-radius:3px;vertical-align:middle;margin-right:6px;'></span>Full H</div>"
            "</div>",
            unsafe_allow_html=True,
        )

        # ---- Pond boxes are now clickable buttons -----------------------
        # Clicking a pond stores its latest saved row in session_state and
        # a details panel is rendered right below the grid, showing:
        # Stocking Density, Last Visited Date, Feed Per Day, ABW,
        # Expect Harvest (KG), Grade, Issues, Remark.
        _pond_detail_state_key = f"selected_pond_detail_{farm}"
        _pond_css_rules = []
        _cols_per_row = 6
        _pond_rows_list = [
            _pond_latest.iloc[i:i + _cols_per_row] for i in range(0, len(_pond_latest), _cols_per_row)
        ]

        for _row_df in _pond_rows_list:
            _pond_cols = st.columns(len(_row_df))
            for _col, (_, _prow) in zip(_pond_cols, _row_df.iterrows()):
                _pond_no = _escape_html_pond(_prow.get("Pond Number", ""))
                _box_color = _pond_box_color(_prow)
                _status_box = _pond_status(_prow)
                _species_label = _species_letter(_prow)

                if _status_box == "Full H":
                    _h_date = str(_prow.get("Harvest Date 2", "")).strip() or str(_prow.get("Harvest Date", "")).strip()
                    _line2 = "Full H"
                    _line3 = _h_date or "-"
                elif _status_box == "Soon to be":
                    _line2 = "Soon to be"
                    _line3 = ""
                else:
                    _doc_today_raw = _prow.get("DOC Today", "")
                    try:
                        _started_date = (
                            pd.Timestamp(date.today()) - pd.Timedelta(days=int(float(_doc_today_raw)))
                        ).strftime("%Y-%m-%d")
                        _line3 = f"Started {_started_date}"
                    except (TypeError, ValueError):
                        _line3 = "Started ---"
                    _line2 = f"DOC {_doc_today_raw or '-'}"

                _label_parts = [f"Pond {_pond_no}", _line2]
                if _line3:
                    _label_parts.append(_line3)
                if _species_label:
                    _label_parts.append(f"({_species_label})")
                _btn_label = "\n".join(_label_parts)

                _safe_key_part = re.sub(r"[^A-Za-z0-9_]", "_", f"{farm}_{_pond_no}")
                _box_key = f"pondbox_{_safe_key_part}"
                _pond_css_rules.append(
                    f".st-key-{_box_key} button {{background-color:{_box_color} !important;"
                    "border:2px solid #333 !important;white-space:pre-line !important;"
                    "min-height:90px !important;color:#111 !important;font-weight:600 !important;}}"
                )

                with _col:
                    with st.container(key=_box_key):
                        if st.button(_btn_label, key=f"pondbtn_{_safe_key_part}", use_container_width=True):
                            st.session_state[_pond_detail_state_key] = _prow.to_dict()

        # Colors the buttons per pond status (requires Streamlit >= 1.34
        # for st.container(key=...) support). If your Streamlit version is
        # older, the buttons still work — they just show default styling.
        if _pond_css_rules:
            st.markdown(f"<style>{''.join(_pond_css_rules)}</style>", unsafe_allow_html=True)

        # ---- Details panel for the last-clicked pond ---------------------
        _selected_pond_data = st.session_state.get(_pond_detail_state_key)
        if _selected_pond_data:
            st.markdown(f"##### 🔍 Pond {_selected_pond_data.get('Pond Number', '')} Details")
            _detail_fields = [
                ("Stocking Density", _selected_pond_data.get("Density", "")),
                ("Last Visited Date", _selected_pond_data.get("Date", "")),
                ("Feed Per Day", _selected_pond_data.get("Feed Per Day", "")),
                ("ABW", _selected_pond_data.get("ABW", "")),
                ("Expect Harvest (KG)", _selected_pond_data.get("Expect Harvest (KG)", "")),
                ("Grade", _selected_pond_data.get("Grade", "")),
                ("Issues", _selected_pond_data.get("Issues", "")),
                ("Remark", _selected_pond_data.get("Remark", "")),
            ]
            _detail_cols = st.columns(4)
            for _idx, (_flabel, _fval) in enumerate(_detail_fields):
                with _detail_cols[_idx % 4]:
                    st.metric(_flabel, str(_fval).strip() if str(_fval).strip() else "-")
else:
    st.info(f"No saved records yet for {farm}.")

# =========================================================================
# SALES DETAILS — read-only from the Google Sheet's perspective (nothing
# here ever writes back to the Sales Details spreadsheet), filtered to the
# Customer Code belonging to the Customer + Farm selected above. Shown as
# individual sales line items with a per-row "Delete" checkbox — checking
# it just hides that row from THIS view (and from the totals below) for
# the current session; it never touches the Google Sheet, so a refresh or
# a new session brings the row right back.
# =========================================================================
st.markdown("---")
st.markdown(f"#### 🧾 Sales Details — {farm}")

try:
    df_sales = load_sales_data()
except Exception as e:
    df_sales = None
    st.error(f"❌ Could not connect to the Sales Details Google Sheet. Check sharing settings.\n\n{e}")

if df_sales is not None:
    if not selected_customer_code:
        st.info(
            "No Customer Code found for this farm in 'Customer List.xlsx', so Sales Details "
            "can't be filtered. Add a 'Customer Code' column to the customer list to enable this."
        )
    elif len(df_sales) == 0:
        st.info("No sales records found in the Sales Details sheet.")
    else:
        df_sales_farm = df_sales[
            df_sales["Customer Code"].astype(str).str.strip().str.lower()
            == selected_customer_code.strip().lower()
        ].copy()

        if len(df_sales_farm) == 0:
            st.info(f"No sales records found for Customer Code '{selected_customer_code}'.")
        else:
            df_sales_farm["Quantity"] = pd.to_numeric(df_sales_farm["Quantity"], errors="coerce").fillna(0)
            df_sales_farm["Sales Amt"] = pd.to_numeric(df_sales_farm["Sales Amt"], errors="coerce").fillna(0)
            df_sales_farm["Settle"] = df_sales_farm["Settle"].astype(str)

            # Same pivot as before: one row per Date, one column per Item
            # Description, cell value = summed Quantity for that date/item.
            pivot_sales = df_sales_farm.pivot_table(
                index="Date",
                columns="Item Description",
                values="Quantity",
                aggfunc="sum",
                fill_value=0,
            )
            pivot_sales = pivot_sales.sort_index()
            pivot_sales.index.name = "Date"

            # Dates already marked Settle = 'Yes' in the Sheet are treated
            # as permanently removed — they're excluded before the table
            # is even built, so a refresh doesn't bring them back.
            _settled_by_date = (
                df_sales_farm.groupby("Date")["Settle"]
                .apply(lambda s: s.str.strip().str.lower().eq("yes").all())
            )

            pivot_display = pivot_sales.reset_index()
            pivot_display = pivot_display[
                ~pivot_display["Date"].map(_settled_by_date).fillna(False)
            ].reset_index(drop=True)
            sales_editor_key = f"sales_delete_editor_{selected_customer_code}"

            st.caption(
                "🗑️ Select a row's checkbox (left edge) then click the recycle-bin icon "
                "above the table to remove that date — this writes 'Yes' to the Settle "
                "column in the Google Sheet, so it stays removed after a refresh."
            )
            edited_pivot = st.data_editor(
                pivot_display,
                use_container_width=True,
                hide_index=True,
                key=sales_editor_key,
                num_rows="dynamic",
                disabled=list(pivot_display.columns),
            )

            # A date missing from edited_pivot was just removed via the
            # recycle bin — mark every one of that date's line items as
            # Settle = 'Yes' in the Sheet so the removal persists.
            removed_dates = set(pivot_display["Date"]) - set(edited_pivot["Date"].dropna())
            if removed_dates:
                sales_ws = get_sales_worksheet()
                settle_col_idx = get_or_create_column(sales_ws, "Settle")
                for _removed_date in removed_dates:
                    _rows_for_date = df_sales_farm[df_sales_farm["Date"] == _removed_date]
                    for _, _r in _rows_for_date.iterrows():
                        _current_settle = str(_r.get("Settle", "")).strip().lower()
                        if _current_settle != "yes":
                            sales_ws.update_cell(int(_r["_RowNumber"]), settle_col_idx, "Yes")
                st.rerun()

            kept_dates = edited_pivot["Date"].dropna()
            _num_hidden = len(pivot_display) - len(kept_dates)

            df_sales_farm_visible = df_sales_farm[df_sales_farm["Date"].isin(kept_dates)]

            _caption = (
                f"{len(df_sales_farm_visible)} sales line item(s) across "
                f"{len(kept_dates)} date(s) for Customer Code '{selected_customer_code}'."
            )
            if _num_hidden:
                _caption += f" ({_num_hidden} row(s) hidden in this view.)"
            st.caption(_caption)

            # Total Quantity / Total Sales Amt reflect FEED items only —
            # rows whose "Item No." starts with "FEED" (same prefix check
            # used by the Running List's feed-purchase rollup below), so
            # non-feed line items no longer inflate these two totals.
            _sales_totals_feed_mask = (
                df_sales_farm_visible["Item No."].astype(str).str.strip().str.upper().str.startswith("FEED")
            )
            df_sales_farm_visible_feed_only = df_sales_farm_visible[_sales_totals_feed_mask]

            total_qty = df_sales_farm_visible_feed_only["Quantity"].sum() if len(df_sales_farm_visible_feed_only) else 0
            total_amt = df_sales_farm_visible_feed_only["Sales Amt"].sum() if len(df_sales_farm_visible_feed_only) else 0
            st.markdown(
                f"**Total Quantity: {total_qty:,.0f}  |  Total Sales Amt: {total_amt:,.2f}**"
            )

            # =================================================================
            # FEED ORDER STATUS — two rows of rectangles, one per feed size,
            # in a fixed order: NANAMI FEED sizes, then EGO FEED sizes. Each
            # box shows that size's total purchased Quantity in the middle
            # and its latest purchase Date below. Boxes are colored one way
            # if any quantity has been purchased, another if not yet.
            #
            # NANAMI_LIMIT_FACTORS / EGO_LIMIT_FACTORS hold the "Do not
            # exceed" formula for each brand's sizes. Each size's limit =
            # factor * that farm's Total Density for the species the brand
            # is fed to — NANAMI uses Vannamei Ponds Total Density, EGO
            # uses Monodon Ponds Total Density (both shown above in
            # "🧮 ... Ponds Total Density").
            # =================================================================
            NANAMI_FEED_ORDER = [
                "NANAMI 1", "NANAMI 1S", "NANAMI 2S", "NANAMI 3S",
                "NANAMI 3M", "NANAMI 3L", "NANAMI 4",
            ]
            EGO_FEED_ORDER = [
                "EGO - 01", "EGO - 01S", "EGO - 02S", "EGO - 03S",
                "EGO - 03M", "EGO - 03L", "EGO - 04L",
            ]
            NANAMI_LIMIT_FACTORS = {
                "NANAMI 1": 50 / 100000,
                "NANAMI 1S": 150 / 100000,
                "NANAMI 2S": 150 / 100000,
                "NANAMI 3S": 150 / 100000,
                "NANAMI 3M": 750 / 100000,
                "NANAMI 3L": 1000 / 100000
            }
            EGO_LIMIT_FACTORS = {
                "EGO - 01": 50 / 100000,
                "EGO - 01S": 150 / 100000,
                "EGO - 02S": 150 / 100000,
                "EGO - 03S": 200 / 100000,
                "EGO - 03M": 500 / 100000,
                "EGO - 03L": 750 / 100000, "EGO - 04L": 1000 / 100000
            }

            # The Vannamei / Monodon entries from total_density_by_species
            # (computed in the "All Saved Records" section above, for this
            # same Customer + Farm). Matched case-insensitively since the
            # exact text comes from whatever the Species Culture column
            # contains (e.g. "Vannamei", "L. vannamei", "Monodon", etc.).
            _vannamei_total_density = 0.0
            _monodon_total_density = 0.0
            for _density_species_key, _density_species_val in total_density_by_species.items():
                _density_key_lower = str(_density_species_key).lower()
                if "vannamei" in _density_key_lower:
                    _vannamei_total_density = _density_species_val
                elif "monodon" in _density_key_lower:
                    _monodon_total_density = _density_species_val

            def _escape_html_feed(v):
                return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

            def _render_feed_row(title, size_labels, limit_factors=None, limit_density=0.0):
                _desc_upper = df_sales_farm_visible["Item Description"].astype(str).str.strip().str.upper()
                boxes_html = ""
                for _label in size_labels:
                    _subset = df_sales_farm_visible[_desc_upper == _label.upper()]
                    _qty = _subset["Quantity"].sum() if len(_subset) else 0
                    _latest_date = _subset["Date"].max() if len(_subset) else "-"
                    _has_qty = _qty > 0

                    # "Do not exceed" limit for this size, if it has one —
                    # factor * limit_density (Vannamei Ponds Total Density
                    # for NANAMI, Monodon Ponds Total Density for EGO).
                    _limit_val = None
                    if limit_factors and _label in limit_factors:
                        _limit_val = limit_factors[_label] * limit_density

                    if _limit_val is not None and _qty > _limit_val:
                        _box_color = "#ff4d4d"  # red — exceeded its "Do not exceed" limit
                    elif _has_qty:
                        _box_color = "#d4edda"  # green if purchased, gray if not yet
                    else:
                        _box_color = "#e2e2e2"

                    _qty_label = f"{_qty:,.0f}" if _has_qty else "-"

                    # Limit caption sits BELOW the box, outside its border —
                    # only shown for sizes that have a limit_factors entry.
                    _limit_caption_html = ""
                    if _limit_val is not None:
                        _limit_caption_html = (
                            "<div style='font-size:0.65rem;color:#555;text-align:center;"
                            f"margin-top:3px;max-width:120px;'>Do not exceed: {_limit_val:,.2f}</div>"
                        )

                    boxes_html += (
                        "<div style='display:flex;flex-direction:column;align-items:center;margin:5px;'>"
                        "<div style='width:120px;height:80px;border:2px solid #333;border-radius:6px;"
                        "display:flex;flex-direction:column;align-items:center;justify-content:center;"
                        f"background:{_box_color};'>"
                        f"<div style='font-size:0.7rem;color:#555;'>{_escape_html_feed(_label)}</div>"
                        f"<div style='font-size:1.2rem;font-weight:bold;color:#111;'>{_qty_label}</div>"
                        f"<div style='font-size:0.65rem;color:#777;'>{_escape_html_feed(_latest_date)}</div>"
                        "</div>"
                        f"{_limit_caption_html}"
                        "</div>"
                    )
                st.markdown(f"**{title}**")
                st.markdown(
                    f"<div style='display:flex;flex-wrap:wrap;justify-content:center;'>{boxes_html}</div>",
                    unsafe_allow_html=True,
                )

            st.markdown("---")
            _render_feed_row("🐟 NANAMI FEED", NANAMI_FEED_ORDER, limit_factors=NANAMI_LIMIT_FACTORS,
                              limit_density=_vannamei_total_density)
            st.markdown("")
            _render_feed_row("🐟 EGO FEED", EGO_FEED_ORDER, limit_factors=EGO_LIMIT_FACTORS,
                              limit_density=_monodon_total_density)

# =========================================================================
# ALL HARVEST DETAILS — every row (across ALL customers/farms/ponds in the
# Google Sheet, not just the one selected above) that has a non-blank
# value in EITHER harvest slot: Harvest Date/Type (the first harvest) or
# Harvest Date 2/Type 2 (a second harvest for the same pond row). Read-only,
# straight from the Sheet.
# =========================================================================
st.markdown("---")
st.markdown("#### 🌾 All Harvest Details")

df_all_records = load_data()
_harvest_cols_needed = {"Harvest Date", "Harvest Type", "Harvest Date 2", "Harvest Type 2"}
if len(df_all_records) > 0 and _harvest_cols_needed.issubset(df_all_records.columns):
    _harvest_mask = (
        (df_all_records["Harvest Date"].astype(str).str.strip() != "")
        | (df_all_records["Harvest Type"].astype(str).str.strip() != "")
        | (df_all_records["Harvest Date 2"].astype(str).str.strip() != "")
        | (df_all_records["Harvest Type 2"].astype(str).str.strip() != "")
    )
    df_harvest_all = df_all_records[_harvest_mask].copy()
else:
    df_harvest_all = pd.DataFrame(columns=COLUMN_ORDER)

if len(df_harvest_all) > 0:
    # The "DOC" column shown in the All Harvest Details table (and the
    # Zone Wise breakdown further down, which reuses this same dataframe)
    # is DOC-as-of-today for a row that hasn't reached Full Harvest yet —
    # but for a row whose Harvest Type (checking the more recent Harvest
    # Type 2 first, then Harvest Type) says "Full", DOC instead STOPS
    # ADVANCING at that row's Full Harvest date, since the pond was fully
    # harvested there and DOC shouldn't keep counting up to today. Uses
    # the row's own saved "Date" field (captured here before the
    # Timestamp-based "Date" override just below) as the elapsed-days
    # starting point — the same basis "DOC" is always computed from
    # elsewhere in this file. This only affects these two harvest tables;
    # every other table/section in this file keeps using the row's saved
    # "DOC" value untouched.
    def _harvest_doc_display(row):
        try:
            _doc_num = int(float(row.get("DOC")))
        except (TypeError, ValueError):
            return row.get("DOC", "")
        _row_date = pd.to_datetime(row.get("Date"), errors="coerce")
        if pd.isna(_row_date):
            return row.get("DOC", "")
        _t2 = str(row.get("Harvest Type 2", "")).strip().lower()
        _t1 = str(row.get("Harvest Type", "")).strip().lower()
        _full_harvest_date_str = ""
        if "full" in _t2:
            _full_harvest_date_str = str(row.get("Harvest Date 2", "")).strip()
        elif "full" in _t1:
            _full_harvest_date_str = str(row.get("Harvest Date", "")).strip()
        if _full_harvest_date_str:
            _full_harvest_date = pd.to_datetime(_full_harvest_date_str, errors="coerce")
            if pd.notna(_full_harvest_date):
                return str(_doc_num + (_full_harvest_date - _row_date).days)
        return str(_doc_num + (pd.Timestamp(date.today()) - _row_date).days)

    if "DOC" in df_harvest_all.columns:
        df_harvest_all["DOC"] = df_harvest_all.apply(_harvest_doc_display, axis=1)

    # The "Date" column shown in the All Harvest Details table (and the
    # Zone Wise breakdown further down, which reuses this same dataframe)
    # is the LATEST DATE THE USER ACTUALLY SUBMITTED the harvest record —
    # i.e. the date portion of "Timestamp" — rather than the row's saved
    # "Date" field. Falls back to the original "Date" value if a row's
    # Timestamp can't be parsed. This only affects these two harvest
    # tables; every other table/section in this file keeps using the
    # original "Date" field untouched.
    if "Timestamp" in df_harvest_all.columns:
        _harvest_timestamp_date = pd.to_datetime(
            df_harvest_all["Timestamp"], errors="coerce"
        ).dt.strftime("%Y-%m-%d")
        df_harvest_all["Date"] = _harvest_timestamp_date.fillna(df_harvest_all["Date"])

    if "Date" in df_harvest_all.columns:
        df_harvest_all["_ParsedDate"] = pd.to_datetime(df_harvest_all["Date"], errors="coerce")
        _harvest_sort_cols = [c for c in ["Customer", "Farm Name with Code", "Pond Number"]
                               if c in df_harvest_all.columns] + ["_ParsedDate"]
        df_harvest_all = df_harvest_all.sort_values(by=_harvest_sort_cols).drop(columns=["_ParsedDate"])
    _harvest_display_cols = ["Customer", "Farm Name with Code", "Pond Number", "Date", "DOC",
                              "Species Culture", "Cycle Type", "Harvest Date", "Harvest Type",
                              "Harvest KG", "Harvest ABW", "Harvest Date 2", "Harvest Type 2",
                              "Harvest KG 2", "Harvest ABW 2", "Harvest Submitted Date", "Technician"]
    # "Harvest Submitted Date" is an existing column read straight from the
    # Sheet (see load_data() above) — added here only so it shows up as a
    # "Sort by" option; nothing computes or writes to it.
    _harvest_display_cols = [c for c in _harvest_display_cols if c in df_harvest_all.columns]

    st.caption(
        "🗑️ Select a row's checkbox (left edge) then click the recycle-bin icon above the "
        "table to remove that record — this writes 'H' to a Harvest Status column in the "
        "Google Sheet, so it stays removed after a refresh."
    )
    _harvest_full_cols = ["Timestamp"] + _harvest_display_cols
    df_harvest_editor_source = df_harvest_all[_harvest_full_cols].reset_index(drop=True)

    # Streamlit turns OFF its built-in click-to-sort on data_editor tables
    # whenever num_rows="dynamic" is set (needed just below for the
    # recycle-bin delete) — that's a Streamlit-level constraint, not
    # something togglable from here. So sorting is offered manually via
    # these two controls instead, applied to the data before it's handed
    # to the editor.
    _hsort_col1, _hsort_col2 = st.columns(2)
    with _hsort_col1:
        _default_sort_idx = _harvest_display_cols.index("Date") if "Date" in _harvest_display_cols else 0
        _harvest_sort_by = st.selectbox(
            "Sort by", options=_harvest_display_cols, index=_default_sort_idx, key="harvest_sort_by"
        )
    with _hsort_col2:
        _harvest_sort_order = st.selectbox(
            "Order", options=["Ascending", "Descending"], index=1, key="harvest_sort_order"
        )

    def _harvest_sort_key(series):
        numeric = pd.to_numeric(series, errors="coerce")
        if numeric.notna().sum() >= max(1, len(series) * 0.5):
            return numeric
        parsed_dt = pd.to_datetime(series, errors="coerce")
        if parsed_dt.notna().sum() >= max(1, len(series) * 0.5):
            return parsed_dt
        return series.astype(str).str.lower()

    df_harvest_editor_source = (
        df_harvest_editor_source.assign(_SortKey=_harvest_sort_key(df_harvest_editor_source[_harvest_sort_by]))
        .sort_values(by="_SortKey", ascending=(_harvest_sort_order == "Ascending"), na_position="last")
        .drop(columns=["_SortKey"])
        .reset_index(drop=True)
    )

    edited_harvest_all = st.data_editor(
        df_harvest_editor_source,
        use_container_width=True,
        hide_index=True,
        key="harvest_all_editor",
        num_rows="dynamic",
        column_order=_harvest_display_cols,
        disabled=_harvest_display_cols,
    )

    # A Timestamp missing from edited_harvest_all was just removed via the
    # recycle bin — mark that row's Harvest Status = 'H' in the main Sheet
    # so the removal persists.
    removed_timestamps = set(df_harvest_editor_source["Timestamp"]) - set(edited_harvest_all["Timestamp"].dropna())
    if removed_timestamps:
        try:
            ws_main = get_worksheet()
            harvest_status_col_idx = get_or_create_column(ws_main, "Harvest Status")
            for _ts in removed_timestamps:
                _cell = ws_main.find(str(_ts), in_column=1)
                if _cell:
                    ws_main.update_cell(_cell.row, harvest_status_col_idx, "H")
            st.rerun()
        except gspread.exceptions.APIError as e:
            st.error(f"❌ Could not save that removal to the Google Sheet. Please try again.\n\n{e}")

    _num_harvest_hidden = len(df_harvest_editor_source) - len(edited_harvest_all)
    _harvest_caption = f"{len(edited_harvest_all)} harvested record(s) shown."
    if _num_harvest_hidden:
        _harvest_caption += f" ({_num_harvest_hidden} row(s) hidden in this view.)"
    st.caption(_harvest_caption)
else:
    st.info("No harvest details recorded yet.")

# =========================================================================
# ALL HARVEST DETAILS — ZONE WISE. Same rows as "All Harvest Details"
# above, grouped by Zone into their own tables, so the manager can review
# harvests zone-by-zone without changing anything else on this page.
# =========================================================================
st.markdown("---")
st.markdown("#### 🌍 All Harvest Details — Zone Wise")

if len(df_harvest_all) > 0 and "Zone" in df_harvest_all.columns:
    _zones_present = sorted(
        {str(z).strip() for z in df_harvest_all["Zone"].tolist() if str(z).strip() and str(z).strip().lower() != "nan"}
    )
    if _zones_present:
        _selected_zones_harvest = st.multiselect(
            "Select Zone(s)", options=_zones_present, default=_zones_present, key="harvest_zone_filter"
        )
        if not _selected_zones_harvest:
            st.info("Select at least one zone above to display harvest details.")
        else:
            for _zone in _selected_zones_harvest:
                _zone_df = df_harvest_all[df_harvest_all["Zone"].astype(str).str.strip() == _zone]
                st.markdown(f"**{_zone}** ({len(_zone_df)} record(s))")
                st.dataframe(
                    _zone_df[_harvest_display_cols], use_container_width=True, hide_index=True
                )
    else:
        st.info("No Zone information found on the harvest records.")
else:
    st.info("No harvest details recorded yet.")

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>KMN Aqua Services - Water Quality Monitoring System "
    "(View — read only)</p>",
    unsafe_allow_html=True,
)

# =========================================================================
# RUNNING LIST — ZONE WISE. "Running" = every Customer/Farm that does NOT
# yet have all of its ponds Full H (i.e. still has at least one pond that
# is Running or Partial H). Grouped by Zone. For each running farm shows:
#   Customer Name | Farm Name with Code | No of Ponds |
#   Full Harvested Ponds | Partial H Ponds |
#   Latest Harvest Date | Harvest Quantity |
#   Last Feed Purchase Date | Due date last Purchase | Last Order
#
# The pond counts come from the same WaterQualityData sheet used above
# (load_data() + the same "latest record per pond" / "Partial H sticks if
# it ever happened" rules used by the Pond Layout section). Latest Harvest
# Date / Harvest Quantity are rolled up ONLY from that farm's Full H /
# Partial H ponds (2nd harvest slot wins over the 1st when both are
# filled in, same as the Full H box label in Pond Layout). The last three
# columns are mapped from the Sales Details Google Sheet via Customer Code
# (from Customer List.xlsx), using the same FEED-item logic as the "Last
# Feed Purchase Date Report" app: Item No. starts with "FEED", Quantity >
# 0, latest Date per Customer Code wins, Last Order = every item bought on
# that latest date combined into one string. Entirely read-only — this
# section never writes anything back to either Google Sheet.
# =========================================================================
st.markdown("---")
st.markdown("#### 🏃 Running List — Zone Wise")

RUNNING_FEED_PREFIX = "FEED"

def _running_customer_code(cust_name, farm_name):
    _match = customer_df[
        (customer_df["Customer Name"] == cust_name) & (customer_df["Farm Name with Code"] == farm_name)
    ]
    if len(_match) == 0:
        return ""
    for _cand in _CUSTOMER_CODE_COLUMN_CANDIDATES:
        if _cand in customer_df.columns:
            _val = str(_match.iloc[0].get(_cand, "")).strip()
            if _val and _val.lower() != "nan":
                return _val
    return ""

df_all_for_running = load_data()
_running_required = {"Customer", "Farm Name with Code", "Pond Number", "Date",
                      "Harvest Type", "Harvest Type 2"}
if len(df_all_for_running) > 0 and _running_required.issubset(df_all_for_running.columns):
    df_all_for_running = df_all_for_running.copy()
    df_all_for_running["_ParsedDate"] = pd.to_datetime(df_all_for_running["Date"], errors="coerce")

    # Latest saved record per Customer+Farm+Pond — same "latest per pond"
    # rule used by the Pond Layout section above.
    _latest_per_pond_all = (
        df_all_for_running.dropna(subset=["_ParsedDate"])
        .sort_values("_ParsedDate")
        .groupby(["Customer", "Farm Name with Code", "Pond Number"], as_index=False)
        .last()
    )

    # A pond keeps counting as Partial H if ANY of its saved records ever
    # had a Partial harvest (same rule as the Pond Layout section above).
    _partial_hist_all = (
        df_all_for_running.assign(
            _HasPartial=(
                df_all_for_running.get("Harvest Type", pd.Series("", index=df_all_for_running.index))
                .astype(str).str.lower().str.contains("partial")
                | df_all_for_running.get("Harvest Type 2", pd.Series("", index=df_all_for_running.index))
                .astype(str).str.lower().str.contains("partial")
            )
        )
        .groupby(["Customer", "Farm Name with Code", "Pond Number"])["_HasPartial"]
        .any()
    )

    def _pond_status_all(prow):
        _h_type = (str(prow.get("Harvest Type 2", "")).strip()
                   or str(prow.get("Harvest Type", "")).strip()).lower()
        _key = (prow.get("Customer", ""), prow.get("Farm Name with Code", ""), prow.get("Pond Number", ""))
        _has_partial = bool(_partial_hist_all.get(_key, False))
        if "full" in _h_type:
            return "Full H"
        elif "partial" in _h_type or _has_partial:
            return "Partial H"
        else:
            return "Running"

    _latest_per_pond_all["_PondStatus"] = _latest_per_pond_all.apply(_pond_status_all, axis=1)

    # --- FIX: Partial H date/KG lookback (Running List tables only) -----
    # A pond's "_PondStatus" can be Partial H even when its overall LATEST
    # saved row has both harvest slots blank (e.g. Partial H happened on
    # day X, then day X+1 got a routine new row with no harvest fields
    # filled in — the pond still reads as Partial H thanks to the history
    # check above, but that latest row itself carries no harvest date/KG).
    # For Full H this isn't an issue, since Full H status is only ever set
    # from a pond's true latest row in the first place. So: for Partial H
    # ponds only, look back through that pond's own history to the most
    # recent row where a harvest slot's Type actually says "partial", and
    # source the date/KG from THAT row instead of the pond's overall
    # latest row. This affects only the Running List's "Latest Harvest
    # Date" / "Harvest Quantity" rollups below — nothing else in the file.
    def _latest_partial_row(group):
        _g = group.dropna(subset=["_ParsedDate"]).sort_values("_ParsedDate", ascending=False)
        for _, _r in _g.iterrows():
            _t1 = str(_r.get("Harvest Type", "")).strip().lower()
            _t2 = str(_r.get("Harvest Type 2", "")).strip().lower()
            if "partial" in _t2 or "partial" in _t1:
                return _r
        return None

    _partial_source_by_pond = {}
    for _key, _grp in df_all_for_running.groupby(["Customer", "Farm Name with Code", "Pond Number"]):
        _row = _latest_partial_row(_grp)
        if _row is not None:
            _partial_source_by_pond[_key] = _row

    # Per-pond harvest date/quantity — prefer the 2nd harvest slot (Harvest
    # Date 2 / Harvest KG 2) when it's filled in, else fall back to the 1st
    # slot (Harvest Date / Harvest KG). Same "2nd slot wins" rule already
    # used for the Full H box label in the Pond Layout section above. For
    # Partial H ponds, these fields are sourced from the pond's actual last
    # "partial" row (found above) rather than the overall latest row.
    def _pond_harvest_date_str(prow):
        _source = prow
        if prow.get("_PondStatus") == "Partial H":
            _key = (prow.get("Customer", ""), prow.get("Farm Name with Code", ""), prow.get("Pond Number", ""))
            _src = _partial_source_by_pond.get(_key)
            if _src is not None:
                _source = _src
        return str(_source.get("Harvest Date 2", "")).strip() or str(_source.get("Harvest Date", "")).strip()

    _latest_per_pond_all["_PondHarvestDateStr"] = _latest_per_pond_all.apply(_pond_harvest_date_str, axis=1)
    _latest_per_pond_all["_PondHarvestDateParsed"] = pd.to_datetime(
        _latest_per_pond_all["_PondHarvestDateStr"], errors="coerce"
    )

    # Harvest Quantity per pond: for a Full H pond, use whichever slot's
    # Harvest Type actually says "Full" (checking the more recent 2nd slot
    # first, then the 1st slot) — that's the true full-harvest weight, not
    # just "whichever KG field happens to be filled in". For a pond that
    # hasn't reached Full H yet (still Partial H), the same check now runs
    # against that pond's actual last "partial" row (found above) instead
    # of the overall latest row. If neither slot's Type text matches (e.g.
    # a blank Type but a KG value was still entered), fall back to the 2nd
    # slot's KG, else the 1st — same safety fallback as before.
    def _pond_harvest_kg(prow):
        _wanted = "full" if prow.get("_PondStatus") == "Full H" else "partial"
        _source = prow
        if prow.get("_PondStatus") == "Partial H":
            _key = (prow.get("Customer", ""), prow.get("Farm Name with Code", ""), prow.get("Pond Number", ""))
            _src = _partial_source_by_pond.get(_key)
            if _src is not None:
                _source = _src
        _t1 = str(_source.get("Harvest Type", "")).strip().lower()
        _t2 = str(_source.get("Harvest Type 2", "")).strip().lower()
        _kg1 = pd.to_numeric(_source.get("Harvest KG", ""), errors="coerce")
        _kg2 = pd.to_numeric(_source.get("Harvest KG 2", ""), errors="coerce")
        if _wanted in _t2:
            return _kg2
        elif _wanted in _t1:
            return _kg1
        else:
            return _kg2 if pd.notna(_kg2) else _kg1

    _latest_per_pond_all["_PondHarvestKG"] = _latest_per_pond_all.apply(_pond_harvest_kg, axis=1)
    # --- end fix ---------------------------------------------------------

    # DOC Today per pond — same formula as the Pond Layout section's
    # "DOC Today" (saved DOC + days elapsed since that row's Date; stays 0
    # for a pond whose Cycle Type is "Soon to be"). Used only to build the
    # "DOC Today Values" rollup column below.
    def _pond_doc_today_running(prow):
        if str(prow.get("Cycle Type") or "").strip() == "Soon to be":
            return "0"
        _parsed = pd.to_datetime(prow.get("Date"), errors="coerce")
        if pd.isna(_parsed):
            return ""
        try:
            _doc_num = int(float(prow.get("DOC")))
        except (TypeError, ValueError):
            return ""
        _days_passed = (pd.Timestamp(date.today()) - _parsed).days
        return str(_doc_num + _days_passed)

    _latest_per_pond_all["_PondDocToday"] = _latest_per_pond_all.apply(_pond_doc_today_running, axis=1)

    # Groups a farm's per-pond values into "count-(value), count-(value)"
    # form, e.g. 3 ponds at DOC Today 39 and 1 pond at 23 -> "3-(39), 1-(23)".
    # Blank/empty pond values are skipped; used for DOC Today Values,
    # Issues, and Grade below. Sorted by count (most ponds first).
    def _grouped_value_counts(series):
        _clean = series.astype(str).str.strip()
        _clean = _clean[_clean != ""]
        if len(_clean) == 0:
            return "-"
        _counts = _clean.value_counts()
        return ", ".join(f"{_cnt}-({_val})" for _val, _cnt in _counts.items())

    # Roll pond statuses up to one row per Customer+Farm.
    _farm_pond_summary = (
        _latest_per_pond_all.groupby(["Customer", "Farm Name with Code"])
        .agg(
            **{
                "No of Ponds": ("Pond Number", "nunique"),
                "Full Harvested Ponds": ("_PondStatus", lambda s: (s == "Full H").sum()),
                "Partial H Ponds": ("_PondStatus", lambda s: (s == "Partial H").sum()),
            }
        )
        .reset_index()
    )

    # Latest Harvest Date / Harvest Quantity — rolled up ONLY from that
    # farm's ponds currently sitting at Full H or Partial H (a "Running"
    # pond has no harvest yet, so it's excluded from both).
    _harvested_ponds_all = _latest_per_pond_all[
        _latest_per_pond_all["_PondStatus"].isin(["Full H", "Partial H"])
    ]
    _farm_harvest_rollup = (
        _harvested_ponds_all.groupby(["Customer", "Farm Name with Code"])
        .agg(
            **{
                "_LatestHarvestDateParsed": ("_PondHarvestDateParsed", "max"),
                "Harvest Quantity": ("_PondHarvestKG", "sum"),
            }
        )
        .reset_index()
    )
    _farm_harvest_rollup["Latest Harvest Date"] = _farm_harvest_rollup["_LatestHarvestDateParsed"].dt.strftime(
        "%Y-%m-%d"
    ).fillna("")
    _farm_harvest_rollup = _farm_harvest_rollup.drop(columns=["_LatestHarvestDateParsed"])

    _farm_pond_summary = _farm_pond_summary.merge(
        _farm_harvest_rollup, on=["Customer", "Farm Name with Code"], how="left"
    )
    _farm_pond_summary["Latest Harvest Date"] = _farm_pond_summary["Latest Harvest Date"].fillna("")
    _farm_pond_summary["Harvest Quantity"] = _farm_pond_summary["Harvest Quantity"].fillna(0)

    # DOC Today Values / Issues / Grade — each farm's ponds grouped into
    # "count-(value)" form (see _grouped_value_counts above), across ALL
    # of that farm's ponds regardless of harvest status.
    _farm_pond_detail_rollup = (
        _latest_per_pond_all.groupby(["Customer", "Farm Name with Code"])
        .agg(
            **{
                "DOC Today Values": ("_PondDocToday", _grouped_value_counts),
                "Issues": ("Issues", _grouped_value_counts),
                "Grade": ("Grade", _grouped_value_counts),
            }
        )
        .reset_index()
    )
    _farm_pond_summary = _farm_pond_summary.merge(
        _farm_pond_detail_rollup, on=["Customer", "Farm Name with Code"], how="left"
    )

    # Running = farms where NOT every pond is Full H yet.
    _farm_pond_summary = _farm_pond_summary[
        _farm_pond_summary["Full Harvested Ponds"] < _farm_pond_summary["No of Ponds"]
    ].reset_index(drop=True)

    if len(_farm_pond_summary) == 0:
        st.info("No running farms — every farm's ponds are fully harvested.")
    else:
        # Attach Zone (from Customer List.xlsx) for grouping. The lookup's
        # "Customer Name" column is renamed to "Customer" BEFORE the merge
        # so the merge key names line up exactly (on=...) — merging with
        # mismatched left_on/right_on names would keep both "Customer" and
        # "Customer Name" as separate columns, and the later rename below
        # would then collide with that leftover "Customer Name" column,
        # producing a dataframe with two columns of the same name (which
        # Streamlit's table renderer cannot display).
        _zone_lookup = customer_df[["Customer Name", "Farm Name with Code", "Zone"]].drop_duplicates(
            subset=["Customer Name", "Farm Name with Code"]
        ).rename(columns={"Customer Name": "Customer"})
        _farm_pond_summary = _farm_pond_summary.merge(
            _zone_lookup,
            on=["Customer", "Farm Name with Code"],
            how="left",
        )

        # Pull Last Feed Purchase Date / Due date last Purchase / Last Order
        # per Customer Code from the Sales Details sheet (same FEED-item
        # logic as the Last Feed Purchase Date Report app / "2nd code").
        _running_feed_info = {}
        try:
            df_sales_running = load_sales_data()
        except Exception:
            df_sales_running = None

        # "Last Order" label per feed item — just the words "NANAMI" and
        # "EGO" swapped for a single letter (N / E) within the description,
        # e.g. "NANAMI 3M" -> "N 3M", "EGO - 01S" -> "E - 01S". Everything
        # else in the description (sizes, dashes, spacing) stays exactly
        # as-is — this only shortens the column width, same as before.
        def _running_order_item_label(desc):
            _label = str(desc)
            _label = re.sub(r"(?i)\bnanami\b", "N", _label)
            _label = re.sub(r"(?i)\bego\b", "E", _label)
            return _label

        if df_sales_running is not None and len(df_sales_running) > 0:
            _sales_r = df_sales_running.copy()
            _sales_r["Quantity"] = pd.to_numeric(_sales_r["Quantity"], errors="coerce").fillna(0)
            _sales_r["_ParsedDate"] = pd.to_datetime(_sales_r["Date"], errors="coerce")
            _feed_r = _sales_r[
                _sales_r["Item No."].astype(str).str.strip().str.upper().str.startswith(RUNNING_FEED_PREFIX)
                & (_sales_r["Quantity"] > 0)
            ]
            _last_feed_date_r = _feed_r.dropna(subset=["_ParsedDate"]).groupby("Customer Code")["_ParsedDate"].max()
            _today_r = pd.Timestamp(date.today())
            for _code, _last_date in _last_feed_date_r.items():
                _same_day = _feed_r[
                    (_feed_r["Customer Code"] == _code) & (_feed_r["_ParsedDate"] == _last_date)
                ]
                _order_parts = [
                    f"{_running_order_item_label(d)} ({q:g})"
                    for d, q in zip(_same_day["Item Description"], _same_day["Quantity"])
                ]
                _running_feed_info[_code] = {
                    "Last Feed Purchase Date": _last_date.strftime("%Y-%m-%d"),
                    "Due date last Purchase": (_today_r - _last_date).days,
                    "Last Order": ", ".join(_order_parts),
                }

        def _running_feed_field(row, field, default=""):
            _code = _running_customer_code(row["Customer"], row["Farm Name with Code"])
            return _running_feed_info.get(_code, {}).get(field, default)

        _farm_pond_summary["Last Feed Purchase Date"] = _farm_pond_summary.apply(
            lambda r: _running_feed_field(r, "Last Feed Purchase Date"), axis=1
        )
        _farm_pond_summary["Due date last Purchase"] = _farm_pond_summary.apply(
            lambda r: _running_feed_field(r, "Due date last Purchase"), axis=1
        )
        _farm_pond_summary["Last Order"] = _farm_pond_summary.apply(
            lambda r: _running_feed_field(r, "Last Order"), axis=1
        )

        _farm_pond_summary = _farm_pond_summary.rename(columns={"Customer": "Customer Name"})
        _running_display_cols = [
            "Customer Name", "Farm Name with Code", "No of Ponds", "Full Harvested Ponds",
            "Partial H Ponds", "Latest Harvest Date", "Harvest Quantity",
            "DOC Today Values", "Issues", "Grade",
            "Last Feed Purchase Date", "Due date last Purchase", "Last Order",
        ]

        _zones_running = sorted(
            {str(z).strip() for z in _farm_pond_summary["Zone"].tolist() if str(z).strip() and str(z).strip().lower() != "nan"}
        )
        if _zones_running:
            _selected_zones_running = st.multiselect(
                "Select Zone(s)  ", options=_zones_running, default=_zones_running, key="running_zone_filter"
            )
            if not _selected_zones_running:
                st.info("Select at least one zone above to display the running list.")
            else:
                for _zone_r in _selected_zones_running:
                    _zone_running_df = _farm_pond_summary[
                        _farm_pond_summary["Zone"].astype(str).str.strip() == _zone_r
                    ]
                    st.markdown(f"**{_zone_r}** ({len(_zone_running_df)} farm(s) running)")
                    st.dataframe(
                        _zone_running_df[_running_display_cols], use_container_width=True, hide_index=True
                    )
        else:
            st.dataframe(_farm_pond_summary[_running_display_cols], use_container_width=True, hide_index=True)
            st.caption("No Zone information found on the customer list — showing unfiltered.")
else:
    st.info("No records available yet to build the running list.")

# =========================================================================
# SPECIES-WISE POND SUMMARY — ZONE WISE. Bottom-of-page Species Culture
# selector, followed by zone-wise tables showing, for every Customer+Farm
# that has at least one pond of the selected species:
#   Customer Name | Code with Farm Name | His Total Ponds |
#   Number of Selected Species Ponds | DOC of these selected ponds
#
# "His Total Ponds" is that farm's pond count across ALL species (every
# pond it has ever had a saved record for) — unlike the Running List
# above, this section isn't limited to farms still running (Full-H-only
# farms are included too, as long as they have a pond of the selected
# species). "DOC of these selected ponds" uses the same "count-(value)"
# grouping as the Running List's DOC Today Values column (e.g. 3 ponds at
# DOC Today 39 and 1 at 23 -> "3-(39), 1-(23)"), computed only from the
# ponds matching the selected species. Entirely read-only.
# =========================================================================
st.markdown("---")
st.markdown("#### 🦐 Species-wise Pond Summary — Zone Wise")

df_all_for_species = load_data()
_species_required = {"Customer", "Farm Name with Code", "Pond Number", "Date", "Species Culture"}
if len(df_all_for_species) > 0 and _species_required.issubset(df_all_for_species.columns):
    df_all_for_species = df_all_for_species.copy()
    df_all_for_species["_ParsedDate"] = pd.to_datetime(df_all_for_species["Date"], errors="coerce")

    # Latest saved record per Customer+Farm+Pond, same rule used elsewhere
    # in this file.
    _latest_per_pond_species = (
        df_all_for_species.dropna(subset=["_ParsedDate"])
        .sort_values("_ParsedDate")
        .groupby(["Customer", "Farm Name with Code", "Pond Number"], as_index=False)
        .last()
    )

    # DOC Today per pond — same formula as the Pond Layout / Running List
    # sections above.
    def _pond_doc_today_species(prow):
        if str(prow.get("Cycle Type") or "").strip() == "Soon to be":
            return "0"
        _parsed = pd.to_datetime(prow.get("Date"), errors="coerce")
        if pd.isna(_parsed):
            return ""
        try:
            _doc_num = int(float(prow.get("DOC")))
        except (TypeError, ValueError):
            return ""
        _days_passed = (pd.Timestamp(date.today()) - _parsed).days
        return str(_doc_num + _days_passed)

    _latest_per_pond_species["_PondDocToday"] = _latest_per_pond_species.apply(_pond_doc_today_species, axis=1)

    # Groups a farm's per-pond values into "count-(value), count-(value)"
    # form (self-contained copy of the same helper used by the Running
    # List section above, kept local so this section works on its own).
    def _grouped_value_counts_species(series):
        _clean = series.astype(str).str.strip()
        _clean = _clean[_clean != ""]
        if len(_clean) == 0:
            return "-"
        _counts = _clean.value_counts()
        return ", ".join(f"{_cnt}-({_val})" for _val, _cnt in _counts.items())

    _species_options = sorted(
        [s for s in _latest_per_pond_species["Species Culture"].astype(str).str.strip().unique()
         if s and s.lower() != "nan"]
    )
    if not _species_options:
        st.info("No Species Culture values found on any saved record.")
    else:
        _selected_species = st.selectbox(
            "Species Culture", options=_species_options, key="species_summary_select"
        )

        # His Total Ponds — every pond that farm has ever had a saved
        # record for, any species.
        _farm_total_ponds = (
            _latest_per_pond_species.groupby(["Customer", "Farm Name with Code"])["Pond Number"]
            .nunique()
            .reset_index(name="His Total Ponds")
        )

        # Ponds matching the selected species only.
        _species_ponds = _latest_per_pond_species[
            _latest_per_pond_species["Species Culture"].astype(str).str.strip() == _selected_species
        ]
        _farm_species_summary = (
            _species_ponds.groupby(["Customer", "Farm Name with Code"])
            .agg(
                **{
                    "Number of Selected Species Ponds": ("Pond Number", "nunique"),
                    "DOC of these selected ponds": ("_PondDocToday", _grouped_value_counts_species),
                }
            )
            .reset_index()
        )

        _farm_species_table = _farm_total_ponds.merge(
            _farm_species_summary, on=["Customer", "Farm Name with Code"], how="left"
        )
        _farm_species_table["Number of Selected Species Ponds"] = (
            _farm_species_table["Number of Selected Species Ponds"].fillna(0).astype(int)
        )
        _farm_species_table["DOC of these selected ponds"] = (
            _farm_species_table["DOC of these selected ponds"].fillna("-")
        )

        # Only farms that actually have at least one pond of the selected
        # species show up in the table.
        _farm_species_table = _farm_species_table[
            _farm_species_table["Number of Selected Species Ponds"] > 0
        ].reset_index(drop=True)

        if len(_farm_species_table) == 0:
            st.info(f"No farms currently have any ponds running '{_selected_species}'.")
        else:
            _zone_lookup_species = customer_df[["Customer Name", "Farm Name with Code", "Zone"]].drop_duplicates(
                subset=["Customer Name", "Farm Name with Code"]
            ).rename(columns={"Customer Name": "Customer"})
            _farm_species_table = _farm_species_table.merge(
                _zone_lookup_species, on=["Customer", "Farm Name with Code"], how="left"
            )
            _farm_species_table = _farm_species_table.rename(
                columns={"Customer": "Customer Name", "Farm Name with Code": "Code with Farm Name"}
            )
            _species_display_cols = [
                "Customer Name", "Code with Farm Name", "His Total Ponds",
                "Number of Selected Species Ponds", "DOC of these selected ponds",
            ]

            _zones_species = sorted(
                {str(z).strip() for z in _farm_species_table["Zone"].tolist() if str(z).strip() and str(z).strip().lower() != "nan"}
            )
            if _zones_species:
                _selected_zones_species = st.multiselect(
                    "Select Zone(s)   ", options=_zones_species, default=_zones_species,
                    key="species_zone_filter"
                )
                if not _selected_zones_species:
                    st.info("Select at least one zone above to display the species summary.")
                else:
                    for _zone_s in _selected_zones_species:
                        _zone_species_df = _farm_species_table[
                            _farm_species_table["Zone"].astype(str).str.strip() == _zone_s
                        ]
                        st.markdown(f"**{_zone_s}** ({len(_zone_species_df)} farm(s))")
                        st.dataframe(
                            _zone_species_df[_species_display_cols], use_container_width=True, hide_index=True
                        )
            else:
                st.dataframe(_farm_species_table[_species_display_cols], use_container_width=True, hide_index=True)
                st.caption("No Zone information found on the customer list — showing unfiltered.")
else:
    st.info("No records available yet to build the species summary.")
