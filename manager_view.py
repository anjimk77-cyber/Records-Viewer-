import pandas as pd
import streamlit as st
from datetime import date

import gspread
from google.oauth2.service_account import Credentials

try:
    from streamlit_autorefresh import st_autorefresh
    _AUTOREFRESH_AVAILABLE = True
except ImportError:
    _AUTOREFRESH_AVAILABLE = False

# =========================================================================
# CONFIG
#
# This is a SEPARATE app meant for the manager. It shares the same Google
# Sheet and the same "Customer List.xlsx" as the data-entry app (app.py).
# It has no Pond Details editor and no Harvest Details form — the only
# write it ever performs is the "Settle" Yes/blank flag in the Sales
# Details sheet (set when a date row's tick is toggled below), so that
# tick survives a refresh. Everything else here is read-only:
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
# LIVE HARVEST ALERTS — auto-reruns this page every 30s (while it's open
# in the browser) via streamlit-autorefresh, so a pond that turns Partial H
# / Full H shows a toast alert without the manager needing to click
# anything. Needs `pip install streamlit-autorefresh`; if it's not
# installed, the page just runs normally without auto-refreshing.
# =========================================================================
if _AUTOREFRESH_AVAILABLE:
    st_autorefresh(interval=30_000, key="pond_status_autorefresh")
else:
    st.caption(
        "ℹ️ Install `streamlit-autorefresh` (`pip install streamlit-autorefresh`) "
        "to enable live harvest-status alerts on this page."
    )

# =========================================================================
# GOOGLE SHEETS BACKEND — READ-ONLY. This app has no append/update/delete
# functions at all, so there is no way for it to modify the Sheet.
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

def get_or_create_settle_column(ws):
    """Returns the 1-based column number of the 'Settle' header in the
    Sales Details sheet, creating that header (in the first empty column)
    if it isn't there yet."""
    headers = ws.row_values(1)
    if "Settle" in headers:
        return headers.index("Settle") + 1
    new_col_idx = len(headers) + 1
    ws.update_cell(1, new_col_idx, "Settle")
    return new_col_idx

def load_data():
    """Always reads fresh from the Google Sheet (no caching), so the
    manager always sees the latest saved records — including anything
    just added by the data-entry app or edited directly in the Sheet.
    Soft-deleted rows (Deleted = Yes) are filtered out, same as the
    data-entry app."""
    ws = get_worksheet()
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for c in COLUMN_ORDER:
        if c not in df.columns:
            df[c] = ""
    if len(df) > 0:
        df = df[COLUMN_ORDER]
    df = df.astype(str).replace("nan", "")
    if "Deleted" in df.columns:
        is_deleted = df["Deleted"].astype(str).str.strip().str.lower().isin(["yes", "true", "1"])
        df = df[~is_deleted].reset_index(drop=True)
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

df_farm_summary = load_data()
_farm_required = {"Customer", "Farm Name with Code"}
if len(df_farm_summary) > 0 and _farm_required.issubset(df_farm_summary.columns):
    df_farm_summary = df_farm_summary[
        (df_farm_summary["Customer"] == customer) & (df_farm_summary["Farm Name with Code"] == farm)
    ].copy()
else:
    df_farm_summary = pd.DataFrame(columns=COLUMN_ORDER)

if len(df_farm_summary) > 0:
    total_expect_harvest_kg = None
    if "Date" in df_farm_summary.columns:
        df_farm_summary["_ParsedDate"] = pd.to_datetime(df_farm_summary["Date"], errors="coerce")

        # Total Expect Harvest (KG) for the farm = each pond's MOST RECENT
        # saved record's "Expect Harvest (KG)" value, summed across every
        # pond on this farm (not every historical row, which would double
        # count a pond's earlier daily estimates). Same logic as app.py.
        if {"Pond Number", "Expect Harvest (KG)"}.issubset(df_farm_summary.columns):
            _latest_per_pond = (
                df_farm_summary.dropna(subset=["_ParsedDate"])
                .sort_values("_ParsedDate")
                .groupby("Pond Number", as_index=False)
                .last()
            )
            _harvest_vals = pd.to_numeric(_latest_per_pond["Expect Harvest (KG)"], errors="coerce").dropna()
            if len(_harvest_vals) > 0:
                total_expect_harvest_kg = float(_harvest_vals.sum())

        sort_cols = [c for c in ["Pond Number"] if c in df_farm_summary.columns] + ["_ParsedDate"]
        df_farm_summary = df_farm_summary.sort_values(by=sort_cols).drop(columns=["_ParsedDate"])

    # "DOC Today" = this row's saved DOC + however many days have passed
    # between its Date and today (i.e. what the DOC would be right now).
    # It's a live, always-changing number rather than something actually
    # saved in the Sheet, so it's shown in red/bold to stand out. Rows
    # whose Cycle Type is "Soon to be" haven't actually started yet, so
    # DOC Today just stays 0 for them instead of counting elapsed days.
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

    st.markdown(
        _render_highlighted_table(df_farm_summary, _farm_display_cols, "DOC Today"),
        unsafe_allow_html=True,
    )
    st.caption(f"{len(df_farm_summary)} saved record(s) across all ponds for {farm}.")

    if total_expect_harvest_kg is not None:
        st.markdown(
            f"**🌾 Total Expect Harvest (KG) — {farm}: {total_expect_harvest_kg:,.2f} kg** "
            "(sum of each pond's latest Expect Harvest (KG) estimate)"
        )

    # =========================================================================
    # POND LAYOUT — one rectangle per Pond Number (using that pond's most
    # recent saved record). Running / Partial H ponds show DOC Today
    # centered inside the box; Full H ponds show "Full H" and its Harvest
    # Date instead; ponds whose Cycle Type is "Soon to be" show "Soon to be"
    # instead, in their own distinct color.
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
            _cycle_type_status = str(prow.get("Cycle Type", "")).strip()
            _h_type_lower = _pond_harvest_type(prow).lower()
            _has_partial_history = bool(_partial_history_by_pond.get(_pond_no_status, False))
            if _cycle_type_status == "Soon to be":
                return "Soon to be"
            elif "full" in _h_type_lower:
                return "Full H"
            elif "partial" in _h_type_lower or _has_partial_history:
                return "Partial H"
            else:
                return "Running"

        def _pond_box_color(prow):
            _status = _pond_status(prow)
            if _status == "Soon to be":
                return "#e6d9f7"  # purple — Soon to be
            elif _status == "Partial H":
                return "#fff3cd"  # yellow — Partial Harvest
            elif _status == "Full H":
                return "#d4edda"  # green — Full Harvest
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

        # ---------------------------------------------------------------
        # LIVE ALERT: compare each pond's current status against the last
        # status seen this session — fire a toast only when a pond newly
        # becomes Partial H or Full H (not on every autorefresh tick).
        # ---------------------------------------------------------------
        for _, _prow in _pond_latest.iterrows():
            _pond_no_alert = str(_prow.get("Pond Number", "")).strip()
            _status_now = _pond_status(_prow)

            _status_key = f"pond_status::{customer}::{farm}::{_pond_no_alert}"
            _prev_status = st.session_state.get(_status_key)

            if _status_now in ("Full H", "Partial H") and _prev_status != _status_now:
                st.toast(
                    f"🚨 {farm} — Pond {_pond_no_alert}: {_status_now}",
                    icon="🚨",
                )
            st.session_state[_status_key] = _status_now

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
                # Ponds not yet started: show "Soon to be" instead of DOC Today.
                _box_middle_html = (
                    "<div style='font-size:1.1rem;font-weight:bold;color:#7a3fc4;'>Soon to be</div>"
                )
            else:
                # Running / Partial H ponds: keep showing DOC Today, as before.
                _doc_today_val = _escape_html_pond(_prow.get("DOC Today", "") or "-")
                _box_middle_html = (
                    f"<div style='font-size:1.4rem;font-weight:bold;color:red;'>{_doc_today_val}</div>"
                    "<div style='font-size:0.7rem;color:#777;'>DOC Today</div>"
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
            "<div><span style='display:inline-block;width:14px;height:14px;background:#e6d9f7;"
            "border:1px solid #333;border-radius:3px;vertical-align:middle;margin-right:6px;'></span>Soon to be</div>"
            "</div>",
            unsafe_allow_html=True,
        )

        st.markdown(
            f"<div style='display:flex;flex-wrap:wrap;justify-content:center;'>{_pond_boxes_html}</div>",
            unsafe_allow_html=True,
        )
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

            # A tick is only "on" for a date once every line item on that
            # date has Settle = "Yes" in the Sales Details sheet — this is
            # what makes the tick survive a refresh instead of resetting.
            _settled_by_date = (
                df_sales_farm.groupby("Date")["Settle"]
                .apply(lambda s: s.str.strip().str.lower().eq("yes").all())
            )

            pivot_display = pivot_sales.reset_index()
            pivot_display.insert(1, "Delete", pivot_display["Date"].map(_settled_by_date).fillna(False))
            sales_editor_key = f"sales_delete_editor_{selected_customer_code}"

            edited_pivot = st.data_editor(
                pivot_display,
                use_container_width=True,
                hide_index=True,
                key=sales_editor_key,
                column_config={
                    "Delete": st.column_config.CheckboxColumn(
                        "Financial Status",
                        help="Remove this date's row from this view only — it stays in the Google Sheet.",
                        default=False,
                    )
                },
                disabled=[c for c in pivot_display.columns if c != "Delete"],
            )

            # Persist any tick changes: every line item belonging to that
            # date gets its 'Settle' cell in the Sales Details Google Sheet
            # set to "Yes" (ticked) or cleared (unticked). Only cells whose
            # value actually needs to change are written.
            sales_ws = get_sales_worksheet()
            settle_col_idx = get_or_create_settle_column(sales_ws)
            _wrote_change = False
            for _, prow in edited_pivot.iterrows():
                d = prow["Date"]
                checked = bool(prow["Delete"])
                rows_for_date = df_sales_farm[df_sales_farm["Date"] == d]
                for _, r in rows_for_date.iterrows():
                    current_settle = str(r.get("Settle", "")).strip().lower()
                    if checked and current_settle != "yes":
                        sales_ws.update_cell(int(r["_RowNumber"]), settle_col_idx, "Yes")
                        _wrote_change = True
                    elif not checked and current_settle == "yes":
                        sales_ws.update_cell(int(r["_RowNumber"]), settle_col_idx, "")
                        _wrote_change = True

            # Without this, the checkbox you just ticked still shows its
            # OLD value until a second click, because the sheet read that
            # built this table happened before the write above. Rerunning
            # immediately after a successful save re-reads the sheet so the
            # box reflects the saved state right away — one click, not two.
            if _wrote_change:
                st.rerun()

            kept_dates = edited_pivot.loc[~edited_pivot["Delete"], "Date"]
            _num_hidden = len(edited_pivot) - len(kept_dates)

            df_sales_farm_visible = df_sales_farm[df_sales_farm["Date"].isin(kept_dates)]

            _caption = (
                f"{len(df_sales_farm_visible)} sales line item(s) across "
                f"{len(kept_dates)} date(s) for Customer Code '{selected_customer_code}'."
            )
            if _num_hidden:
                _caption += f" ({_num_hidden} row(s) hidden in this view.)"
            st.caption(_caption)

            total_qty = df_sales_farm_visible["Quantity"].sum() if len(df_sales_farm_visible) else 0
            total_amt = df_sales_farm_visible["Sales Amt"].sum() if len(df_sales_farm_visible) else 0
            st.markdown(
                f"**Total Quantity: {total_qty:,.0f}  |  Total Sales Amt: {total_amt:,.2f}**"
            )

            # =================================================================
            # FEED ORDER STATUS — two rows of rectangles, one per feed size,
            # in a fixed order: NANAMI FEED sizes, then EGO FEED sizes. Each
            # box shows that size's total purchased Quantity in the middle
            # and its latest purchase Date below. Boxes are colored one way
            # if any quantity has been purchased, another if not yet.
            # =================================================================
            NANAMI_FEED_ORDER = [
                "NANAMI 1", "NANAMI 1S", "NANAMI 2S", "NANAMI 3S",
                "NANAMI 3M", "NANAMI 3L", "NANAMI 4",
            ]
            EGO_FEED_ORDER = [
                "EGO - 01", "EGO - 01S", "EGO - 02S", "EGO - 03S",
                "EGO - 03M", "EGO - 03L", "EGO - 04L",
            ]

            def _escape_html_feed(v):
                return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

            def _render_feed_row(title, size_labels):
                _desc_upper = df_sales_farm_visible["Item Description"].astype(str).str.strip().str.upper()
                boxes_html = ""
                for _label in size_labels:
                    _subset = df_sales_farm_visible[_desc_upper == _label.upper()]
                    _qty = _subset["Quantity"].sum() if len(_subset) else 0
                    _latest_date = _subset["Date"].max() if len(_subset) else "-"
                    _has_qty = _qty > 0
                    _box_color = "#d4edda" if _has_qty else "#e2e2e2"  # green if purchased, gray if not yet
                    _qty_label = f"{_qty:,.0f}" if _has_qty else "-"
                    boxes_html += (
                        "<div style='width:120px;height:80px;border:2px solid #333;border-radius:6px;"
                        "display:flex;flex-direction:column;align-items:center;justify-content:center;"
                        f"background:{_box_color};margin:5px;'>"
                        f"<div style='font-size:0.7rem;color:#555;'>{_escape_html_feed(_label)}</div>"
                        f"<div style='font-size:1.2rem;font-weight:bold;color:#111;'>{_qty_label}</div>"
                        f"<div style='font-size:0.65rem;color:#777;'>{_escape_html_feed(_latest_date)}</div>"
                        "</div>"
                    )
                st.markdown(f"**{title}**")
                st.markdown(
                    f"<div style='display:flex;flex-wrap:wrap;justify-content:center;'>{boxes_html}</div>",
                    unsafe_allow_html=True,
                )

            st.markdown("---")
            _render_feed_row("🐟 NANAMI FEED", NANAMI_FEED_ORDER)
            st.markdown("")
            _render_feed_row("🐟 EGO FEED", EGO_FEED_ORDER)

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
    if "Date" in df_harvest_all.columns:
        df_harvest_all["_ParsedDate"] = pd.to_datetime(df_harvest_all["Date"], errors="coerce")
        _harvest_sort_cols = [c for c in ["Customer", "Farm Name with Code", "Pond Number"]
                               if c in df_harvest_all.columns] + ["_ParsedDate"]
        df_harvest_all = df_harvest_all.sort_values(by=_harvest_sort_cols).drop(columns=["_ParsedDate"])
    _harvest_display_cols = ["Customer", "Farm Name with Code", "Pond Number", "Date", "DOC",
                              "Species Culture", "Cycle Type", "Harvest Date", "Harvest Type",
                              "Harvest KG", "Harvest ABW", "Harvest Date 2", "Harvest Type 2",
                              "Harvest KG 2", "Harvest ABW 2", "Technician"]
    _harvest_display_cols = [c for c in _harvest_display_cols if c in df_harvest_all.columns]
    st.dataframe(df_harvest_all[_harvest_display_cols], use_container_width=True, hide_index=True)
    st.caption(f"{len(df_harvest_all)} harvested record(s) found across the Google Sheet.")
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
        [z for z in df_harvest_all["Zone"].astype(str).str.strip().unique() if z and z.lower() != "nan"]
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
