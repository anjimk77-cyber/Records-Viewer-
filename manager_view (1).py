import pandas as pd
import streamlit as st

import gspread
from google.oauth2.service_account import Credentials

# =========================================================================
# CONFIG
#
# This is a SEPARATE, READ-ONLY app meant for the manager. It shares the
# same Google Sheet and the same "Customer List.xlsx" as the data-entry
# app (app.py), but it never writes anything — there is no Pond Details
# editor and no Harvest Details form here, only:
#   1) "📋 Enter Water Quality Data" — Customer / Farm selection (used only
#      to choose which farm's records to view)
#   2) "📊 All Saved Records" — a live, read-only view of that farm's data
#      straight from the Google Sheet.
# Deploy this file as its own Streamlit app (its own URL/link) so the
# manager gets a separate link from the data-entry app.
# =========================================================================
st.set_page_config(page_title="Water Quality Report - Manager View", layout="wide", page_icon="📊")

CUSTOMER_FILE = "Customer List.xlsx"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Must match app.py's COLUMN_ORDER exactly, since both apps read/write the
# same Google Sheet.
COLUMN_ORDER = [
    "Timestamp", "Customer", "Farm Name with Code", "Zone", "Area",
    "Pond Number", "Date", "Species Culture", "Cycle Type",
    "DOC", "Density", "Feed Per Day", "ABW",
    "Issues", "Water Color", "Grade", "Remark", "Technician",
    "Harvest Date", "Harvest Type", "Harvest KG", "Harvest ABW",
    "Deleted",
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

st.markdown("<h1 style='text-align: center;'>Water Quality Report - Manager View</h1>",
            unsafe_allow_html=True)
st.subheader("KMN Aqua Services")
st.markdown("---")

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
st.subheader("📋 Enter Water Quality Data")

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
    if "Date" in df_farm_summary.columns:
        df_farm_summary["_ParsedDate"] = pd.to_datetime(df_farm_summary["Date"], errors="coerce")
        sort_cols = [c for c in ["Pond Number"] if c in df_farm_summary.columns] + ["_ParsedDate"]
        df_farm_summary = df_farm_summary.sort_values(by=sort_cols).drop(columns=["_ParsedDate"])
    _farm_display_cols = ["Pond Number", "Date", "Species Culture", "Cycle Type", "DOC", "Density",
                           "Feed Per Day", "ABW", "Issues", "Water Color", "Grade", "Remark", "Technician",
                           "Harvest Date", "Harvest Type", "Harvest KG", "Harvest ABW"]
    _farm_display_cols = [c for c in _farm_display_cols if c in df_farm_summary.columns]
    st.dataframe(df_farm_summary[_farm_display_cols], use_container_width=True, hide_index=True)
    st.caption(f"{len(df_farm_summary)} saved record(s) across all ponds for {farm}.")
else:
    st.info(f"No saved records yet for {farm}.")

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>KMN Aqua Services - Water Quality Monitoring System "
    "(Manager View — read only)</p>",
    unsafe_allow_html=True,
)
