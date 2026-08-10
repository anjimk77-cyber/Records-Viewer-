# Water Quality Report — Manager View (read-only)

A read-only Streamlit app for viewing saved Water Quality records. It
shares the same Google Sheet as the main data-entry app, but this app has
no ability to add, edit, or delete anything — it only has:

1. **📋 Enter Water Quality Data** — Customer / Farm selection (used only
   to filter which farm's records to view).
2. **📊 All Saved Records** — a live, read-only table pulled straight from
   the Google Sheet, with a Refresh button.

## Files in this repo

- `manager_view.py` — the app itself.
- `requirements.txt` — Python dependencies.
- `.streamlit/secrets.toml.example` — template for the secrets this app
  needs. Copy its contents into real secrets (see below) — never commit
  the real, filled-in version.
- `Customer List.xlsx` — **you need to add this yourself.** Copy the exact
  same `Customer List.xlsx` used by your data-entry app into this repo's
  root folder (same file name). It's what populates the Customer / Farm
  dropdowns.

## Setup

1. **Add `Customer List.xlsx`** to this repo's root folder (same file the
   data-entry app uses).
2. **Push this repo to GitHub** (a separate repo from the data-entry app
   is fine — this app only needs `manager_view.py`, `requirements.txt`,
   and `Customer List.xlsx` to run).
3. **Deploy on Streamlit Community Cloud** (share.streamlit.io):
   - New app → point it at this repo.
   - Main file path: `manager_view.py`.
   - This gives you a separate link/URL from your data-entry app.
4. **Add secrets**: in the app's Settings → Secrets, paste the same
   `[gcp_service_account]` and `[gsheet]` values used by your data-entry
   app (see `.streamlit/secrets.toml.example` for the shape). Both apps
   must point at the *same* `sheet_id` and `worksheet_name` so the manager
   sees the same data the technicians are entering.
5. The Google Sheet must already be shared (Editor access is fine, or
   Viewer since this app never writes) with the service account's
   `client_email`.

## Notes

- This app never writes to the Google Sheet — there are no
  append/update/delete functions in `manager_view.py` at all.
- Soft-deleted rows (`Deleted = Yes`) are filtered out automatically, same
  as in the data-entry app.
- If you add new columns to the Sheet later (via the data-entry app), also
  update the `COLUMN_ORDER` list in `manager_view.py` here to match, or
  the new columns just won't be shown in this view.
