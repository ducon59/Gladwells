# 📊 Gladwells Management Accounts Dashboard

## How it works

The dashboard uses a tiny **Google Sheet as an index** — it holds the Drive file ID of the current month's accounts file. Each month you update one cell in that sheet. The dashboard reads it automatically.

No scraping. No API keys. No folder permissions issues. Works forever.

---

## One-time setup (10 minutes)

### Step 1 — Upload the .xlsm to Google Drive

1. Go to [drive.google.com](https://drive.google.com)
2. Upload the management accounts `.xlsm` file
3. Right-click the file → **Share** → **Anyone with the link → Viewer**
4. Right-click again → **Get link** — copy the URL, which looks like:
   ```
   https://drive.google.com/file/d/1aBcDeFgHiJkLmNoPqRsTuVwXyZ/view
   ```
5. Your **File ID** is the part between `/d/` and `/view`:
   ```
   1aBcDeFgHiJkLmNoPqRsTuVwXyZ
   ```

### Step 2 — Create the index Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank sheet
2. In **cell A1**: paste the File ID from Step 1 (e.g. `1aBcDeFgHiJkLmNoPqRsTuVwXyZ`)
3. In **cell B1**: type the filename (e.g. `Gladwells_Feb26_Management_Accounts.xlsm`)
4. Go to **File → Share → Publish to web**
5. Under "Link", set it to **Sheet1** and format **Comma-separated values (.csv)**
6. Click **Publish** → copy the URL, which looks like:
   ```
   https://docs.google.com/spreadsheets/d/SHEET_ID/pub?output=csv
   ```

### Step 3 — Deploy to Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app** → select repo → `app.py`
3. Under **Advanced settings → Secrets**, add:
   ```toml
   GDRIVE_FOLDER_ID = "https://docs.google.com/spreadsheets/d/SHEET_ID/pub?output=csv"
   ```
   (paste the full CSV URL from Step 2)
4. Deploy — you get a permanent URL, e.g. `https://gladwells.streamlit.app`

---

## Monthly update (2 minutes)

1. Upload the new `.xlsm` to Google Drive and share it ("Anyone with the link → Viewer")
2. Copy its File ID from the share URL
3. Open the index Google Sheet and **update cell A1** with the new File ID (and B1 with the new filename)
4. Done — the dashboard picks it up automatically on next load

No code changes. No redeployment. No touching Streamlit secrets.

---

## What the index sheet looks like

| A | B |
|---|---|
| `1aBcDeFgHiJkLmNoPqRsTuVwXyZ` | `Gladwells_Feb26_Management_Accounts.xlsm` |

Just two cells. Update A1 each month with the new file's ID.

---

## Troubleshooting

| Error | Fix |
|-------|-----|
| "Index sheet is empty" | Make sure cell A1 in the Google Sheet has the file ID |
| "Could not read the index sheet" | Re-publish: File → Share → Publish to web → CSV → Publish |
| "Failed to parse workbook" | Confirm sheet names: `P&L_Data`, `BS_Data`, `Stats`, `Manual`, `Overview` |
| Shows old data | Click **🔄 Refresh data** in sidebar (cached 1 hour) |
| inotify error on Streamlit Cloud | `fileWatcherType = "none"` is already set in `.streamlit/config.toml` |

---

## Local development

Create `.streamlit/secrets.toml` (never commit this):
```toml
GDRIVE_FOLDER_ID = "https://docs.google.com/spreadsheets/d/SHEET_ID/pub?output=csv"
```
Then: `streamlit run app.py`
