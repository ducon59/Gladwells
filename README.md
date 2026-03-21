# 📊 Gladwells Management Accounts Dashboard

An interactive Streamlit dashboard that reads directly from your password-protected monthly management accounts file stored in Google Drive.

---

## How it works

Each month you receive a new `.xlsm` file. You simply **replace the file in Google Drive** and click **Refresh** in the sidebar — the dashboard updates automatically. No code changes needed.

---

## Quick start

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Set up Google Drive

1. Upload your management accounts `.xlsm` file to **Google Drive**
2. Right-click the file → **Share** → **Anyone with the link → Viewer**
3. Copy the share link — it looks like:
   ```
   https://drive.google.com/file/d/1aBcDeFgHiJkLmNoPqRsTuVwXyZ/view
   ```
4. Your **File ID** is the long string between `/d/` and `/view`:
   ```
   1aBcDeFgHiJkLmNoPqRsTuVwXyZ
   ```

### 3. Run locally

```bash
streamlit run app.py
```

Paste the File ID into the sidebar when prompted.

---

## Monthly update process

1. Download the new monthly `.xlsm` file from your accountant
2. In Google Drive, **right-click** the existing file → **Manage versions** → **Upload new version**  
   *(The File ID stays the same — the dashboard link never changes)*
3. Open the dashboard and click **🔄 Refresh data**

That's it. The dashboard will show the latest figures automatically.

---

## Deploying to Streamlit Cloud (free hosting)

1. Push this folder to a **GitHub repository**
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in
3. Click **New app** → select your repo → set `app.py` as the main file
4. Under **Advanced settings → Secrets**, add:
   ```toml
   GDRIVE_FILE_ID = "your_file_id_here"
   ```
5. Deploy — you'll get a permanent URL you can bookmark and share

### Setting the secret on Streamlit Cloud

In your app settings, add this to the Secrets section:
```toml
GDRIVE_FILE_ID = "1aBcDeFgHiJkLmNoPqRsTuVwXyZ"
```

This means the File ID is pre-loaded and you only need to click Refresh when new data arrives.

---

## Dashboard tabs

| Tab | Contents |
|-----|----------|
| 🏠 Overview | KPI cards, management commentary, 12-month trend chart |
| 📈 P&L Detail | Waterfall chart, monthly P&L table, overhead breakdown |
| 💰 Cash & Balance Sheet | Cash flow chart, balance sheet KPIs and trends |
| 🛒 Revenue Mix | Donut chart, category table, stacked bar trend |

---

## Troubleshooting

**"Could not fetch from Google Drive"**  
→ Make sure the file is shared as "Anyone with the link can view"  
→ Double-check the File ID (just the ID, not the full URL)

**"Failed to parse workbook"**  
→ The file structure may have changed. Check the sheet names match: `P&L_Data`, `BS_Data`, `Stats`, `Manual`, `Overview`, `Text`

**Dashboard shows old data after uploading new file**  
→ Click "🔄 Refresh data" in the sidebar (data is cached for 1 hour)

---

## File password

The file is automatically decrypted using the stored password. If the password changes, update `FILE_PASSWORD` at the top of `app.py`.
