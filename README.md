# 📊 Gladwells Management Accounts Dashboard

An always-on Streamlit dashboard that automatically loads the **most recently added file** from a shared Google Drive folder. No file IDs to paste, no links to share — just drop the new monthly file into the folder and the dashboard updates itself.

---

## How it works

- You maintain one **shared Google Drive folder**
- Every month, drop the new `.xlsm` file into it (old files stay — nothing gets overwritten)
- The dashboard automatically detects and loads the **most recently modified file** in the folder
- Users just open the permanent dashboard URL — no setup required on their end

---

## One-time setup

### 1. Create the shared Google Drive folder

1. In Google Drive, create a folder — e.g. `Gladwells Management Accounts`
2. Right-click → **Share** → **Anyone with the link → Viewer**
3. Copy the folder URL — it looks like:
   ```
   https://drive.google.com/drive/folders/1aBcDeFgHiJkLmNoPqRsTuVwXyZ
   ```
4. Your **Folder ID** is the string at the end after `/folders/`
5. Upload the first `.xlsm` file into this folder

### 2. Deploy to Streamlit Cloud (free, permanent URL)

1. Push this folder to a **GitHub repository**
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app** → select your repo → set `app.py` as main file
3. Under **Advanced settings → Secrets**, add:
   ```toml
   GDRIVE_FOLDER_ID = "your_folder_id_here"
   ```
4. Deploy — you get a permanent URL to bookmark and share

### 3. Local development

Create `.streamlit/secrets.toml`:
```toml
GDRIVE_FOLDER_ID = "your_folder_id_here"
```
Then run: `streamlit run app.py`

---

## Monthly update process (30 seconds)

1. Receive the new `.xlsm` file from your accountant
2. Upload it to the shared Google Drive folder
3. Done — the dashboard picks it up automatically on next load

No renaming, no deleting old files, no code changes needed.

---

## Troubleshooting

| Error | Fix |
|-------|-----|
| "No spreadsheet files found" | Upload at least one .xlsm into the folder |
| "Drive API returned 403" | Share the folder as "Anyone with the link → Viewer" |
| "Failed to parse workbook" | Check sheet names: P&L_Data, BS_Data, Stats, Manual, Overview |
| Shows old data | Click "🔄 Check for new file" in sidebar (cached 1 hour) |
