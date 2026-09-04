# Daily Work Archive — v2

A personal work archive that stores data in **Google Sheets** and provides a simple UI for daily entry, search, editing, deletion, and reporting.

Authentication uses **your Google account (OAuth)** — no service account key.

## Features

- Daily entry form based on the supplied template
- Existing values automatically loaded when editing a day
- Search by date range / category / keyword
- Edit or delete an entire day
- Quick report ranges + Detailed / Executive styles
- Markdown, CSV, and Excel snapshot downloads
- Persistent storage via Google Sheets (works with Streamlit Cloud)

## Google setup (OAuth — no service account key)

### 1. Create the Sheet
1. Create a Google Sheet you own (or can edit).
2. Copy the spreadsheet ID from the URL:
   `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit`

### 2. Google Cloud project + APIs
1. Open [Google Cloud Console](https://console.cloud.google.com/) and create/select a project.
2. Enable **Google Sheets API** (APIs & Services → Library).
3. Configure the **OAuth consent screen**:
   - For a work Google account, prefer **Internal** if available.
   - App name: e.g. Daily Work Archive
   - Add your email as a test user if the app is in Testing / External mode.
4. Create credentials: **APIs & Services → Credentials → Create credentials → OAuth client ID**
   - Application type: **Web application**
   - Authorized redirect URIs:
     - Local: `http://localhost:8501/`
     - Streamlit Cloud (later): `https://YOUR-APP.streamlit.app/`
5. Copy the **Client ID** and **Client secret**.

### 3. App secrets
```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Fill in:
- `spreadsheet_id`
- `google_oauth.client_id`
- `google_oauth.client_secret`
- `google_oauth.redirect_uri` (`http://localhost:8501/` locally)

### 4. Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

Click **Sign in with Google**, approve Sheets access, and continue.

Tokens are saved locally to `.streamlit/google_token.json` (gitignored) so you are not prompted every restart.

### Streamlit Cloud notes
1. Add the Cloud app URL as an Authorized redirect URI on the OAuth client.
2. Set the same values in **App settings → Secrets**, including `redirect_uri` for your Cloud URL.
3. Optional: after one successful local login, copy `refresh_token` from `.streamlit/google_token.json` into secrets as `google_oauth.refresh_token` so Cloud can reuse it without clicking Sign in every cold start.

## Data model

Worksheet name: **Daily Log**

| Date | Category | Field | Details | Created At | Updated At |
|---|---|---|---|---|---|

## Important behavior

- Saving a date replaces that date's previous form values. Blank fields are removed.
- You cannot create a second entry for a day that already exists; edit or delete it first.
- At least one field must be filled to save/update.
- Google Sheets is the source of truth.
