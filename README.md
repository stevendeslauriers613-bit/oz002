# OZ Store Cloud Accountant Package

Streamlit app that stores all data in Dropbox so multiple users can share the same state on Streamlit Cloud.

## Setup

1. Create a Dropbox app and generate an access token.
2. Add the token **only** in Streamlit secrets (never commit it):

   **Streamlit Cloud → Settings → Secrets:**
   ```toml
   DROPBOX_ACCESS_TOKEN = "sl.u.your_token_here"
   ```

   **Locally:** create `.streamlit/secrets.toml` with the same line.

3. Deploy this repo on Streamlit Cloud (or run locally with `streamlit run oz_store_cloud.py`).

## Files
- `oz_store_cloud.py` — main app
- `requirements.txt` — dependencies

Data is stored in Dropbox under `/OZ_Store_Uploads/master_store_state.json`.
