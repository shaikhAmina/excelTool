"""
Feedback storage — appends submissions to a Google Sheet.

How it works:
  1. Reads credentials from st.secrets["gcp_service_account"]
  2. Opens the sheet named in st.secrets["feedback"]["sheet_name"]
  3. Appends a row: [timestamp, idea_text]

If the secrets are not configured (local dev, or not yet set up),
it fails silently so the app still works — the thank-you message
still shows, the submission just isn't stored.
"""

import datetime

import streamlit as st


def save_feedback(idea: str) -> bool:
    """
    Append one feedback row to Google Sheets.
    Returns True on success, False if skipped / failed.
    """
    # ── 1. Check secrets exist ────────────────────────
    try:
        creds_info = st.secrets["gcp_service_account"]
        sheet_name = st.secrets["feedback"]["sheet_name"]
    except (KeyError, FileNotFoundError):
        # Secrets not configured — silently skip storage
        return False

    # ── 2. Authenticate ───────────────────────────────
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(
            dict(creds_info),
            scopes=scopes,
        )
        client = gspread.authorize(creds)
    except Exception:
        return False

    # ── 3. Open sheet and append row ──────────────────
    try:
        sheet = client.open(sheet_name).sheet1

        # Add header row if the sheet is completely empty
        if sheet.row_count == 0 or sheet.cell(1, 1).value is None:
            sheet.append_row(
                ["Timestamp", "Feedback"],
                value_input_option="USER_ENTERED",
            )

        timestamp = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
        sheet.append_row(
            [timestamp, idea.strip()],
            value_input_option="USER_ENTERED",
        )
        return True

    except Exception:
        return False
