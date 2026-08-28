"""
Feedback storage — appends submissions to a Google Sheet.
"""

import datetime
import streamlit as st


def save_feedback(idea: str) -> bool:
    """
    Append one feedback row to Google Sheets.
    Returns True on success, False + shows error detail if something fails.
    """
    # ── 1. Check secrets exist ────────────────────────
    try:
        creds_info = st.secrets["gcp_service_account"]
        sheet_name = st.secrets["feedback"]["sheet_name"]
    except (KeyError, FileNotFoundError) as e:
        st.warning(f"⚠️ Secrets not found: {e}")
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
    except Exception as e:
        st.warning(f"⚠️ Google auth failed: {e}")
        return False

    # ── 3. Open sheet and append row ──────────────────
    try:
        sheet = client.open(sheet_name).sheet1

        # Add header row if the sheet is completely empty
        if sheet.cell(1, 1).value is None:
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

    except Exception as e:
        st.warning(f"⚠️ Could not write to sheet: {e}")
        return False
