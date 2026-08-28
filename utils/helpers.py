"""
Shared UI helpers for Excel Toolbox.
Every tool imports from here to keep the UX consistent.
"""
import streamlit as st
import pandas as pd
from io import BytesIO


# ──────────────────────────────────────────────
# FILE INFO CARD
# ──────────────────────────────────────────────
def show_file_info(uploaded_file, df: pd.DataFrame):
    """Display a compact file-info card after upload."""
    size_kb = round(uploaded_file.size / 1024, 1)
    size_str = f"{size_kb} KB" if size_kb < 1024 else f"{round(size_kb/1024, 2)} MB"

    sheets = 1  # default; callers can pass sheet count separately if needed
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📄 File", uploaded_file.name[:24] + ("…" if len(uploaded_file.name) > 24 else ""))
    col2.metric("📏 Rows", f"{len(df):,}")
    col3.metric("📊 Columns", len(df.columns))
    col4.metric("💾 Size", size_str)


def show_file_info_multi(files: list):
    """Show a compact summary card for a list of uploaded files."""
    total_size = sum(f.size for f in files)
    size_kb = round(total_size / 1024, 1)
    size_str = f"{size_kb} KB" if size_kb < 1024 else f"{round(size_kb/1024, 2)} MB"
    col1, col2 = st.columns(2)
    col1.metric("📁 Files uploaded", len(files))
    col2.metric("💾 Total size", size_str)


# ──────────────────────────────────────────────
# DATA PREVIEW
# ──────────────────────────────────────────────
def show_data_preview(df: pd.DataFrame, n: int = 10, label: str = "Data Preview"):
    """Show first N rows with a collapsible expander."""
    with st.expander(f"👀 {label} — first {min(n, len(df))} of {len(df):,} rows", expanded=True):
        st.dataframe(df.head(n), use_container_width=True)


# ──────────────────────────────────────────────
# RESULT BANNER
# ──────────────────────────────────────────────
def result_banner(title: str, stats: dict):
    """
    Green success banner with key stats.
    stats = {"Files processed": 5, "Total rows": 12482}
    """
    st.markdown(
        f"""
        <div style="
            background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
            border: 1px solid #28a745;
            border-radius: 10px;
            padding: 16px 20px;
            margin: 12px 0 8px 0;
        ">
            <div style="font-size:1.15rem; font-weight:700; color:#155724;">
                ✅ {title}
            </div>
            {"".join(f'<div style="font-size:0.9rem; color:#155724; margin-top:4px;">• {k}: <strong>{v}</strong></div>' for k, v in stats.items())}
        </div>
        """,
        unsafe_allow_html=True,
    )


# ──────────────────────────────────────────────
# FRIENDLY ERROR MESSAGES
# ──────────────────────────────────────────────
_ERROR_MAP = {
    "No columns to parse": "The file appears to be empty or has no data rows. Please check the file and try again.",
    "Excel file format cannot be determined": "The file doesn't look like a valid Excel file (.xlsx). Please re-export and try again.",
    "Worksheet named": "The sheet name you selected doesn't exist in this file.",
    "JSONDecodeError": "The JSON is not valid. Check for missing commas, unclosed brackets, or unquoted keys.",
    "json.decoder": "The JSON is not valid. Check for missing commas, unclosed brackets, or unquoted keys.",
    "ConnectionError": "Could not reach the API. Check your internet connection and the URL.",
    "HTTPError": "The API returned an error. Check that the URL is correct and your auth token is valid.",
    "KeyError": "A required column was not found. Make sure you selected the right columns.",
    "MergeError": "The merge failed — the key columns may have different data types. Try enabling numeric conversion.",
    "Permission": "The file is open in another program (like Excel). Please close it and try again.",
    "UnicodeDecodeError": "The file contains characters that couldn't be read. Try saving it as UTF-8 first.",
    "openpyxl": "The Excel file may be corrupted or in an older .xls format. Please re-save as .xlsx.",
}


def friendly_error(e: Exception):
    """Map a Python exception to a human-readable Streamlit error message."""
    raw = str(e)
    for key, msg in _ERROR_MAP.items():
        if key.lower() in raw.lower():
            st.error(f"⚠️ {msg}")
            return
    # Fallback — show something readable but not the raw traceback
    st.error(f"⚠️ Something went wrong: {raw[:200]}")


# ──────────────────────────────────────────────
# BEFORE / AFTER PREVIEW
# ──────────────────────────────────────────────
def show_before_after(df_before: pd.DataFrame, df_after: pd.DataFrame, n: int = 5):
    """Side-by-side before / after preview for cleaning/conversion tools."""
    left, right = st.columns(2)
    with left:
        st.markdown("**Before**")
        st.dataframe(df_before.head(n), use_container_width=True)
    with right:
        st.markdown("**After**")
        st.dataframe(df_after.head(n), use_container_width=True)


# ──────────────────────────────────────────────
# EXCEL BYTES HELPER
# ──────────────────────────────────────────────
def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


# ──────────────────────────────────────────────
# PRIVACY NOTICE
# ──────────────────────────────────────────────
def privacy_notice():
    st.info(
        "🔒 **Privacy** — Files you upload are processed entirely in your browser session "
        "and are never stored on any server. They are discarded as soon as you close or "
        "refresh the tab.",
        icon=None,
    )


# ──────────────────────────────────────────────
# FEEDBACK WIDGET
# ──────────────────────────────────────────────
def feedback_widget():
    from utils.feedback import save_feedback

    st.markdown("---")
    with st.expander("💬 Need another Excel tool? Tell us!", expanded=False):
        idea = st.text_area(
            "Describe the tool you need",
            placeholder="e.g. 'A tool that removes blank rows and trims whitespace from all cells'",
            height=100,
            key="feedback_idea",
        )
        if st.button("Submit idea", key="feedback_submit"):
            if idea.strip():
                saved = save_feedback(idea)
                if saved:
                    st.success("Thanks! Your idea has been saved. 🙌")
                else:
                    st.success("Thanks! We've noted your idea. 🙌")
            else:
                st.warning("Please describe your idea before submitting.")
