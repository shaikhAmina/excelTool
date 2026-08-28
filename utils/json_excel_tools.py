import streamlit as st
import pandas as pd
import ujson as json
import requests
from io import BytesIO

from utils.helpers import (
    show_file_info,
    show_data_preview,
    show_before_after,
    result_banner,
    friendly_error,
    df_to_excel_bytes,
    privacy_notice,
    feedback_widget,
)


# ─────────────────────────────────────────────────────
# SHARED HELPER: parse raw data → DataFrame
# ─────────────────────────────────────────────────────
def _data_to_df(data) -> pd.DataFrame:
    """Normalise a JSON value (list / dict / nested) into a flat DataFrame."""
    if isinstance(data, dict):
        # Try to unwrap {"data": [...]} or {"results": [...]} etc.
        list_val = next((v for v in data.values() if isinstance(v, list)), None)
        if list_val is not None:
            data = list_val
        else:
            data = [data]
    if not isinstance(data, list):
        raise ValueError("Cannot flatten this JSON structure into a table.")
    return pd.json_normalize(data)


# ─────────────────────────────────────────────────────
# MAIN UI DISPATCHER
# ─────────────────────────────────────────────────────
def json_excel_tool_ui():
    st.markdown("### 🔄 JSON ↔ Excel Converter")
    st.caption(
        "Convert JSON files to Excel, export Excel data as JSON, "
        "pull live API data into a spreadsheet, or convert a text file containing JSON."
    )
    privacy_notice()
    st.markdown("---")

    tool = st.radio(
        "What would you like to do?",
        ["JSON → Excel", "Excel → JSON", "API → Excel", "Text File → Excel"],
        horizontal=True,
        key="je_mode",
    )

    if tool == "JSON → Excel":
        _json_to_excel_ui()
    elif tool == "Excel → JSON":
        _excel_to_json_ui()
    elif tool == "API → Excel":
        _api_to_excel_ui()
    else:
        _text_to_excel_ui()


# ─────────────────────────────────────────────────────
# JSON → EXCEL
# ─────────────────────────────────────────────────────
def _json_to_excel_ui():
    st.markdown("#### Step 1 — Upload or Paste JSON")

    uploaded = st.file_uploader("Upload a JSON file", type=["json"], key="je_json_up")
    json_text = st.text_area(
        "Or paste JSON directly",
        height=160,
        placeholder='[{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}]',
        key="je_json_text",
    )

    if not uploaded and not json_text.strip():
        st.info("Upload a .json file or paste JSON above to get started.")
        feedback_widget()
        return

    st.markdown("#### Step 2 — Configure")
    sheet_name = st.text_input("Sheet name in output Excel", value="Data", key="je_sheet")

    st.markdown("#### Step 3 — Preview")
    try:
        if uploaded:
            raw = uploaded.read()
            data = json.loads(raw)
        else:
            data = json.loads(json_text)

        df = _data_to_df(data)
    except Exception as e:
        friendly_error(e)
        return

    show_data_preview(df, label=f"{len(df):,} rows, {len(df.columns)} columns")

    st.markdown("#### Step 4 — Process")
    if st.button("🔄 Convert to Excel", type="primary", key="je_btn_je"):
        try:
            result_banner(
                "Conversion complete",
                {
                    "Rows": f"{len(df):,}",
                    "Columns": len(df.columns),
                    "Sheet": sheet_name,
                },
            )
            st.download_button(
                "⬇️ Download Excel",
                df_to_excel_bytes(df, sheet_name),
                file_name="converted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            friendly_error(e)

    feedback_widget()


# ─────────────────────────────────────────────────────
# EXCEL → JSON
# ─────────────────────────────────────────────────────
def _excel_to_json_ui():
    st.markdown("#### Step 1 — Upload")
    file = st.file_uploader("Upload an Excel file (.xlsx)", type=["xlsx"], key="je_xl_up")
    if not file:
        st.info("Upload an .xlsx file to get started.")
        feedback_widget()
        return

    try:
        df = pd.read_excel(file)
    except Exception as e:
        friendly_error(e)
        return

    show_file_info(file, df)

    st.markdown("#### Step 2 — Configure")
    orient = st.radio(
        "JSON structure",
        ["records  →  [{col: val, …}]", "columns  →  {col: [val, …]}"],
        horizontal=True,
        key="je_orient",
    )
    orient_val = "records" if "records" in orient else "columns"
    indent = st.slider("Indentation (spaces)", 0, 4, 2, key="je_indent")
    ensure_ascii = st.checkbox("Escape non-ASCII characters", value=False, key="je_ascii")

    st.markdown("#### Step 3 — Preview")
    show_data_preview(df)

    st.markdown("#### Step 4 — Process")
    if st.button("🔄 Convert to JSON", type="primary", key="je_btn_ej"):
        try:
            json_data = df.to_dict(orient=orient_val)
            json_str = json.dumps(json_data, indent=indent if indent > 0 else None)

            result_banner(
                "Conversion complete",
                {
                    "Source rows": f"{len(df):,}",
                    "Source columns": len(df.columns),
                    "Output size": f"{len(json_str):,} chars",
                },
            )

            # Show a capped preview
            preview_chars = 1500
            with st.expander("👀 JSON preview", expanded=True):
                st.code(json_str[:preview_chars] + ("…" if len(json_str) > preview_chars else ""), language="json")

            st.download_button(
                "⬇️ Download JSON",
                json_str,
                file_name=f"{file.name.replace('.xlsx', '')}.json",
                mime="application/json",
            )
        except Exception as e:
            friendly_error(e)

    feedback_widget()


# ─────────────────────────────────────────────────────
# API → EXCEL
# ─────────────────────────────────────────────────────
def _api_to_excel_ui():
    st.markdown("#### Step 1 — Configure API Request")

    api_url = st.text_input(
        "API endpoint URL (GET only)",
        placeholder="https://api.example.com/v1/users",
        key="je_api_url",
    )
    auth_header = st.text_input(
        "Authorization header value (optional)",
        placeholder="Bearer eyJhbGci…",
        type="password",
        key="je_api_auth",
    )
    extra_headers_raw = st.text_area(
        "Additional headers — one per line as  Key: Value  (optional)",
        placeholder="X-Api-Version: 2\nAccept: application/json",
        height=80,
        key="je_api_headers",
    )
    timeout = st.slider("Request timeout (seconds)", 5, 60, 15, key="je_api_timeout")

    if not api_url.strip():
        st.info("Enter an API URL above.")
        feedback_widget()
        return

    st.markdown("#### Step 2 — Process")
    if st.button("🌐 Fetch & Convert", type="primary", key="je_btn_api"):
        try:
            headers = {"Accept": "application/json"}
            if auth_header.strip():
                headers["Authorization"] = auth_header.strip()
            for line in extra_headers_raw.strip().splitlines():
                if ":" in line:
                    k, v = line.split(":", 1)
                    headers[k.strip()] = v.strip()

            with st.spinner("Fetching data…"):
                resp = requests.get(api_url.strip(), headers=headers, timeout=timeout)
                resp.raise_for_status()
                data = resp.json()

            df = _data_to_df(data)

            result_banner(
                "API fetch complete",
                {
                    "Rows fetched": f"{len(df):,}",
                    "Columns": len(df.columns),
                    "Status": resp.status_code,
                },
            )
            show_data_preview(df, label="API data preview")

            st.download_button(
                "⬇️ Download Excel",
                df_to_excel_bytes(df, "API_Data"),
                file_name="api_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except requests.exceptions.ConnectionError:
            st.error("⚠️ Could not connect to that URL. Check your internet connection and the address.")
        except requests.exceptions.HTTPError as e:
            st.error(f"⚠️ The API returned an error: {e.response.status_code} {e.response.reason}. Check your URL and auth token.")
        except requests.exceptions.Timeout:
            st.error(f"⚠️ The request timed out after {timeout}s. Try increasing the timeout or check if the API is reachable.")
        except Exception as e:
            friendly_error(e)

    feedback_widget()


# ─────────────────────────────────────────────────────
# TEXT FILE → EXCEL
# ─────────────────────────────────────────────────────
def _text_to_excel_ui():
    st.markdown("#### Step 1 — Upload")
    txt_file = st.file_uploader(
        "Upload a .txt or .json file containing JSON data",
        type=["txt", "json"],
        key="je_txt_up",
    )
    if not txt_file:
        st.info("Upload a text file that contains JSON array or object data.")
        feedback_widget()
        return

    st.markdown("#### Step 2 — Configure")
    encoding = st.selectbox("File encoding", ["utf-8", "utf-8-sig", "latin-1", "cp1252"], key="je_enc")
    sheet_name = st.text_input("Sheet name", value="Data", key="je_txt_sheet")

    st.markdown("#### Step 3 — Preview & Process")
    if st.button("🔄 Convert to Excel", type="primary", key="je_btn_txt"):
        try:
            raw = txt_file.read().decode(encoding).strip()
            data = json.loads(raw)
            df = _data_to_df(data)

            result_banner(
                "Conversion complete",
                {
                    "Rows": f"{len(df):,}",
                    "Columns": len(df.columns),
                    "Source file": txt_file.name,
                },
            )
            show_data_preview(df)

            st.download_button(
                "⬇️ Download Excel",
                df_to_excel_bytes(df, sheet_name),
                file_name=f"{txt_file.name.rsplit('.', 1)[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except UnicodeDecodeError:
            st.error(f"⚠️ Could not read the file with {encoding} encoding. Try a different encoding option above.")
        except Exception as e:
            friendly_error(e)

    feedback_widget()
