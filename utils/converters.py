import streamlit as st
import pandas as pd
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
# NUMBER FORMATTER  (was "Convert to Number")
# ─────────────────────────────────────────────────────
def convert_to_number_ui():
    st.markdown("### 🔢 Number Formatter")
    st.caption(
        "Convert text-formatted numbers to real numeric values so Excel formulas, "
        "filters, and pivot tables work correctly."
    )
    privacy_notice()
    st.markdown("---")

    # ── STEP 1: Upload ────────────────────────────────
    st.markdown("#### Step 1 — Upload")
    file = st.file_uploader("Upload an Excel file (.xlsx)", type=["xlsx"], key="num_upload")
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

    # ── STEP 2: Configure ─────────────────────────────
    st.markdown("#### Step 2 — Configure")

    # Auto-detect likely text-numeric columns
    suspect_cols = [
        c for c in df.columns
        if df[c].dtype == object and pd.to_numeric(df[c], errors="coerce").notna().mean() > 0.5
    ]
    default_sel = suspect_cols if suspect_cols else []

    columns = st.multiselect(
        "Select columns to convert to numbers",
        list(df.columns),
        default=default_sel,
        help="Columns are auto-suggested if they look like they contain numbers stored as text.",
    )
    if suspect_cols:
        st.caption(f"💡 Auto-detected likely text-number columns: **{', '.join(suspect_cols)}**")

    on_error = st.radio(
        "What to do with values that can't be converted?",
        ["Leave as-is", "Replace with blank (NaN)"],
        horizontal=True,
    )
    errors_mode = "ignore" if on_error == "Leave as-is" else "coerce"

    if not columns:
        st.warning("Select at least one column above.")
        feedback_widget()
        return

    # ── STEP 3: Preview ───────────────────────────────
    st.markdown("#### Step 3 — Preview")
    show_data_preview(df, label="Source file")

    # ── STEP 4: Process ───────────────────────────────
    st.markdown("#### Step 4 — Process")
    if st.button("🔢 Convert Columns", type="primary"):
        try:
            df_before = df[columns].copy()
            df_out = df.copy()
            converted = []
            skipped = []

            for col in columns:
                new_col = pd.to_numeric(df_out[col], errors=errors_mode)
                if new_col.dtype != df_out[col].dtype:
                    df_out[col] = new_col
                    converted.append(col)
                else:
                    skipped.append(col)

            df_after = df_out[columns].copy()

            # ── STEP 5: Result ────────────────────────
            result_banner(
                "Conversion complete",
                {
                    "Columns converted": ", ".join(converted) if converted else "none",
                    "Columns unchanged": ", ".join(skipped) if skipped else "none",
                },
            )

            st.markdown("**Before / After**")
            show_before_after(df_before, df_after)

            # ── STEP 6: Download ──────────────────────
            st.download_button(
                "⬇️ Download Converted Excel",
                df_to_excel_bytes(df_out),
                file_name=f"numbers_{file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            friendly_error(e)

    feedback_widget()


# ─────────────────────────────────────────────────────
# COMMA / QUOTE STRING FORMATTER
# ─────────────────────────────────────────────────────
def comma_string_formatter_ui():
    st.markdown("### 🧩 String List Formatter")
    st.caption(
        "Paste a list of values and instantly format them as a comma-separated string — "
        "perfect for SQL IN clauses, spreadsheet filters, or API calls."
    )
    st.markdown("---")

    # ── STEP 1: Input ─────────────────────────────────
    st.markdown("#### Step 1 — Enter Values")
    input_text = st.text_area(
        "One value per line, or comma-separated",
        placeholder="12345\n67890\nJohn\nDoe\n\n— or —\n\n12345, 67890, John, Doe",
        height=180,
        key="fmt_input",
    )

    # ── STEP 2: Configure ─────────────────────────────
    st.markdown("#### Step 2 — Choose Format")
    output_type = st.radio(
        "Output format",
        [
            "Plain  →  A, B, C",
            "Single quotes  →  'A', 'B', 'C'",
            "Double quotes  →  \"A\", \"B\", \"C\"",
            "Parentheses  →  (A, B, C)",
            "SQL IN clause  →  IN ('A', 'B', 'C')",
        ],
        key="fmt_type",
    )
    trim_values = st.checkbox("Trim whitespace from each value", value=True)

    # ── STEP 3: Process ───────────────────────────────
    st.markdown("#### Step 3 — Convert")
    if st.button("🔄 Format", type="primary"):
        if not input_text.strip():
            st.warning("Please enter some values first.")
            return

        raw_items = input_text.replace(",", "\n").split("\n")
        items = [x.strip() if trim_values else x for x in raw_items if x.strip()]

        if not items:
            st.warning("No valid values found. Please check your input.")
            return

        # Build formatted string
        if "Single quotes" in output_type:
            formatted = ", ".join(f"'{x}'" for x in items)
        elif "Double quotes" in output_type:
            formatted = ", ".join(f'"{x}"' for x in items)
        elif "Parentheses" in output_type:
            formatted = f"({', '.join(items)})"
        elif "SQL IN" in output_type:
            formatted = "IN (" + ", ".join(f"'{x}'" for x in items) + ")"
        else:
            formatted = ", ".join(items)

        # ── STEP 4: Result ────────────────────────────
        result_banner(
            "Formatting complete",
            {"Values processed": len(items), "Output length": f"{len(formatted)} chars"},
        )
        st.code(formatted, language="text")

        try:
            import pyperclip
            pyperclip.copy(formatted)
            st.success("📋 Copied to clipboard!")
        except Exception:
            st.info("💡 Select the text above and copy it manually (Ctrl+C / Cmd+C).")

        # ── STEP 5: Download ──────────────────────────
        st.download_button(
            "⬇️ Download as .txt",
            formatted,
            file_name="formatted_values.txt",
            mime="text/plain",
        )

    feedback_widget()
