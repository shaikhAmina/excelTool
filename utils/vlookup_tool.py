import streamlit as st
import pandas as pd

from utils.helpers import (
    show_file_info,
    show_data_preview,
    result_banner,
    friendly_error,
    df_to_excel_bytes,
    privacy_notice,
    feedback_widget,
)


def vlookup_tool_ui():
    st.markdown("### 🔗 VLOOKUP")
    st.caption(
        "Match rows across two spreadsheets by a common key column and pull in extra columns — "
        "just like Excel's VLOOKUP, but without formula limits."
    )
    privacy_notice()
    st.markdown("---")

    # ── STEP 1: Upload ────────────────────────────────
    st.markdown("#### Step 1 — Upload")
    col_a, col_b = st.columns(2)
    with col_a:
        file_main = st.file_uploader(
            "Main file (the one you want to enrich)",
            type=["xlsx"],
            key="vl_main",
        )
    with col_b:
        file_lookup = st.file_uploader(
            "Lookup / reference file",
            type=["xlsx"],
            key="vl_lookup",
        )

    if not file_main or not file_lookup:
        st.info("Upload both files above to continue.")
        feedback_widget()
        return

    try:
        df_main = pd.read_excel(file_main)
        df_lookup = pd.read_excel(file_lookup)
    except Exception as e:
        friendly_error(e)
        return

    # File info cards
    col_a2, col_b2 = st.columns(2)
    with col_a2:
        st.markdown(f"**Main:** `{file_main.name}`")
        show_file_info(file_main, df_main)
    with col_b2:
        st.markdown(f"**Lookup:** `{file_lookup.name}`")
        show_file_info(file_lookup, df_lookup)

    # ── STEP 2: Configure ─────────────────────────────
    st.markdown("#### Step 2 — Configure")
    c1, c2 = st.columns(2)
    with c1:
        main_key = st.selectbox(
            "Key column in Main file",
            df_main.columns,
            help="This is the column that links the two files (e.g. ID, Order number).",
        )
    with c2:
        lookup_key = st.selectbox(
            "Matching column in Lookup file",
            df_lookup.columns,
        )

    lookup_values = st.multiselect(
        "Columns to fetch from the Lookup file",
        [c for c in df_lookup.columns if c != lookup_key],
        help="These columns will be appended to your Main file.",
    )

    convert_numbers = st.checkbox(
        "Convert key columns to numeric before matching",
        value=False,
        help="Useful when one file stores IDs as text and the other as numbers.",
    )

    join_type = st.radio(
        "Join type",
        ["Left join (keep all main rows)", "Inner join (only matched rows)"],
        horizontal=True,
    )
    how = "left" if "Left" in join_type else "inner"

    if not lookup_values:
        st.warning("Select at least one column to fetch from the Lookup file.")
        feedback_widget()
        return

    # ── STEP 3: Preview ───────────────────────────────
    st.markdown("#### Step 3 — Preview")
    tab1, tab2 = st.tabs([f"Main — {file_main.name}", f"Lookup — {file_lookup.name}"])
    with tab1:
        show_data_preview(df_main, label="Main file")
    with tab2:
        show_data_preview(df_lookup, label="Lookup file")

    # ── STEP 4: Process ───────────────────────────────
    st.markdown("#### Step 4 — Process")
    if st.button("🔗 Run VLOOKUP", type="primary"):
        try:
            df_m = df_main.copy()
            df_l = df_lookup.copy()

            if convert_numbers:
                df_m[main_key] = pd.to_numeric(df_m[main_key], errors="coerce")
                df_l[lookup_key] = pd.to_numeric(df_l[lookup_key], errors="coerce")

            original_cols = set(df_m.columns)
            merged = pd.merge(
                df_m,
                df_l[[lookup_key] + lookup_values].drop_duplicates(subset=[lookup_key]),
                left_on=main_key,
                right_on=lookup_key,
                how=how,
            )
            if lookup_key != main_key:
                merged.drop(columns=[lookup_key], inplace=True, errors="ignore")

            new_cols = [c for c in merged.columns if c not in original_cols]
            matched = int(merged[new_cols[0]].notna().sum()) if new_cols else 0
            unmatched = len(merged) - matched

            # ── STEP 5: Result ────────────────────────
            result_banner(
                "VLOOKUP completed",
                {
                    "Rows in result": f"{len(merged):,}",
                    "Matched rows": f"{matched:,}",
                    "Unmatched rows": f"{unmatched:,}",
                    "Columns added": ", ".join(new_cols) if new_cols else "none",
                },
            )

            # Highlight new columns in preview
            def _highlight(col):
                return ["background-color: #fff3cd"] * len(col) if col.name in new_cols else [""] * len(col)

            with st.expander("👀 Result preview — new columns highlighted", expanded=True):
                st.dataframe(merged.head(10).style.apply(_highlight), use_container_width=True)

            # ── STEP 6: Download ──────────────────────
            out_name = f"vlookup_{file_main.name.replace('.xlsx','')}.xlsx"
            st.download_button(
                "⬇️ Download Result Excel",
                df_to_excel_bytes(merged),
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            friendly_error(e)

    feedback_widget()
