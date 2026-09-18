import streamlit as st
import pandas as pd

from io import BytesIO

from utils.data_workflow import data_workflow_ui

from utils.helpers import (
    show_file_info,
    show_data_preview,
    result_banner,
    friendly_error,
    df_to_excel_bytes,
    privacy_notice,
    feedback_widget,
)


# =========================================================
# HELPER — CREATE MULTI-SHEET EXCEL
# =========================================================

def dataframes_to_excel_bytes(dataframes: dict) -> bytes:
    """
    Convert multiple DataFrames into one Excel file.

    Example:
        {
            "Sheet1": df1,
            "Sheet2": df2
        }
    """

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        for sheet_name, df in dataframes.items():

            # Excel sheet names cannot exceed 31 characters
            safe_sheet_name = str(sheet_name)[:31]

            # Excel does not allow some special characters
            for char in ["\\", "/", "*", "?", ":", "[", "]"]:
                safe_sheet_name = safe_sheet_name.replace(
                    char,
                    "_",
                )

            if not safe_sheet_name:
                safe_sheet_name = "Sheet"

            # Prevent duplicate sheet names
            original_name = safe_sheet_name
            counter = 1

            while safe_sheet_name in [
                ws.title for ws in writer.book.worksheets
            ]:

                suffix = f"_{counter}"

                safe_sheet_name = (
                    original_name[: 31 - len(suffix)]
                    + suffix
                )

                counter += 1

            df.to_excel(
                writer,
                sheet_name=safe_sheet_name,
                index=False,
            )

    output.seek(0)

    return output.getvalue()


# =========================================================
# HELPER — SAFE COLUMN NAME
# =========================================================

def get_safe_lookup_column_name(
    column_name: str,
    existing_columns: set,
) -> str:
    """
    If lookup column already exists in Main,
    rename it to:

        amount_lookup

    If that already exists:

        amount_lookup_1
        amount_lookup_2
        ...
    """

    if column_name not in existing_columns:
        return column_name

    new_name = f"{column_name}_lookup"

    counter = 1

    while new_name in existing_columns:

        new_name = (
            f"{column_name}_lookup_{counter}"
        )

        counter += 1

    return new_name


# =========================================================
# HELPER — PREPARE LOOKUP DATA
# =========================================================

def prepare_lookup_dataframe(
    df_lookup: pd.DataFrame,
    lookup_key: str,
    lookup_values: list,
    main_columns: list,
    convert_numbers: bool,
):
    """
    Prepare lookup dataframe before merge.
    Handles duplicate/conflicting column names.
    """

    df_l = df_lookup.copy()

    # -----------------------------------------------------
    # Normalize column names
    # -----------------------------------------------------

    df_l.columns = [
        str(col).strip()
        for col in df_l.columns
    ]

    # -----------------------------------------------------
    # Convert lookup key
    # -----------------------------------------------------

    if convert_numbers:

        df_l[lookup_key] = pd.to_numeric(
            df_l[lookup_key],
            errors="coerce",
        )

    # -----------------------------------------------------
    # Select required columns
    # -----------------------------------------------------

    lookup_columns = [
        lookup_key
    ] + lookup_values

    df_lookup_selected = df_l[
        lookup_columns
    ].copy()

    # -----------------------------------------------------
    # Rename conflicting columns
    # -----------------------------------------------------

    rename_map = {}

    existing_columns = set(
        main_columns
    )

    # Lookup key itself can be same as Main key.
    # It is handled by merge later.

    for column in lookup_values:

        new_name = get_safe_lookup_column_name(
            column,
            existing_columns,
        )

        if new_name != column:

            rename_map[column] = new_name

        existing_columns.add(new_name)

    if rename_map:

        df_lookup_selected = (
            df_lookup_selected.rename(
                columns=rename_map
            )
        )

    # -----------------------------------------------------
    # Remove duplicate lookup keys
    # -----------------------------------------------------

    df_lookup_selected = (
        df_lookup_selected.drop_duplicates(
            subset=[lookup_key],
            keep="first",
        )
    )

    return (
        df_lookup_selected,
        rename_map,
    )


# =========================================================
# MAIN VLOOKUP UI
# =========================================================

def vlookup_tool_ui():

    st.markdown("### 🔗 VLOOKUP")

    st.caption(
        "Match rows across two spreadsheets by a common key column "
        "and pull in extra columns — just like Excel's VLOOKUP, "
        "but without formula limits."
    )

    privacy_notice()

    st.markdown("---")

    # =====================================================
    # MODE
    # =====================================================

    mode = st.radio(
        "Choose a lookup style",
        [
            "Simple lookup",
            "Custom workflow — concatenate, convert, then lookup",
        ],
        horizontal=True,
        help=(
            "Use Custom workflow when you need to prepare "
            "columns before matching the files."
        ),
        key="vlookup_mode",
    )

    if mode.startswith("Custom workflow"):

        st.info(
            "Build the steps in the order you need. "
            "For concatenation, select the source columns "
            "and enter a separator such as `-`, a space, "
            "or leave it blank to join directly."
        )

        data_workflow_ui()

        return

    # =====================================================
    # STEP 1 — UPLOAD
    # =====================================================

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

        st.info(
            "Upload both files above to continue."
        )

        feedback_widget()

        return

    # =====================================================
    # READ ALL SHEETS
    # =====================================================

    try:

        main_sheets = pd.read_excel(
            file_main,
            sheet_name=None,
        )

        lookup_sheets = pd.read_excel(
            file_lookup,
            sheet_name=None,
        )

    except Exception as e:

        friendly_error(e)

        return

    main_sheet_names = list(
        main_sheets.keys()
    )

    lookup_sheet_names = list(
        lookup_sheets.keys()
    )

    # =====================================================
    # STEP 1A — MAIN SHEET SELECTION
    # =====================================================

    st.markdown("#### 📄 Select Main File Sheets")

    if len(main_sheet_names) == 1:

        main_sheet_mode = "One sheet"

        selected_main_sheets = [
            main_sheet_names[0]
        ]

        st.info(
            f"Main file contains one sheet: "
            f"**{main_sheet_names[0]}**"
        )

    else:

        main_sheet_mode = st.radio(
            "How do you want to process the Main file?",
            [
                "One sheet",
                "Multiple sheets",
                "All sheets",
            ],
            horizontal=True,
            key="vl_main_sheet_mode",
        )

        # -------------------------------------------------
        # ONE MAIN SHEET
        # -------------------------------------------------

        if main_sheet_mode == "One sheet":

            selected_main_sheet = st.selectbox(
                "Main file sheet",
                main_sheet_names,
                key="vl_main_sheet_single",
            )

            selected_main_sheets = [
                selected_main_sheet
            ]

        # -------------------------------------------------
        # MULTIPLE MAIN SHEETS
        # -------------------------------------------------

        elif main_sheet_mode == "Multiple sheets":

            selected_main_sheets = st.multiselect(
                "Select Main sheets",
                main_sheet_names,
                default=main_sheet_names,
                key="vl_main_sheets_multiple",
                help=(
                    "Each selected Main sheet will be processed "
                    "separately and will remain a separate sheet "
                    "in the result Excel."
                ),
            )

        # -------------------------------------------------
        # ALL MAIN SHEETS
        # -------------------------------------------------

        else:

            selected_main_sheets = (
                main_sheet_names.copy()
            )

            st.success(
                f"All {len(selected_main_sheets)} Main sheets "
                "will be processed separately."
            )

    if not selected_main_sheets:

        st.warning(
            "Please select at least one Main sheet."
        )

        feedback_widget()

        return

    # =====================================================
    # STEP 1B — LOOKUP SHEET SELECTION
    # =====================================================

    st.markdown("#### 🔎 Select Lookup File Sheets")

    if len(lookup_sheet_names) == 1:

        lookup_sheet_mode = "One sheet"

        selected_lookup_sheets = [
            lookup_sheet_names[0]
        ]

        st.info(
            f"Lookup file contains one sheet: "
            f"**{lookup_sheet_names[0]}**"
        )

    else:

        lookup_sheet_mode = st.radio(
            "How do you want to use the Lookup file?",
            [
                "One sheet",
                "Multiple sheets",
                "All sheets",
            ],
            horizontal=True,
            key="vl_lookup_sheet_mode",
        )

        # -------------------------------------------------
        # ONE LOOKUP SHEET
        # -------------------------------------------------

        if lookup_sheet_mode == "One sheet":

            selected_lookup_sheet = st.selectbox(
                "Lookup file sheet",
                lookup_sheet_names,
                key="vl_lookup_sheet_single",
            )

            selected_lookup_sheets = [
                selected_lookup_sheet
            ]

        # -------------------------------------------------
        # MULTIPLE LOOKUP SHEETS
        # -------------------------------------------------

        elif lookup_sheet_mode == "Multiple sheets":

            selected_lookup_sheets = st.multiselect(
                "Select Lookup sheets",
                lookup_sheet_names,
                default=lookup_sheet_names,
                key="vl_lookup_sheets_multiple",
                help=(
                    "Selected Lookup sheets will be combined "
                    "into one lookup table."
                ),
            )

        # -------------------------------------------------
        # ALL LOOKUP SHEETS
        # -------------------------------------------------

        else:

            selected_lookup_sheets = (
                lookup_sheet_names.copy()
            )

            st.success(
                f"All {len(selected_lookup_sheets)} Lookup sheets "
                "will be combined for lookup."
            )

    if not selected_lookup_sheets:

        st.warning(
            "Please select at least one Lookup sheet."
        )

        feedback_widget()

        return

    # =====================================================
    # CREATE COMBINED LOOKUP DATA
    # =====================================================

    lookup_frames = []

    for sheet_name in selected_lookup_sheets:

        temp_df = lookup_sheets[
            sheet_name
        ].copy()

        # Add source sheet only when using
        # multiple lookup sheets
        if len(selected_lookup_sheets) > 1:

            temp_df["_lookup_source_sheet"] = (
                sheet_name
            )

        lookup_frames.append(
            temp_df
        )

    df_lookup = pd.concat(
        lookup_frames,
        ignore_index=True,
    )

    # =====================================================
    # REMOVE EMPTY ROWS/COLUMNS FROM LOOKUP
    # =====================================================

    df_lookup = df_lookup.dropna(
        axis=0,
        how="all",
    )

    df_lookup = df_lookup.dropna(
        axis=1,
        how="all",
    )

    # =====================================================
    # MAIN DATA FOR CONFIGURATION
    # =====================================================

    # Use the first selected Main sheet
    # to configure key columns.
    configuration_main_sheet = (
        selected_main_sheets[0]
    )

    df_main_config = (
        main_sheets[
            configuration_main_sheet
        ].copy()
    )

    df_main_config = df_main_config.dropna(
        axis=0,
        how="all",
    )

    df_main_config = df_main_config.dropna(
        axis=1,
        how="all",
    )

    # Normalize columns
    df_main_config.columns = [
        str(col).strip()
        for col in df_main_config.columns
    ]

    df_lookup.columns = [
        str(col).strip()
        for col in df_lookup.columns
    ]

    # =====================================================
    # FILE INFORMATION
    # =====================================================

    col_a2, col_b2 = st.columns(2)

    with col_a2:

        st.markdown(
            f"**Main:** `{file_main.name}`"
        )

        st.caption(
            f"Selected sheets: "
            f"{len(selected_main_sheets)}"
        )

        show_file_info(
            file_main,
            df_main_config,
        )

    with col_b2:

        st.markdown(
            f"**Lookup:** `{file_lookup.name}`"
        )

        st.caption(
            f"Selected sheets: "
            f"{len(selected_lookup_sheets)}"
        )

        show_file_info(
            file_lookup,
            df_lookup,
        )

    # =====================================================
    # STEP 2 — CONFIGURE
    # =====================================================

    st.markdown("#### Step 2 — Configure")

    c1, c2 = st.columns(2)

    # =====================================================
    # MAIN KEY
    # =====================================================

    with c1:

        main_key = st.selectbox(
            "Key column in Main file",
            df_main_config.columns,
            help=(
                "This key column must exist in every selected "
                "Main sheet if you are processing multiple sheets."
            ),
            key="vl_main_key",
        )

    # =====================================================
    # LOOKUP KEY
    # =====================================================

    with c2:

        lookup_key = st.selectbox(
            "Matching column in Lookup file",
            df_lookup.columns,
            help=(
                "This key column must exist in the selected "
                "Lookup sheets."
            ),
            key="vl_lookup_key",
        )

    # =====================================================
    # LOOKUP COLUMNS
    # =====================================================

    lookup_available_columns = [
        column
        for column in df_lookup.columns
        if column != lookup_key
    ]

    lookup_values = st.multiselect(
        "Columns to fetch from the Lookup file",
        lookup_available_columns,
        help=(
            "These columns will be appended to every "
            "selected Main sheet."
        ),
        key="vl_lookup_values",
    )

    if not lookup_values:

        st.warning(
            "Select at least one column to fetch "
            "from the Lookup file."
        )

        feedback_widget()

        return

    # =====================================================
    # CONFLICTING COLUMNS
    # =====================================================

    main_columns_set = set(
        df_main_config.columns
    )

    conflicting_columns = [
        column
        for column in lookup_values
        if column in main_columns_set
    ]

    if conflicting_columns:

        st.warning(
            "⚠️ These Lookup columns already exist "
            "in the Main file:"
        )

        st.write(
            ", ".join(
                f"`{column}`"
                for column in conflicting_columns
            )
        )

        st.info(
            "They will automatically be renamed as "
            "`column_lookup`, so your original Main "
            "columns will not be overwritten."
        )

    # =====================================================
    # NUMERIC CONVERSION
    # =====================================================

    convert_numbers = st.checkbox(
        "Convert key columns to numeric before matching",
        value=False,
        help=(
            "Useful when one file stores IDs as text "
            "and the other stores them as numbers."
        ),
        key="vl_convert_numbers",
    )

    # =====================================================
    # JOIN TYPE
    # =====================================================

    join_type = st.radio(
        "Join type",
        [
            "Left join (keep all main rows)",
            "Inner join (only matched rows)",
        ],
        horizontal=True,
        key="vl_join_type",
    )

    how = (
        "left"
        if "Left" in join_type
        else "inner"
    )

    # =====================================================
    # STEP 3 — PREVIEW
    # =====================================================

    st.markdown("#### Step 3 — Preview")

    tab1, tab2 = st.tabs(
        [
            "Main file",
            "Lookup file",
        ]
    )

    with tab1:

        st.write(
            "**Selected Main sheets:**"
        )

        for sheet in selected_main_sheets:

            st.write(
                f"📄 `{sheet}`"
            )

        st.write(
            "Preview of first selected Main sheet:"
        )

        show_data_preview(
            df_main_config,
            label=(
                f"Main — {configuration_main_sheet}"
            ),
        )

    with tab2:

        st.write(
            "**Selected Lookup sheets:**"
        )

        for sheet in selected_lookup_sheets:

            st.write(
                f"🔎 `{sheet}`"
            )

        show_data_preview(
            df_lookup,
            label="Lookup file",
        )

    # =====================================================
    # STEP 4 — PROCESS
    # =====================================================

    st.markdown("#### Step 4 — Process")

    if st.button(
        "🔗 Run VLOOKUP",
        type="primary",
        key="vl_run",
    ):

        try:

            # =================================================
            # RESULT STORAGE
            # =================================================

            result_sheets = {}

            all_rename_maps = {}

            total_rows = 0
            total_matched = 0
            total_unmatched = 0

            failed_sheets = []

            # =================================================
            # PROCESS EACH MAIN SHEET
            # =================================================

            for main_sheet_name in selected_main_sheets:

                try:

                    # -----------------------------------------
                    # LOAD MAIN SHEET
                    # -----------------------------------------

                    df_m = main_sheets[
                        main_sheet_name
                    ].copy()

                    # -----------------------------------------
                    # REMOVE EMPTY ROWS/COLUMNS
                    # -----------------------------------------

                    df_m = df_m.dropna(
                        axis=0,
                        how="all",
                    )

                    df_m = df_m.dropna(
                        axis=1,
                        how="all",
                    )

                    # -----------------------------------------
                    # NORMALIZE COLUMNS
                    # -----------------------------------------

                    df_m.columns = [
                        str(col).strip()
                        for col in df_m.columns
                    ]

                    # -----------------------------------------
                    # CHECK MAIN KEY
                    # -----------------------------------------

                    if main_key not in df_m.columns:

                        failed_sheets.append(
                            {
                                "sheet": main_sheet_name,
                                "reason": (
                                    f"Main key `{main_key}` "
                                    "does not exist."
                                ),
                            }
                        )

                        continue

                    # -----------------------------------------
                    # PREPARE MAIN KEY
                    # -----------------------------------------

                    if convert_numbers:

                        df_m[main_key] = pd.to_numeric(
                            df_m[main_key],
                            errors="coerce",
                        )

                    # -----------------------------------------
                    # PREPARE LOOKUP
                    # -----------------------------------------

                    (
                        df_lookup_selected,
                        rename_map,
                    ) = prepare_lookup_dataframe(
                        df_lookup=df_lookup,
                        lookup_key=lookup_key,
                        lookup_values=lookup_values,
                        main_columns=list(
                            df_m.columns
                        ),
                        convert_numbers=convert_numbers,
                    )

                    # -----------------------------------------
                    # STORE RENAME INFO
                    # -----------------------------------------

                    if rename_map:

                        all_rename_maps[
                            main_sheet_name
                        ] = rename_map

                    # -----------------------------------------
                    # MERGE
                    # -----------------------------------------

                    merged = pd.merge(
                        df_m,
                        df_lookup_selected,
                        left_on=main_key,
                        right_on=lookup_key,
                        how=how,
                    )

                    # -----------------------------------------
                    # REMOVE LOOKUP KEY IF DIFFERENT
                    # -----------------------------------------

                    if lookup_key != main_key:

                        merged.drop(
                            columns=[
                                lookup_key
                            ],
                            inplace=True,
                            errors="ignore",
                        )

                    # -----------------------------------------
                    # FIND NEW COLUMNS
                    # -----------------------------------------

                    original_cols = set(
                        df_m.columns
                    )

                    new_cols = [
                        column
                        for column in merged.columns
                        if column not in original_cols
                    ]

                    # -----------------------------------------
                    # MATCH COUNT
                    # -----------------------------------------

                    if new_cols:

                        # We need to determine whether a
                        # lookup match exists.
                        #
                        # Use the first lookup output column.
                        first_lookup_column = (
                            new_cols[0]
                        )

                        matched = int(
                            merged[
                                first_lookup_column
                            ]
                            .notna()
                            .sum()
                        )

                    else:

                        matched = 0

                    unmatched = (
                        len(merged)
                        - matched
                    )

                    # -----------------------------------------
                    # STORE RESULT
                    # -----------------------------------------

                    result_sheets[
                        main_sheet_name
                    ] = merged

                    # -----------------------------------------
                    # TOTALS
                    # -----------------------------------------

                    total_rows += len(merged)

                    total_matched += matched

                    total_unmatched += unmatched

                except Exception as sheet_error:

                    failed_sheets.append(
                        {
                            "sheet": main_sheet_name,
                            "reason": str(
                                sheet_error
                            ),
                        }
                    )

            # =================================================
            # NO RESULTS
            # =================================================

            if not result_sheets:

                st.error(
                    "No Main sheets could be processed."
                )

                if failed_sheets:

                    st.dataframe(
                        pd.DataFrame(
                            failed_sheets
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )

                return

            # =================================================
            # RESULT SUMMARY
            # =================================================

            result_banner(
                "VLOOKUP completed",
                {
                    "Main sheets processed": (
                        f"{len(result_sheets):,}"
                    ),
                    "Rows in result": (
                        f"{total_rows:,}"
                    ),
                    "Matched rows": (
                        f"{total_matched:,}"
                    ),
                    "Unmatched rows": (
                        f"{total_unmatched:,}"
                    ),
                    "Lookup sheets used": (
                        f"{len(selected_lookup_sheets):,}"
                    ),
                },
            )

            # =================================================
            # FAILED SHEETS
            # =================================================

            if failed_sheets:

                st.warning(
                    "⚠️ Some Main sheets could not be processed."
                )

                st.dataframe(
                    pd.DataFrame(
                        failed_sheets
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

            # =================================================
            # SHOW CONFLICT RENAMES
            # =================================================

            if all_rename_maps:

                st.markdown(
                    "### 🔄 Renamed conflicting columns"
                )

                rename_rows = []

                for (
                    sheet_name,
                    rename_map,
                ) in all_rename_maps.items():

                    for (
                        old_name,
                        new_name,
                    ) in rename_map.items():

                        rename_rows.append(
                            {
                                "Main Sheet": sheet_name,
                                "Original Lookup Column": old_name,
                                "Result Column": new_name,
                            }
                        )

                if rename_rows:

                    st.dataframe(
                        pd.DataFrame(
                            rename_rows
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )

            # =================================================
            # RESULT PREVIEW
            # =================================================

            st.markdown(
                "### 👀 Result Preview"
            )

            # Create tabs for result sheets
            preview_tabs = st.tabs(
                list(
                    result_sheets.keys()
                )
            )

            for (
                tab,
                sheet_name,
            ) in zip(
                preview_tabs,
                result_sheets.keys(),
            ):

                with tab:

                    result_df = result_sheets[
                        sheet_name
                    ]

                    # Find newly added columns
                    original_df = main_sheets[
                        sheet_name
                    ]

                    original_columns = set(
                        str(col).strip()
                        for col in original_df.columns
                    )

                    result_columns = [
                        column
                        for column in result_df.columns
                        if column not in original_columns
                    ]

                    def highlight_new_columns(
                        column
                    ):

                        if (
                            column.name
                            in result_columns
                        ):

                            return [
                                "background-color: #fff3cd"
                            ] * len(column)

                        return [
                            ""
                        ] * len(column)

                    st.dataframe(
                        result_df.head(10).style.apply(
                            highlight_new_columns
                        ),
                        use_container_width=True,
                    )

                    st.caption(
                        f"{len(result_df):,} rows × "
                        f"{len(result_df.columns):,} columns"
                    )

            # =================================================
            # DOWNLOAD MULTI-SHEET EXCEL
            # =================================================

            output_bytes = (
                dataframes_to_excel_bytes(
                    result_sheets
                )
            )

            output_file_name = (
                "vlookup_"
                f"{file_main.name.replace('.xlsx', '')}"
                ".xlsx"
            )

            st.download_button(
                "⬇️ Download Result Excel",
                output_bytes,
                file_name=output_file_name,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                key="vl_download",
            )

        except Exception as e:

            friendly_error(e)

    # =====================================================
    # FEEDBACK
    # =====================================================

    feedback_widget()