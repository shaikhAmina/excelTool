"""Configurable, ordered spreadsheet transformations."""

import pandas as pd
import streamlit as st

from utils.helpers import (
    df_to_excel_bytes,
    feedback_widget,
    friendly_error,
    privacy_notice,
    result_banner,
    show_data_preview,
    show_file_info,
)


STEP_OPTIONS = ["Concatenate columns", "Convert to number", "Lookup columns"]


def _text_value(value) -> str:
    """Turn cell values into joinable text without turning blanks into 'nan'."""
    return "" if pd.isna(value) else str(value)


def _concat_columns(df: pd.DataFrame, columns: list[str], output: str, separator: str) -> pd.DataFrame:
    if not output.strip():
        raise ValueError("Give the concatenated column a name.")
    if len(columns) < 2:
        raise ValueError("Choose at least two columns to concatenate.")
    result = df.copy()
    result[output.strip()] = result[columns].apply(
        lambda row: separator.join(_text_value(value) for value in row), axis=1
    )
    return result


def _convert_to_number(df: pd.DataFrame, column: str) -> pd.DataFrame:
    result = df.copy()
    # Commas and surrounding whitespace are common in Excel exports.
    cleaned = result[column].astype("string").str.replace(",", "", regex=False).str.strip()
    result[column] = pd.to_numeric(cleaned, errors="coerce")
    return result


def _lookup_columns(
    main: pd.DataFrame,
    lookup: pd.DataFrame,
    main_key: str,
    lookup_key: str,
    columns: list[str],
    how: str,
) -> tuple[pd.DataFrame, list[str]]:
    if not columns:
        raise ValueError("Choose at least one lookup column to add.")
    conflicts = [column for column in columns if column in main.columns]
    if conflicts:
        raise ValueError(
            "Lookup output already has these column names: " + ", ".join(conflicts)
            + ". Rename them in the lookup file first."
        )
    lookup_subset = lookup[[lookup_key] + columns].drop_duplicates(subset=[lookup_key])
    result = pd.merge(main, lookup_subset, left_on=main_key, right_on=lookup_key, how=how)
    if lookup_key != main_key:
        result = result.drop(columns=[lookup_key])
    return result, columns


def data_workflow_ui():
    st.markdown("### ⚙️ Data Workflow")
    st.caption(
        "Build an ordered recipe for your spreadsheet: combine columns, convert values to "
        "numbers, and enrich rows from a second Excel file. Steps run from top to bottom."
    )
    privacy_notice()
    st.markdown("---")

    st.markdown("#### Step 1 — Upload files")
    left, right = st.columns(2)
    with left:
        main_file = st.file_uploader("Main Excel file", type=["xlsx"], key="workflow_main")
    with right:
        lookup_file = st.file_uploader(
            "Lookup Excel file (only needed for a lookup step)", type=["xlsx"], key="workflow_lookup"
        )

    if not main_file:
        st.info("Upload a main .xlsx file to start building a workflow.")
        feedback_widget()
        return

    try:
        main_df = pd.read_excel(main_file)
        lookup_df = pd.read_excel(lookup_file) if lookup_file else None
    except Exception as error:
        friendly_error(error)
        return

    show_file_info(main_file, main_df)
    if lookup_file:
        show_file_info(lookup_file, lookup_df)

    st.markdown("#### Step 2 — Build the workflow")
    step_count = st.select_slider(
        "How many operations should run, in order?", options=[1, 2, 3, 4, 5, 6], value=1,
        help="The output of each operation becomes the input for the next one.",
    )

    # Track the columns that each configured step will make available to later steps.
    # This lets a composite key be built in *both* files before a lookup.
    main_columns = list(main_df.columns)
    lookup_columns = list(lookup_df.columns) if lookup_df is not None else []
    workflow = []
    for index in range(step_count):
        with st.expander(f"Operation {index + 1}", expanded=True):
            operation = st.selectbox(
                "Action", STEP_OPTIONS, key=f"workflow_action_{index}"
            )
            config = {"operation": operation}

            if operation == "Concatenate columns":
                target = st.radio(
                    "Apply to", ["Main file", "Lookup file"], horizontal=True,
                    key=f"workflow_concat_target_{index}", disabled=lookup_df is None,
                    help="Choose Lookup file to build the same composite key in the reference data.",
                )
                available_columns = main_columns if target == "Main file" else lookup_columns
                selected = st.multiselect(
                    "Columns to combine", available_columns, key=f"workflow_concat_cols_{index}",
                    help="Their values are joined in the order selected.",
                )
                output = st.text_input(
                    "New column name", value=f"combined_{index + 1}", key=f"workflow_concat_name_{index}"
                )
                separator = st.text_input(
                    "Separator", value="", key=f"workflow_concat_sep_{index}",
                    help="Examples: a space, a hyphen (-), or leave empty to join directly.",
                )
                config.update(target=target, columns=selected, output=output, separator=separator)
                if output.strip() and output.strip() not in available_columns:
                    available_columns.append(output.strip())

            elif operation == "Convert to number":
                target = st.radio(
                    "Apply to", ["Main file", "Lookup file"], horizontal=True,
                    key=f"workflow_num_target_{index}", disabled=lookup_df is None,
                )
                available_columns = main_columns if target == "Main file" else lookup_columns
                column = st.selectbox("Column to convert", available_columns, key=f"workflow_num_col_{index}")
                config.update(target=target, column=column)

            else:
                if lookup_df is None:
                    st.warning("Upload a lookup Excel file above before using this operation.")
                    config["missing_lookup_file"] = True
                else:
                    c1, c2 = st.columns(2)
                    with c1:
                        main_key = st.selectbox(
                            "Key in the current main data", main_columns, key=f"workflow_main_key_{index}"
                        )
                    with c2:
                        lookup_key = st.selectbox(
                            "Matching key in lookup file", lookup_columns, key=f"workflow_lookup_key_{index}"
                        )
                    values = st.multiselect(
                        "Columns to add from lookup file",
                        [column for column in lookup_df.columns if column != lookup_key],
                        key=f"workflow_lookup_values_{index}",
                    )
                    join_label = st.radio(
                        "Rows to keep", ["Keep every main row", "Keep matched rows only"],
                        horizontal=True, key=f"workflow_join_{index}",
                    )
                    config.update(
                        main_key=main_key,
                        lookup_key=lookup_key,
                        columns=values,
                        how="left" if join_label.startswith("Keep every") else "inner",
                    )
                    for column in values:
                        if column not in main_columns:
                            main_columns.append(column)
            workflow.append(config)

    st.markdown("#### Step 3 — Input preview")
    tabs = st.tabs(["Main file", "Lookup file"] if lookup_df is not None else ["Main file"])
    with tabs[0]:
        show_data_preview(main_df, label="Starting data")
    if lookup_df is not None:
        with tabs[1]:
            show_data_preview(lookup_df, label="Lookup data")

    st.markdown("#### Step 4 — Run & download")
    if st.button("▶ Run workflow", type="primary"):
        try:
            result = main_df.copy()
            lookup_result = lookup_df.copy() if lookup_df is not None else None
            summaries = []
            for position, config in enumerate(workflow, start=1):
                operation = config["operation"]
                if operation == "Concatenate columns":
                    if config["target"] == "Main file":
                        result = _concat_columns(result, config["columns"], config["output"], config["separator"])
                    else:
                        lookup_result = _concat_columns(
                            lookup_result, config["columns"], config["output"], config["separator"]
                        )
                    summaries.append(f"{position}. Created {config['output'].strip()} in {config['target']}")
                elif operation == "Convert to number":
                    if config["target"] == "Main file":
                        result = _convert_to_number(result, config["column"])
                    else:
                        lookup_result = _convert_to_number(lookup_result, config["column"])
                    summaries.append(f"{position}. Converted {config['column']} to numbers in {config['target']}")
                else:
                    if config.get("missing_lookup_file"):
                        raise ValueError(f"Operation {position} needs a lookup Excel file.")
                    result, added = _lookup_columns(
                        result, lookup_result, config["main_key"], config["lookup_key"],
                        config["columns"], config["how"],
                    )
                    matched = int(result[added[0]].notna().sum()) if added else 0
                    summaries.append(f"{position}. Added {', '.join(added)} ({matched:,} matched rows)")

            result_banner(
                "Workflow completed",
                {"Rows in result": f"{len(result):,}", "Columns": len(result.columns), "Operations": step_count},
            )
            st.success(" → ".join(summaries))
            with st.expander("👀 Result preview", expanded=True):
                st.dataframe(result.head(20), use_container_width=True)
            st.download_button(
                "⬇️ Download Result Excel", df_to_excel_bytes(result),
                file_name=f"workflow_{main_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as error:
            friendly_error(error)

    feedback_widget()
