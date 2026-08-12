"""ExcelFlow Streamlit UI entry point.

Business logic lives in the ``excel_flow`` package.
Phase 4 adds aggregation for Step 3.
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

from excel_flow.aggregator import (
    AGGREGATION_GROUP_KEY,
    AGGREGATION_LABELS,
    AGGREGATION_METHOD_KEY,
    AGGREGATION_SORT_KEY,
    AGGREGATION_VALUE_KEY,
    aggregate_data,
    describe_aggregation,
    get_numeric_columns,
    reset_aggregation_state,
    result_value_column_name,
    sort_aggregated,
)
from excel_flow.data_cleaner import clean_dataframe, summarize_missing_values
from excel_flow.file_loader import (
    detect_file_format,
    get_excel_sheet_names,
    load_file,
    summarize_columns,
)
from excel_flow.validators import ExcelFlowError, validate_rename_mapping

st.set_page_config(
    page_title="ExcelFlow",
    page_icon="📊",
    layout="wide",
)

st.title("ExcelFlow")
st.caption("Excel / CSV 業務自動化ツール")
st.info(
    "現在の開発状況：**Phase 4（集計）** — "
    "ファイル読込・データ整形・集計が利用できます。"
)

AGGREGATION_OPTIONS = {
    "合計": "sum",
    "平均": "mean",
    "件数": "count",
    "最大": "max",
    "最小": "min",
}

AGGREGATION_RESULT_WIDTH = 800


def _handle_aggregation_reset() -> None:
    """Button callback: runs before widgets are created on the next run."""
    reset_aggregation_state(st.session_state)
    st.session_state["aggregation_reset_notice"] = True


def _reset_cleaning_state() -> None:
    for key in (
        "cleaned_df",
        "cleaning_summary",
        "rename_mapping",
        "pending_rename_source",
        "pending_rename_target",
    ):
        st.session_state.pop(key, None)
    reset_aggregation_state(st.session_state)


def _init_rename_mapping() -> dict[str, str]:
    if "rename_mapping" not in st.session_state:
        st.session_state.rename_mapping = {}
    return st.session_state.rename_mapping


# --- STEP 1 ---
st.header("① ファイルを読み込む")
st.write("ExcelまたはCSVファイルを選択してください。")

uploaded_file = st.file_uploader(
    "Excel（.xlsx）または CSV（.csv）",
    type=["xlsx", "csv"],
    help="対応形式: .xlsx / .csv",
)

original_df = None
selected_sheet: str | None = None
file_format: str | None = None

if uploaded_file is not None:
    try:
        file_format = detect_file_format(uploaded_file.name)

        if file_format == "xlsx":
            sheet_names = get_excel_sheet_names(uploaded_file)
            if len(sheet_names) == 1:
                selected_sheet = sheet_names[0]
                st.caption(f"シート: **{selected_sheet}**")
            else:
                selected_sheet = st.selectbox("読み込むシートを選択", sheet_names)

        loaded_df = load_file(uploaded_file, uploaded_file.name, selected_sheet)
        file_key = f"{uploaded_file.name}:{uploaded_file.size}:{selected_sheet or ''}"

        if st.session_state.get("loaded_file_key") != file_key:
            st.session_state.loaded_file_key = file_key
            st.session_state.original_df = loaded_df.copy()
            _reset_cleaning_state()

        original_df = st.session_state.original_df
        row_count = len(original_df)
        column_count = len(original_df.columns)

        st.success(f"{row_count:,}行 × {column_count:,}列を読み込みました")

        info_cols = st.columns(4 if selected_sheet else 3)
        info_cols[0].metric("ファイル名", uploaded_file.name)
        info_cols[1].metric("ファイル形式", file_format.upper())
        info_cols[2].metric("サイズ", f"{row_count:,} × {column_count:,}")
        if selected_sheet:
            info_cols[3].metric("シート", selected_sheet)

        with st.expander("元データのプレビュー / 列情報", expanded=False):
            st.caption("先頭30行を表示しています。")
            st.dataframe(original_df.head(30), use_container_width=True)
            st.dataframe(
                summarize_columns(original_df),
                use_container_width=True,
                hide_index=True,
            )

    except ExcelFlowError as exc:
        st.error(exc.user_message)
        original_df = None
    except Exception:
        st.error("ファイルの読み込み中に予期しないエラーが発生しました。")
        original_df = None
else:
    st.session_state.pop("loaded_file_key", None)
    st.session_state.pop("original_df", None)
    _reset_cleaning_state()

st.divider()

# --- STEP 2 ---
st.header("② データを整える")
st.write("不要列の削除、空白行・重複行の除去、文字列の前後空白除去などを行います。")

if original_df is None:
    st.info("先にファイルを読み込んでください。")
else:
    rename_mapping = _init_rename_mapping()

    columns_to_drop = st.multiselect(
        "削除する列",
        options=list(original_df.columns.astype(str)),
        default=[],
        help="削除したい列を選択します。すべての列は削除できません。",
    )

    with st.expander("列名を変更する", expanded=False):
        available_sources = [
            column
            for column in original_df.columns.astype(str)
            if column not in rename_mapping
        ]
        col_a, col_b, col_c = st.columns([2, 2, 1])
        with col_a:
            rename_source = st.selectbox(
                "変更する列",
                options=available_sources or [""],
                disabled=not available_sources,
                key="rename_source_select",
            )
        with col_b:
            rename_target = st.text_input("新しい列名", key="rename_target_input")
        with col_c:
            st.write("")
            st.write("")
            add_rename = st.button("変更を追加", use_container_width=True)

        if add_rename:
            try:
                if not available_sources:
                    st.warning("変更できる列がありません。")
                else:
                    candidate = {**rename_mapping, str(rename_source): rename_target}
                    validated = validate_rename_mapping(original_df, candidate)
                    st.session_state.rename_mapping = validated
                    st.rerun()
            except ExcelFlowError as exc:
                st.error(exc.user_message)

        if rename_mapping:
            st.caption("変更予定一覧")
            for source, target in list(rename_mapping.items()):
                item_cols = st.columns([4, 1])
                item_cols[0].write(f"`{source}` → `{target}`")
                if item_cols[1].button("削除", key=f"remove_rename_{source}"):
                    rename_mapping.pop(source, None)
                    st.session_state.rename_mapping = rename_mapping
                    st.rerun()

    check_cols = st.columns(3)
    with check_cols[0]:
        strip_whitespace = st.checkbox("文字列の前後空白を削除", value=False)
    with check_cols[1]:
        remove_blank_rows = st.checkbox("空白行を削除", value=False)
    with check_cols[2]:
        remove_duplicates = st.checkbox("重複行を削除", value=False)

    action_cols = st.columns(2)
    run_cleaning = action_cols[0].button("データ整形を実行", type="primary")
    reset_cleaning = action_cols[1].button("整形をリセット", type="secondary")

    if reset_cleaning:
        _reset_cleaning_state()
        st.success("整形結果をリセットしました。元データからやり直せます。")
        st.rerun()

    if run_cleaning:
        try:
            active_renames = st.session_state.get("rename_mapping", {})
            effective_drop = [active_renames.get(column, column) for column in columns_to_drop]
            cleaned_df, summary = clean_dataframe(
                original_df,
                rename_mapping=active_renames,
                columns_to_drop=effective_drop,
                strip_whitespace=strip_whitespace,
                remove_blank_rows=remove_blank_rows,
                remove_duplicates=remove_duplicates,
            )
            st.session_state.cleaned_df = cleaned_df
            st.session_state.cleaning_summary = summary
            reset_aggregation_state(st.session_state)
            st.success("データ整形が完了しました。")
        except ExcelFlowError as exc:
            st.error(exc.user_message)
        except Exception:
            st.error("データ整形中に予期しないエラーが発生しました。")

    cleaned_df = st.session_state.get("cleaned_df")
    summary = st.session_state.get("cleaning_summary")

    if cleaned_df is not None and summary is not None:
        with st.expander("整形結果", expanded=True):
            st.write(
                f"整形前：{summary['original_rows']:,}行 × {summary['original_columns']:,}列  →  "
                f"整形後：{summary['cleaned_rows']:,}行 × {summary['cleaned_columns']:,}列"
            )

            metric_cols = st.columns(4)
            metric_cols[0].metric("削除行数", f"{summary['removed_rows']:,}")
            metric_cols[1].metric("削除列数", f"{summary['removed_columns']:,}")
            metric_cols[2].metric("空白行削除", f"{summary['blank_rows_removed']:,}")
            metric_cols[3].metric("重複削除", f"{summary['duplicates_removed']:,}")

            st.caption("整形済データのプレビュー（先頭30行）")
            st.dataframe(cleaned_df.head(30), use_container_width=True)

            result_tabs = st.tabs(["欠損値一覧", "列名 / データ型"])
            with result_tabs[0]:
                st.dataframe(
                    summarize_missing_values(cleaned_df),
                    use_container_width=True,
                    hide_index=True,
                )
                st.caption("欠損は NaN / None / NaT のみを集計しています（空白文字列は含みません）。")
            with result_tabs[1]:
                st.dataframe(
                    summarize_columns(cleaned_df),
                    use_container_width=True,
                    hide_index=True,
                )

st.divider()

# --- STEP 3 ---
st.header("③ 集計する")
st.write("グループ項目と集計方法を指定して、データを集計します。")

if original_df is None:
    st.info("先にファイルを読み込んでください。")
else:
    cleaned_df = st.session_state.get("cleaned_df")
    if cleaned_df is not None:
        source_df = cleaned_df
        source_label = "整形済データ"
    else:
        source_df = original_df
        source_label = "元データ"

    if st.session_state.pop("aggregation_reset_notice", False):
        st.success("集計結果をリセットしました。")

    st.info(f"集計対象：**{source_label}**（{len(source_df):,}行 × {len(source_df.columns):,}列）")

    # Keep Phase 4 controls compact (not full-page width).
    controls_left, _controls_spacer = st.columns([2, 1])
    with controls_left:
        group_columns = st.multiselect(
            "グループ項目（最大2列）",
            options=list(source_df.columns.astype(str)),
            default=[],
            max_selections=2,
            help="集計の軸となる列を1〜2つ選択します。",
            key=AGGREGATION_GROUP_KEY,
        )

        method_col, value_col_ui = st.columns(2)
        with method_col:
            method_label = st.selectbox(
                "集計方法",
                options=list(AGGREGATION_OPTIONS.keys()),
                index=0,
                key=AGGREGATION_METHOD_KEY,
            )
        aggregation = AGGREGATION_OPTIONS[method_label]

        value_column: str | None = None
        numeric_columns = get_numeric_columns(source_df)

        with value_col_ui:
            if aggregation == "count":
                st.caption("件数：グループごとの行数を集計します。")
            else:
                if not numeric_columns:
                    st.warning("数値列がないため、合計・平均・最大・最小は実行できません。")
                value_column = st.selectbox(
                    "集計対象列",
                    options=numeric_columns or [""],
                    disabled=not numeric_columns,
                    key=AGGREGATION_VALUE_KEY,
                )

        sort_label = st.radio(
            "並べ替え（集計値）",
            options=["降順", "昇順"],
            index=0,
            horizontal=True,
            key=AGGREGATION_SORT_KEY,
        )

        agg_action_cols = st.columns(2)
        run_aggregation = agg_action_cols[0].button("集計を実行", type="primary")
        # on_click runs before widgets are created, avoiding StreamlitAPIException.
        agg_action_cols[1].button(
            "集計をリセット",
            type="secondary",
            on_click=_handle_aggregation_reset,
        )

    if run_aggregation:
        try:
            result_df = aggregate_data(
                source_df,
                group_columns=group_columns,
                aggregation=aggregation,
                value_column=None if aggregation == "count" else value_column,
            )
            result_column = result_value_column_name(
                aggregation,
                None if aggregation == "count" else value_column,
            )
            st.session_state.aggregated_df = result_df
            st.session_state.aggregation_config = {
                "source_label": source_label,
                "group_columns": list(group_columns),
                "aggregation": aggregation,
                "aggregation_label": AGGREGATION_LABELS[aggregation],
                "value_column": None if aggregation == "count" else value_column,
                "result_column": result_column,
                "description": describe_aggregation(
                    group_columns=list(group_columns),
                    aggregation=aggregation,
                    value_column=None if aggregation == "count" else value_column,
                ),
            }
            st.success("集計が完了しました。")
        except ExcelFlowError as exc:
            st.error(exc.user_message)
        except Exception:
            st.error("集計中に予期しないエラーが発生しました。")

    aggregated_df = st.session_state.get("aggregated_df")
    aggregation_config = st.session_state.get("aggregation_config")

    if aggregated_df is not None and aggregation_config is not None:
        ascending = sort_label == "昇順"
        display_df = sort_aggregated(
            aggregated_df,
            aggregation_config["result_column"],
            ascending=ascending,
        )

        st.subheader("集計結果")
        st.write(f"集計条件：{aggregation_config['description']}")
        st.write(f"集計対象：{aggregation_config['source_label']}")
        st.write(f"結果：{len(display_df):,}グループ")

        # Equal-ish column widths inside the fixed 800px table; numbers left-aligned.
        # display_df is view-only; aggregated_df numeric types remain unchanged.
        equal_width = max(160, AGGREGATION_RESULT_WIDTH // max(len(display_df.columns), 1))
        column_config = {}
        for column_name in display_df.columns:
            if pd.api.types.is_numeric_dtype(display_df[column_name]):
                column_config[column_name] = st.column_config.NumberColumn(
                    str(column_name),
                    width=equal_width,
                    alignment="left",
                )
            else:
                column_config[column_name] = st.column_config.TextColumn(
                    str(column_name),
                    width=equal_width,
                    alignment="left",
                )

        st.dataframe(
            display_df,
            width=AGGREGATION_RESULT_WIDTH,
            hide_index=True,
            column_config=column_config,
        )

st.divider()

with st.expander("④〜⑤ 今後の機能（未実装）", expanded=False):
    st.write("④ グラフで確認する — Phase 5 で実装予定")
    st.write("⑤ Excelへ出力する — Phase 6 で実装予定")

st.caption("ExcelFlow Ver1.0 — Phase 4: 集計")
