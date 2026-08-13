"""ExcelFlow Streamlit UI entry point.

Business logic lives in the ``excel_flow`` package.
Phase 7 focuses on UI/UX polish for the demo experience.
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
from excel_flow.chart_builder import (
    CHART_TITLE_KEY,
    CHART_TYPE_KEY,
    CHART_TYPE_OPTIONS,
    CHART_X_KEY,
    CHART_Y_KEY,
    build_chart,
    chart_settings_from_aggregation,
    reset_chart_state,
)
from excel_flow.data_cleaner import clean_dataframe, summarize_missing_values
from excel_flow.excel_exporter import (
    EXPORT_BYTES_KEY,
    EXPORT_FILENAME_KEY,
    EXPORT_INCLUDE_AGG_KEY,
    EXPORT_INCLUDE_DATA_KEY,
    EXPORT_INCLUDE_HISTORY_KEY,
    EXPORT_READY_FILENAME_KEY,
    build_excel_workbook,
    build_process_history_from_session,
    clear_export_artifacts,
    generate_default_filename,
    is_export_ready,
    normalize_export_filename,
    reset_export_state,
    store_export_result,
)
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
st.markdown(
    "Excel / CSV の**読込・整形・集計・グラフ・Excel出力**を、"
    "ブラウザ上で行える業務自動化ツールです。"
)
st.caption(
    "操作の流れ：Step 1 ファイル読込 → Step 2 データ整形 → "
    "Step 3 集計 → Step 4 グラフ → Step 5 Excel出力"
)
st.info("動作確認には `sample_data` フォルダ内のサンプルデータを利用できます。")

AGGREGATION_OPTIONS = {
    "合計": "sum",
    "平均": "mean",
    "件数": "count",
    "最大": "max",
    "最小": "min",
}

AGGREGATION_RESULT_WIDTH = 800


def _reset_cleaning_state() -> None:
    for key in (
        "cleaned_df",
        "cleaning_summary",
        "cleaning_config",
        "rename_mapping",
        "pending_rename_source",
        "pending_rename_target",
    ):
        st.session_state.pop(key, None)
    reset_aggregation_state(st.session_state)


def _handle_cleaning_reset() -> None:
    """Button callback: clear cleaning results and dependent downstream state."""
    _reset_cleaning_state()
    st.session_state["cleaning_reset_notice"] = True


def _handle_aggregation_reset() -> None:
    """Button callback: runs before widgets are created on the next run."""
    reset_aggregation_state(st.session_state)
    st.session_state["aggregation_reset_notice"] = True


def _handle_chart_reset() -> None:
    """Button callback: clear chart state only, keep aggregation results."""
    reset_chart_state(st.session_state)
    st.session_state["chart_reset_notice"] = True


def _handle_export_reset() -> None:
    """Button callback: clear export state only, keep upstream results."""
    reset_export_state(st.session_state)
    st.session_state["export_reset_notice"] = True


def _ensure_chart_widget_defaults(aggregation_config: dict) -> None:
    """Set chart widget defaults before widgets are instantiated."""
    defaults = chart_settings_from_aggregation(aggregation_config)
    if CHART_TYPE_KEY not in st.session_state:
        st.session_state[CHART_TYPE_KEY] = "棒グラフ"
    if CHART_X_KEY not in st.session_state:
        st.session_state[CHART_X_KEY] = defaults["x_column"]
    if CHART_Y_KEY not in st.session_state:
        st.session_state[CHART_Y_KEY] = defaults["y_column"]
    if CHART_TITLE_KEY not in st.session_state:
        st.session_state[CHART_TITLE_KEY] = defaults["title"]


def _init_rename_mapping() -> dict[str, str]:
    if "rename_mapping" not in st.session_state:
        st.session_state.rename_mapping = {}
    return st.session_state.rename_mapping


# --- STEP 1 ---
st.header("Step 1　ファイル読込")
st.write("Excel（.xlsx）または CSV（.csv）ファイルを選択してください。")

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
            st.session_state.loaded_file_meta = {
                "filename": uploaded_file.name,
                "file_format": file_format,
                "sheet_name": selected_sheet,
            }
            _reset_cleaning_state()

        original_df = st.session_state.original_df
        row_count = len(original_df)
        column_count = len(original_df.columns)

        st.success("ファイルを読み込みました。")

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
        st.error("ファイルを読み込めませんでした。")
        original_df = None
else:
    st.session_state.pop("loaded_file_key", None)
    st.session_state.pop("original_df", None)
    st.session_state.pop("loaded_file_meta", None)
    _reset_cleaning_state()

st.divider()

# --- STEP 2 ---
st.header("Step 2　データ整形")
st.write("不要列の削除、列名変更、空白行・重複行の除去、文字列の前後空白除去を行います。")

if original_df is None:
    st.info("データ整形を行うには、先にStep 1でファイルを読み込んでください。")
else:
    if st.session_state.pop("cleaning_reset_notice", False):
        st.success("整形結果をリセットしました。")

    rename_mapping = _init_rename_mapping()

    clean_controls, _clean_spacer = st.columns([2, 1])
    with clean_controls:
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
        action_cols[1].button(
            "整形をリセット",
            type="secondary",
            on_click=_handle_cleaning_reset,
        )

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
            st.session_state.cleaning_config = {
                "columns_to_drop": list(columns_to_drop),
                "rename_mapping": dict(active_renames),
                "strip_whitespace": strip_whitespace,
                "remove_blank_rows": remove_blank_rows,
                "remove_duplicates": remove_duplicates,
            }
            reset_aggregation_state(st.session_state)
            st.success("データ整形を実行しました。")
        except ExcelFlowError as exc:
            st.error(exc.user_message)
        except Exception:
            st.error("データ整形を実行できませんでした。")

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
st.header("Step 3　集計")
st.write("グループ項目と集計方法を指定して、データを集計します。")

if original_df is None:
    st.info("集計を行うには、先にStep 1でファイルを読み込んでください。")
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

    st.caption(f"集計対象：{source_label}（{len(source_df):,}行 × {len(source_df.columns):,}列）")

    # Keep aggregation controls compact (not full-page width).
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
        if not group_columns:
            st.caption("集計する項目を選択してください。")

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
                    st.warning("対象となる数値列がありません。")
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
            reset_chart_state(st.session_state)
            st.success("集計を実行しました。")
        except ExcelFlowError as exc:
            st.error(exc.user_message)
        except Exception:
            st.error("集計を実行できませんでした。")

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

# --- STEP 4 ---
st.header("Step 4　グラフ")
st.write("Step 3の集計結果をもとに、棒グラフ・折れ線グラフ・円グラフを作成します。")

aggregated_df = st.session_state.get("aggregated_df")
aggregation_config = st.session_state.get("aggregation_config")

if original_df is None:
    st.info("グラフを作成するには、先にStep 1でファイルを読み込んでください。")
elif aggregated_df is None or aggregation_config is None:
    st.info("グラフを作成するには、先にStep 3で集計を実行してください。")
else:
    if st.session_state.pop("chart_reset_notice", False):
        st.success("グラフ設定をリセットしました。")

    _ensure_chart_widget_defaults(aggregation_config)

    group_columns = list(aggregation_config["group_columns"])
    result_column = aggregation_config["result_column"]
    color_column = group_columns[1] if len(group_columns) >= 2 else None
    x_candidates = group_columns
    y_candidates = [result_column]

    chart_controls, _chart_spacer = st.columns([2, 1])
    with chart_controls:
        type_col, x_col, y_col = st.columns(3)
        with type_col:
            chart_type_label = st.selectbox(
                "グラフ種類",
                options=list(CHART_TYPE_OPTIONS.keys()),
                key=CHART_TYPE_KEY,
            )
        chart_type = CHART_TYPE_OPTIONS[chart_type_label]

        with x_col:
            x_label = "カテゴリ" if chart_type == "pie" else "X軸"
            x_column = st.selectbox(
                x_label,
                options=x_candidates,
                key=CHART_X_KEY,
            )
        with y_col:
            y_label = "値" if chart_type == "pie" else "Y軸"
            y_column = st.selectbox(
                y_label,
                options=y_candidates,
                key=CHART_Y_KEY,
            )

        chart_title = st.text_input("グラフタイトル", key=CHART_TITLE_KEY)

        if color_column and chart_type != "pie":
            st.caption(f"色分け：`{color_column}`（第2グループ項目）")
        if chart_type == "pie" and color_column:
            st.warning("円グラフは1つのグループ項目で集計した場合に利用できます。")

        chart_action_cols = st.columns(2)
        run_chart = chart_action_cols[0].button("グラフを作成", type="primary")
        chart_action_cols[1].button(
            "グラフをリセット",
            type="secondary",
            on_click=_handle_chart_reset,
        )

    if run_chart:
        try:
            if chart_type == "pie" and color_column:
                raise ExcelFlowError("円グラフは1つのグループ項目で集計した場合に利用できます。")

            ascending = st.session_state.get(AGGREGATION_SORT_KEY, "降順") == "昇順"
            chart_source = sort_aggregated(
                aggregated_df,
                result_column,
                ascending=ascending,
            )
            # Validate by building once; store settings for reruns.
            build_chart(
                chart_source,
                chart_type=chart_type,
                x_column=x_column,
                y_column=y_column,
                title=chart_title,
                color_column=None if chart_type == "pie" else color_column,
            )
            st.session_state.chart_config = {
                "chart_type": chart_type,
                "x_column": x_column,
                "y_column": y_column,
                "title": chart_title,
                "color_column": None if chart_type == "pie" else color_column,
            }
            st.session_state.chart_generated = True
            reset_export_state(st.session_state)
            st.success("グラフを作成しました。")
        except ExcelFlowError as exc:
            st.error(exc.user_message)
        except Exception:
            st.error("グラフを作成できませんでした。")

    if st.session_state.get("chart_generated") and st.session_state.get("chart_config"):
        try:
            ascending = st.session_state.get(AGGREGATION_SORT_KEY, "降順") == "昇順"
            chart_source = sort_aggregated(
                aggregated_df,
                result_column,
                ascending=ascending,
            )
            cfg = st.session_state.chart_config
            figure = build_chart(
                chart_source,
                chart_type=cfg["chart_type"],
                x_column=cfg["x_column"],
                y_column=cfg["y_column"],
                title=cfg.get("title", ""),
                color_column=cfg.get("color_column"),
            )
            st.plotly_chart(figure, use_container_width=True)
        except ExcelFlowError as exc:
            st.error(exc.user_message)
        except Exception:
            st.error("グラフを作成できませんでした。")

st.divider()

# --- STEP 5 ---
st.header("Step 5　Excel出力")
st.write("整形済データ・集計結果・処理履歴を Excel ファイルとしてダウンロードします。")

if original_df is None:
    st.info("Excel出力を行うには、先にStep 1でファイルを読み込んでください。")
else:
    if st.session_state.pop("export_reset_notice", False):
        st.success("出力設定をリセットしました。")

    cleaned_df = st.session_state.get("cleaned_df")
    aggregated_df = st.session_state.get("aggregated_df")
    aggregation_config = st.session_state.get("aggregation_config")
    export_source_df = cleaned_df if cleaned_df is not None else original_df
    export_source_label = "整形済データ" if cleaned_df is not None else "元データ"

    st.caption(
        f"出力対象：{export_source_label}"
        f"（{len(export_source_df):,}行 × {len(export_source_df.columns):,}列）"
    )
    if aggregated_df is None:
        st.caption("集計結果シートは、Step 3で集計を実行すると出力できます。")

    if EXPORT_FILENAME_KEY not in st.session_state:
        st.session_state[EXPORT_FILENAME_KEY] = generate_default_filename()
    if EXPORT_INCLUDE_DATA_KEY not in st.session_state:
        st.session_state[EXPORT_INCLUDE_DATA_KEY] = True
    if EXPORT_INCLUDE_AGG_KEY not in st.session_state:
        st.session_state[EXPORT_INCLUDE_AGG_KEY] = aggregated_df is not None
    if EXPORT_INCLUDE_HISTORY_KEY not in st.session_state:
        st.session_state[EXPORT_INCLUDE_HISTORY_KEY] = True

    export_controls, _export_spacer = st.columns([2, 1])
    with export_controls:
        st.text_input("出力ファイル名", key=EXPORT_FILENAME_KEY)

        include_cols = st.columns(3)
        with include_cols[0]:
            include_data = st.checkbox("整形済データ", key=EXPORT_INCLUDE_DATA_KEY)
        with include_cols[1]:
            include_aggregated = st.checkbox(
                "集計結果",
                key=EXPORT_INCLUDE_AGG_KEY,
                disabled=aggregated_df is None,
            )
        with include_cols[2]:
            include_history = st.checkbox("処理履歴", key=EXPORT_INCLUDE_HISTORY_KEY)

        export_action_cols = st.columns(2)
        run_export = export_action_cols[0].button("Excelファイルを作成", type="primary")
        export_action_cols[1].button(
            "出力をリセット",
            type="secondary",
            on_click=_handle_export_reset,
        )

    if run_export:
        # Prevent stale downloadables from surviving a failed recreate.
        clear_export_artifacts(st.session_state)
        try:
            filename = normalize_export_filename(st.session_state.get(EXPORT_FILENAME_KEY))
            ascending = st.session_state.get(AGGREGATION_SORT_KEY, "降順") == "昇順"
            sorted_aggregated = None
            if aggregated_df is not None and aggregation_config is not None:
                sorted_aggregated = sort_aggregated(
                    aggregated_df,
                    aggregation_config["result_column"],
                    ascending=ascending,
                )

            history_df = build_process_history_from_session(
                {
                    "file_meta": st.session_state.get("loaded_file_meta") or {},
                    "original_df": original_df,
                    "cleaned_df": cleaned_df,
                    "cleaning_summary": st.session_state.get("cleaning_summary"),
                    "cleaning_config": st.session_state.get("cleaning_config"),
                    "aggregated_df": sorted_aggregated,
                    "aggregation_config": aggregation_config,
                    "chart_config": st.session_state.get("chart_config"),
                    "chart_generated": st.session_state.get("chart_generated"),
                    "data_source_label": export_source_label,
                    "sort_label": st.session_state.get(AGGREGATION_SORT_KEY, "降順"),
                }
            )

            buffer = build_excel_workbook(
                export_source_df,
                aggregated_df=sorted_aggregated,
                process_history=history_df,
                include_data=include_data,
                include_aggregated=include_aggregated and sorted_aggregated is not None,
                include_history=include_history,
            )
            store_export_result(
                st.session_state,
                export_bytes=buffer.getvalue(),
                filename=filename,
                export_config={
                    "include_data": include_data,
                    "include_aggregated": include_aggregated and sorted_aggregated is not None,
                    "include_history": include_history,
                    "source_label": export_source_label,
                },
            )
            st.success("Excelファイルを作成しました。")
        except ExcelFlowError as exc:
            clear_export_artifacts(st.session_state)
            st.error(exc.user_message)
        except Exception:
            clear_export_artifacts(st.session_state)
            st.error("Excelファイルを作成できませんでした。")

    if is_export_ready(st.session_state):
        st.download_button(
            label="Excelファイルをダウンロード",
            data=st.session_state[EXPORT_BYTES_KEY],
            file_name=st.session_state[EXPORT_READY_FILENAME_KEY],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.caption("ExcelFlow Ver1.0")
