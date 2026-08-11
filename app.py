"""ExcelFlow Streamlit UI entry point.

Business logic lives in the ``excel_flow`` package.
Phase 3 adds data cleaning for Step 2.
"""

from __future__ import annotations

import streamlit as st

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
    "現在の開発状況：**Phase 3（データ整形）** — "
    "ファイル読込とデータ整形が利用できます。"
)


def _reset_cleaning_state() -> None:
    for key in (
        "cleaned_df",
        "cleaning_summary",
        "rename_mapping",
        "pending_rename_source",
        "pending_rename_target",
    ):
        st.session_state.pop(key, None)


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

        with st.expander("元データのプレビュー / 列情報", expanded=True):
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
            # Multiselect uses original column names; map them after renames.
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
            st.success("データ整形が完了しました。")
        except ExcelFlowError as exc:
            st.error(exc.user_message)
        except Exception:
            st.error("データ整形中に予期しないエラーが発生しました。")

    cleaned_df = st.session_state.get("cleaned_df")
    summary = st.session_state.get("cleaning_summary")

    if cleaned_df is not None and summary is not None:
        st.subheader("整形結果")
        st.write(
            f"整形前：{summary['original_rows']:,}行 × {summary['original_columns']:,}列  →  "
            f"整形後：{summary['cleaned_rows']:,}行 × {summary['cleaned_columns']:,}列"
        )

        metric_cols = st.columns(4)
        metric_cols[0].metric("削除行数", f"{summary['removed_rows']:,}")
        metric_cols[1].metric("削除列数", f"{summary['removed_columns']:,}")
        metric_cols[2].metric("空白行削除", f"{summary['blank_rows_removed']:,}")
        metric_cols[3].metric("重複削除", f"{summary['duplicates_removed']:,}")

        st.subheader("整形済データのプレビュー")
        st.caption("先頭30行を表示しています。元データとは別の結果です。")
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

# --- Later phases (compact) ---
with st.expander("③〜⑤ 今後の機能（未実装）", expanded=False):
    st.write("③ 集計する — Phase 4 で実装予定")
    st.write("④ グラフで確認する — Phase 5 で実装予定")
    st.write("⑤ Excelへ出力する — Phase 6 で実装予定")

st.caption("ExcelFlow Ver1.0 — Phase 3: データ整形")
