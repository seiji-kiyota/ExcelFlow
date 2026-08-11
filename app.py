"""ExcelFlow Streamlit UI entry point.

Business logic lives in the ``excel_flow`` package.
Phase 2 implements file loading and preview for Step 1.
"""

from __future__ import annotations

import streamlit as st

from excel_flow.file_loader import (
    detect_file_format,
    get_excel_sheet_names,
    load_file,
    summarize_columns,
)
from excel_flow.validators import ExcelFlowError

st.set_page_config(
    page_title="ExcelFlow",
    page_icon="📊",
    layout="wide",
)

st.title("ExcelFlow")
st.caption("Excel / CSV 業務自動化ツール")
st.info("現在の開発状況：**Phase 2（ファイル読込）** — Excel / CSV の読込とプレビューが利用できます。")

st.divider()

# --- STEP 1 ---
st.header("① ファイルを読み込む")
st.write("ExcelまたはCSVファイルを選択してください。")

uploaded_file = st.file_uploader(
    "Excel（.xlsx）または CSV（.csv）",
    type=["xlsx", "csv"],
    help="対応形式: .xlsx / .csv",
)

if uploaded_file is not None:
    try:
        file_format = detect_file_format(uploaded_file.name)
        selected_sheet: str | None = None

        if file_format == "xlsx":
            sheet_names = get_excel_sheet_names(uploaded_file)
            if len(sheet_names) == 1:
                selected_sheet = sheet_names[0]
                st.caption(f"シート: **{selected_sheet}**")
            else:
                selected_sheet = st.selectbox("読み込むシートを選択", sheet_names)

        dataframe = load_file(uploaded_file, uploaded_file.name, selected_sheet)

        row_count = len(dataframe)
        column_count = len(dataframe.columns)

        st.success(f"{row_count:,}行 × {column_count:,}列を読み込みました")

        info_cols = st.columns(4 if selected_sheet else 3)
        info_cols[0].metric("ファイル名", uploaded_file.name)
        info_cols[1].metric("ファイル形式", file_format.upper())
        info_cols[2].metric("サイズ", f"{row_count:,} × {column_count:,}")
        if selected_sheet:
            info_cols[3].metric("シート", selected_sheet)

        st.subheader("データプレビュー")
        st.caption("先頭30行を表示しています。")
        st.dataframe(dataframe.head(30), use_container_width=True)

        st.subheader("列情報")
        st.dataframe(summarize_columns(dataframe), use_container_width=True, hide_index=True)

    except ExcelFlowError as exc:
        st.error(exc.user_message)
    except Exception:
        st.error("ファイルの読み込み中に予期しないエラーが発生しました。")

st.divider()

# --- STEP 2 ---
st.header("② データを整える")
st.write("不要列の削除、空白行・重複行の除去、文字列の前後空白除去などを行います。")
st.warning("Phase 3 で実装予定 — 現在準備中")

st.divider()

# --- STEP 3 ---
st.header("③ 集計する")
st.write("部署・商品・地域などのグループ項目で、合計・平均・件数などを集計します。")
st.warning("Phase 4 で実装予定 — 現在準備中")

st.divider()

# --- STEP 4 ---
st.header("④ グラフで確認する")
st.write("集計結果をもとに棒グラフ・折れ線グラフ・円グラフを表示します。")
st.warning("Phase 5 で実装予定 — 現在準備中")

st.divider()

# --- STEP 5 ---
st.header("⑤ Excelへ出力する")
st.write("整形済データ・集計結果・処理履歴を Excel ファイルとしてダウンロードします。")
st.warning("Phase 6 で実装予定 — 現在準備中")

st.divider()
st.caption("ExcelFlow Ver1.0 — Phase 2: ファイル読込")
