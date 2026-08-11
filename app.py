"""ExcelFlow Streamlit UI entry point.

Business logic lives in the ``excel_flow`` package.
Phase 1 provides the screen shell and step placeholders only.
"""

import streamlit as st

st.set_page_config(
    page_title="ExcelFlow",
    page_icon="📊",
    layout="wide",
)

st.title("ExcelFlow")
st.caption("Excel / CSV 業務自動化ツール")
st.info("現在の開発状況：**Phase 1（プロジェクト基盤）** — 各機能はこれから実装します。")

st.divider()

# --- STEP 1 ---
st.header("① ファイルを読み込む")
st.write("Excel（.xlsx）または CSV をアップロードし、内容をプレビューします。")
st.warning("Phase 2 で実装予定 — 現在準備中")

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
st.caption("ExcelFlow Ver1.0 — Phase 1: プロジェクト基盤")
