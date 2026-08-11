# ExcelFlow

Excel / CSV 形式の業務データをブラウザ上で読み込み、整形・集計・グラフ化・Excel出力まで一連で実行する業務自動化ツールです。

## 主な予定機能（Ver1.0）

1. ファイルを読み込む（Excel / CSV）
2. データを整える（列削除、空白・重複除去など）
3. 集計する（グループ集計）
4. グラフで確認する（棒・折れ線・円）
5. Excelへ出力する（整形済データ・集計結果・処理履歴）

## 使用技術

- Python
- Streamlit
- pandas
- openpyxl
- Plotly
- pytest

## 現在の開発状況

**Phase 1：プロジェクト基盤** — 完了  
**Phase 2：ファイル読込** — 実装済み

現在利用可能な機能:

- Excel（.xlsx）読込
- CSV読込（UTF-8 / UTF-8 BOM / CP932）
- Excelシート選択
- 行数・列数の表示
- データプレビュー
- 列名 / データ型の確認

Phase 3以降（データ整形・集計・グラフ・Excel出力）は未実装です。

## ローカル起動方法

```bash
# 依存関係のインストール
pip install -r requirements.txt

# アプリ起動
streamlit run app.py
```

サンプルデータ: `sample_data/sample_sales.xlsx` / `sample_data/sample_sales.csv`

## テスト実行方法

```bash
python -m pytest -q
```
