# ExcelFlow

## 概要

Excel / CSV の読込・整形・集計・グラフ・Excel出力を行う、Streamlit ベースの業務自動化デモツールです。

## デモ

Streamlit Community Cloud で公開しています。

https://excelflow-demo.streamlit.app

Excel / CSV ファイルをアップロードして、読込・整形・集計・グラフ・Excel出力をブラウザ上で試せます。

## 主な機能

- Excel / CSV 読込
- データプレビュー
- データ整形
- 集計
- グラフ
- Excel 出力
- 処理履歴

## 操作フロー

1. ファイル読込
2. データ整形
3. 集計
4. グラフ
5. Excel 出力

Step 2（データ整形）は任意です。スキップした場合、集計・出力は元データを対象にします。

## 対応形式

- `.xlsx`（Excel）
- `.csv`（UTF-8 / UTF-8 BOM / CP932）

## ローカル起動

```bash
pip install -r requirements.txt
streamlit run app.py
```

Windows では `run_ExcelFlow.bat` からも起動できます。

## テスト

```bash
python -m pytest -q
```

## サンプルデータ

`sample_data` 内のサンプルを利用できます。

- `sample_data/sample_sales.xlsx`
- `sample_data/sample_sales.csv`

サンプルデータには、デモ用に意図的に空白・重複・前後空白を含めています。個人情報や実在の会社情報は含みません。

## 技術

- Python
- Streamlit
- pandas
- openpyxl
- Plotly
- pytest
