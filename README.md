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
**Phase 2：ファイル読込** — 完了  
**Phase 3：データ整形** — 完了  
**Phase 4：集計** — 実装済み

現在利用可能な機能:

- Excel（.xlsx）読込
- CSV読込（UTF-8 / UTF-8 BOM / CP932）
- Excelシート選択
- データプレビュー
- 不要列削除
- 列名変更
- 空白行削除
- 重複行削除
- 文字列前後空白除去
- 欠損値確認
- 整形前後サマリー
- 最大2項目のグループ集計
- 合計 / 平均 / 件数 / 最大 / 最小
- 昇順 / 降順
- 元データ / 整形済データの自動選択

Phase 5以降（グラフ・Excel出力）は未実装です。

## ローカル起動方法

```bash
# 依存関係のインストール
pip install -r requirements.txt

# アプリ起動
streamlit run app.py

# Windowsではダブルクリックでも起動できます
run_ExcelFlow.bat
```

サンプルデータ: `sample_data/sample_sales.xlsx` / `sample_data/sample_sales.csv`

## テスト実行方法

```bash
python -m pytest -q
```
