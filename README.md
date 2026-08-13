# ExcelFlow

Excel / CSV の読込・整形・集計・グラフ・Excel出力を、ブラウザ上で行える業務自動化ツールです。

## 主な機能

1. **ファイル読込** — Excel（.xlsx） / CSV（UTF-8 / UTF-8 BOM / CP932）、シート選択、プレビュー
2. **データ整形** — 列削除、列名変更、空白行・重複行削除、文字列前後空白除去、欠損値確認
3. **集計** — 最大2項目のグループ集計（合計 / 平均 / 件数 / 最大 / 最小）、昇順・降順
4. **グラフ** — 棒グラフ / 折れ線グラフ / 円グラフ（Plotly）
5. **Excel出力** — 整形済データ・集計結果・処理履歴を `.xlsx` でダウンロード

## 操作フロー

```
Step 1 ファイル読込
  → Step 2 データ整形（任意）
  → Step 3 集計
  → Step 4 グラフ
  → Step 5 Excel出力
```

Step 2 をスキップした場合、集計・出力は元データを対象にします。

## 対応ファイル

| 形式 | 拡張子 | 備考 |
|------|--------|------|
| Excel | `.xlsx` | 複数シート時は画面で選択 |
| CSV | `.csv` | UTF-8 / UTF-8 BOM / CP932 |

サンプルデータ: `sample_data/sample_sales.xlsx` / `sample_data/sample_sales.csv`

## 使用技術

- Python
- Streamlit
- pandas
- openpyxl
- Plotly
- pytest

## 起動方法

```bash
# 依存関係のインストール
pip install -r requirements.txt

# アプリ起動
streamlit run app.py
```

Windows では `run_ExcelFlow.bat` をダブルクリックしても起動できます。

## テスト方法

```bash
python -m pytest -q
```
