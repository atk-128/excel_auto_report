# Excel 売上レポート自動作成ツール

## 概要
複数の Excel（.xlsx）ファイルを読み込み、  
日別売上・商品別売上・売上TOP5 を自動で集計し、  
ExcelおよびCSV形式でレポートを出力する Python ツールです。

Excelでの手作業集計を自動化することを目的としています。

---

## フォルダ構成
---

## 機能

- 複数のExcelファイル（.xlsx）を自動で結合
- 日別売上の集計
- 商品別売上の集計
- 売上TOP5商品の自動抽出
- 集計結果をExcel（複数シート）で出力
- 集計結果をCSV形式でも出力

---

## 出力内容

### Excelファイル
- 結合データ
- 日別売上
- 商品別売上
- 売上TOP5

### CSVファイル
- daily_sales.csv（日別売上）
- product_sales.csv（商品別売上）
- top5_products.csv（売上TOP5）
---

## 売上TOP5グラフ（自動生成）

![売上TOP5](output/top5_products.png)

---

## 売上TOP5グラフ（自動生成）

![売上TOP5](output/top5_products.png)

---

## 使い方

1. `input/` フォルダに集計したいExcel（.xlsx）を入れる
2. 依存ライブラリをインストール
※ output フォルダにレポートファイルが生成されていれば正常に実行されています。

```bash
pip3 install pandas openpyxl matplotlib