# Excel 売上レポート自動作成ツール

## 概要
`input/` にある複数のExcel（.xlsx）を読み込み・結合し、
日別売上・商品別売上を自動で集計して `output/` にレポートを出力します。

## 使い方
1. `input/` に集計したいExcel（.xlsx）を入れる
2. 依存ライブラリをインストール
   ```bash
   pip3 install pandas openpyxl
   