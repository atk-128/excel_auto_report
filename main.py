import pandas as pd
import glob
import os
from datetime import datetime

# =========================
# 設定（ここだけ触ればOK）
# =========================
INPUT_DIR = 'input'
OUTPUT_DIR = 'output'
OUTPUT_NAME = f"report_all_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
REQUIRED_COLUMNS = {'日付', '商品名', '単価', '数量', '売上'}

# =========================
# 準備
# =========================
os.makedirs(OUTPUT_DIR, exist_ok=True)

files = glob.glob(os.path.join(INPUT_DIR, '*.xlsx'))
if not files:
    raise FileNotFoundError(f'❌ Excelが見つかりません: {INPUT_DIR} に .xlsx を入れてください')

dfs = []

# =========================
# 読み込み＆検証
# =========================
for f in files:
    df = pd.read_excel(f)

    # 列チェック（安全装置）
    if not REQUIRED_COLUMNS.issubset(df.columns):
        raise ValueError(f'❌ 列が不足しています: {os.path.basename(f)}')

    df['元ファイル'] = os.path.basename(f)
    df['日付'] = pd.to_datetime(df['日付']).dt.date
    dfs.append(df)

# =========================
# 結合
# =========================
df_all = pd.concat(dfs, ignore_index=True)

# =========================
# 集計
# =========================
daily = (
    df_all.groupby('日付', as_index=False)['売上']
          .sum()
          .sort_values('日付')
)

product = (
    df_all.groupby('商品名', as_index=False)['売上']
          .sum()
          .sort_values('売上', ascending=False)
)

# =========================
# 出力
# =========================
output_path = os.path.join(OUTPUT_DIR, OUTPUT_NAME)
with pd.ExcelWriter(output_path) as writer:
    df_all.to_excel(writer, sheet_name='結合データ', index=False)
    daily.to_excel(writer, sheet_name='日別売上', index=False)
    product.to_excel(writer, sheet_name='商品別売上', index=False)

print('✅ 完了')
print('出力ファイル:', output_path)
print('処理ファイル数:', len(files))