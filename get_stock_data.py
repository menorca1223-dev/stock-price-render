import yfinance as yf
import pandas as pd
from datetime import datetime
import os

# ✅ 指定された銘柄コードとラベル
stocks = {
    "7203.T": "トヨタ自動車",
    "6460.T": "セガサミーHD",
    "3765.T": "ガンホー"
}

# 出力フォルダの作成
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)

# 日付付きファイル名
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
filename = f"{output_dir}/stock_data_{timestamp}.xlsx"

# ExcelWriterで複数シートに保存
with pd.ExcelWriter(filename, engine="openpyxl") as writer:
    for code, label in stocks.items():
        print(f"取得中: {label}（{code}）")

        # データ取得（過去1日分の1時間足）
        df = yf.Ticker(code).history(interval="60m", period="1d")

        if df.empty:
            print(f"⚠️ データなし: {label}（{code}）")
            continue

        df.reset_index(inplace=True)
        df.to_excel(writer, sheet_name=label, index=False)

print(f"✅ 完了: {filename}")
