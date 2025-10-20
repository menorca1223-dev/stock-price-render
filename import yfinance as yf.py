import os
from datetime import datetime
import pandas as pd
import yfinance as yf

# 取得対象の銘柄（Yahoo Financeのコード）
stocks = {
    "7203.T": "トヨタ自動車",
    "6460.T": "セガサミーHD",
    "3765.T": "ガンホー"
}

# 今日の日付をファイル名に使用
today = datetime.today().strftime('%Y_%m_%d')
filename = f"株価_{today}.xlsx"

# 保存先ディレクトリ（WSL形式のパスに修正）
save_dir = "/mnt/c/Users/atush/OneDrive/デスクトップ/株価"
os.makedirs(save_dir, exist_ok=True)  # フォルダがなければ作成

# フルパスでファイル名を指定
filepath = os.path.join(save_dir, filename)

# Excelファイルに各銘柄のデータを書き込む
with pd.ExcelWriter(filepath, engine='openpyxl', mode='a' if os.path.exists(filepath) else 'w') as writer:
    for code, label in stocks.items():
        try:
            df = yf.Ticker(code).history(interval="60m", period="1d")
            df.reset_index(inplace=True)
            df.to_excel(writer, sheet_name=label, index=False)
        except Exception as e:
            print(f"{label}（{code}）の取得に失敗しました: {e}")
