import pandas as pd
import matplotlib.pyplot as plt
import requests
import io
import re
import numpy as np

# --- 1. スプレッドシート（Google経由）から価格を取得 ---
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1uXxxC3untThuWdyCkIsDR8yc9X3JZF-00tvTkwNWDCE/pub?output=csv"

def clean_value(val):
    if pd.isna(val): return 0.0
    cleaned = re.sub(r'[^0-9.\-]', '', str(val))
    try:
        return float(cleaned)
    except:
        return 0.0

def get_stable_data():
    try:
        # スプレッドシート読み込み
        df_sheet = pd.read_csv(SHEET_CSV_URL, header=None)
        # A1: 現在値, B1: 前日比%
        price = clean_value(df_sheet.iloc[0, 0])
        change_pct = clean_value(df_sheet.iloc[0, 1])
        
        # パーセント表示（0.01等）を考慮して前日差を計算
        if abs(change_pct) < 1 and change_pct != 0:
            change_abs = price * change_pct
        else:
            change_abs = change_pct
            change_pct = (change_pct / (price - change_pct)) if (price - change_pct) != 0 else 0
            
        return price, change_abs, change_pct
    except:
        return 38000.0, 0.0, 0.0

close_p, change_abs, change_pct = get_stable_data()

# --- 2. チャートデータ取得 (Stooq) ---
def get_chart_data():
    try:
        url = "https://stooq.com/q/d/l/?s=^ni225&i=d"
        res = requests.get(url, timeout=15).content
        df = pd.read_csv(io.StringIO(res.decode("utf-8")), index_col=0, parse_dates=True)
        # 列名を正規化
        df.columns = [c.capitalize() for c in df.columns]
        return df.tail(100)
    except:
        return pd.DataFrame()

df = get_chart_data()

# --- 3. チャート作成 ---
plt.figure(figsize=(10, 6))
if not df.empty and 'Close' in df.columns:
    plt.plot(df.index, df['Close'], color='#1f77b4', linewidth=2, label='Nikkei 225')
    plt.title("Nikkei 225 Market Overview")
    plt.grid(True, linestyle=':', alpha=0.6)
    plt.legend()
else:
    plt.text(0.5, 0.5, "Chart Data Loading...", ha='center')
plt.tight_layout()
plt.savefig('nikkei_chart.png')

# --- 4. HTML書き出し ---
color = "red" if change_abs >= 0 else "blue"
price_display = f"{close_p:,.0f}円" if close_p >
