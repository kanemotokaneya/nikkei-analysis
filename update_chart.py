import pandas as pd
import matplotlib.pyplot as plt
import requests
import io
import numpy as np

# --- 1. スプレッドシート（Google経由）から価格を取得 ---
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1uXxxC3untThuWdyCkIsDR8yc9X3JZF-00tvTkwNWDCE/pub?output=csv"

def get_stable_data():
    try:
        # スプレッドシートから現在値と前日比を読み込む
        df_sheet = pd.read_csv(SHEET_CSV_URL, header=None)
        # スプレッドシートのA1に現在値、B1に前日比(%)が入っている想定
        price = float(df_sheet.iloc[0, 0])
        change_pct = float(df_sheet.iloc[0, 1])
        change_abs = price * change_pct
        return price, change_abs
    except Exception as e:
        print(f"Sheet Error: {e}")
        return 38000.0, 0.0

# --- 2. チャートデータ取得 ---
def get_chart_data():
    try:
        url = "https://stooq.com/q/d/l/?s=^ni225&i=d"
        res = requests.get(url, timeout=15).content
        df = pd.read_csv(io.StringIO(res.decode("utf-8")), index_col=0, parse_dates=True)
        df.columns = [c.capitalize() for c in df.columns]
        return df.tail(100)
    except:
        return pd.DataFrame()

close_p, change_p = get_stable_data()
df = get_chart_data()

# --- 3. チャート作成 ---
plt.figure(figsize=(12, 7))
if not df.empty:
    plt.plot(df.index, df['Close'], color='#1f77b4', linewidth=2)
    plt.title("Market Overview")
    plt.grid(True, linestyle=':', alpha=0.6)
plt.tight_layout()
plt.savefig('nikkei_chart.png')

# --- 4. HTML書き出し ---
top_html = f"""
<div class='analysis-box'>
    <h2 style='font-size:2.8em; margin:0;'>{close_p:,.0f}円</h2>
    <p style='font-size:1.5em; margin-top:5px;'>前日比: <span style='color:{"red" if change_p >= 0 else "blue"}'>{change_p:+.0f}円</span></p>
</div>
"""
with open("info.html", "w", encoding="utf-8") as f: f.write(top_html)
