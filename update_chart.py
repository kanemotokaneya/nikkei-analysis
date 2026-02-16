import pandas as pd
import matplotlib.pyplot as plt
import requests
import io
import re

# スプレッドシートのURL
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1uXxxC3untThuWdyCkIsDR8yc9X3JZF-00tvTkwNWDCE/pub?output=csv"

def clean_val(val):
    if pd.isna(val): return 0.0
    c = re.sub(r'[^0-9.\-]', '', str(val))
    try: return float(c)
    except: return 0.0

# 1. データ取得
try:
    df_s = pd.read_csv(SHEET_CSV_URL, header=None)
    price = clean_val(df_s.iloc[0, 0])
    change = clean_val(df_s.iloc[0, 1])
    # パーセント表記か判定
    if abs(change) < 1 and change != 0:
        c_abs = price * change
        c_pct = change
    else:
        c_abs = change
        c_pct = (change / (price - change)) if (price - change) != 0 else 0
except:
    price, c_abs, c_pct = 0.0, 0.0, 0.0

# 2. チャート作成
try:
    url = "https://stooq.com/q/d/l/?s=^ni225&i=d"
    res = requests.get(url, timeout=15).content
    df = pd.read_csv(io.StringIO(res.decode("utf-8")), index_col=0, parse_dates=True)
    df.columns = [c.capitalize() for c in df.columns]
    plt.figure(figsize=(10, 6))
    plt.plot(df.index.tail(100), df['Close'].tail(100), color='#1f77b4')
    plt.grid(True, alpha=0.3)
    plt.savefig('nikkei_chart.png')
except:
    pass

# 3. HTML出力 (エラーが起きにくい形式)
color = "red" if c_abs >= 0 else "blue"
p_text = f"{price:,.0f}円" if price > 0 else "データ取得中"

top_html = f"""
<div class='analysis-box'>
    <h2 style='font-size:2.8em; margin:0;'>{p_text}</h2>
    <p style='font-size:1.4em; margin-top:5px;'>
        前日比: <span style='color:{color}; font-weight:bold;'>
            {c_abs:+.0f}円 ({c_pct:+.2%})
        </span>
    </p>
</div>
"""
with open("info.html", "w", encoding="utf-8") as f:
    f.write(top_html)
