import pandas as pd
import matplotlib.pyplot as plt
import requests
import io
import re

# --- スプレッドシートURL ---
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1uXxxC3untThuWdyCkIsDR8yc9X3JZF-00tvTkwNWDCE/export?format=csv&gid=0"

def clean_val(val):
    if pd.isna(val): return 0.0
    # 数字、ドット、マイナス以外を排除
    c = re.sub(r'[^0-9.\-]', '', str(val))
    try:
        return float(c) if c else 0.0
    except:
        return 0.0

# 1. データの取得
try:
    df_s = pd.read_csv(SHEET_CSV_URL, header=None)
    # A1: 価格, B1: 前日比, C1: 日経VI
    price = clean_val(df_s.iloc[0, 0])
    change = clean_val(df_s.iloc[0, 1])
    vi_val = clean_val(df_s.iloc[0, 2])
    
    # 取得失敗時のガード
    if price == 0: price = 38000.0
    if vi_val == 0: vi_val = 20.0
except:
    price, change, vi_val = 38000.0, 0.0, 20.0

# 2. チャート作成（5日・25日移動平均線付き）
try:
    url = "https://stooq.com/q/d/l/?s=^ni225&i=d"
    res = requests.get(url, timeout=15).content
    df = pd.read_csv(io.StringIO(res.decode("utf-8")), index_col=0, parse_dates=True)
    df.columns = [c.capitalize() for c in df.columns]
    
    # 移動平均の計算
    df['MA5'] = df['Close'].rolling(window=5).mean()
    df['MA25'] = df['Close'].rolling(window=25).mean()
    
    plt.figure(figsize=(10, 6))
    plot_df = df.tail(60) # 直近60日分
    
    plt.plot(plot_df.index, plot_df['Close'], label='日経平均', color='#1f77b4', linewidth=2)
    plt.plot(plot_df.index, plot_df['MA5'], label='5日線', color='green', alpha=0.7)
    plt.plot(plot_df.index, plot_df['MA25'], label='25日線', color='orange', alpha=0.7)
    
    plt.title("Nikkei 225 & Moving Average")
    plt.grid(True, alpha=0.3)
    plt.legend()
    plt.savefig('nikkei_chart.png')
except:
    pass

# 3. HTML出力
color = "red" if change >= 0 else "blue"
top_html = f"""
<div class='analysis-box'>
    <h2 style='font-size:2.8em; margin:0;'>{price:,.0f}円</h2>
    <p style='font-size:1.4em; margin-top:5px;'>
        前日比: <span style='color:{color}; font-weight:bold;'>{change:+.0f}円</span>
    </p>
    <p><b>日経VI:</b> {vi_val:.2f}</p>
</div>
"""
with open("info.html", "w", encoding="utf-8") as f:
    f.write(top_html)

    
