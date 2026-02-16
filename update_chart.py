import pandas as pd
import matplotlib.pyplot as plt
import requests
import io
import re

# スプレッドシートURL
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1uXxxC3untThuWdyCkIsDR8yc9X3JZF-00tvTkwNWDCE/export?format=csv&gid=0"

def clean_val(val):
    if pd.isna(val): return 0.0
    c = re.sub(r'[^0-9.\-]', '', str(val))
    try: return float(c)
    except: return 0.0

# --- 1. 価格とVIの取得 ---
try:
    df_s = pd.read_csv(SHEET_CSV_URL, header=None)
    price = clean_val(df_s.iloc[0, 0])
    change = clean_val(df_s.iloc[0, 1])
    
    # スプレッドシート(C1)からVI取得を試みる
    vi_val = clean_val(df_s.iloc[0, 2])
    
    # スプレッドシートが0ならPythonで直接「みんかぶ」からVIを取得
    if vi_val == 0:
        headers = {"User-Agent": "Mozilla/5.0"}
        vi_res = requests.get("https://minkabu.jp/stock/NI225VI", headers=headers, timeout=10)
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(vi_res.text, "html.parser")
        vi_text = soup.find("div", class_="stock_price").text
        vi_val = clean_val(vi_text)
except:
    price, change, vi_val = 0.0, 0.0, 20.0

# --- 2. チャート作成（移動平均線） ---
[attachment_0](attachment)
try:
    url = "https://stooq.com/q/d/l/?s=^ni225&i=d"
    res = requests.get(url, timeout=15).content
    df = pd.read_csv(io.StringIO(res.decode("utf-8")), index_col=0, parse_dates=True)
    df.columns = [c.capitalize() for c in df.columns]
    
    # 移動平均線(5日, 25日)を計算
    df['MA5'] = df['Close'].rolling(window=5).mean()
    df['MA25'] = df['Close'].rolling(window=25).mean()
    
    plt.figure(figsize=(10, 6))
    p_df = df.tail(60) # 直近2ヶ月
    plt.plot(p_df.index, p_df['Close'], label='Nikkei 225', color='#1f77b4', linewidth=2)
    plt.plot(p_df.index, p_df['MA5'], label='5MA', color='green', alpha=0.7)
    plt.plot(p_df.index, p_df['MA25'], label='25MA', color='orange', alpha=0.7)
    
    plt.title("Nikkei 225 Market Dashboard")
    plt.grid(True, linestyle=':', alpha=0.5)
    plt.legend()
    plt.savefig('nikkei_chart.png')
except:
    pass

# --- 3. HTML書き出し ---
color = "red" if change >= 0 else "blue"
top_html = f"""
<div class='analysis-box'>
    <h2 style='font-size:2.8em; margin:0;'>{price:,.0f}円</h2>
    <p style='font-size:1.4em; margin-top:5px;'>前日比: <span style='color:{color};'>{change:+.0f}円</span></p>
    <hr>
    <p><b>日経VI:</b> {vi_val:.2f} ({'警戒' if vi_val > 25 else '安定'})</p>
    <p><b>テクニカル:</b> 5日線/25日線表示中</p>
</div>
"""
with open("info.html", "w", encoding="utf-8") as f:
    f.write(top_html)
