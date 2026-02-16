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

# 1. データの取得
try:
    df_s = pd.read_csv(SHEET_CSV_URL, header=None)
    price = clean_val(df_s.iloc[0, 0])
    # もし価格が異常に大きい（1329 ETFなどの可能性）場合は調整を試みる
    if price > 50000: price = price / 1.45 # 暫定的な調整（必要に応じて修正）
    
    change = clean_val(df_s.iloc[0, 1])
    # C1セルから日経VIを取得
    vi_val = clean_val(df_s.iloc[0, 2])
    if vi_val == 0: vi_val = 20.0 # 失敗時のデフォルト
except:
    price, change, vi_val = 0.0, 0.0, 20.0

# 2. チャート作成（移動平均線 5日・25日）
try:
    url = "https://stooq.com/q/d/l/?s=^ni225&i=d"
    res = requests.get(url, timeout=15).content
    df = pd.read_csv(io.StringIO(res.decode("utf-8")), index_col=0, parse_dates=True)
    df.columns = [c.capitalize() for c in df.columns]
    
    # 移動平均線の計算
    df['MA5'] = df['Close'].rolling(window=5).mean()
    df['MA25'] = df['Close'].rolling(window=25).mean()
    
    plt.figure(figsize=(10, 6))
    plot_df = df.tail(60) # 直近60日分を表示
    
    plt.plot(plot_df.index, plot_df['Close'], label='Nikkei 225', color='#1f77b4', linewidth=2)
    plt.plot(plot_df.index, plot_df['MA5'], label='5-Day MA', color='green', linestyle='--', alpha=0.8)
    plt.plot(plot_df.index, plot_df['MA25'], label='25-Day MA', color='orange', linestyle='--', alpha=0.8)
    
    plt.title("Nikkei 225 & Moving Averages")
    plt.grid(True, alpha=0.3)
    plt.legend()
    plt.tight_layout()
    plt.savefig('nikkei_chart.png')
except:
    pass

# 3. HTML出力
color = "red" if change >= 0 else "blue"
p_display = f"{price:,.0f}円" if price > 0 else "取得エラー"

top_html = f"""
<div class='analysis-box'>
    <h2 style='font-size:2.8em; margin:0;'>{p_display}</h2>
    <p style='font-size:1.4em; margin-top:5px;'>
        前日比: <span style='color:{color}; font-weight:bold;'>{change:+.0f}円</span>
    </p>
    <p><b>日経VI:</b> {vi_val:.2f} | <b>ボラティリティ判定:</b> {'高' if vi_val > 25 else '安定'}</p>
</div>
"""
with open("info.html", "w", encoding="utf-8") as f:
    f.write(top_html)
