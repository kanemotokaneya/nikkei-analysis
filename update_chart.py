import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import numpy as np
import requests
from bs4 import BeautifulSoup
import io

# --- 1. データの取得（YahooがダメならStooqで試行） ---
def get_data():
    symbols = ["^N225", "NKY.JP"] # Yahoo用とStooq用
    for sym in symbols:
        try:
            print(f"Trying to download {sym}...")
            if sym == "^N225":
                data = yf.download(sym, period="1y", interval="1d")
            else:
                # Yahooがダメな場合のバックアップ（Stooq経由）
                url = f"https://stooq.com/q/d/l/?s={sym}&i=d"
                res = requests.get(url).content
                data = pd.read_csv(io.StringIO(res.decode("utf-8")), index_col=0, parse_dates=True)
            
            if not data.empty:
                # 徹底洗浄
                for col in ['Open', 'High', 'Low', 'Close']:
                    data[col] = pd.to_numeric(data[col], errors='coerce')
                return data.dropna().astype(float)
        except:
            continue
    return pd.DataFrame()

df = get_data()

# 最新値の確定
if not df.empty:
    close_p = float(df['Close'].iloc[-1])
    yesterday_p = float(df['Close'].iloc[-2])
else:
    # 最悪の事態のダミー
    close_p, yesterday_p = 38000.0, 38000.0

# 日経VI（取得失敗時は20.0固定）
vi_value = 20.0
try:
    df_vi = yf.download("^JNIV", period="5d", interval="1d")
    if not df_vi.empty:
        vi_value = float(df_vi['Close'].dropna().iloc[-1])
except:
    pass

# --- 2. 予測値幅の計算 ---
daily_range = close_p * (vi_value / 100) / np.sqrt(250)

# --- 3. JPXデータの抽出 ---
oi = {'large_all': '-', 'large_mar': '-', 'mini_all': '-', 'mini_mar': '-', 'topix_all': '-', 'topix_mar': '-'}
try:
    jpx_url = "https://www.jpx.co.jp/markets/derivatives/trading-volume/index.html"
    res = requests.get(jpx_url, timeout=10)
    soup = BeautifulSoup(res.text, "html.parser")
    link = soup.find("a", href=lambda h: h and "open_interest" in h and h.endswith(".xlsx"))
    if link:
        excel_res = requests.get("https://www.jpx.co.jp" + link.get("href"), timeout=10)
        df_jpx = pd.read_excel(io.BytesIO(excel_res.content), header=None)
        oi['large_all'] = str(df_jpx.iloc[48, 4])
        oi['large_mar'] = str(df_jpx.iloc[29, 4])
        oi['mini_all']  = str(df_jpx.iloc[51, 11])
        oi['mini_mar']  = str(df_jpx.iloc[35, 11])
        oi['topix_all'] = str(df_jpx.iloc[62, 4])
        oi['topix_mar'] = str(df_jpx.iloc[49, 4])
except:
    pass

# --- 4. チャート作成（移動平均線付き） ---
try:
    plot_df = df.tail(100).copy()
    mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit')
    s  = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)
    
    mpf.plot(plot_df, type='candle', style=s, mav=(5, 25),
             title="Nikkei 225 & MA", figsize=(12, 8), savefig='nikkei_chart.png')
except:
    plt.figure(figsize=(12, 8))
    plt.plot(df['Close'].tail(50))
    plt.savefig('nikkei_chart.png')

# --- 5. HTML書き出し ---
top_html = f"""
<div class='analysis-box'>
    <h2 style='font-size:2.5em; margin:0;'>{close_p:,.0f}円</h2>
    <p>前日比: <span style='color:{"red" if close_p >= yesterday_p else "blue"}'>{close_p - yesterday_p:+.0f}円</span></p>
    <p><b>日経VI:</b> {vi_value:.2f} | <b>予測値幅:</b> ±{daily_range:,.0f}円</p>
</div>
"""
with open("info.html", "w", encoding="utf-8") as f: f.write(top_html)

detail_html = f"""
<div class='analysis-box'>
    <h3>先物建玉</h3>
    <table border='1' style='width:100%; border-collapse:collapse; text-align:center;'>
        <tr style='background:#eee;'><th>銘柄</th><th>全体</th><th>3月限</th></tr>
        <tr><td>ラージ</td><td>{oi['large_all']}</td><td>{oi['large_mar']}</td></tr>
        <tr><td>mini</td><td>{oi['mini_all']}</td><td>{oi['mini_mar']}</td></tr>
        <tr><td>TOPIX</td><td>{oi['topix_all']}</td><td>{oi['topix_mar']}</td></tr>
    </table>
</div>
"""
with open("details_info.html", "w", encoding="utf-8") as f: f.write(detail_html)
