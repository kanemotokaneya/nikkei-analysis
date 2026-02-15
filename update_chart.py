import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import numpy as np
import requests
from bs4 import BeautifulSoup
import io

# --- 1. データの取得と洗浄 ---
try:
    df = yf.download("^N225", period="6mo", interval="1d")
    df = df.dropna().copy()
    for col in ['Open', 'High', 'Low', 'Close']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna().astype(float)
    close_p = float(df['Close'].iloc[-1])
    yesterday_p = float(df['Close'].iloc[-2])
except:
    close_p, yesterday_p = 38000.0, 38000.0

vi_value = 20.0
try:
    df_vi = yf.download("^JNIV", period="5d", interval="1d")
    if not df_vi.empty:
        vi_value = float(df_vi['Close'].dropna().iloc[-1])
except:
    pass

# --- 2. 予測値幅の計算 ---
daily_range = close_p * (vi_value / 100) / np.sqrt(250)
weekly_range = close_p * (vi_value / 100) / np.sqrt(52)

# --- 3. JPXデータの抽出 ---
oi = {}
opt_rows = ""
try:
    jpx_url = "https://www.jpx.co.jp/markets/derivatives/trading-volume/index.html"
    res = requests.get(jpx_url, timeout=10)
    soup = BeautifulSoup(res.text, "html.parser")
    link = soup.find("a", href=lambda h: h and "open_interest" in h and h.endswith(".xlsx"))
    
    if link:
        excel_url = "https://www.jpx.co.jp" + link.get("href")
        excel_res = requests.get(excel_url, timeout=10)
        df_jpx = pd.read_excel(io.BytesIO(excel_res.content), header=None)
        
        oi['large_all'] = str(df_jpx.iloc[48, 4])
        oi['large_mar'] = str(df_jpx.iloc[29, 4])
        oi['mini_all']  = str(df_jpx.iloc[51, 11])
        oi['mini_mar']  = str(df_jpx.iloc[35, 11])
        oi['topix_all'] = str(df_jpx.iloc[62, 4])
        oi['topix_mar'] = str(df_jpx.iloc[49, 4])

        atm = int(round(close_p / 500) * 500)
        strike_range = range(atm + 5000, atm - 5500, -500)
        for strike in strike_range:
            bg_color = "#fff5f5" if strike == atm else "white"
            opt_rows += f"<tr style='background-color:{bg_color};'><td>{strike:,}</td><td>-</td><td>-</td></tr>"
except:
    for k in ['large_all','large_mar','mini_all','mini_mar','topix_all','topix_mar']:
        oi[k] = "取得失敗"
    opt_rows = "<tr><td colspan='3'>データ抽出エラー</td></tr>"

# --- 4. チャート作成 ---
plot_df = df.tail(50)
mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit')
s = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)
mpf.plot(plot_df, type='candle', style=s, figsize=(12, 8), savefig='nikkei_chart.png')

# --- 5. ページ別にHTMLを書き出し ---
# ① Topページ用データ
top_html = f"""
    <div class='analysis-box'>
        <h3>市場概況</h3>
        <p><b>日経平均:</b> {close_p:,.0f}円 (<span style='color:{"red" if close_p >= yesterday_p else "blue"}'>{close_p - yesterday_p:+.0f}円</span>)</p>
        <p><b>日経VI:</b> {vi_value:.2f}</p>
        <p><b>デイリー予測:</b> {close_p - daily_range:,.0f} ～ {close_p + daily_range:,.0f}円</p>
    </div>
"""
with open("info.html", "w", encoding="utf-8") as f: f.write(top_html)

# ② 詳細ページ用データ
detail_html = f"""
    <div class='analysis-box'>
        <h3>先物建玉残高（前日比）</h3>
        <table style='width:100%; border-collapse: collapse; margin-bottom: 20px;'>
            <tr style='background:#eee;'><th>銘柄</th><th>全体建玉</th><th>3月限</th></tr>
            <tr><td>日経225(ラージ)</td><td>{oi['large_all']}</td><td>{oi['large_mar']}</td></tr>
            <tr><td>日経225 mini</td><td>{oi['mini_all']}</td><td>{oi['mini_mar']}</td></tr>
            <tr><td>TOPIX</td><td>{oi['topix_all']}</td><td>{oi['topix_mar']}</td></tr>
        </table>
        <h3>オプション建玉分布</h3>
        <table style='width:100%; border-collapse: collapse;'>
            <tr style='background:#4a90e2; color:white;'><th>権利行使価格</th><th>プット</th><th>コール</th></tr>
            {opt_rows}
        </table>
    </div>
"""
with open("details_info.html", "w", encoding="utf-8") as f: f.write(detail_html)
