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

# --- 3. JPXデータの抽出 (先物・オプション) ---
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
        
        # 座標指定抽出
        oi['large_all'] = str(df_jpx.iloc[48, 4])
        oi['large_mar'] = str(df_jpx.iloc[29, 4])
        oi['mini_all']  = str(df_jpx.iloc[51, 11])
        oi['mini_mar']  = str(df_jpx.iloc[35, 11])
        oi['topix_all'] = str(df_jpx.iloc[62, 4])
        oi['topix_mar'] = str(df_jpx.iloc[49, 4])

        # オプション分布(±5000円)
        atm = int(round(close_p / 500) * 500)
        strike_range = range(atm + 5000, atm - 5500, -500)
        for strike in strike_range:
            bg_color = "#fff5f5" if strike == atm else "white"
            opt_rows += f"<tr style='background-color:{bg_color};'><td>{strike:,}</td><td>-</td><td>-</td></tr>"
except:
    for k in ['large_all','large_mar','mini_all','mini_mar','topix_all','topix_mar']:
        oi[k] = "取得失敗"
    opt_rows = "<tr><td colspan='3'>データ抽出中</td></tr>"

# --- 4. チャート作成 ---
mpf.plot(df.tail(50), type='candle', style='charles', figsize=(12, 8), savefig='nikkei_chart.png')

# --- 5. 2つのHTML用データを出力 ---
# ① info.html (TOP概況用)
top_html = f"""
    <div class='analysis-box'>
        <p><b>現在値:</b> {close_p:,.0f}円 (<span style='color:{"red" if close_p >= yesterday_p else "blue"}'>{close_p - yesterday_p:+.0f}円</span>)</p>
        <p><b>日経VI:</b> {vi_value:.2f}</p>
        <p><b>デイリー予測レンジ:</b> {close_p - daily_range:,.0f} ～ {close_p + daily_range:,.0f}円</p>
    </div>
"""
with open("info.html", "w", encoding="utf-8") as f: f.write(top_html)

# ② details_info.html (詳細ページ用)
detail_html = f"""
    <div class='analysis-box'>
        <h3 style='color:#4a90e2;'>■ 先物建玉（前日比）</h3>
        <table style='width:100%; border-collapse: collapse; margin-bottom: 20px;'>
            <tr style='background:#eee;'><th>銘柄</th><th>全体</th><th>3月限</th></tr>
            <tr><td>日経225(ラージ)</td><td>{oi['large_all']}</td><td>{oi['large_mar']}</td></tr>
            <tr><td>日経225 mini</td><td>{oi['mini_all']}</td><td>{oi['mini_mar']}</td></tr>
            <tr><td>TOPIX</td><td>{oi['topix_all']}</td><td>{oi['topix_mar']}</td></tr>
        </table>
        <h3 style='color:#4a90e2;'>■ オプション分布（±5000円）</h3>
        <table style='width:100%; border-collapse: collapse;'>
            <tr style='background:#4a90e2; color:white;'><th>権利行使価格</th><th>P建玉</th><th>C建玉</th></tr>
            {opt_rows}
        </table>
    </div>
"""
with open("details_info.html", "w", encoding="utf-8") as f: f.write(detail_html)
