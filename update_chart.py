import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import numpy as np
import datetime
import requests
from bs4 import BeautifulSoup
import io

# --- 1. データの取得 ---
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

# --- 3. JPXデータの抽出 (座標指定) ---
oi = {}
try:
    jpx_url = "https://www.jpx.co.jp/markets/derivatives/trading-volume/index.html"
    res = requests.get(jpx_url, timeout=10)
    soup = BeautifulSoup(res.text, "html.parser")
    link = soup.find("a", href=lambda h: h and "open_interest" in h and h.endswith(".xlsx"))
    
    if link:
        excel_url = "https://www.jpx.co.jp" + link.get("href")
        excel_res = requests.get(excel_url, timeout=10)
        df_jpx = pd.read_excel(io.BytesIO(excel_res.content), header=None)
        
        # 指定座標から抽出 (Pythonのインデックスは 0から始まるため、列-1, 行-1 で指定)
        # 日経225（ラージ）E列=4
        oi['large_all'] = str(df_jpx.iloc[48, 4]) # E列49行
        oi['large_mar'] = str(df_jpx.iloc[29, 4]) # E列30行
        
        # 日経225 mini L列=11
        oi['mini_all'] = str(df_jpx.iloc[51, 11]) # L列52行
        oi['mini_mar'] = str(df_jpx.iloc[35, 11]) # L列36行
        
        # TOPIX E列=4
        oi['topix_all'] = str(df_jpx.iloc[62, 4]) # E列63行
        oi['topix_mar'] = str(df_jpx.iloc[49, 4]) # E列50行
except:
    # 失敗した時の初期値
    keys = ['large_all', 'large_mar', 'mini_all', 'mini_mar', 'topix_all', 'topix_mar']
    for k in keys: oi[k] = "取得失敗"

# --- 4. ローソク足チャートの作成 ---
plot_df = df.tail(50)
mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit')
s = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)
mpf.plot(plot_df, type='candle', style=s, figsize=(12, 8), savefig='nikkei_chart.png')

# --- 5. 解析結果をHTML書き出し ---
report_content = f"""
    <div class='analysis-box'>
        <h3 style='color: #2c3e50; margin-top:0;'>① 日経平均・VI分析</h3>
        <p><b>終値:</b> {close_p:,.0f}円 (前日比：{close_p - yesterday_p:+.0f}円)</p>
        <p><b>予測(デイリー):</b> {close_p - daily_range:,.0f} ～ {close_p + daily_range:,.0f}円</p>
        
        <h3 style='color: #2c3e50;'>② 先物建玉残高（前日比）</h3>
        <table style='width:100%; border-collapse: collapse;'>
            <tr style='background:#eee;'><th>銘柄</th><th>全体建玉</th><th>3月限</th></tr>
            <tr><td>日経225(ラージ)</td><td>{oi['large_all']}</td><td>{oi['large_mar']}</td></tr>
            <tr><td>日経225 mini</td><td>{oi['mini_all']}</td><td>{oi['mini_mar']}</td></tr>
            <tr><td>TOPIX</td><td>{oi['topix_all']}</td><td>{oi['topix_mar']}</td></tr>
        </table>
    </div>
"""
with open("info.html", "w", encoding="utf-8") as f:
    f.write(report_content)
