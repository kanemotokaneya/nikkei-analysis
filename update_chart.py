import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import numpy as np
import datetime
import requests
from bs4 import BeautifulSoup
import io
import os

# --- 1. データの取得と洗浄 ---
try:
    df = yf.download("^N225", period="6mo", interval="1d")
    df = df.dropna().copy()
    # データの型を強制的に数値に変換（ローソク足エラー対策）
    for col in ['Open', 'High', 'Low', 'Close']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna().astype(float)
    
    close_p = float(df['Close'].iloc[-1])
    yesterday_p = float(df['Close'].iloc[-2])
except Exception as e:
    print(f"Data Error: {e}")
    close_p, yesterday_p = 38000.0, 38000.0

# 日経VI（取れない場合は20.0固定）
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
oi_all, oi_mar = "更新待ち", "更新待ち"
try:
    jpx_url = "https://www.jpx.co.jp/markets/derivatives/trading-volume/index.html"
    res = requests.get(jpx_url, timeout=10)
    soup = BeautifulSoup(res.text, "html.parser")
    link = soup.find("a", href=lambda h: h and "open_interest" in h and h.endswith(".xlsx"))
    if link:
        excel_url = "https://www.jpx.co.jp" + link.get("href")
        excel_res = requests.get(excel_url, timeout=10)
        df_jpx = pd.read_excel(io.BytesIO(excel_res.content), header=None)
        oi_all = str(df_jpx.iloc[48, 4]) # E列49行
        oi_mar = str(df_jpx.iloc[29, 4]) # E列30行
except:
    pass

# --- 4. ローソク足チャートの作成 ---
try:
    plot_df = df.tail(50)
    mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit')
    s = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)
    mpf.plot(plot_df, type='candle', style=s, 
             title=f"Nikkei 225 (VI: {vi_value:.2f})",
             ylabel='Price (JPY)', figsize=(12, 8), savefig='nikkei_chart.png')
except:
    # 失敗した場合は単純なラインチャートで代用
    plt.figure(figsize=(12, 8))
    plt.plot(df['Close'].tail(50))
    plt.savefig('nikkei_chart.png')

# --- 5. 解析結果をHTML形式で書き出し ---
# ここで info.html を確実に作成します
report_content = f"""
    <div class='analysis-box'>
        <h3 style='color: #2c3e50; margin-top:0;'>① 日経平均・VI分析</h3>
        <p><b>終値:</b> {close_p:,.0f}円 (<span style='color:{"red" if close_p >= yesterday_p else "blue"}'>{close_p - yesterday_p:+.0f}円</span>)</p>
        <p><b>日経VI:</b> {vi_value:.2f}</p>
        <p><b>デイリー予測:</b> {close_p - daily_range:,.0f} ～ {close_p + daily_range:,.0f}円</p>
        <p><b>ウィークリー予測:</b> {close_p - weekly_range:,.0f} ～ {close_p + weekly_range:,.0f}円</p>
        
        <h3 style='color: #2c3e50;'>② 先物建玉残高（ラージ）</h3>
        <p><b>全体建玉:</b> {oi_all}</p>
        <p><b>3月限建玉:</b> {oi_mar}</p>
    </div>
"""
with open("info.html", "w", encoding="utf-8") as f:
    f.write(report_content)
