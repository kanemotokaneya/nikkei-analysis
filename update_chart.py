import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import numpy as np
import datetime
import requests
from bs4 import BeautifulSoup
import io

# --- 1. データの取得と強力な洗浄 ---
try:
    # 日経平均(現物)を取得
    df = yf.download("^N225", period="6mo", interval="1d")
    # ローソク足描画のために、不純物（空欄や文字列）を徹底排除
    df = df.dropna().copy()
    for col in ['Open', 'High', 'Low', 'Close', 'Volume']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna().astype(float)
    
    close_p = float(df['Close'].iloc[-1])
    yesterday_p = float(df['Close'].iloc[-2])
except Exception as e:
    print(f"Stock Data Error: {e}")
    # 万が一取れなかった場合のダミー（エラー停止防止）
    close_p, yesterday_p = 38000.0, 38000.0

# 日経平均VI（予備の数値20.0を用意）
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

# --- 3. JPXデータの抽出 (エラー回避強化) ---
oi_data = {'large_all': '更新待ち', 'large_mar': '更新待ち'}
try:
    jpx_url = "https://www.jpx.co.jp/markets/derivatives/trading-volume/index.html"
    res = requests.get(jpx_url, timeout=10)
    soup = BeautifulSoup(res.text, "html.parser")
    # リンクの探し方をより柔軟に変更
    link = soup.find("a", href=lambda h: h and "open_interest" in h and h.endswith(".xlsx"))
    
    if link:
        excel_url = "https://www.jpx.co.jp" + link.get("href")
        excel_res = requests.get(excel_url, timeout=10)
        df_jpx = pd.read_excel(io.BytesIO(excel_res.content), header=None)
        # 指定座標から抽出（E列=4, 49行=48 / E列=4, 30行=29）
        oi_data['large_all'] = str(df_jpx.iloc[48, 4])
        oi_data['large_mar'] = str(df_jpx.iloc[29, 4])
except Exception as e:
    print(f"JPX Extraction Error: {e}")

# --- 4. ローソク足チャートの作成 ---
try:
    plot_df = df.tail(50)
    mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit')
    s  = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)
    
    # チャート作成と保存
    mpf.plot(plot_df, type='candle', style=s, 
             title=f"Nikkei 225 (VI Reference: {vi_value:.2f})",
             ylabel='Price (JPY)',
             figsize=(12, 8),
             savefig='nikkei_chart.png')
except Exception as e:
    print(f"Chart Creation Error: {e}")

# --- 5. 解析結果をHTML用に保存 ---
report_html = f"""
<div class='analysis-box'>
    <h3>① 日経平均・VI分析</h3>
    <p>日経平均終値：{close_p:,.0f}円 (前日比：{close_p - yesterday_p:+.0f}円)</p>
    <p>日経VI（参考値）：{vi_value:.2f}</p>
    <p>1日の予測値幅：{close_p - daily_range:,.0f}円 ～ {close_p + daily_range:,.0f}円</p>
    <p>1週間の予測値幅：{close_p - weekly_range:,.0f}円 ～ {close_p + weekly_range:,.0f}円</p>
    <h3>② 先物建玉残高（ラージ）</h3>
    <p>建玉残高(前日比)：{oi_data['large_all']}</p>
    <p>3月限(前日比)：{oi_data['large_mar']}</p>
</div>
"""
with open("info.html", "w", encoding="utf-8") as f:
    f.write(report_html)
