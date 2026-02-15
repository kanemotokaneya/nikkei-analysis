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
# 日経平均(現物)
df = yf.download("^N225", period="6mo", interval="1d")

# 日経平均VI（エラー回避処理付き）
try:
    df_vi = yf.download("^JNIV", period="5d", interval="1d")
    if len(df_vi) > 0:
        vi_value = float(df_vi['Close'].iloc[-1])
    else:
        vi_value = 20.0  # 取れない場合は標準的な20.0を使用
except:
    vi_value = 20.0

# 最新値の抽出
close_p = float(df['Close'].iloc[-1])
yesterday_p = float(df['Close'].iloc[-2])

# --- 2. 予測値幅の計算 (ご提示のロジック) ---
daily_range = close_p * (vi_value / 100) / np.sqrt(250)
weekly_range = close_p * (vi_value / 100) / np.sqrt(52)

# --- 3. JPXデータの抽出 (座標指定の準備) ---
# プロンプト②の「日経225(ラージ) E列49行」などを狙い撃ちします
oi_data = {}
try:
    jpx_url = "https://www.jpx.co.jp/markets/derivatives/trading-volume/index.html"
    res = requests.get(jpx_url)
    soup = BeautifulSoup(res.text, "html.parser")
    link = soup.find("a", string=lambda t: t and "建玉残高表" in t)
    excel_url = "https://www.jpx.co.jp" + link.get("href")
    
    # Excel読み込み（header=Noneで座標を数えやすくします）
    excel_res = requests.get(excel_url)
    df_jpx = pd.read_excel(io.BytesIO(excel_res.content), header=None)
    
    # 座標指定抽出（pandasは0から数えるので、E列は4、49行目は48と指定）
    # ※Excelの構成によって微調整が必要な場合があります
    oi_data['large_all'] = df_jpx.iloc[48, 4] # E列49行
    oi_data['large_mar'] = df_jpx.iloc[29, 4] # E列30行
except:
    oi_data['large_all'] = "取得中"
    oi_data['large_mar'] = "取得中"

# --- 4. ローソク足チャートの作成 ---
plot_df = df.tail(50)
mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit')
s  = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)

mpf.plot(plot_df, type='candle', style=s, 
         title=f"Nikkei 225 (VI: {vi_value:.2f})",
         ylabel='Price (JPY)',
         figsize=(12, 8),
         savefig='nikkei_chart.png')

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
