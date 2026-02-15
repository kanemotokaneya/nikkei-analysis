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
# 日経平均(現物)と日経VI(現物指数)
df = yf.download("^N225", period="6mo", interval="1d")
df_vi = yf.download("^JNIV", period="1mo", interval="1d")

# 最新値の抽出
close_p = float(df['Close'].iloc[-1])
yesterday_p = float(df['Close'].iloc[-2])
vi_value = float(df_vi['Close'].iloc[-1])

# --- 2. 予測値幅の計算 (ご提示のロジック) ---
# 1日の予測値幅 = 現物終値 × (日経VI / 100) / √250
daily_range = close_p * (vi_value / 100) / np.sqrt(250)
# 1週間の予測値幅 = 現物終値 × (日経VI / 100) / √52
weekly_range = close_p * (vi_value / 100) / np.sqrt(52)

# --- 3. ローソク足チャートの作成 ---
plot_df = df.tail(50) # 直近50日
mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit')
s  = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)

# 保存用
mpf.plot(plot_df, type='candle', style=s, 
         title=f"Nikkei 225 & Predictive Range (VI: {vi_value:.2f})",
         ylabel='Price (JPY)',
         figsize=(12, 8),
         savefig='nikkei_chart.png')

# --- 4. JPXデータの抽出 (将来の座標指定への土台) ---
# ここに今後、プロンプトの②〜⑦の座標抽出コードを追加していきます。
