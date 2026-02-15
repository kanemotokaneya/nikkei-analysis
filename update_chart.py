import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import numpy as np
import datetime
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
except:
    close_p = 38000.0

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

# --- 3. オプション建玉分布の生成 (±5000円範囲) ---
opt_rows = ""
try:
    # ATM（現在の価格）を500円刻みに丸める
    atm = int(round(close_p / 500) * 500)
    # 現在価格 ± 5000円の範囲を500円刻みでリスト化
    strike_range = range(atm + 5000, atm - 5500, -500)

    for strike in strike_range:
        bg_color = "#fff5f5" if strike == atm else "white"
        # 現時点では価格リストを生成。ここに将来Excelデータを結合します
        opt_rows += f"""
        <tr style='background-color: {bg_color};'>
            <td>{strike:,}</td>
            <td>-</td>
            <td>-</td>
        </tr>"""
except:
    opt_rows = "<tr><td colspan='3'>データ生成エラー</td></tr>"

# --- 4. チャート作成 ---
try:
    plot_df = df.tail(50)
    mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit')
    s = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)
    mpf.plot(plot_df, type='candle', style=s, figsize=(12, 8), savefig='nikkei_chart.png')
except:
    plt.figure(figsize=(12, 8))
    plt.plot(df['Close'].tail(50))
    plt.savefig('nikkei_chart.png')

# --- 5. HTML書き出し (最後の引用符まで確実にコピペしてください) ---
report_content = f"""
<div class='analysis-box'>
    <h3 style='margin-top:0;'>① 市場概況・予測</h3>
    <p><b>日経平均現在値:</b> {close_p:,.0f}円</p>
    <p><b>デイリー予測:</b> {close_p - daily_range:,.0f} ～ {close_p + daily_range:,.0f}円</p>
    
    <h3>⑥ オプション建玉分布 (直近限月)</h3>
    <table style='width:100%; border-collapse: collapse; margin-top:10px;'>
        <thead>
            <tr style='background:#4a90e2; color:white;'>
                <th>権利行使価格</th><th>プット建玉</th><th>コール建玉</th>
            </tr>
        </thead>
        <tbody>
            {opt_rows}
        </tbody>
    </table>
    <p style='font-size:0.8em; color:gray; margin-top:10px;'>※現在値±5000円の範囲を500円刻みで算出</p>
</div>
"""

with open("info.html", "w", encoding="utf-8") as f:
    f.write(report_content)
