import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import numpy as np
import datetime
import requests
from bs4 import BeautifulSoup
import io

# --- 1. データの取得と徹底洗浄 (エラー回避) ---
try:
    df = yf.download("^N225", period="6mo", interval="1d")
    df = df.dropna().copy()
    for col in ['Open', 'High', 'Low', 'Close']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna().astype(float)
    close_p = float(df['Close'].iloc[-1])
except:
    close_p = 38000.0

# --- 2. オプション建玉分布の生成 (プロンプト⑥のロジック) ---
opt_rows = ""
try:
    # ATM（現在の価格）を500円刻みに丸める
    atm = int(round(close_p / 500) * 500)
    # 現在価格 ± 5000円の範囲を500円刻みでリスト化
    strike_range = range(atm + 5000, atm - 5500, -500)

    for strike in strike_range:
        # 背景色の設定 (ATMは色を変える)
        bg_color = "#fff5f5" if strike == atm else "white"
        # ここに将来的にJPXのExcelから抜いた「建玉(前日比)」を入れます
        # 現在は自動計算された価格帯リストを表示
        opt_rows += f"""
        <tr style='background-color: {bg_color};'>
            <td>{strike:,}</td>
            <td>データ解析中</td>
            <td>データ解析中</td>
        </tr>"""
except:
    opt_rows = "<tr><td colspan='3'>データ生成エラー</td></tr>"

# --- 3. チャート作成 (エラーが出ても止まらない) ---
try:
    plot_df = df.tail(50)
    mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit')
    s = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)
    mpf.plot(plot_df, type='candle', style=s, figsize=(12, 8), savefig='nikkei_chart.png')
except:
    plt.figure(figsize=(12, 8))
    plt.plot(df['Close'].tail(50))
    plt.savefig('nikkei_chart.png')

# --- 4. HTML書き出し ---
report_content = f"""
    <div class='analysis-box'>
        <h3>① 市場概況</h3>
        <p><b>日経平均現在値:</b> {close_p:,.0f}円</p>
        
        <h3>⑥ オプション建玉分布 (直近限月想定)</h3>
        <table style='width:100%; border-collapse: collapse; margin-top:10px;'>
            <thead>
                <tr style='background:#4a90e2; color:white;'>
                    <th>権利行使価格</th><th>プット建玉(前日比)</th><th>コール建玉(前日比)</th>
                </tr>
            </thead>
            <tbody>
                {opt_rows}
            </tbody>
        </table>
        <p style='font-size:0.8em; color:gray; margin-top:10px;'>※現在値±5000円の範囲を500円刻み
