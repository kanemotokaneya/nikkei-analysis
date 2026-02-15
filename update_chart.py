    import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import numpy as np
import requests
from bs4 import BeautifulSoup
import io

# --- 1. データの取得と強力な洗浄 ---
try:
    # 期間を2年に広げ、十分な計算用データを確保
    df = yf.download("^N225", period="2y", interval="1d")
    df.index = pd.to_datetime(df.index)
    
    # データの徹底洗浄
    for col in ['Open', 'High', 'Low', 'Close', 'Volume']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna().astype(float)
    
    # 最新値の取得
    if not df.empty:
        close_p = float(df['Close'].iloc[-1])
        yesterday_p = float(df['Close'].iloc[-2])
    else:
        close_p, yesterday_p = 38000.0, 38000.0
except Exception as e:
    print(f"Data Error: {e}")
    close_p, yesterday_p = 38000.0, 38000.0

# 日経VI (エラー時は20.0固定)
vi_value = 20.0
try:
    df_vi = yf.download("^JNIV", period="5d", interval="1d")
    if not df_vi.empty:
        vi_value = float(df_vi['Close'].dropna().iloc[-1])
except:
    pass

# --- 2. 予測値幅の計算 ---
daily_range = close_p * (vi_value / 100) / np.sqrt(250)

# --- 3. JPXデータの抽出 (先物) ---
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

# --- 4. ローソク足と移動平均線の作成 ---
try:
    # 描画用に直近100日分をコピー
    plot_df = df.tail(100).copy()
    
    # 描画スタイルの設定（上昇:赤 / 下落:青）
    mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit')
    s  = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)
    
    # 移動平均線の計算と描画設定
    # 5日(緑), 25日(オレンジ), 75日(水色)の3本
    mpf.plot(plot_df, type='candle', style=s,
             mav=(5, 25, 75), # 移動平均線を自動計算
             title=f"Nikkei 225 & MA",
             ylabel='Price (JPY)',
             figsize=(12, 8), 
             savefig='nikkei_chart.png')
except Exception as e:
    print(f"Chart Error: {e}")
    # 失敗時のバックアップ（折れ線）
    plt.figure(figsize=(12, 8))
    plt.plot(df['Close'].tail(50))
    plt.savefig('nikkei_chart.png')

# --- 5. HTML書き出し ---
# info.html (価格が0円にならないよう再確認)
top_html = f"""
<div class='analysis-box' style='background:#f8faff; padding:20px; border-radius:10px;'>
    <h2 style='color:#2c3e50; font-size:2.5em; margin:0;'>{close_p:,.0f}円</h2>
    <p style='font-size:1.4em; margin-top:5px;'>前日比: <span style='color:{"red" if close_p >= yesterday_p else "blue"}'>{close_p - yesterday_p:+.0f}円</span></p>
    <hr style='border:0; border-top:1px solid #eee;'>
    <p><b>日経VI:</b> {vi_value:.2f}</p>
    <p><b>本日の予測レンジ:</b> {close_p - daily_range:,.0f} ～ {close_p + daily_range:,.0f}円</p>
</div>
"""
with open("info.html", "w", encoding="utf-8") as f: f.write(top_html)

# details_info.html
detail_html = f"""
<div class='analysis-box'>
    <h3>■ 先物建玉状況</h3>
    <table style='width:100%; border-collapse: collapse; text-align:center;'>
        <tr style='background:#eee;'><th>銘柄</th><th>全体</th><th>3月限</th></tr>
        <tr><td>日経225(ラージ)</td><td>{oi['large_all']}</td><td>{oi['large_mar']}</td></tr>
        <tr><td>日経225 mini</td><td>{oi['mini_all']}</td><td>{oi['mini_mar']}</td></tr>
        <tr><td>TOPIX</td><td>{oi['topix_all']}</td><td>{oi['topix_mar']}</td></tr>
    </table>
</div>
"""
with open("details_info.html", "w", encoding="utf-8") as f: f.write(detail_html)
