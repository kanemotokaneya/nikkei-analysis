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
    # 日経平均を取得。最新の価格を確実に反映させるため期間を調整
    df = yf.download("^N225", period="1y", interval="1d")
    df.index = pd.to_datetime(df.index)
    
    # データの型を強制的に数値(float)にし、不純物を排除
    for col in ['Open', 'High', 'Low', 'Close', 'Volume']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna().astype(float)
    
    # 最新値の取得
    close_p = float(df['Close'].iloc[-1])
    yesterday_p = float(df['Close'].iloc[-2])
except Exception as e:
    print(f"Stock Data Error: {e}")
    close_p, yesterday_p = 0.0, 0.0

# 日経VIの取得
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
oi = {'large_all': '取得中', 'large_mar': '取得中', 'mini_all': '取得中', 'mini_mar': '取得中', 'topix_all': '取得中', 'topix_mar': '取得中'}
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
    # 直近50日分のデータを抽出
    plot_df = df.tail(60).copy()
    
    # 移動平均線(5日, 25日)を計算して追加
    plot_df['MA5'] = plot_df['Close'].rolling(window=5).mean()
    plot_df['MA25'] = plot_df['Close'].rolling(window=25).mean()
    
    # 描画スタイルの設定
    mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit', volume='inherit')
    s  = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)
    
    # 5日線(緑)と25日線(オレンジ)を追加プロットとして設定
    apds = [
        mpf.make_addplot(plot_df['MA5'], color='green', width=1.0),
        mpf.make_addplot(plot_df['MA25'], color='orange', width=1.5)
    ]
    
    # チャート保存
    mpf.plot(plot_df, type='candle', style=s, addplot=apds,
             title=f"Nikkei 225 & Moving Average",
             ylabel='Price (JPY)',
             figsize=(12, 8), savefig='nikkei_chart.png')
except Exception as e:
    print(f"Chart Error: {e}")
    # 失敗時は単純なグラフ
    plt.figure(figsize=(12, 8))
    plt.plot(df['Close'].tail(50))
    plt.savefig('nikkei_chart.png')

# --- 5. HTML書き出し ---
# info.html (TOP用：価格を大きく表示)
top_html = f"""
<div class='analysis-box'>
    <h2 style='color:#2c3e50; font-size:2em; margin-bottom:5px;'>{close_p:,.0f}円</h2>
    <p style='font-size:1.2em; margin-top:0;'>前日比: <span style='color:{"red" if close_p >= yesterday_p else "blue"}'>{close_p - yesterday_p:+.0f}円</span></p>
    <hr>
    <p><b>日経VI:</b> {vi_value:.2f}</p>
    <p><b>本日の予測レンジ:</b> {close_p - daily_range:,.0f} ～ {close_p + daily_range:,.0f}円</p>
</div>
"""
with open("info.html", "w", encoding="utf-8") as f: f.write(top_html)

# details_info.html (詳細用)
detail_html = f"""
<div class='analysis-box'>
    <h3>■ 先物建玉状況</h3>
    <table style='width:100%; border-collapse: collapse;'>
        <tr style='background:#eee;'><th>銘柄</th><th>全体</th><th>3月限</th></tr>
        <tr><td>日経225(ラージ)</td><td>{oi['large_all']}</td><td>{oi['large_mar']}</td></tr>
        <tr><td>日経225 mini</td><td>{oi['mini_all']}</td><td>{oi['mini_mar']}</td></tr>
        <tr><td>TOPIX</td><td>{oi['topix_all']}</td><td>{oi['topix_mar']}</td></tr>
    </table>
</div>
"""
with open("details_info.html", "w", encoding="utf-8") as f: f.write(detail_html)
