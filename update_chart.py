import pandas as pd
import matplotlib.pyplot as plt
import requests
from bs4 import BeautifulSoup
import io
import numpy as np

# --- 1. 日経平均株価を「株探」から直接取得 ---
def get_nikkei_price():
    try:
        url = "https://kabutan.jp/stock/kabuexe?code=0000"
        headers = {"User-Agent": "Mozilla/5.0"}
        res = requests.get(url, headers=headers)
        soup = BeautifulSoup(res.text, "html.parser")
        
        # 株価・前日比を抽出
        price_text = soup.find("span", class_="kabuka").text.replace(",", "").replace("円", "")
        change_text = soup.find("dd", class_="zenjitsu_henka").text.replace(",", "").replace("円", "")
        return float(price_text), float(change_text)
    except:
        return 38000.0, 0.0 # 失敗時のダミー

# --- 2. チャート用データをStooqのCSVから直接取得 ---
def get_chart_data():
    try:
        # StooqのCSVダウンロードリンク（日経225）
        url = "https://stooq.com/q/d/l/?s=^ni225&i=d"
        res = requests.get(url).content
        df = pd.read_csv(io.StringIO(res.decode("utf-8")), index_col=0, parse_dates=True)
        return df.tail(100) # 直近100日分
    except:
        return pd.DataFrame()

close_p, change_p = get_nikkei_price()
df = get_chart_data()

# --- 3. JPXデータの抽出 ---
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

# --- 4. チャート作成（ライブラリ依存を減らす） ---
plt.figure(figsize=(12, 8))
if not df.empty:
    # 5日・25日移動平均を計算
    ma5 = df['Close'].rolling(window=5).mean()
    ma25 = df['Close'].rolling(window=25).mean()
    
    plt.plot(df.index, df['Close'], label='Close', color='#2c3e50', linewidth=2)
    plt.plot(df.index, ma5, label='MA5', color='green', alpha=0.7)
    plt.plot(df.index, ma25, label='MA25', color='orange', alpha=0.7)
    plt.fill_between(df.index, df['Low'], df['High'], color='gray', alpha=0.1)
    plt.title("Nikkei 225 Market Overview")
    plt.grid(True, linestyle='--', alpha=0.5)
    plt.legend()
plt.savefig('nikkei_chart.png')

# --- 5. HTML書き出し ---
top_html = f"""
<div class='analysis-box'>
    <h2 style='font-size:2.8em; margin:0;'>{close_p:,.0f}円</h2>
    <p style='font-size:1.5em;'>前日比: <span style='color:{"red" if change_p >= 0 else "blue"}'>{change_p:+.0f}円</span></p>
</div>
"""
with open("info.html", "w", encoding="utf-8") as f: f.write(top_html)

detail_html = f"""
<div class='analysis-box'>
    <h3>先物建玉状況</h3>
    <table border='1' style='width:100%; border-collapse:collapse; text-align:center;'>
        <tr style='background:#eee;'><th>銘柄</th><th>全体</th><th>3月限</th></tr>
        <tr><td>ラージ</td><td>{oi['large_all']}</td><td>{oi['large_mar']}</td></tr>
        <tr><td>mini</td><td>{oi['mini_all']}</td><td>{oi['mini_mar']}</td></tr>
        <tr><td>TOPIX</td><td>{oi['topix_all']}</td><td>{oi['topix_mar']}</td></tr>
    </table>
</div>
"""
with open("details_info.html", "w", encoding="utf-8") as f: f.write(detail_html)
