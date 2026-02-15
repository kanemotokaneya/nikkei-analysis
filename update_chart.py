import yfinance as yf
import matplotlib.pyplot as plt
import pandas as pd
import requests
from bs4 import BeautifulSoup
import datetime
import io

# --- 1. 株価データの取得 ---
ticker = "^N225"
data = yf.download(ticker, period="1y", interval="1d")
latest_price = float(data['Close'].iloc[-1])

# --- 2. JPXから建玉残高を取得 (pandas使用) ---
oi_text = "取得失敗"
try:
    # JPXのサイトから最新ExcelのURLを特定
    jpx_url = "https://www.jpx.co.jp/markets/derivatives/trading-volume/index.html"
    res = requests.get(jpx_url)
    soup = BeautifulSoup(res.text, "html.parser")
    link = soup.find("a", string=lambda t: t and "建玉残高表" in t)
    excel_url = "https://www.jpx.co.jp" + link.get("href")
    
    # Excelを読み込む
    excel_res = requests.get(excel_url)
    # pandasでExcelを読み込み（1つ目のシートを選択）
    df = pd.read_excel(io.BytesIO(excel_res.content))
    
    # 「日経225先物」と「合計」が含まれる行を探して、建玉残高を抜く
    # ※Excelの構成によって列の位置(11番目など)を調整
    row = df[df.iloc[:, 1].astype(str).str.contains("日経225先物") & df.iloc[:, 2].astype(str).str.contains("合計")]
    if not row.empty:
        oi_value = row.iloc[0, 11] # 11番目の列に建玉があることが多い
        oi_text = f"{int(oi_value):,}"
except Exception as e:
    print(f"Error: {e}")
    oi_text = "更新待ち"

# --- 3. グラフの作成 ---
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10), gridspec_kw={'height_ratios': [4, 1]})

# 上段：チャート
ax1.plot(data.tail(60).index, data['Close'].tail(60), color='black', label='Price')
ax1.set_title(f"Nikkei 225 & Open Interest ({datetime.date.today()})")
ax1.grid(True)

# 下段：建玉情報の表示
ax2.axis('off')
ax2.text(0.5, 0.6, f"日経225先物 最新建玉残高 (JPX合計)", fontsize=14, ha='center')
ax2.text(0.5, 0.3, f"{oi_text} 枚", fontsize=24, ha='center', color='blue', weight='bold')

plt.tight_layout()
plt.savefig("nikkei_chart.png")
