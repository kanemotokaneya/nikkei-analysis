import yfinance as yf
import matplotlib.pyplot as plt
import datetime

# 1. 日経225先物（あるいは指数）のデータを取得
# ^N225 は日経平均株価のシンボルです
data = yf.download("^N225", period="1mo", interval="1d")

# 2. グラフの作成
plt.figure(figsize=(10, 5))
plt.plot(data['Close'], label='Nikkei 225', color='#2ca02c')
plt.title(f"Nikkei 225 Analysis - {datetime.date.today()}")
plt.legend()
plt.grid(True)

# 3. 画像として保存
plt.savefig("nikkei_chart.png")
