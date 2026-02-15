import yfinance as yf
import matplotlib.pyplot as plt
import datetime

# 1. データの取得（移動平均を計算するために少し長めに取得します）
ticker = "^N225"
data = yf.download(ticker, period="1y", interval="1d")

# 2. 移動平均線の計算
data['MA5'] = data['Close'].rolling(window=5).mean()
data['MA25'] = data['Close'].rolling(window=25).mean()
data['MA75'] = data['Close'].rolling(window=75).mean()

# 3. 最新価格と前日比の計算
latest_price = data['Close'].iloc[-1]
yesterday_price = data['Close'].iloc[-2]
diff = latest_price - yesterday_price
diff_pct = (diff / yesterday_price) * 100

# 前日比の表示用テキスト
diff_text = f"{'+' if diff > 0 else ''}{diff:.2f} ({'+' if diff > 0 else ''}{diff_pct:.2f}%)"

# 4. グラフの作成
plt.figure(figsize=(12, 6))
plt.plot(data.index[-60:], data['Close'][-60:], label='Price', color='black', linewidth=1.5) # 直近60日分を表示
plt.plot(data.index[-60:], data['MA5'][-60:], label='5-day MA', alpha=0.8)
plt.plot(data.index[-60:], data['MA25'][-60:], label='25-day MA', alpha=0.8)
plt.plot(data.index[-60:], data['MA75'][-60:], label='75-day MA', alpha=0.8)

# グラフ内に最新情報をテキストで表示
info_text = f"Latest: {latest_price:,.0f} JPY\nChange: {diff_text}"
plt.text(0.02, 0.95, info_text, transform=plt.gca().transAxes, fontsize=12, verticalalignment='top', bbox=dict(boxstyle='round', facecolor='white', alpha=0.5))

plt.title(f"Nikkei 225 Analysis - {datetime.date.today()}")
plt.legend()
plt.grid(True, linestyle='--', alpha=0.6)

# 5. 画像として保存
plt.savefig("nikkei_chart.png")

# HTMLに渡すための数値情報をテキストファイルに保存（オプション）
with open("latest_info.txt", "w") as f:
    f.write(f"最新価格: {latest_price:,.0f}円 / 前日比: {diff_text}")
