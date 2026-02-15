import yfinance as yf
import matplotlib.pyplot as plt
import datetime

# 1. データの取得（2年分取得して計算を安定させます）
ticker = "^N225"
data = yf.download(ticker, period="2y", interval="1d")

# 2. 移動平均線の計算
data['MA5'] = data['Close'].rolling(window=5).mean()
data['MA25'] = data['Close'].rolling(window=25).mean()
data['MA75'] = data['Close'].rolling(window=75).mean()

# 3. 最新価格と前日比の計算（エラー回避版）
# 最新の2日分を取り出し、単純な数値(float)に変換します
latest_price = float(data['Close'].iloc[-1])
yesterday_price = float(data['Close'].iloc[-2])

diff = latest_price - yesterday_price
diff_pct = (diff / yesterday_price) * 100

# 前日比の表示文字列を作成
sign = "+" if diff > 0 else ""
diff_text = f"{sign}{diff:.2f} ({sign}{diff_pct:.2f}%)"

# 4. グラフの作成
plt.figure(figsize=(12, 6))
# 直近100日分を表示
plot_data = data.tail(100)

plt.plot(plot_data.index, plot_data['Close'], label='Price', color='black', linewidth=1.5)
plt.plot(plot_data.index, plot_data['MA5'], label='5-day MA', linewidth=1)
plt.plot(plot_data.index, plot_data['MA25'], label='25-day MA', linewidth=1)
plt.plot(plot_data.index, plot_data['MA75'], label='75-day MA', linewidth=1)

# グラフ内に情報を表示
info_text = f"Latest: {latest_price:,.0f} JPY\nChange: {diff_text}"
plt.text(0.02, 0.95, info_text, transform=plt.gca().transAxes, 
         fontsize=12, verticalalignment='top', 
         bbox=dict(boxstyle='round', facecolor='white', alpha=0.8))

plt.title(f"Nikkei 225 Analysis - {datetime.date.today()}")
plt.legend(loc='lower right')
plt.grid(True, linestyle='--', alpha=0.5)

# 5. 保存
plt.savefig("nikkei_chart.png")
