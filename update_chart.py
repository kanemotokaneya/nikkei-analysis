import yfinance as yf
import matplotlib.pyplot as plt
import datetime

# データの取得（確実に計算できるよう期間を広めに設定）
ticker = "^N225"
data = yf.download(ticker, period="2y", interval="1d")

# データが空でないかチェック
if len(data) < 75:
    print("データが足りません")
    exit(1)

# 移動平均線の計算
data['MA5'] = data['Close'].rolling(window=5).mean()
data['MA25'] = data['Close'].rolling(window=25).mean()
data['MA75'] = data['Close'].rolling(window=75).mean()

# 最新価格と前日比（最新の2行を取得）
latest = data.tail(2)
latest_price = latest['Close'].iloc[-1]
yesterday_price = latest['Close'].iloc[-2]

diff = latest_price - yesterday_price
diff_pct = (diff / yesterday_price) * 100
diff_text = f"{'+' if diff > 0 else ''}{float(diff):.2f} ({'+' if diff > 0 else ''}{float(diff_pct):.2f}%)"

# グラフ作成
plt.figure(figsize=(12, 6))
# 直近100日分に絞って表示
plot_data = data.tail(100)
plt.plot(plot_data.index, plot_data['Close'], label='Price', color='black', linewidth=1.5)
plt.plot(plot_data.index, plot_data['MA5'], label='5-day MA')
plt.plot(plot_data.index, plot_data['MA25'], label='25-day MA')
plt.plot(plot_data.index, plot_data['MA75'], label='75-day MA')

# テキスト表示
info_text = f"Latest: {float(latest_price):,.0f} JPY\nChange: {diff_text}"
plt.text(0.02, 0.95, info_text, transform=plt.gca().transAxes, fontsize=12, verticalalignment='top', bbox=dict(boxstyle='round', facecolor='white', alpha=0.5))

plt.title(f"Nikkei 225 Analysis - {datetime.date.today()}")
plt.legend()
plt.grid(True, linestyle='--', alpha=0.6)
plt.savefig("nikkei_chart.png")
