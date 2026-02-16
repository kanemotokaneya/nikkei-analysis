import pandas as pd
import matplotlib.pyplot as plt
import requests
import io

def update_market_board():
    try:
        # 1. 信頼性の高いデータサイトから日経平均の履歴を取得
        url = "https://stooq.com/q/d/l/?s=^ni225&i=d"
        res = requests.get(url, timeout=15).content
        df = pd.read_csv(io.StringIO(res.decode("utf-8")), index_col=0, parse_dates=True)
        df.columns = [c.capitalize() for c in df.columns]
        
        # 2. 移動平均線の計算 (5日・25日)
        df['MA5'] = df['Close'].rolling(window=5).mean()
        df['MA25'] = df['Close'].rolling(window=25).mean()
        
        # 3. 最新データと前日比の取得
        latest = df.iloc[-1]
        prev_close = df.iloc[-2]['Close']
        price = latest['Close']
        diff = price - prev_close
        
        # 4. チャート画像の作成
        plt.figure(figsize=(10, 6))
        plot_df = df.tail(60) # 直近2ヶ月分
        plt.plot(plot_df.index, plot_df['Close'], label='Price', color='#1f77b4', linewidth=2)
        plt.plot(plot_df.index, plot_df['MA5'], label='5-Day MA', color='green', linestyle='--')
        plt.plot(plot_df.index, plot_df['MA25'], label='25-Day MA', color='orange', linestyle='--')
        plt.title("Nikkei 225 Market Overview")
        plt.grid(True, alpha=0.3)
        plt.legend()
        plt.savefig('nikkei_chart.png')
        
        # 5. HTML（表示部分）の作成
        color = "red" if diff >= 0 else "blue"
        html = f"""
        <div style='background:#1a1a1a; color:white; padding:20px; border-radius:10px; font-family:sans-serif;'>
            <h1 style='margin:0; font-size:2.5em;'>{price:,.0f}円</h1>
            <p style='color:{color}; font-size:1.5em; margin:5px 0;'>前日比: {diff:+.0f}円</p>
            <div style='border-top:1px solid #333; margin-top:10px; padding-top:10px;'>
                <p>5日線: {latest['MA5']:,.0f}円 / 25日線: {latest['MA25']:,.0f}円</p>
                <p style='font-size:0.8em; color:gray;'>更新日: {df.index[-1].strftime('%Y-%m-%d')}</p>
            </div>
        </div>
        """
        with open("info.html", "w", encoding="utf-8") as f:
            f.write(html)
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    update_market_board()
