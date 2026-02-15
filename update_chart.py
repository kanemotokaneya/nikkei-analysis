import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import numpy as np
import datetime
import requests
from bs4 import BeautifulSoup
import io

# --- 0. 変数の初期化（これがエラー防止の安全装置です） ---
# 万が一データ取得に失敗しても、この数値が使われるのでエラーになりません
close_p = 0.0
yesterday_p = 0.0
vi_value = 20.0 # VIが取れない場合の仮置き数値
daily_range = 0.0
weekly_range = 0.0
oi_data = {
    'large_total': '取得失敗',
    'large_march': '取得失敗'
}
plot_df = pd.DataFrame() # 空のデータフレーム

# --- 1. 株価・VIデータの取得 ---
print("Fetching Market Data...")
try:
    # 日経平均 (yfinanceの仕様変更に対応)
    df = yf.download("^N225", period="6mo", interval="1d")
    
    # データが空でないか確認
    if not df.empty:
        # マルチインデックス対策（念のため）
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)
            
        df = df.dropna()
        
        # 終値の取得
        close_p = float(df['Close'].iloc[-1])
        yesterday_p = float(df['Close'].iloc[-2])
        plot_df = df.tail(60) # チャート用データを保存

        # 日経VI
        try:
            df_vi = yf.download("^JNIV", period="5d", interval="1d")
            if not df_vi.empty:
                if isinstance(df_vi.columns, pd.MultiIndex):
                    df_vi.columns = df_vi.columns.get_level_values(0)
                vi_value = float(df_vi['Close'].dropna().iloc[-1])
        except Exception as e:
            print(f"VI Download Error: {e}")
            # 失敗しても vi_value は初期値(20.0)のまま進む

        # 予測値幅の計算
        daily_range = close_p * (vi_value / 100) / np.sqrt(250)
        weekly_range = close_p * (vi_value / 100) / np.sqrt(52)
    else:
        print("Market Data is empty.")

except Exception as e:
    print(f"Market Data Critical Error: {e}")
    # エラー起きても初期値があるので止まらない

# --- 2. JPX建玉データの取得 ---
print("Fetching JPX Data...")
try:
    jpx_url = "https://www.jpx.co.jp/markets/derivatives/trading-volume/index.html"
    res = requests.get(jpx_url, timeout=10)
    soup = BeautifulSoup(res.text, "html.parser")
    link = soup.find("a", href=lambda h: h and "open_interest" in h and h.endswith(".xlsx"))
    
    if link:
        excel_url = "https://www.jpx.co.jp" + link.get("href")
        print(f"Downloading Excel: {excel_url}")
        excel_res = requests.get(excel_url, timeout=10)
        df_jpx = pd.read_excel(io.BytesIO(excel_res.content), header=None)
        
        # 戦略A: 座標指定 (E列=4, 49行目=48)
        try:
            val_total = df_jpx.iloc[48, 4]
            val_march = df_jpx.iloc[29, 4]
            oi_data['large_total'] = f"{val_total:,}"
            oi_data['large_march'] = f"{val_march:,}"
        except:
            pass

        # 戦略B: 検索モード（座標がズレてたら文字を探す）
        for i, row in df_jpx.iterrows():
            row_str = str(row.values)
            if "日経225先物" in row_str and "合計" in row_str:
                val = df_jpx.iloc[i, 4]
                # 数字っぽいものがあれば取得
                if pd.notna(val):
                    try:
                        oi_data['large_total'] = f"{int(val):,}"
                    except:
                        pass
except Exception as e:
    print(f"JPX Error: {e}")

# --- 3. HTMLファイルの生成 ---
# ここで変数が定義されていないエラーはもう起きません
print("Generating HTML...")
html_content = f"""
<div class='analysis-box'>
    <h3 style='border-bottom: 2px solid #336; padding-bottom: 5px;'>① 日経平均・VI分析</h3>
    <p><b>現物終値:</b> {close_p:,.0f}円 (<span style='color:{"red" if close_p >= yesterday_p else "blue"}'>{close_p - yesterday_p:+.0f}円</span>)</p>
    <p><b>日経VI:</b> {vi_value:.2f}</p>
    <p><b>デイリー予測:</b> {close_p - daily_range:,.0f}円 ～ {close_p + daily_range:,.0f}円</p>
    
    <h3 style='border-bottom: 2px solid #336; padding-bottom: 5px; margin-top: 20px;'>② 日経225先物 建玉残高（前日比）</h3>
    <table style='width:100%; border-collapse: collapse;'>
        <tr style='background-color: #f0f0f0;'>
            <th style='border: 1px solid #ddd; padding: 8px;'>項目</th>
            <th style='border: 1px solid #ddd; padding: 8px;'>建玉残高</th>
        </tr>
        <tr>
            <td style='border: 1px solid #ddd; padding: 8px;'>ラージ合計</td>
            <td style='border: 1px solid #ddd; padding: 8px; font-weight:bold;'>{oi_data['large_total']}</td>
        </tr>
        <tr>
            <td style='border: 1px solid #ddd; padding: 8px;'>3月限</td>
            <td style='border: 1px solid #ddd; padding: 8px;'>{oi_data['large_march']}</td>
        </tr>
    </table>
    <p style='font-size:0.8em; color:gray;'>更新日時: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
</div>
"""

with open("info.html", "w", encoding="utf-8") as f:
    f.write(html_content)

# --- 4. チャート生成 ---
try:
    if not plot_df.empty:
        mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit')
        s = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)
        mpf.plot(plot_df, type='candle', style=s, 
                 title=f"Nikkei 225 (VI: {vi_value:.2f})",
                 ylabel='Price (JPY)', figsize=(12, 8), savefig='nikkei_chart.png')
    else:
        print("No data to plot chart.")
        # 空の画像を作ってエラー回避
        plt.figure()
        plt.text(0.5, 0.5, 'No Data', ha='center')
        plt.savefig('nikkei_chart.png')
except Exception as e:
    print(f"Chart Error: {e}")
