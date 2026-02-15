import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import numpy as np
import datetime
import requests
from bs4 import BeautifulSoup
import io

# --- 1. 株価・VIデータの取得 ---
print("Fetching Market Data...")
try:
    # 日経平均
    df = yf.download("^N225", period="6mo", interval="1d")
    df = df.dropna().copy()
    for col in ['Open', 'High', 'Low', 'Close']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna().astype(float)
    
    close_p = float(df['Close'].iloc[-1])
    yesterday_p = float(df['Close'].iloc[-2])
    
    # 日経VI (取得失敗時は20.0)
    vi_value = 20.0
    try:
        df_vi = yf.download("^JNIV", period="5d", interval="1d")
        if not df_vi.empty:
            vi_value = float(df_vi['Close'].dropna().iloc[-1])
    except:
        pass

    # 予測値幅
    daily_range = close_p * (vi_value / 100) / np.sqrt(250)
    weekly_range = close_p * (vi_value / 100) / np.sqrt(52)
    
except Exception as e:
    print(f"Market Data Error: {e}")
    close_p, yesterday_p = 0.0, 0.0
    daily_range, weekly_range = 0.0, 0.0

# --- 2. JPX建玉データの取得（検索機能付き） ---
print("Fetching JPX Data...")
oi_data = {
    'large_total': '取得失敗',
    'large_march': '取得失敗'
}

try:
    jpx_url = "https://www.jpx.co.jp/markets/derivatives/trading-volume/index.html"
    res = requests.get(jpx_url, timeout=10)
    soup = BeautifulSoup(res.text, "html.parser")
    # 最新の「建玉残高表」Excelリンクを探す
    link = soup.find("a", href=lambda h: h and "open_interest" in h and h.endswith(".xlsx"))
    
    if link:
        excel_url = "https://www.jpx.co.jp" + link.get("href")
        print(f"Downloading Excel: {excel_url}")
        excel_res = requests.get(excel_url, timeout=10)
        
        # Excel読み込み（ヘッダーなしで読み込む）
        df_jpx = pd.read_excel(io.BytesIO(excel_res.content), header=None)
        
        # --- 戦略A: 座標指定（まずは指定の場所を見る） ---
        # E列(index 4), 49行目(index 48) -> ラージ合計
        try:
            val_total = df_jpx.iloc[48, 4]
            val_march = df_jpx.iloc[29, 4] # 3月限 (30行目)
            oi_data['large_total'] = f"{val_total:,}" if pd.notna(val_total) else "データなし"
            oi_data['large_march'] = f"{val_march:,}" if pd.notna(val_march) else "データなし"
        except:
            print("座標指定での取得に失敗しました。検索モードに移行します。")

        # --- 戦略B: 検索モード（文字を探す） ---
        # もし座標取得が変なら、文字検索で上書きする
        # 「日経225先物」がある行を探す
        for i, row in df_jpx.iterrows():
            row_str = str(row.values)
            if "日経225先物" in row_str:
                if "合計" in row_str:
                    # 合計行のE列(4)を取得
                    val = df_jpx.iloc[i, 4]
                    if pd.notna(val) and str(val).replace('.','').replace('-','').isdigit():
                         oi_data['large_total'] = f"{int(val):,}"
                
                # 限月（例：202603）などを探すロジックは必要に応じて追加
                
except Exception as e:
    print(f"JPX Error: {e}")

# --- 3. HTMLファイルの生成 ---
print("Generating HTML...")
html_content = f"""
<div class='analysis-box'>
    <h3 style='border-bottom: 2px solid #336; padding-bottom: 5px;'>① 日経平均・VI分析</h3>
    <p><b>現物終値:</b> {close_p:,.0f}円 (<span style='color:{"red" if close_p >= yesterday_p else "blue"}'>{close_p - yesterday_p:+.0f}円</span>)</p>
    <p><b>日経VI:</b> {vi_value:.2f}</p>
    <p><b>デイリー予測レンジ:</b> {close_p - daily_range:,.0f}円 ～ {close_p + daily_range:,.0f}円</p>
    
    <h3 style='border-bottom: 2px solid #336; padding-bottom: 5px; margin-top: 20px;'>② 日経225先物 建玉残高（前日比）</h3>
    <table style='width:100%; border-collapse: collapse;'>
        <tr style='background-color: #f0f0f0;'>
            <th style='border: 1px solid #ddd; padding: 8px;'>項目</th>
            <th style='border: 1px solid #ddd; padding: 8px;'>建玉変化</th>
        </tr>
        <tr>
            <td style='border: 1px solid #ddd; padding: 8px;'>ラージ合計</td>
            <td style='border: 1px solid #ddd; padding: 8px; font-weight:bold;'>{oi_data['large_total']}</td>
        </tr>
        <tr>
            <td style='border: 1px solid #ddd; padding: 8px;'>3月限（直近）</td>
            <td style='border: 1px solid #ddd; padding: 8px;'>{oi_data['large_march']}</td>
        </tr>
    </table>
    <p style='font-size:0.8em; color:gray;'>※JPX「建玉残高表」より自動抽出</p>
</div>
"""

with open("info.html", "w", encoding="utf-8") as f:
    f.write(html_content)

# --- 4. チャート生成（変更なし） ---
try:
    plot_df = df.tail(60)
    mc = mpf.make_marketcolors(up='red', down='blue', edge='inherit', wick='inherit')
    s = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', y_on_right=False)
    mpf.plot(plot_df, type='candle', style=s, 
             title=f"Nikkei 225 (VI: {vi_value:.2f})",
             ylabel='Price (JPY)', figsize=(12, 8), savefig='nikkei_chart.png')
except Exception as e:
    print(f"Chart Error: {e}")
