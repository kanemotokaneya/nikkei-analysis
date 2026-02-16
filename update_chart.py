import pandas as pd
import matplotlib.pyplot as plt
import requests
import io
import re

# スプレッドシートのCSV URL
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1uXxxC3untThuWdyCkIsDR8yc9X3JZF-00tvTkwNWDCE/pub?output=csv"

def clean_value(val):
    """どんな形式のセルデータも数値に変換する"""
    if pd.isna(val): return 0.0
    cleaned = re.sub(r'[^0-9.\-]', '', str(val))
    try:
        return float(cleaned)
    except:
        return 0.0

def get_stable_data():
    try:
        df_sheet = pd.read_csv(SHEET_CSV_URL, header=None)
        # A1: 現在値, B1: 前日比
        price = clean_value(df_sheet.iloc[0, 0])
        change = clean_value(df_sheet.iloc[0, 1])
        
        # B1がパーセント（0.01など1未満）の場合、絶対額に変換
        if abs(change) < 1 and change != 0:
            change_abs = price * change
            change_pct = change
        else:
            change_abs = change
            change_pct = (change / (price - change)) if (price - change) != 0 else 0
            
        return price, change_abs, change_pct
    except:
        return 0.0, 0.0, 0.0

close_p, change_abs, change_pct = get_stable_data()

# --- 2. チャート作成 (安定ソース Stooq) ---
try:
    url = "https://stooq.com/q/d/l/?s=^ni225&i=d"
