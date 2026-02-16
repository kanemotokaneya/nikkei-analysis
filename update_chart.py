import pandas as pd
import requests
import io
import re

# 公開されたCSV URL
URL = "https://docs.google.com/spreadsheets/d/1uXxxC3untThuWdyCkIsDR8yc9X3JZF-00tvTkwNWDCE/pub?output=csv"

def get_data():
    try:
        res = requests.get(URL, timeout=10)
        df = pd.read_csv(io.StringIO(res.text), header=None)
        # 数字だけ抜き出す掃除
        def c(v): return float(re.sub(r'[^0-9.\-]', '', str(v))) if pd.notnull(v) else 0.0
        return c(df.iloc[0, 0]), c(df.iloc[0, 1]), c(df.iloc[0, 2])
    except:
        return 38000.0, 0.0, 20.0

p, diff, vi = get_data()

# サイトの文字を作る
html = f"""
<div style='background:#1a1a1a; color:white; padding:20px; border-radius:10px; font-family:sans-serif;'>
    <h1 style='margin:0; font-size:2.5em;'>{p:,.0f}円</h1>
    <p style='color:{"#ff4d4d" if diff>=0 else "#4d94ff"}; font-size:1.5em; margin:5px 0;'>前日比: {diff:+.0f}円</p>
    <div style='border-top:1px solid #333; margin-top:10px; padding-top:10px;'>
        <p><b>日経VI:</b> {vi:.2f}</p>
    </div>
</div>
"""
with open("info.html", "w", encoding="utf-8") as f: f.write(html)
