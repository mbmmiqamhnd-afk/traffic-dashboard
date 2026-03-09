import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import os
import io

# --- 1. 頁面設定 ---
st.set_page_config(page_title="取締砂石車專案勤務", layout="wide")
st.title("🚛 取締砂石（大型貨）車重點違規專案勤務規劃表")

SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# ... (預設範本 DEFAULT_CMD, NOTES 等與您原始代碼相同)

# --- 2. 核心：修正後的合併邏輯 ---
def generate_html(month, briefing, df_cmd, df_schedule):
    font_family = "'標楷體', 'DFKai-SB', 'BiauKai', 'KaiTi', serif"
    style = f"""
    <style>
        body {{ font-family: {font_family}; color: #000; font-size: 15px; padding: 20px; }}
        .container {{ max-width: 850px; margin: auto; border: 1px solid #000; padding: 30px; }}
        h2 {{ text-align: center; font-weight: bold; margin-bottom: 20px; }}
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; }}
        th, td {{ border: 1px solid black; padding: 8px; text-align: center; vertical-align: middle; word-break: break-all; }}
        th {{ background-color: #f2f2f2; font-weight: bold; }}
        .left {{ text-align: left; }}
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2>{UNIT}<br>執行{month}「取締砂石（大型貨）車重點違規」專案勤務規劃表</h2>"
    
    # --- 任務編組表 (維持不變) ---
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>"
    for _, r in df_cmd.iterrows():
        html += f"<tr><td><b>{r['職稱']}</b></td><td>{r['代號']}</td><td>{str(r['姓名']).replace('、','<br>')}</td><td class='left'>{r['任務']}</td></tr>"
    html += "</table><br>"

    html += f"<p><b>📢 勤前教育：</b><br>{briefing.replace('\n','<br>')}</p>"

    # --- 警力佈署：自動合併日期儲存格 ---
    html += "<table><tr><th colspan='4'>警　力　佈　署</th></tr><tr><th width='30%'>日期</th><th width='20%'>執行單位</th><th width='10%'>人數</th><th width='40%'>執行路段</th></tr>"
    
    # 將 DataFrame 轉為 list 處理，避免索引混亂
    data = df_schedule.fillna("").values.tolist()
    total_rows = len(data)
    
    i = 0
    while i < total_rows:
        current_date = str(data[i][0]).strip()
        
        # 1. 判斷這一組日期要合併幾列 (Rowspan)
        # 邏輯：從當前列開始，往下找直到「日期欄不為空」或「結束」為止
        span = 1
        if current_date != "": # 只有當第一格有日期時才開始計算合併
            for j in range(i + 1, total_rows):
                next_date = str(data[j][0]).strip()
                if next_date == "": # 如果下一列日期是空的，就併入
                    span += 1
                else:
                    break
        
        # 2. 根據計算出的 span，生成每一列的 HTML
        for k in range(span):
            row_idx = i + k
            html += "<tr>"
            
            # 第一列要加上 rowspan，後續被合併的列則不寫 <td>
            if k == 0:
                # 如果日期本身是空的，但它屬於第一列，給它一個占位符或空白
                display_date = current_date if current_date != "" else "&nbsp;"
                html += f"<td rowspan='{span}'>{display_date}</td>"
            
            # 渲染其他欄位 (單位、人數、路段)
            unit = data[row_idx][1]
            num = data[row_idx][2]
            road = str(data[row_idx][3]).replace('\n', '<br>')
            
            html += f"<td>{unit}</td>"
            html += f"<td>{num}</td>"
            html += f"<td class='left'>{road}</td>"
            html += "</tr>"
            
        i += span # 跳轉到下一組

    html += f"</table><p><b>備註：</b><br>{NOTES}</p></div></body></html>"
    return html

# --- 3. 介面與功能按鈕 ---
# (此處接續您原本的 load_data, save_data 以及 st.data_editor 程式碼)
