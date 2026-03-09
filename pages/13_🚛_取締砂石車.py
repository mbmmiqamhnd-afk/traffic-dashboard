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
NOTES = "※ 加強取締砂石（大型貨）車超載、車速、酒醉駕車、闖紅燈、無照駕車、爭道行駛、違反禁行路線、變更車斗、未使用專用車箱及未裝設行車紀錄器（行車視野輔助器）等違規，以共同消弭不法行為，保障用路人生命財產安全。"

# --- 2. 核心：修正後的 HTML 生成 (解決消失問題) ---
def generate_html(month, briefing, df_cmd, df_schedule):
    font_family = "'標楷體', 'DFKai-SB', 'BiauKai', 'KaiTi', serif"
    style = f"""
    <style>
        body {{ font-family: {font_family}; color: #000; font-size: 15px; padding: 20px; }}
        .container {{ max-width: 850px; margin: auto; border: 1px solid #000; padding: 30px; }}
        h2 {{ text-align: center; font-weight: bold; margin-bottom: 20px; }}
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; table-layout: fixed; }}
        th, td {{ border: 1px solid black; padding: 8px; text-align: center; vertical-align: middle; word-wrap: break-word; }}
        th {{ background-color: #f2f2f2; font-weight: bold; }}
        .left {{ text-align: left; }}
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2>{UNIT}<br>執行{month}「取締砂石（大型貨）車重點違規」專案勤務規劃表</h2>"
    
    # --- 任務編組表 ---
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr><tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, r in df_cmd.iterrows():
        name = str(r.get('姓名', '')).replace('、', '<br>')
        html += f"<tr><td><b>{r.get('職稱','')}</b></td><td>{r.get('代號','')}</td><td>{name}</td><td class='left'>{r.get('任務','')}</td></tr>"
    html += "</table>"

    html += f"<p><b>📢 勤前教育：</b><br>{str(briefing).replace('\\n','<br>').replace('\n','<br>')}</p>"

    # --- 警力佈署：自動合併日期邏輯 ---
    html += "<table><tr><th colspan='4'>警　力　佈　署</th></tr><tr><th width='25%'>日期</th><th width='20%'>執行單位</th><th width='10%'>執行人數</th><th width='45%'>執行路段</th></tr>"
    
    # 確保資料為字串並填補空值
    df_s = df_schedule.astype(str).replace('nan', '')
    data = df_s.values.tolist()
    total_rows = len(data)
    
    if total_rows == 0:
        html += "<tr><td colspan='4'>暫無勤務資料</td></tr>"
    else:
        i = 0
        while i < total_rows:
            # 取得目前的日期
            current_date = data[i][0].strip()
            
            # 判斷這組日期需要合併幾列 (Rowspan)
            # 條件：如果日期不是空的，就往下算有多少個空白列
            span = 1
            if current_date != "":
                for j in range(i + 1, total_rows):
                    if data[j][0].strip() == "":
                        span += 1
                    else:
                        break
            
            # 渲染這一組的所有行
            for k in range(span):
                curr_idx = i + k
                # 確保不超出索引範圍 (防呆)
                if curr_idx >= total_rows: break
                
                html += "<tr>"
                # 只有這組的第一行要畫日期格，並設定 rowspan
                if k == 0:
                    html += f"<td rowspan='{span}'>{current_date if current_date != '' else '&nbsp;'}</td>"
                
                # 單位、人數、路段
                unit = data[curr_idx][1]
                num = data[curr_idx][2]
                road = data[curr_idx][3].replace('\\n', '<br>').replace('\n', '<br>')
                
                html += f"<td>{unit}</td><td>{num}</td><td class='left'>{road}</td>"
                html += "</tr>"
            
            i += span # 跳轉到下一組

    html += f"</table><p><b>備註：</b><br>{NOTES}</p></div></body></html>"
    return html

# --- 3. 介面 (這部分請保留您原有的 Google Sheets 讀取邏輯) ---
# 假設 edited_cmd 與 edited_schedule 是您從 st.data_editor 拿到的資料
# html_out = generate_html(current_month, brief_info, edited_cmd, edited_schedule)
# st.components.v1.html(html_out, height=800, scrolling=True)
