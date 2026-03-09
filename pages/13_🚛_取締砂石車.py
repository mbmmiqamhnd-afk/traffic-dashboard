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

# --- 預設範本 ---
DEFAULT_MONTH = "115年3月份"
DEFAULT_BRIEF = "時間：各單位執行前\n地點：現地勤教"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "建興", "姓名": "駐區督察 孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、保安民防組巡官 陳鵬翔（代理人：警員張庭溱）、人事室警員 陳明祥、行政組警務佐 曾威仁", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期": "115年3月13日（星期五）", "執行單位": "聖亭派出所", "執行人數": "2至4人", "執行路段": "中豐路、聖亭路段"},
    {"日期": "", "執行單位": "龍潭派出所", "執行人數": "", "執行路段": "大昌路、中豐路段"},
    {"日期": "", "執行單位": "中興派出所", "執行人數": "", "執行路段": "中興路、福龍路段"},
    {"日期": "115年3月27日（星期五）", "執行單位": "聖亭派出所", "執行人數": "2至4人", "執行路段": "中豐路、聖亭路段"},
    {"日期": "", "執行單位": "龍潭派出所", "執行人數": "", "執行路段": "大昌路、中豐路段"},
])

NOTES = "※ 加強取締砂石（大型貨）車重點違規事項..."

# --- 2. HTML 生成 (核心：自動向下合併空白日期) ---
def generate_html(month, briefing, df_cmd, df_schedule):
    font_family = "'標楷體', 'DFKai-SB', 'BiauKai', 'KaiTi', serif"
    style = f"""
    <style>
        body {{ font-family: {font_family}; color: #000; font-size: 15px; padding: 20px; }}
        .container {{ max-width: 850px; margin: auto; border: 1px solid #000; padding: 30px; }}
        h2 {{ text-align: center; font-weight: bold; }}
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; }}
        th, td {{ border: 1px solid black; padding: 8px; text-align: center; vertical-align: middle; }}
        th {{ background-color: #f2f2f2; }}
        .left {{ text-align: left; }}
        @media print {{ body {{ padding: 0; }} .container {{ border: none; }} }}
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2>{UNIT}<br>執行{month}「取締砂石（大型貨）車重點違規」專案勤務規劃表</h2>"
    
    # 任務編組表 (略過)
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>"
    for _, r in df_cmd.iterrows():
        html += f"<tr><td><b>{r['職稱']}</b></td><td>{r['代號']}</td><td>{str(r['姓名']).replace('、','<br>')}</td><td class='left'>{r['任務']}</td></tr>"
    html += "</table>"

    html += f"<p><b>📢 勤前教育：</b><br>{briefing.replace('\n','<br>')}</p>"

    # --- 警力佈署：自動合併日期邏輯 ---
    html += "<table><tr><th colspan='4'>警　力　佈　署</th></tr><tr><th width='30%'>日期</th><th width='20%'>執行單位</th><th width='10%'>人數</th><th width='40%'>執行路段</th></tr>"
    
    rows = df_schedule.to_dict('records')
    idx = 0
    while idx < len(rows):
        current_row = rows[idx]
        date_text = str(current_row.get('日期', '')).strip()
        
        # 如果當前日期欄有資料，則往下尋找空白儲存格
        if date_text != "" and date_text != "nan":
            span = 1
            for j in range(idx + 1, len(rows)):
                next_date = str(rows[j].get('日期', '')).strip()
                if next_date == "" or next_date == "nan":
                    span += 1
                else:
                    break
            
            # 渲染這一組 (有資料的日期 + 下方空白)
            for k in range(span):
                html += "<tr>"
                if k == 0:
                    html += f"<td rowspan='{span}'>{date_text}</td>"
                html += f"<td>{rows[idx+k]['執行單位']}</td><td>{rows[idx+k]['執行人數']}</td><td class='left'>{rows[idx+k]['執行路段']}</td></tr>"
            idx += span
        else:
            # 如果一開始就是空白（異常狀況），則單獨渲染
            html += f"<tr><td></td><td>{current_row['執行單位']}</td><td>{current_row['執行人數']}</td><td class='left'>{current_row['執行路段']}</td></tr>"
            idx += 1

    html += f"</table><p><b>備註：</b><br>{NOTES}</p></div></body></html>"
    return html

# --- 3. Streamlit 介面 ---
st.subheader("編輯資訊")
c1, c2 = st.columns(2)
month_val = c1.text_input("月份", DEFAULT_MONTH)
brief_val = c2.text_area("勤前教育", DEFAULT_BRIEF)

st.subheader("編輯警力佈署 (日期下方留白將自動合併)")
ed_sch = st.data_editor(DEFAULT_SCHEDULE, num_rows="dynamic", use_container_width=True)

# 產生預覽
html_preview = generate_html(month_val, brief_val, DEFAULT_CMD, ed_sch)
st.markdown("---")
st.components.v1.html(html_preview, height=500, scrolling=True)

# 下載按鈕
st.download_button(
    label="📥 下載 HTML 報表 (自動合併日期儲存格)",
    data=html_preview.encode("utf-8"),
    file_name=f"砂石車勤務表_{month_val}.html",
    mime="text/html"
)
