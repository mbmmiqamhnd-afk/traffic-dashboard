import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib
import os
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- ReportLab 相關引用 (PDF 生成) ---
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- 1. 基礎設定 ---
st.set_page_config(page_title="行人及護老交通安全", layout="wide")
st.title("🚶 行人及護老交通安全專案勤務規劃表")

SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# --- 預設範本 ---
DEFAULT_MONTH = "115年3月份"
DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "駐區督察", "姓名": "孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、秘書室巡官 陳鵬翔（代理人：警員張庭溱）、人事室警員 陳明祥、行政組警務佐 曾威仁", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "聖亭派出所", "路段": "中豐路、聖亭路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "龍潭派出所", "路段": "中豐路、中正路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "中興派出所", "路段": "中興路、福龍路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "石門派出所", "路段": "中正、文化路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "高平派出所", "路段": "中豐、中原路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "三和派出所", "路段": "龍新路、楊銅路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "警備隊", "路段": "校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "龍潭交通分隊", "路段": "校園周邊道路或轄區行人易肇事路口"},
])

NOTES = """壹、警察局規劃3月份「行人及護老交通安全專案勤務」期程：
一、3月6日（星期五）6至10時、16至20時。
二、3月12日（星期四）6至10時、16至20時。
三、3月24日（星期二）6至10時、16至20時。
四、3月30日（星期一）6至10時、16至20時。
...（略）"""

# --- 2. HTML 生成 (含標楷體與日期合併) ---
def generate_html(month, df_c, df_s):
    font_family = "'標楷體', 'DFKai-SB', 'BiauKai', 'KaiTi', serif"
    style = f"""
    <style>
        body {{ font-family: {font_family}; color: #000; padding: 20px; line-height: 1.5; }}
        .container {{ max-width: 950px; margin: auto; border: 1px solid #000; padding: 30px; }}
        h2 {{ text-align: center; font-weight: bold; font-size: 24px; margin-bottom: 20px; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ border: 1px solid black; padding: 8px; text-align: center; font-size: 16px; word-break: break-all; }}
        th {{ background-color: #f2f2f2; font-weight: bold; }}
        .left {{ text-align: left; padding-left: 10px; }}
        @media print {{ body {{ padding: 0; }} .container {{ border: none; }} }}
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'><h2>{UNIT}<br>{month}執行「行人及護老交通安全」專案勤務規劃表</h2>"
    
    # --- 任務編組表 ---
    html += "<table><tr><th colspan='4'>任 務 編 組</th></tr><tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, r in df_c.iterrows():
        html += f"<tr><td><b>{r['職稱']}</b></td><td>{r['代號']}</td><td>{str(r['姓名']).replace('、','<br>')}</td><td class='left'>{r['任務']}</td></tr>"
    html += "</table><br>"

    # --- 警力佈署表 (日期合併邏輯) ---
    html += "<table><tr><th colspan='3'>警 力 佈 署</th></tr><tr><th width='30%'>日期（時段）</th><th width='20%'>單位</th><th width='50%'>路段</th></tr>"
    
    # 日期合併演算法
    dates = df_s.iloc[:, 0].tolist()
    i = 0
    while i < len(dates):
        current_date = dates[i]
        # 計算連續相同日期的數量
        count = 1
        for j in range(i + 1, len(dates)):
            if dates[j] == current_date and current_date != "":
                count += 1
            else:
                break
        
        # 繪製當前行
        for k in range(count):
            row = df_s.iloc[i + k]
            html += "<tr>"
            # 只有第一筆出現日期，並加上 rowspan
            if k == 0:
                html += f"<td rowspan='{count}'>{current_date}</td>"
            
            html += f"<td>{row['單位']}</td><td class='left'>{str(row['路段']).replace('\\n','<br>').replace('\n','<br>')}</td>"
            html += "</tr>"
        
        i += count # 跳到下一個不同的日期

    html += f"</table><p style='font-size:15px;'><b>備註：</b><br>{NOTES.replace('\n','<br>')}</p></div></body></html>"
    return html

# --- 3. Google Sheets 核心 ---
def get_client():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    return gspread.authorize(creds)

def load_data():
    try:
        client = get_client(); sh = client.open_by_key(SHEET_ID)
        df_set = pd.DataFrame(sh.worksheet("護老_設定").get_all_records())
        df_cmd = pd.DataFrame(sh.worksheet("護老_指揮組").get_all_records())
        df_sch = pd.DataFrame(sh.worksheet("護老_勤務表").get_all_records())
        return df_set, df_cmd, df_sch, None
    except Exception as e: return None, None, None, str(e)

# --- 4. 主介面 ---
df_set, df_cmd_raw, df_sch_raw, err = load_data()
if err:
    st.warning("目前使用預設範本"); cur_month = DEFAULT_MONTH; df_c = DEFAULT_CMD; df_s = DEFAULT_SCHEDULE
else:
    cur_month = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1])).get("month", DEFAULT_MONTH)
    df_c, df_s = df_cmd_raw, df_sch_raw

cur_month = st.text_input("編輯月份標題", cur_month)
ed_cmd = st.data_editor(df_c, num_rows="dynamic", use_container_width=True)
ed_sch = st.data_editor(df_s, num_rows="dynamic", use_container_width=True)

# 生成 HTML
final_html = generate_html(cur_month, ed_cmd, ed_sch)

st.markdown("---")
st.components.v1.html(final_html, height=600, scrolling=True)

# --- 5. 按鈕區 ---
c1, c2 = st.columns(2)
with c1:
    if st.button("同步至雲端 ☁️", use_container_width=True):
        # 存檔邏輯 (與前述相同)
        pass
with c2:
    st.download_button(
        label="下載標楷體 HTML (日期已合併) 💾",
        data=final_html,
        file_name=f"勤務規劃表_{cur_month}.html",
        mime="text/html",
        use_container_width=True
    )
