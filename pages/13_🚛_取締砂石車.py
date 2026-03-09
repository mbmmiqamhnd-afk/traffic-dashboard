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

# --- ReportLab 相關引用 (PDF 生成核心) ---
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- 1. 頁面設定 ---
st.set_page_config(page_title="取締砂石車專案勤務", layout="wide")
st.title("🚛 取締砂石（大型貨）車重點違規專案勤務規劃表")
st.caption("資料與 Google Sheets 即時連線，手機、電腦皆可編輯")

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
    {"日期": "115年3月13日（星期五）00時至24時", "執行單位": "聖亭派出所", "執行人數": "2至4人", "執行路段": "中豐路、聖亭路段等砂石（大型貨）車行經路段"},
    {"日期": "", "執行單位": "龍潭派出所", "執行人數": "", "執行路段": "大昌路、中豐路段砂石（大型貨）車行經路段"},
    {"日期": "", "執行單位": "中興派出所", "執行人數": "", "執行路段": "中興路、福龍路及龍平路段砂石（大型貨）車行經路段"},
    {"日期": "", "執行單位": "石門派出所", "執行人數": "", "執行路段": "中正路、龍源路及民族路段砂石（大型貨）車行經路段"},
    {"日期": "", "執行單位": "高平派出所", "執行人數": "", "執行路段": "中豐路、龍源路段砂石（大型貨）車行經路段"},
    {"日期": "", "執行單位": "三和派出所", "執行人數": "", "執行路段": "楊銅路、龍新路段砂石（大型貨）車行經路段"},
    {"日期": "", "執行單位": "警備隊", "執行人數": "", "執行路段": "中豐路、龍源路、聖亭路段砂石（大型貨）車行經路段"},
    {"日期": "", "執行單位": "龍潭交通分隊", "執行人數": "", "執行路段": "中豐路、龍源路、聖亭路段砂石（大型貨）車行經路段"},
])

NOTES = "※ 加強取締砂石（大型貨）車超載、車速、酒醉駕車、闖紅燈、無照駕車、爭道行駛、違反禁行路線、變更車斗、未使用專用車箱及未裝設行車紀錄器（行車視野輔助器）等違規，以共同消弭不法行為，保障用路人生命財產安全。"

# --- 2. 工具函式 ---
def _get_font():
    fname = "kaiu"
    paths = ['/mount/src/traffic-dashboard/kaiu.ttf', 'kaiu.ttf', './kaiu.ttf']
    for p in paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(fname, p))
                return fname
            except: continue
    return "Helvetica"

def generate_pdf(month, briefing, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    W = A4[0] - 30*mm
    story = []
    s_title = ParagraphStyle("t", fontName=font, fontSize=13, alignment=1, spaceAfter=2, leading=18)
    s_cell = ParagraphStyle("c", fontName=font, fontSize=9, leading=13, alignment=1)
    s_left = ParagraphStyle("l", fontName=font, fontSize=9, leading=13, alignment=0)
    s_note = ParagraphStyle("n", fontName=font, fontSize=9, leading=14)

    def c(txt, style=None):
        txt = str(txt).replace("\n","<br/>").replace("、","<br/>")
        return Paragraph(txt, style or s_cell)

    story.append(Paragraph(f"{UNIT}執行{month}「取締砂石（大型貨）車重點違規」專案勤務規劃表", s_title))
    story.append(Spacer(1, 2*mm))

    # 任務編組
    data1 = [[Paragraph("<b>任　務　編　組</b>", s_title),'','','']]
    data1.append([c("<b>職稱</b>"),c("<b>代號</b>"),c("<b>姓名</b>"),c("<b>任務</b>")])
    for _, row in df_cmd.iterrows():
        data1.append([c(f"<b>{row.get('職稱','')}</b>"),c(row.get('代號','')), c(row.get('姓名','')),c(row.get('任務',''),s_left)])
    t1 = Table(data1, colWidths=[W*0.15, W*0.10, W*0.25, W*0.50], repeatRows=2)
    t1.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.5,colors.black), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')), ('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t1)
    
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(f"<b>📢 勤前教育：</b>{briefing.replace('\n','<br/>')}", s_note))
    story.append(Spacer(1, 3*mm))

    # 警力佈署
    data2 = [[Paragraph("<b>警　力　佈　署</b>", s_title),'','','']]
    data2.append([c("<b>日期</b>"),c("<b>執行單位</b>"),c("<b>執行人數</b>"),c("<b>執行路段</b>")])
    for _, row in df_schedule.iterrows():
        data2.append([c(row.get('日期','')),c(row.get('執行單位','')), c(row.get('執行人數','')),c(row.get('執行路段',''),s_left)])
    t2 = Table(data2, colWidths=[W*0.25, W*0.20, W*0.10, W*0.45], repeatRows=2)
    t2.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.5,colors.black), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')), ('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t2)
    
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(f"<b>備註：</b>{NOTES}", s_note))
    doc.build(story)
    return buf.getvalue()

# --- 3. Google Sheets 連線 ---
def get_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def load_data():
    try:
        client = get_client(); sh = client.open_by_key(SHEET_ID)
        df_set = pd.DataFrame(sh.worksheet("砂石_設定").get_all_records())
        df_cmd = pd.DataFrame(sh.worksheet("砂石_指揮組").get_all_records())
        df_sch = pd.DataFrame(sh.worksheet("砂石_勤務表").get_all_records())
        return df_set, df_cmd, df_sch, None
    except Exception as e: return None, None, None, str(e)

def save_data(month, briefing, df_cmd, df_schedule):
    try:
        client = get_client(); sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("砂石_設定")
        ws_set.clear(); ws_set.update([["Key", "Value"], ["month", month], ["briefing", briefing]])
        ws_cmd = sh.worksheet("砂石_指揮組")
        ws_cmd.clear(); ws_cmd.update([df_cmd.columns.tolist()] + df_cmd.fillna("").values.tolist())
        ws_sch = sh.worksheet("砂石_勤務表")
        ws_sch.clear(); ws_sch.update([df_schedule.columns.tolist()] + df_schedule.fillna("").values.tolist())
        st.toast("✅ 雲端存檔成功！", icon="☁️")
    except Exception as e: st.error(f"❌ 存檔失敗：{e}")

# --- 4. 主介面邏輯 ---
df_set, df_cmd_raw, df_sch_raw, err = load_data()
if err:
    st.info("💡 已載入預設範本。")
    cur_month, cur_brief, df_c, df_s = DEFAULT_MONTH, DEFAULT_BRIEF, DEFAULT_CMD.copy(), DEFAULT_SCHEDULE.copy()
else:
    sd = dict(zip(df_set.iloc[:, 0], df_set.iloc[:, 1]))
    cur_month = sd.get("month", DEFAULT_MONTH)
    cur_brief = sd.get("briefing", DEFAULT_BRIEF)
    df_c, df_s = df_cmd_raw, df_sch_raw

st.subheader("1. 基礎資訊")
c1, c2 = st.columns(2)
cur_month = c1.text_input("月份", value=cur_month)
brief_info = c2.text_area("📢 勤前教育", value=cur_brief, height=80)

st.subheader("2. 任務編組")
ed_cmd = st.data_editor(df_c, num_rows="dynamic", use_container_width=True)

st.subheader("3. 執行任務單位、時間及路段")
ed_sch = st.data_editor(df_s, num_rows="dynamic", use_container_width=True)

# --- 5. HTML 生成 (強制標楷體) ---
def generate_html(month, briefing, df_cmd, df_schedule):
    # 設定標楷體優先順序 (Windows/Mac/Linux)
    font_family = "'標楷體', 'DFKai-SB', 'BiauKai', 'KaiTi', serif"
    style = f"""
    <style>
        body {{ font-family: {font_family}; color: #000; font-size: 15px; line-height: 1.6; padding: 20px; }}
        .container {{ width: 100%; max-width: 850px; margin: 0 auto; border: 1px solid #000; padding: 30px; }}
        h2 {{ text-align: center; margin-bottom: 20px; font-weight: bold; }}
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; }}
        th, td {{ border: 1px solid black; padding: 8px; text-align: center; vertical-align: middle; }}
        th {{ background-color: #f2f2f2; font-weight: bold; }}
        .left {{ text-align: left; }}
        .notes {{ font-size: 14px; white-space: pre-wrap; }}
        @media print {{ body {{ padding: 0; }} .container {{ border: none; }} }}
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2>{UNIT}<br>執行{month}「取締砂石（大型貨）車重點違規」專案勤務規劃表</h2>"
    
    # 任務編組
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr>"
    html += "<tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, r in df_cmd.iterrows():
        name = str(r.get('姓名', '')).replace("、", "<br>")
        html += f"<tr><td><b>{r.get('職稱','')}</b></td><td>{r.get('代號','')}</td><td>{name}</td><td class='left'>{r.get('任務','')}</td></tr>"
    html += "</table>"

    html += f"<p><b>📢 勤前教育：</b><br>{briefing.replace('\\n','<br>').replace('\n','<br>')}</p>"

    # 警力佈署 (原始逐行版本)
    html += "<table><tr><th colspan='4'>警　力　佈　署</th></tr>"
    html += "<tr><th width='25%'>日期</th><th width='20%'>執行單位</th><th width='10%'>人數</th><th width='45%'>執行路段</th></tr>"
    for _, r in df_schedule.iterrows():
        html += f"<tr><td>{r.get('日期','')}</td><td>{r.get('執行單位','')}</td><td>{r.get('執行人數','')}</td><td class='left'>{str(r.get('執行路段','')).replace('\\n','<br>')}</td></tr>"
    html += f"</table><p class='notes'><b>備註：</b><br>{NOTES}</p></div></body></html>"
    return html

html_out = generate_html(cur_month, brief_info, ed_cmd, ed_sch)

st.markdown("---")
st.subheader("📄 標楷體預覽")
st.components.v1.html(html_out, height=600, scrolling=True)

# --- 6. 輸出按鈕 ---
col_dl, col_mail = st.columns(2)
with col_dl:
    if st.download_button("下載 HTML 報表 (標楷體) 💾", data=html_out.encode("utf-8"), 
                          file_name=f"砂石車勤務表_{cur_month}.html", mime="text/html"):
        save_data(cur_month, brief_info, ed_cmd, ed_sch)
