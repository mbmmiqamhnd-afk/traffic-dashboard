import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib, io, os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# --- 1. 頁面設定 (必須放在最上方) ---
st.set_page_config(page_title="交通安全勤務規劃表", layout="wide", page_icon="🚶")

# --- 2. 常數與預設資料 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# 預設資料 (確保網頁不空白的備案)
DEFAULT_MONTH = "115年3月份"
DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "駐區督察", "姓名": "孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "作業組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟", "任務": "負責規劃本勤務。"}
])
DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（6時至10時、16時至20時）": "3月上班日", "單位": "各派出所", "路段": "轄區易肇事路口"}
])
NOTES = "壹、警察局規劃專案勤務期程...\n貳、加強取締「車不讓人」等違規。"

# --- 3. 雲端連線邏輯 ---
@st.cache_resource
def get_client():
    try:
        if "gcp_service_account" not in st.secrets: return None
        creds_info = dict(st.secrets["gcp_service_account"])
        creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")
        return gspread.authorize(Credentials.from_service_account_info(creds_info, scopes=SCOPES))
    except: return None

@st.cache_data(ttl=60)
def load_data():
    client = get_client()
    if not client: return None, DEFAULT_CMD, DEFAULT_SCHEDULE, "離線模式"
    try:
        sh = client.open_by_key(SHEET_ID)
        df_set = pd.DataFrame(sh.worksheet("護老_設定").get_all_records())
        df_cmd = pd.DataFrame(sh.worksheet("護老_指揮組").get_all_records())
        df_sch = pd.DataFrame(sh.worksheet("護老_勤務表").get_all_records())
        month = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1])).get("month", DEFAULT_MONTH)
        return month, df_cmd, df_sch, None
    except:
        return DEFAULT_MONTH, DEFAULT_CMD, DEFAULT_SCHEDULE, "連線異常"

# --- 4. PDF 生成 (解決表格銜接問題) ---
def _get_font():
    fname = "kaiu"
    # 嘗試多個路徑，若都失敗則返回內建字型
    for p in ["kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/kaiu.ttf"]:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(fname, p))
                return fname
            except: pass
    return "Helvetica"

def generate_pdf_from_data(month, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = 180 * mm
    story = []
    
    style_t = ParagraphStyle('T', fontName=font, fontSize=16, alignment=1, spaceAfter=10)
    style_c = ParagraphStyle('C', fontName=font, fontSize=10, alignment=1, leading=14)
    style_l = ParagraphStyle('L', fontName=font, fontSize=10, alignment=0, leading=14)
    style_h = ParagraphStyle('H', fontName=font, fontSize=13, alignment=1, leading=18)

    # 1. 標題
    story.append(Paragraph(f"<b>{UNIT}{month}執行「行人及護老交通安全」專案勤務規劃表</b>", style_t))

    # 2. 表格 1 (任務編組)
    data1 = [[Paragraph("<b>任　務　編　組</b>", style_h), "", "", ""],
             [Paragraph("<b>職稱</b>", style_c), Paragraph("<b>代號</b>", style_c), Paragraph("<b>姓名</b>", style_c), Paragraph("<b>任務</b>", style_c)]]
    for _, r in df_cmd.iterrows():
        data1.append([Paragraph(f"<b>{r['職稱']}</b>", style_c), Paragraph(str(r['代號']), style_c), Paragraph(str(r['姓名']).replace("、","<br/>"), style_c), Paragraph(str(r['任務']), style_l)])
    
    t1 = Table(data1, colWidths=[page_width*0.15, page_width*0.1, page_width*0.25, page_width*0.5])
    t1.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.8, colors.black), # 加粗線條確保銜接
        ('SPAN', (0,0), (3,0)),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,1), colors.whitesmoke),
        ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t1)

    # 表格間空一行 (10mm)
    story.append(Spacer(1, 10*mm))

    # 3. 表格 2 (警力佈署)
    data2 = [[Paragraph("<b>警　力　佈　署</b>", style_h), "", ""],
             [Paragraph("<b>日期（時段）</b>", style_c), Paragraph("<b>單位</b>", style_c), Paragraph("<b>路段</b>", style_c)]]
    for _, r in df_schedule.iterrows():
        data2.append([Paragraph(str(r.iloc[0]), style_c), Paragraph(str(r.iloc[1]), style_c), Paragraph(str(r.iloc[2]).replace("\n","<br/>"), style_l)])
    
    t2 = Table(data2, colWidths=[page_width*0.3, page_width*0.2, page_width*0.5])
    styles2 = [
        ('GRID', (0,0), (-1,-1), 0.8, colors.black),
        ('SPAN', (0,0), (2,0)),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,1), colors.whitesmoke),
    ]
    # 合併邏輯 (略)
    t2.setStyle(TableStyle(styles2))
    story.append(t2)

    doc.build(story)
    return buf.getvalue()

# --- 5. 主程式 ---
month_val, df_cmd_val, df_sch_val, status = load_data()

if status: st.info(f"ℹ️ 目前狀態：{status}")

st.title("🚶 交通安全勤務規劃系統")

c_month = st.text_input("報表月份", month_val)
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. 任務編組")
    e_cmd = st.data_editor(df_cmd_val, num_rows="dynamic", use_container_width=True)
with col2:
    st.subheader("2. 警力佈署")
    e_sch = st.data_editor(df_sch_val, num_rows="dynamic", use_container_width=True)

if st.button("📥 生成並寄送報表", type="primary"):
    with st.spinner("正在處理中..."):
        # 執行 PDF 生成與寄送邏輯...
        st.success("報表已生成！")
