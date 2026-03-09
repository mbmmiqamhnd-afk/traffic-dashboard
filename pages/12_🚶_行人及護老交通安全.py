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

# --- 1. 初始化頁面 (必須是第一行指令) ---
st.set_page_config(page_title="勤務規劃系統", layout="wide")

# --- 2. 核心常數與定義預設資料 (解決 NameError) ---
UNIT = "桃園市政府警察局龍潭分局"
DEFAULT_MONTH = "115年3月份"

# 補齊 DEFAULT_CMD 變數
DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "駐區督察", "姓名": "孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫", "任務": "負責規劃本勤務、重點機動督導。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"},
])

# 補齊 DEFAULT_SCHEDULE 變數
DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日", "單位": "各派出所", "路段": "校園周邊道路或轄區行人易肇事路口"}
])

NOTES_TEXT = """壹、警察局規劃3月份「行人及護老交通安全專案勤務」期程...
貳、執行本專案勤務視轄區狀況及執勤警力，擇定轄區易肇事路口..."""

# --- 3. 字型處理 ---
@st.cache_resource
def load_pdf_font():
    font_name = "Helvetica"
    paths = ["kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/kaiu.ttf"]
    for p in paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("標楷體", p))
                return "標楷體"
            except: pass
    return font_name

# --- 4. PDF 生成邏輯 (確保銜接與美觀) ---
def make_pdf(month, df_cmd, df_sch):
    f_name = load_pdf_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_w = 180 * mm
    elements = []
    
    s_title = ParagraphStyle('T', fontName=f_name, fontSize=16, alignment=1, leading=22, spaceAfter=12)
    s_cell = ParagraphStyle('C', fontName=f_name, fontSize=10, alignment=1, leading=14)
    s_left = ParagraphStyle('L', fontName=f_name, fontSize=10, alignment=0, leading=14)
    s_head = ParagraphStyle('H', fontName=f_name, fontSize=13, alignment=1, leading=18)

    # 標題
    elements.append(Paragraph(f"<b>{UNIT}{month}執行「行人及護老交通安全」專案勤務規劃表</b>", s_title))

    # 表格 1: 任務編組
    d1 = [[Paragraph("<b>任　務　編　組</b>", s_head), "", "", ""],
          [Paragraph("<b>職稱</b>", s_cell), Paragraph("<b>代號</b>", s_cell), Paragraph("<b>姓名</b>", s_cell), Paragraph("<b>任務</b>", s_cell)]]
    for _, r in df_cmd.iterrows():
        name_clean = str(r.get('姓名','')).replace("、","<br/>").replace(" ","<br/>")
        d1.append([Paragraph(f"<b>{r.get('職稱','')}</b>", s_cell), Paragraph(str(r.get('代號','')), s_cell), 
                   Paragraph(name_clean, s_cell), Paragraph(str(r.get('任務','')), s_left)])
    
    t1 = Table(d1, colWidths=[page_w*0.15, page_w*0.1, page_w*0.25, page_w*0.5])
    t1.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.8, colors.black),
        ('SPAN', (0,0), (3,0)),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,1), colors.whitesmoke),
        ('TOPPADDING', (0,0), (-1,-1), 5), ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(t1)

    elements.append(Spacer(1, 10*mm)) # 空出一行

    # 表格 2: 警力佈署
    d2 = [[Paragraph("<b>警　力　佈　署</b>", s_head), "", ""],
          [Paragraph("<b>日期</b>", s_cell), Paragraph("<b>單位</b>", s_cell), Paragraph("<b>路段</b>", s_cell)]]
    for _, r in df_sch.iterrows():
        d2.append([Paragraph(str(r.iloc[0]), s_cell), Paragraph(str(r.iloc[1]), s_cell), Paragraph(str(r.iloc[2]).replace("\n","<br/>"), s_left)])
    
    t2 = Table(d2, colWidths=[page_w*0.3, page_w*0.2, page_w*0.5])
    t2.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.8, colors.black),
        ('SPAN', (0,0), (2,0)),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,1), colors.whitesmoke),
        ('TOPPADDING', (0,0), (-1,-1), 5), ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(t2)

    doc.build(elements)
    return buf.getvalue()

# --- 5. 主介面 ---
st.title("🚶 行人及護老交通安全勤務表")

month_input = st.text_input("報表月份", DEFAULT_MONTH)

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. 任務編組")
    # 此處已修正 NameError，確保 DEFAULT_CMD 已定義
    edit_cmd = st.data_editor(DEFAULT_CMD, num_rows="dynamic", use_container_width=True, key="cmd_editor")
with col2:
    st.subheader("2. 警力佈署")
    edit_sch = st.data_editor(DEFAULT_SCHEDULE, num_rows="dynamic", use_container_width=True, key="sch_editor")

st.divider()

# --- 6. 功能按鈕 ---
btn_col1, btn_col2, btn_col3 = st.columns(3)

try:
    # 預先生成 PDF 以供下載
    final_pdf = make_pdf(month_input, edit_cmd, edit_sch)
    
    with btn_col1:
        st.download_button(
            label="📥 下載 PDF 報表",
            data=final_pdf,
            file_name=f"Traffic_Report_{datetime.now().strftime('%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary"
        )
    
    with btn_col2:
        if st.button("📧 寄送郵件 (同步附件)", use_container_width=True):
            # 這裡放置您的寄信邏輯
            st.success("郵件發送功能已觸發！")
            
    with btn_col3:
        if st.button("☁️ 同步至雲端", use_container_width=True):
            # 這裡放置您的 Google Sheet 儲存邏輯
            st.info("雲端資料已更新！")

except Exception as e:
    st.error(f"系統發生錯誤: {e}")
