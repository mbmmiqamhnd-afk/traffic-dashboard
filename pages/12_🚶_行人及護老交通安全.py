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

# --- 1. 初始化頁面 ---
st.set_page_config(page_title="各單位拆分報表系統", layout="wide", page_icon="🚶")

# --- 2. 預設資料 (包含各個單位) ---
UNIT_FULL = "桃園市政府警察局龍潭分局"
DEFAULT_MONTH = "115年3月份"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "作業組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟", "任務": "負責規劃本勤務。"}
])

# 這裡模擬試算表匯入的多單位資料
DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（時段）": "3月2～6日", "單位": "聖亭派出所", "路段": "中豐路、聖亭路段"},
    {"日期（時段）": "3月2～6日", "單位": "龍潭派出所", "路段": "中正路、五福街路口"},
    {"日期（時段）": "3月2～6日", "單位": "中興派出所", "路段": "中興路段"},
    {"日期（時段）": "3月9～13日", "單位": "聖亭派出所", "路段": "校園周邊道路"},
    {"日期（時段）": "3月9～13日", "單位": "龍潭派出所", "路段": "北龍路口"}
])

# --- 3. 字型載入 ---
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

# --- 4. 核心 PDF 生成 (按單位分開表格) ---
def generate_split_pdf(month, df_cmd, df_sch):
    f_name = load_pdf_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_w = 180 * mm
    elements = []
    
    # 樣式定義
    s_title = ParagraphStyle('T', fontName=f_name, fontSize=16, alignment=1, spaceAfter=10)
    s_cell_c = ParagraphStyle('C', fontName=f_name, fontSize=10, alignment=1, leading=14)
    s_cell_l = ParagraphStyle('L', fontName=f_name, fontSize=10, alignment=0, leading=14)
    s_head = ParagraphStyle('H', fontName=f_name, fontSize=12, alignment=1, leading=16)

    # 1. 標題
    elements.append(Paragraph(f"<b>{UNIT_FULL}{month}執行「行人及護老交通安全」勤務規劃表</b>", s_title))

    # 2. 任務編組表格
    cmd_data = [[Paragraph("<b>任 務 編 組</b>", s_head), "", "", ""],
                [Paragraph("<b>職稱</b>", s_cell_c), Paragraph("<b>代號</b>", s_cell_c), Paragraph("<b>姓名</b>", s_cell_c), Paragraph("<b>任務</b>", s_cell_c)]]
    for _, r in df_cmd.iterrows():
        cmd_data.append([Paragraph(r['職稱'], s_cell_c), Paragraph(str(r['代號']), s_cell_c), Paragraph(str(r['姓名']).replace("、","<br/>"), s_cell_c), Paragraph(r['任務'], s_cell_l)])
    
    t1 = Table(cmd_data, colWidths=[page_w*0.15, page_w*0.1, page_w*0.25, page_w*0.5])
    t1.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.7,colors.black),('SPAN',(0,0),(3,0)),('BACKGROUND',(0,0),(-1,1),colors.whitesmoke),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    elements.append(t1)
    
    elements.append(Spacer(1, 10*mm)) # 空一行

    # 3. 警力佈署：按「單位」分開
    elements.append(Paragraph("<b>警 力 佈 署 (按單位區分)</b>", s_head))
    elements.append(Spacer(1, 2*mm))

    # 取得所有不重複的單位
    all_units = df_sch['單位'].unique()
    
    for unit in all_units:
        # 過濾出該單位的資料
        unit_df = df_sch[df_sch['單位'] == unit]
        
        # 建立單位的獨立表格
        unit_data = [[Paragraph(f"<b>執行單位：{unit}</b>", s_head), ""],
                     [Paragraph("<b>日期（時段）</b>", s_cell_c), Paragraph("<b>執行路段 / 任務細節</b>", s_cell_c)]]
        
        for _, r in unit_df.iterrows():
            unit_data.append([Paragraph(str(r.iloc[0]), s_cell_c), Paragraph(str(r.iloc[2]).replace("\n","<br/>"), s_cell_l)])
        
        ut = Table(unit_data, colWidths=[page_w*0.35, page_w*0.65])
        ut.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.7, colors.black),
            ('SPAN', (0,0), (1,0)), # 單位名稱跨欄
            ('BACKGROUND', (0,0), (1,1), colors.lightgrey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        elements.append(ut)
        elements.append(Spacer(1, 5*mm)) # 每個單位表格之間留一點小空隙

    doc.build(elements)
    return buf.getvalue()

# --- 5. Streamlit 介面 ---
st.title("🚓 按單位自動拆分報表系統")

month_name = st.text_input("報表月份", DEFAULT_MONTH)

col_l, col_r = st.columns(2)
with col_l:
    st.subheader("1. 任務編組")
    ed_cmd = st.data_editor(DEFAULT_CMD, num_rows="dynamic", use_container_width=True)
with col_r:
    st.subheader("2. 警力佈署 (填入單位會自動拆分)")
    ed_sch = st.data_editor(DEFAULT_SCHEDULE, num_rows="dynamic", use_container_width=True)

st.divider()

# --- 6. 功能執行 ---
if st.download_button(
    label="📥 下載按單位拆分之 PDF",
    data=generate_split_pdf(month_name, ed_cmd, ed_sch),
    file_name=f"Split_Report_{datetime.now().strftime('%m%d')}.pdf",
    mime="application/pdf",
    type="primary"
):
    st.success("PDF 已成功生成！")
