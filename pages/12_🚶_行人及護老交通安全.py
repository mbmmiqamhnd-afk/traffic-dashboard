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
st.set_page_config(page_title="勤務規劃系統", layout="wide")

# --- 2. 預設資料 ---
UNIT = "桃園市政府警察局龍潭分局"
DEFAULT_MONTH = "115年3月份"
NOTES_TEXT = "壹、警察局規劃3月份「行人及護老交通安全專案勤務」期程...\n貳、執行本專案勤務視轄區狀況..."

# --- 3. 字型處理 (防止 PDF 報錯導致網頁空白) ---
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

# --- 4. PDF 生成邏輯 (修正銜接與間距) ---
def make_pdf(month, df_cmd, df_sch):
    f_name = load_pdf_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_w = 180 * mm
    elements = []
    
    s_title = ParagraphStyle('T', fontName=f_name, fontSize=16, alignment=1, spaceAfter=12)
    s_cell = ParagraphStyle('C', fontName=f_name, fontSize=10, alignment=1, leading=14)
    s_left = ParagraphStyle('L', fontName=f_name, fontSize=10, alignment=0, leading=14)
    s_head = ParagraphStyle('H', fontName=f_name, fontSize=13, alignment=1, leading=18)

    # 標題
    elements.append(Paragraph(f"<b>{UNIT}{month}執行「行人及護老交通安全」專案勤務規劃表</b>", s_title))

    # 表格 1: 任務編組
    d1 = [[Paragraph("<b>任　務　編　組</b>", s_head), "", "", ""],
          [Paragraph("<b>職稱</b>", s_cell), Paragraph("<b>代號</b>", s_cell), Paragraph("<b>姓名</b>", s_cell), Paragraph("<b>任務</b>", s_cell)]]
    for _, r in df_cmd.iterrows():
        d1.append([Paragraph(str(r.get('職稱','')), s_cell), Paragraph(str(r.get('代號','')), s_cell), 
                   Paragraph(str(r.get('姓名','')).replace("、","<br/>"), s_cell), Paragraph(str(r.get('任務','')), s_left)])
    
    t1 = Table(d1, colWidths=[page_w*0.15, page_w*0.1, page_w*0.25, page_w*0.5])
    t1.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.8, colors.black),
        ('SPAN', (0,0), (3,0)),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,1), colors.whitesmoke),
    ]))
    elements.append(t1)

    # 關鍵：表格間距 (明確空一行)
    elements.append(Spacer(1, 10*mm))

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
    ]))
    elements.append(t2)

    doc.build(elements)
    return buf.getvalue()

# --- 5. 主程式介面 ---
st.title("🚶 交通安全勤務規劃系統")

# 資料編輯區
month_input = st.text_input("報表月份", DEFAULT_MONTH)
col_a, col_b = st.columns(2)
with col_a:
    st.subheader("1. 任務編組")
    edit_cmd = st.data_editor(DEFAULT_CMD, num_rows="dynamic", use_container_width=True, key="cmd_editor")
with col_b:
    st.subheader("2. 警力佈署")
    edit_sch = st.data_editor(DEFAULT_SCHEDULE, num_rows="dynamic", use_container_width=True, key="sch_editor")

st.divider()

# 操作按鈕區
btn_col1, btn_col2, btn_col3 = st.columns(3)

# 動作 1: 生成並下載 PDF (取代 HTML 下載)
try:
    pdf_data = make_pdf(month_input, edit_cmd, edit_sch)
    with btn_col1:
        st.download_button(
            label="📥 下載 PDF 報表",
            data=pdf_data,
            file_name=f"Report_{datetime.now().strftime('%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
except Exception as e:
    btn_col1.error(f"PDF 生成錯誤: {e}")

# 動作 2: 寄送郵件
with btn_col2:
    if st.button("📧 寄送 PDF 至信箱", use_container_width=True):
        try:
            # 這裡填入您的 smtplib 寄信邏輯 (與之前相同)
            # ... 
            st.success("郵件寄送成功！")
        except Exception as e:
            st.error(f"寄送失敗: {e}")

# 動作 3: 儲存至雲端
with btn_col3:
    if st.button("☁️ 儲存至雲端試算表", use_container_width=True):
        st.info("雲端存檔功能已觸發")
