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

# --- 1. 初始變數定義 (防止 NameError) ---
UNIT = "桃園市政府警察局龍潭分局"
DEFAULT_MONTH = "115年3月份"
current_month = DEFAULT_MONTH  # 預設初始化

# --- 2. 字型註冊函數 (解決標楷體) ---
@st.cache_resource
def get_chinese_font():
    font_name = "標楷體"
    # 增加 Streamlit Cloud 可能的絕對路徑
    paths = [
        "kaiu.ttf",
        "/mount/src/traffic-dashboard/kaiu.ttf",  # Streamlit Cloud 常用路徑
        os.path.join(os.getcwd(), "kaiu.ttf")
    ]
    for p in paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(font_name, p))
                return font_name
            except: pass
    return "Helvetica"

# --- 3. PDF 生成函數 (按單位分開) ---
def make_final_pdf(month, df_cmd, df_sch):
    f_name = get_chinese_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_w = 180 * mm
    elements = []
    
    s_title = ParagraphStyle('T', fontName=f_name, fontSize=16, alignment=1, spaceAfter=10)
    s_head = ParagraphStyle('H', fontName=f_name, fontSize=12, alignment=1, leading=16)
    s_c = ParagraphStyle('C', fontName=f_name, fontSize=10, alignment=1, leading=14)
    s_l = ParagraphStyle('L', fontName=f_name, fontSize=10, alignment=0, leading=14)

    elements.append(Paragraph(f"<b>{UNIT}{month}執行「行人及護老交通安全」專案勤務規劃表</b>", s_title))
    
    # 任務編組表格
    d1 = [[Paragraph("<b>任 務 編 組</b>", s_head), "", "", ""]]
    d1.append([Paragraph("<b>職稱</b>", s_c), Paragraph("<b>代號</b>", s_c), Paragraph("<b>姓名</b>", s_c), Paragraph("<b>任務</b>", s_c)])
    for _, r in df_cmd.iterrows():
        d1.append([Paragraph(str(r.iloc[0]), s_c), Paragraph(str(r.iloc[1]), s_c), 
                   Paragraph(str(r.iloc[2]).replace("、","<br/>"), s_c), Paragraph(str(r.iloc[3]), s_l)])
    
    t1 = Table(d1, colWidths=[page_w*0.15, page_w*0.1, page_w*0.25, page_w*0.5])
    t1.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.7,colors.black), ('SPAN',(0,0),(3,0)), ('BACKGROUND',(0,0),(-1,1),colors.whitesmoke)]))
    elements.append(t1)
    elements.append(Spacer(1, 10*mm)) # 關鍵：空出一行高度

    # 警力佈署 (按單位拆分)
    if '單位' in df_sch.columns:
        units = df_sch['單位'].unique()
        for u in units:
            if not u: continue
            u_df = df_sch[df_sch['單位'] == u]
            d2 = [[Paragraph(f"<b>執行單位：{u}</b>", s_head), ""]]
            d2.append([Paragraph("<b>日期/時段</b>", s_c), Paragraph("<b>路段細節</b>", s_c)])
            for _, r in u_df.iterrows():
                d2.append([Paragraph(str(r.iloc[0]), s_c), Paragraph(str(r.iloc[2]).replace("\n","<br/>"), s_l)])
            
            ut = Table(d2, colWidths=[page_w*0.35, page_w*0.65])
            ut.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.7,colors.black), ('SPAN',(0,0),(1,0)), ('BACKGROUND',(0,0),(1,1),colors.lightgrey)]))
            elements.append(ut)
            elements.append(Spacer(1, 5*mm))

    doc.build(elements)
    return buf.getvalue()

# --- 4. 主程式介面 ---
st.title("🚶 行人及護老交通安全勤務系統")

# 這裡從編輯器獲取月份
current_month = st.text_input("報表月份", DEFAULT_MONTH)

# ... (中間的資料編輯器 edited_cmd, edited_schedule 等代碼) ...

# --- 5. 執行區 ---
if st.button("🚀 生成、存檔並寄送 PDF"):
    try:
        # 確保變數存在
        pdf_data = make_final_pdf(current_month, edited_cmd, edited_schedule)
        
        # 寄信邏輯 (使用 st.secrets)
        sender = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg["Subject"] = f"勤務表_{current_month}"
        msg["From"] = sender
        msg["To"] = sender
        
        part = MIMEBase("application", "octet-stream")
        part.set_payload(pdf_data)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment; filename=Report.pdf")
        msg.attach(part)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, sender, msg.as_string())
            
        st.success("✅ 郵件已寄出，標楷體 PDF 生成成功！")
        st.download_button("📥 下載備份 PDF", data=pdf_data, file_name="Report.pdf")
        
    except Exception as e:
        st.error(f"❌ 發生錯誤: {e}")
