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

# --- 1. 字型與路徑診斷 (解決標楷體顯示問題) ---
def get_font_name():
    """註冊並回傳字型名稱"""
    font_name = "標楷體"
    # 搜尋順序：1. 專案目錄 2. Linux 系統目錄 3. Windows 目錄
    paths = [
        "kaiu.ttf", 
        "/mount/src/traffic-dashboard/kaiu.ttf", 
        "./kaiu.ttf",
        "/usr/share/fonts/truetype/kaiu.ttf",
        "C:/Windows/Fonts/kaiu.ttf"
    ]
    for p in paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(font_name, p))
                return font_name
            except:
                continue
    return "Helvetica" # 最終回退

# --- 2. 寄信功能 (修正連線與附件) ---
def send_email_via_gmail(pdf_data, subject):
    try:
        # 取得 Secrets
        sender_email = st.secrets["email"]["user"]
        # 注意：此處必須使用「應用程式密碼」
        app_password = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = sender_email # 寄給自己
        msg["Subject"] = subject
        
        msg.attach(MIMEText("附件為自動生成的交通安全勤務規劃表 PDF。", "plain", "utf-8"))
        
        # 處理 PDF 附件
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_data)
        encoders.encode_base64(part)
        # 解決中文檔名亂碼問題
        part.add_header("Content-Disposition", f"attachment; filename=Report.pdf")
        msg.attach(part)
        
        # 使用 SSL 連線
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, app_password)
            server.sendmail(sender_email, sender_email, msg.as_string())
        return True, "OK"
    except Exception as e:
        return False, str(e)

# --- 3. 核心 PDF 生成 (支援單位拆分) ---
def make_final_pdf(month, df_cmd, df_sch):
    f_name = get_font_name()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_w = 180 * mm
    elements = []
    
    # 樣式
    s_title = ParagraphStyle('T', fontName=f_name, fontSize=16, alignment=1, spaceAfter=10)
    s_head = ParagraphStyle('H', fontName=f_name, fontSize=12, alignment=1, leading=16)
    s_c = ParagraphStyle('C', fontName=f_name, fontSize=10, alignment=1, leading=14)
    s_l = ParagraphStyle('L', fontName=f_name, fontSize=10, alignment=0, leading=14)

    # 內容
    elements.append(Paragraph(f"<b>{UNIT}{month}執行「行人及護老交通安全」專案勤務規劃表</b>", s_title))
    
    # 任務編組
    d1 = [[Paragraph("<b>任　務　編　組</b>", s_head), "", "", ""]]
    d1.append([Paragraph("<b>職稱</b>", s_c), Paragraph("<b>代號</b>", s_c), Paragraph("<b>姓名</b>", s_c), Paragraph("<b>任務</b>", s_c)])
    for _, r in df_cmd.iterrows():
        d1.append([Paragraph(str(r[0]), s_c), Paragraph(str(r[1]), s_c), 
                   Paragraph(str(r[2]).replace("、","<br/>"), s_c), Paragraph(str(r[3]), s_l)])
    
    t1 = Table(d1, colWidths=[page_w*0.15, page_w*0.1, page_w*0.25, page_w*0.5])
    t1.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.7,colors.black), ('SPAN',(0,0),(3,0)), ('BACKGROUND',(0,0),(-1,1),colors.whitesmoke)]))
    elements.append(t1)
    
    elements.append(Spacer(1, 10*mm)) # 空出一行

    # 警力佈署 (依單位拆分)
    units = df_sch['單位'].unique()
    for u in units:
        u_df = df_sch[df_sch['單位'] == u]
        d2 = [[Paragraph(f"<b>執行單位：{u}</b>", s_head), ""],
              [Paragraph("<b>日期/時段</b>", s_c), Paragraph("<b>執行路段/細節</b>", s_c)]]
        for _, r in u_df.iterrows():
            d2.append([Paragraph(str(r.iloc[0]), s_c), Paragraph(str(r.iloc[2]).replace("\n","<br/>"), s_l)])
        
        ut = Table(d2, colWidths=[page_w*0.35, page_w*0.65])
        ut.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.7,colors.black), ('SPAN',(0,0),(1,0)), ('BACKGROUND',(0,0),(1,1),colors.lightgrey)]))
        elements.append(ut)
        elements.append(Spacer(1, 5*mm))

    doc.build(elements)
    return buf.getvalue()

# --- 4. 主流程與介面 ---
st.title("🚶 交通安全勤務規劃系統")

# ... (中間載入 Google Sheets 資料邏輯與編輯器請保持原狀) ...

if st.button("🚀 生成、寄信並存檔", type="primary"):
    with st.spinner("系統處理中..."):
        # 1. 生成 PDF
        pdf_out = make_final_pdf(current_month, edited_cmd, edited_schedule)
        
        # 2. 寄信
        ok, msg = send_email_via_gmail(pdf_out, f"交通勤務表_{current_month}")
        
        if ok:
            st.success("✅ 郵件已寄送成功！附件已包含標楷體報表。")
        else:
            st.error(f"❌ 寄信失敗。原因：{msg}")
            st.info("💡 請確認您的 Secrets 中的密碼是否為 16 位元的「應用程式密碼」。")

        # 3. 提供下載按鈕 (作為備援)
        st.download_button("📥 下載 PDF 報表", data=pdf_out, file_name="Report.pdf", mime="application/pdf")
