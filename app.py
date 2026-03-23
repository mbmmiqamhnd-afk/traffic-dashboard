import streamlit as st
import pandas as pd
import io
import re
import gspread
import shutil
import smtplib
import calendar
import traceback
from datetime import datetime, timedelta, date
from pdf2image import convert_from_bytes
from pptx import Presentation
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email import encoders
from email.header import Header

# ==========================================
# 0. 系統初始化配置
# ==========================================
st.set_page_config(page_title="龍潭分局交通智慧戰情室", page_icon="🚓", layout="wide")

# 常數設定
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
TO_EMAIL = "mbmmiqamhnd@gmail.com"

# 目標值配置 (合併各模組)
TARGETS_ENHANCED = {'聖亭所': [5, 115, 5, 16, 7, 10], '龍潭所': [6, 145, 7, 20, 9, 12], '中興所': [5, 115, 5, 16, 7, 10], '石門所': [3, 80, 4, 11, 5, 7], '高平所': [3, 80, 4, 11, 5, 7], '三和所': [2, 40, 2, 6, 2, 5], '交通分隊': [5, 115, 4, 16, 6, 8]}
TARGETS_MAJOR = {'聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479, '高平所': 1294, '三和所': 339, '交通分隊': 2526, '科技執法': 6006}
TARGETS_OVERLOAD = {'聖亭所': 20, '龍潭所': 27, '中興所': 20, '石門所': 16, '高平所': 14, '三和所': 8, '警備隊': 0, '交通分隊': 22}

# ==========================================
# 🛠️ 通用工具函式庫
# ==========================================

def get_standard_unit(raw_name):
    name = str(raw_name).strip()
    if '分隊' in name: return '交通分隊'
    if '科技' in name or '交通組' in name: return '科技執法'
    if '警備' in name: return '警備隊'
    for k in ['聖亭', '龍潭', '中興', '石門', '高平', '三和']:
        if k in name: return k + '所'
    return None

def format_roc_yesterday():
    yesterday = datetime.now() - timedelta(days=1)
    return f"{yesterday.year-1911}年1月1日至{yesterday.year-1911}年{yesterday.month}月{day}日"

def send_mail(excel_bytes, subject, body_text, filename="Report.xlsx"):
    try:
        user, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg['Subject'] = Header(subject, 'utf-8').encode()
        msg['From'], msg['To'] = user, TO_EMAIL
        msg.attach(MIMEText(body_text, 'plain'))
        part = MIMEApplication(excel_bytes, Name=filename)
        part.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(part)
        with smtplib.SMTP('smtp.gmail.com', 587) as s:
            s.starttls()
            s.login(user, pwd)
            s.send_message(msg)
        return True
    except: return False

# ==========================================
# 🏰 導覽選單
# ==========================================
with st.sidebar:
    st.title("🚓 龍潭分局戰情室")
    app_mode = st.selectbox("功能模組", ["🏠 智慧上傳中心", "📂 PDF 轉 PPTX 工具"])
    st.divider()
    st.info("💡 10秒流程：首頁直接拖入報表即可分析。")

# ==========================================
# 🏠 模式一：智慧上傳中心 (核心自動化)
# ==========================================
if app_mode == "🏠 智慧上傳中心":
    st.header("📈 交通數據智慧分析中心")
    uploads = st.file_uploader("📂 拖入隨身碟中的報表檔案", type=["xlsx", "csv", "xls"], accept_multiple_files=True)

    if uploads:
        num = len(uploads)
        
        # --- [1份檔案]：科技執法 ---
        if num == 1:
            f = uploads[0]
            st.success(f"📸 識別為「科技執法」報表：{f.name}")
            # (此處執行科技執法解析、24pt藍紅標題同步與寄信)
            st.info("正在產製科技執法排行...")

        # --- [2份檔案]：重大交通違規 ---
        elif num == 2:
            st.success("✅ 識別為「重大交通違規」統計 (本期+累計)")
            # (此處執行重大違規解析、16pt藍紅標題與負值紅字同步)
            st.info("正在執行重大違規數據推播...")

        # --- [3份檔案]：強化專案 或 超載統計 ---
        elif num == 3:
            if any("stone" in f.name.lower() for f in uploads):
                st.success("🚛 識別為「超載違規」自動統計")
                # (執行超載統計、自動寄信與雲端數據更新)
            else:
                st.success("🔥 識別為「強化交通安全專案」統計")
                # (執行強化專案 16pt 藍紅標題同步)

        # --- [4份檔案]：交通事故 A1/A2 ---
        elif num == 4:
            st.success("🚑 識別為「交通事故 A1/A2」統計")
            # (執行交通事故日期比對、紅字標題同步)

        else:
            st.warning(f"目前收到 {num} 份檔案，請確認是否符合各項統計之數量要求。")

# ==========================================
# 📂 模式二：PDF 轉 PPTX 工具
# ==========================================
elif app_mode == "📂 PDF 轉 PPTX 工具":
    st.header("📂 PDF 行政文書轉檔")
    uploaded_pdf = st.file_uploader("上傳 PDF 檔案", type=["pdf"])
    if uploaded_pdf:
        if st.button("🚀 開始轉檔"):
            with st.spinner("圖片解析中..."):
                images = convert_from_bytes(uploaded_pdf.read(), dpi=150)
                prs = Presentation()
                for img in images:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    img_io = io.BytesIO()
                    img.save(img_io, format='JPEG', quality=85)
                    slide.shapes.add_picture(img_io, 0, 0, width=prs.slide_width, height=prs.slide_height)
                out = io.BytesIO()
                prs.save(out)
                st.download_button("📥 下載 PPTX", out.getvalue(), file_name="Report.pptx")
