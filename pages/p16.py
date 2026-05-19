import streamlit as st
import pandas as pd
import io
import re
import os
import smtplib
import pytesseract
from pdf2image import convert_from_bytes
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime, timedelta

# 🌟 自動偵測環境路徑 (解決雲端路徑錯誤)
if os.path.exists('/usr/bin/tesseract'):
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

st.set_page_config(page_title="勤務督導報告自動生成系統", page_icon="🚓", layout="wide")

if "unit_reports" not in st.session_state: st.session_state.unit_reports = {}

# ==========================================
# 核心函式與解析邏輯
# ==========================================
def send_gmail(subject, body, receiver_email):
    try:
        sender_email = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        msg = MIMEText(body, 'plain', 'utf-8')
        msg['Subject'] = Header(subject, 'utf-8')
        msg['From'] = f"督導助手 <{sender_email}>"
        msg['To'] = receiver_email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        return True
    except Exception as e:
        st.error(f"寄信失敗：{e}")
        return False

# 🌟 OCR 智能解析函式 (針對掃描檔優化)
def parse_police_report(pdf_file, roster_names):
    extracted_data = []
    try:
        pdf_file.seek(0)
        # dpi=100 及 grayscale=True 降低記憶體負荷，防止頁面空白
        images = convert_from_bytes(pdf_file.read(), dpi=100, grayscale=True)
        
        # 內建中興所核心名單作為保底字庫
        active_roster = list(set(roster_names + ['薛德祥', '蕭漢祥', '董德亨', '蔡震東', '廖佩祺', '王清正', '顏利玲', '洪祥浩']))
        
        for i, img in enumerate(images):
            text = pytesseract.image_to_string(img, lang='chi_tra', config='--psm 6')
            clean_text = re.sub(r'[\s\|｜「」_—\-:：,，。、"”’‘\(\)]', '', text)
            
            time_match = re.search(r'(\d{2,3}年\d{1,2}月\d{1,2}日\d{1,2}時\d{1,2}分)', clean_text)
            
            # 勤務人員比對 (模糊匹配)
            officers = [n for n in active_roster if n in clean_text or any(part in clean_text for part in [n[1:], n[:2]])]
            
            if time_match:
                extracted_data.append({
                    "查獲時間": time_match.group(1),
                    "查獲員警": "、".join(list(set(officers))) if officers else "未解析"
                })
            del img
    except Exception as e:
        st.error(f"解析發生異常: {e}")
    return extracted_data

# ==========================================
# 主介面 UI
# ==========================================
st.header("📋 勤務督導報告自動生成系統")
insp_date = st.date_input("選擇督導日期", datetime.now())
num_units = st.number_input("待督導單位數量", 1, 8, 3)
u_tabs = st.tabs([f"🏢 單位 {i+1}" for i in range(num_units)] + ["📄 總匯整報告"])

for i in range(num_units):
    with u_tabs[i]:
        col1, col2, col3 = st.columns(3)
        u_duty = col1.file_uploader(f"單位 {i+1} 勤務表", type=['xlsx'], key=f"ud_{i}")
        u_eq = col2.file_uploader(f"單位 {i+1} 交接簿", type=['xlsx'], key=f"ue_{i}")
        u_pdf = col3.file_uploader(f"刑案呈報單(PDF)", type=['pdf'], accept_multiple_files=True, key=f"updf_{i}")
        
        if u_duty and u_eq:
            # 這裡簡化呼叫，您可以將您原本的 extract_duty_v2 邏輯放回這裡
            # 核心重點：解析 PDF
            if u_pdf:
                with st.spinner("正在進行 OCR 智能掃描..."):
                    all_cases = []
                    for pdf in u_pdf:
                        all_cases.extend(parse_police_report(pdf, []))
                    
                    for case in all_cases:
                        st.write(f"✅ 成功辨識：{case['查獲時間']} - 員警：{case['查獲員警']}")
                        # 在此處將解析結果串接進您的報告清單
