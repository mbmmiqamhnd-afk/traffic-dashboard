import streamlit as st
import pandas as pd
import io
import re
import smtplib
import pytesseract
import cv2
import numpy as np
from pdf2image import convert_from_bytes
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime, timedelta
import os

# 🌟 環境路徑校正 (解決雲端路徑錯誤)
if os.path.exists('/usr/bin/tesseract'):
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

st.set_page_config(page_title="勤務督導報告系統", layout="wide")

if "unit_reports" not in st.session_state: st.session_state.unit_reports = {}

# ==========================================
# 1. 寄信與輔助函式
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

# ==========================================
# 2. 🌟 終極穩定版 OCR 解析函式
# ==========================================
def parse_police_report(pdf_file, roster_names):
    extracted_data = []
    try:
        pdf_file.seek(0)
        # 讀取並進行影像二值化處理 (去除表格線條干擾)
        images = convert_from_bytes(pdf_file.read(), dpi=200)
        
        # 核心名單保底
        active_roster = list(set(roster_names + ['薛德祥', '蕭漢祥', '董德亨', '蔡震東', '廖佩祺', '王清正', '顏利玲', '洪祥浩', '董亦文', '何昀融']))
        
        for i, img in enumerate(images):
            # OpenCV 預處理：去噪、轉灰階、二值化
            open_cv_image = np.array(img.convert('RGB'))
            gray = cv2.cvtColor(open_cv_image, cv2.COLOR_RGB2GRAY)
            _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # OCR 辨識 (使用 --psm 6 針對表格區塊最佳化)
            text = pytesseract.image_to_string(thresh, lang='chi_tra', config='--psm 6')
            clean_text = re.sub(r'[\s\|｜「」_—\-:：,，。、"”’‘\(\)]', '', text)
            
            # 抓資料
            time_match = re.search(r'(\d{2,3}年\d{1,2}月\d{1,2}日\d{1,2}時\d{1,2}分)', clean_text)
            loc_match = re.search(r'查獲地點(.*?)(?:觸犯法條|案類)', clean_text)
            suspect_match = re.search(r'嫌疑人([\u4e00-\u9fa5]{2,3})', clean_text)
            
            # 模糊比對名單 (對抗 OCR 錯字)
            found_officers = set()
            for name in active_roster:
                if name in clean_text or any(part in clean_text for part in [name[1:], name[:2]]):
                    found_officers.add(name)
            
            extracted_data.append({
                "查獲時間": time_match.group(1) if time_match else "未解析",
                "查獲地點": loc_match.group(1)[:15] if loc_match else "未解析",
                "嫌疑人": suspect_match.group(1) if suspect_match else "未解析",
                "查獲員警": "、".join(list(found_officers)) if found_officers else "名單校正中"
            })
    except Exception as e:
        st.error(f"解析發生錯誤: {e}")
    return extracted_data

# ==========================================
# 3. 主 UI 介面 (整合解析結果)
# ==========================================
st.header("📋 勤務督導報告自動生成系統")
# ... (保留您原本的 extract_duty_v2 和 extract_equip_v2 以及 UI 邏輯) ...
# 注意：在 PDF 上傳邏輯中呼叫時：
if u_pdf:
    with st.spinner("正在進行 AI 影像降噪與名單校正..."):
        merit_lines = []
        for pdf_file in u_pdf:
            cases = parse_police_report(pdf_file, dr.get('roster', []))
            for case in cases:
                merit_text = f"優劣蹟紀錄：{dr['term']}同仁 {case['查獲員警']} 勤務落實，於 {case['查獲時間']} 在「{case['查獲地點']}」查獲嫌疑人 {case['嫌疑人']}，表現優良，建議列優蹟註記。"
                merit_lines.append(merit_text)
        if merit_lines: lns.extend(merit_lines)
