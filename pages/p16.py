import streamlit as st
import pandas as pd
import io
import re
import smtplib
import pytesseract
from pdf2image import convert_from_bytes
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime, timedelta
import os

# 🌟 極致穩定：自動尋找雲端環境的 Tesseract 路徑
if os.path.exists('/usr/bin/tesseract'):
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

st.set_page_config(page_title="勤務督導報告系統", layout="wide")

# ... (請保留您原有的 extract_duty_v2 和 extract_equip_v2 函式) ...

# 🌟 輕量版解析函式 (降低記憶體消耗，防止頁面空白)
def parse_police_report(pdf_file, roster_names):
    extracted_data = []
    try:
        pdf_file.seek(0)
        # 設定 dpi=100，避免過大圖片導致當機
        images = convert_from_bytes(pdf_file.read(), dpi=100)
        
        for img in images:
            text = pytesseract.image_to_string(img, lang='chi_tra')
            if not text.strip(): continue
            
            # 簡易擷取
            time_m = re.search(r'(\d{2,3}年\d{1,2}月\d{1,2}日)', text)
            officers = [n for n in roster_names if n in text]
            
            extracted_data.append({
                "查獲時間": time_m.group(1) if time_m else "日期未解析",
                "查獲員警": "、".join(officers) if officers else "未解析"
            })
            del img # 強制釋放記憶體
    except Exception as e:
        st.error(f"解析錯誤: {e}")
    return extracted_data

# ... (請保留您原有的主介面 UI 邏輯) ...
