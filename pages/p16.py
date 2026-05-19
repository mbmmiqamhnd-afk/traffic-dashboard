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

# 🌟 強化設定：強制指定 Tesseract OCR 路徑 (解決模組找不到的問題)
try:
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
except:
    pass

st.set_page_config(page_title="勤務督導報告自動生成系統", page_icon="🚓", layout="wide")

if "unit_reports" not in st.session_state: st.session_state.unit_reports = {}

# ==========================================
# 核心邏輯 (為節省空間，與前述一致)
# ... [此處放入您原本的 send_gmail, extract_duty_v2, extract_equip_v2 函式] ...
# ==========================================

# ==========================================
# 🌟 4.5 新增：PDF 刑案呈報單解析功能 (極致穩定版)
# ==========================================
def parse_police_report(pdf_file, roster_names):
    extracted_data = []
    try:
        pdf_file.seek(0)
        # 設定 dpi=100 以大幅降低記憶體需求
        images = convert_from_bytes(pdf_file.read(), dpi=100)
        
        default_roster = ['薛德祥', '蕭漢祥', '董德亨', '蔡震東', '廖佩祺', '王清正', '顏利玲', '洪祥浩', '董亦文', '何昀融']
        active_roster = list(set(roster_names + default_roster))
        
        for img in images:
            # 使用更快的模式
            text = pytesseract.image_to_string(img, lang='chi_tra', config='--psm 6')
            clean_text = re.sub(r'[\s\|｜「」_—\-:：,，。、"”’‘\(\)]', '', text)
            
            # 抓資料
            time_match = re.search(r'(\d{2,3}年\d{1,2}月\d{1,2}日\d{1,2}時\d{1,2}分)', clean_text)
            time_str = time_match.group(1) if time_match else "未解析"
            
            law_match = re.search(r'觸犯法條(.*?)(?:違反|附送|案件)', clean_text)
            law_str = law_match.group(1)[:15] if law_match else "未解析"
            
            officers = [name for name in active_roster if name in clean_text]
            officer_str = "、".join(list(set(officers))) if officers else "未解析"

            if time_str != "未解析":
                extracted_data.append({
                    "查獲時間": time_str,
                    "觸犯法條": law_str,
                    "查獲員警": officer_str
                })
        return extracted_data
    except Exception as e:
        st.error(f"❌ 系統錯誤: {e}")
        return []

# ==========================================
# 5. 主介面 UI (保留原本邏輯)
# ==========================================
# ... [此處放入您原本的 UI 迴圈邏輯] ...
