import streamlit as st
import pandas as pd
import io
import re
import traceback
import smtplib
import pytesseract
from pdf2image import convert_from_bytes
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime, timedelta

# ==========================================
# 0. 系統初始化與狀態管理
# ==========================================
st.set_page_config(page_title="勤務督導報告自動生成系統", page_icon="🚓", layout="wide")

if "unit_reports" not in st.session_state:
    st.session_state.unit_reports = {}

try:
    from menu import show_sidebar
    show_sidebar()
except:
    pass

st.markdown("""
    <style>
    @font-face { font-family: 'Kaiu'; src: url('kaiu.ttf'); }
    .stTextArea textarea {
        font-family: 'Kaiu', "標楷體", sans-serif !important;
        font-size: 19px !important;
        line-height: 1.7 !important;
        color: #1c1c1c !important;
    }
    .stTabs [data-baseweb="tab-list"] button { font-size: 18px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 1. 寄信功能
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
# 2. 解析工具函式 (省略部分重複邏輯以簡化)
# ==========================================
def safe_int(val):
    try: return int(float(str(val).split('.')[0].replace(',', '')))
    except: return 0

def build_fmap(df):
    fmap = {}
    for r in range(len(df)):
        if '代號' in str(df.iloc[r, 0]) and '職稱' in str(df.iloc[r, 0]):
            for rr in range(r, min(r + 20, len(df))):
                c = 1
                while c < len(df.columns) - 1:
                    code = str(df.iloc[rr, c]).strip()
                    name = str(df.iloc[rr, c + 1]).strip()
                    if code and name and re.match(r'^[A-Za-z0-9甲乙丙丁]{1,3}$', code):
                        fmap[code] = name
                    c += 6
    return fmap

# ==========================================
# 3. 刑案 OCR 智能解析 (交叉比對版)
# ==========================================
def parse_police_report(pdf_file, roster_names):
    extracted_data = []
    try:
        pdf_file.seek(0)
        images = convert_from_bytes(pdf_file.read())
        
        # 建立保底字庫 (當 Excel 沒抓到時，這是最後防線)
        default_roster = ['薛德祥', '蕭漢祥', '董德亨', '蔡震東', '廖佩祺', '王清正', '顏利玲', '洪祥浩', '董亦文', '何昀融']
        active_roster = list(set(roster_names + default_roster))
        
        for i, img in enumerate(images):
            text = pytesseract.image_to_string(img, lang='chi_tra')
            clean_text = re.sub(r'[\s\|｜「」_—\-:：,，。、"”’‘\(\)]', '', text)
            
            # 抓取關鍵資訊
            time_match = re.search(r'(\d{2,3}年\d{1,2}月\d{1,2}日\d{1,2}時\d{1,2}分)', clean_text)
            suspect_match = re.search(r'嫌疑人([\u4e00-\u9fa5]{2,3})', clean_text)
            loc_match = re.search(r'查獲地點(.*?)(?:觸犯法條|案類)', clean_text)
            law_match = re.search(r'觸犯法條(.*?)(?:違反|達反|連反|附送|案件)', clean_text)
            
            # 名單匹配
            officers = set()
            for r_name in active_roster:
                if len(r_name) >= 2 and r_name in clean_text:
                    officers.add(r_name)
            
            extracted_data.append({
                "查獲時間": time_match.group(1) if time_match else "未解析",
                "嫌疑人": suspect_match.group(1) if suspect_match else "未詳",
                "查獲地點": loc_match.group(1)[:15] if loc_match else "未詳",
                "觸犯法條": law_match.group(1)[:15] if law_match else "未解析",
                "查獲員警": "、".join(officers) if officers else "未解析"
            })
    except Exception as e:
        st.error(f"解析失敗: {e}")
    return extracted_data

# ==========================================
# 4. 主介面 (省略重複邏輯，與您原版邏輯一致)
# ==========================================
# ... (此處保留原本 extract_duty_v2, extract_equip_v2 與 UI 邏輯) ...
# 注意：在 UI 中處理 PDF 的那段，請確保將 parse_police_report 呼叫改為：
# cases = parse_police_report(pdf_file, dr.get('roster', []))
