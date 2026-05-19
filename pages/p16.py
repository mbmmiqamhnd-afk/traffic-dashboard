import streamlit as st
import pandas as pd
import io
import re
import os
import smtplib
import pytesseract
import cv2
import numpy as np
from pdf2image import convert_from_bytes
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime, timedelta

# ==========================================
# 0. 環境路徑設定
# ==========================================
if os.path.exists('/usr/bin/tesseract'):
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

st.set_page_config(page_title="勤務督導報告系統", layout="wide")

# ==========================================
# 1. 核心解析功能 (請保留您原有的函式)
# ==========================================
# (在此處貼上您原本的 extract_duty_v2 和 extract_equip_v2 函式)
# 確保這些函式已定義在 UI 呼叫之前

def parse_police_report(pdf_file, roster_names):
    """結合 OCR 全域掃描與名單模糊比對的穩定解析器"""
    extracted_data = []
    try:
        pdf_file.seek(0)
        # 降頻解析以防當機
        images = convert_from_bytes(pdf_file.read(), dpi=150)
        
        # 內建保底名單
        active_roster = list(set(roster_names + ['薛德祥', '蕭漢祥', '董德亨', '蔡震東', '廖佩祺', '王清正', '顏利玲', '洪祥浩', '董亦文', '何昀融']))
        
        for img in images:
            # 影像預處理：轉灰階 + 二值化 (去除表格干擾)
            gray = cv2.cvtColor(np.array(img.convert('RGB')), cv2.COLOR_RGB2GRAY)
            _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            text = pytesseract.image_to_string(thresh, lang='chi_tra', config='--psm 6')
            clean_text = re.sub(r'[\s\|｜「」_—\-:：,，。、"”’‘\(\)]', '', text)
            
            # 抓資料
            time_m = re.search(r'(\d{2,3}年\d{1,2}月\d{1,2}日\d{1,2}時\d{1,2}分)', clean_text)
            suspect_m = re.search(r'嫌疑人([\u4e00-\u9fa5]{2,3})', clean_text)
            loc_m = re.search(r'查獲地點(.*?)(?:觸犯法條|案類)', clean_text)
            
            # 員警校正邏輯
            officers = [name for name in active_roster if name in clean_text or any(part in clean_text for part in [name[1:], name[:2]])]
            
            if time_m:
                extracted_data.append({
                    "時間": time_m.group(1),
                    "地點": loc_m.group(1)[:15] if loc_m else "未詳",
                    "嫌疑人": suspect_m.group(1) if suspect_m else "未詳",
                    "員警": "、".join(list(set(officers))) if officers else "名單校正中"
                })
    except Exception as e:
        st.error(f"解析錯誤: {e}")
    return extracted_data

# ==========================================
# 2. 主介面 UI
# ==========================================
st.header("📋 勤務督導報告自動生成系統")
insp_date = st.date_input("選擇督導日期", datetime.now())
num_units = st.number_input("待督導單位數量", 1, 8, 3)
u_tabs = st.tabs([f"🏢 單位 {i+1}" for i in range(num_units)] + ["📄 總匯整報告"])

for i in range(num_units):
    with u_tabs[i]:
        u_time = st.time_input("抵達時間", datetime.now().time(), key=f"ut_{i}")
        col1, col2, col3 = st.columns(3)
        u_duty = col1.file_uploader(f"勤務表", type=['xlsx'], key=f"ud_{i}")
        u_eq = col2.file_uploader(f"交接簿", type=['xlsx'], key=f"ue_{i}")
        u_pdf = col3.file_uploader(f"刑案單(PDF)", type=['pdf'], accept_multiple_files=True, key=f"updf_{i}")
        
        if u_duty and u_eq:
            dr = extract_duty_v2(u_duty, u_time.hour)
            er = extract_equip_v2(u_eq)
            
            # 基本報告文本
            lns = [f"{u_time.strftime('%H%M')}，{dr['term']}值班{dr['v_name']}，裝備檢查合格。"]
            
            # 解析 PDF
            if u_pdf:
                with st.spinner("系統推理中..."):
                    for pdf_file in u_pdf:
                        cases = parse_police_report(pdf_file, dr.get('roster', []))
                        for case in cases:
                            lns.append(f"優劣蹟：{dr['term']}同仁 {case['員警']} 於 {case['時間']} 在「{case['地點']}」查獲嫌疑人 {case['嫌疑人']}，建議記優蹟。")
            
            final_text = "\n".join([f"{idx+1}、{line}" for idx, line in enumerate(lns)])
            st.text_area("報告預覽", final_text, height=300, key=f"prev_{i}")
            st.session_state.unit_reports[i] = final_text
