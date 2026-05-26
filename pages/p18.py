import streamlit as st
import pandas as pd
import io
import sys
import os
import re
import smtplib
import numpy as np
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# 自動將上層目錄加入路徑
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
try:
    from app import show_sidebar
except ImportError:
    def show_sidebar():
        pass 

def send_report_email_auto(files, year, month):
    try:
        if "email" not in st.secrets:
            return False, "找不到 st.secrets 中的 email 設定"
            
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = sender
        msg['Subject'] = f"【系統備份】龍潭分局 {year}年{month}月 獎勵金點數統計表暨印領清冊"
        
        body = f"郭同仁您好：\n\n系統已自動完成 {year}年{month}月份的獎勵金點數彙整與印領清冊產出。\n本次附件包含「點數統計表」與「印領清冊」共兩份 Excel 檔案，請查收。"
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        for file_data, filename in files:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file_data)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(filename)}")
            msg.attach(part)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

def p18_page():
    show_sidebar()

    st.title("💰 龍潭分局 - 獎勵金點數統計表暨印領清冊產生器")
    st.info("權重已固定 (A2:10, A3:5, 交整:5)。系統支援【管考72% / 督導20% / 其他8%】自動拆分算錢！")

    P_A2, P_A3, P_TRAF = 10.0, 5.0, 5.0

    # 1. 檔案上傳區
    st.subheader("📂 1. 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)", type=['xlsx'], accept_multiple_files=True)
    
    # 2. 印領清冊參數與名單設定
    st.subheader("📝 2. 印領清冊與獎金分配設定")
    point_value = st.number_input("💵 直接執行人員 - 每點獎金金額", value=1.905, format="%.3f", step=0.001)

    st.markdown("##### 👥 共同作業及配合人員 - 分配模式")
    alloc_mode = st.radio(
        "請選擇「共同作業及配合人員」的獎金計算方式：",
        ["🤖 系統自動依比例分配 (管考72%、督導20%、其他8%)", "✍️ 手動輸入固定金額 (僅於總表分類顯示)"]
    )
    
    # 根據選擇模式切換顯示與設定
    if "系統自動" in alloc_mode:
        st.info("💡 系統會依您設定的總預算切成三塊 (72/20/8)，再按下方名單的「分配基數」全自動精算金額，解決四捨五入零頭問題。")
        budget_type = st.selectbox("請選擇預算輸入方式：", [
            "A. 直接輸入【共同作業人員】的總分配預算", 
            "B. 輸入【全分局】本月核撥總預算 (系統會自動先扣掉直接執行人員的總獎金)"
        ])
        
        if "A" in budget_type:
            budget_input = st.number_input("💰 輸入【共同作業人員】總預算 (元)", value=10000, step=100)
        else:
            budget_input = st.number_input("💰 輸入【全分局】核撥總預算 (元)", value=50000, step=100)
            
        col_name_display = "分配基數(權重)"
    else:
        st.info("💡 系統將直接使用您在下方表格填寫的實際金額進行發放。")
        budget_input = 0
        budget_type = ""
        col_name_display = "設定金額"

    st.markdown(f"**共同作業名單 (已根據「僅會計/秘書/人事屬8%」規則更新預設值)**")
    
    # 完整 66 人名單，依照 72% -> 8% -> 20% 重新整理排列
    default_coworkers_data = [
        # --- 負責管考 (72%) ---
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "業務單位主管", "姓名": "陳維明", "設定數值": 298, "蓋章": ""},
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "盧冠仁", "設定數值": 298, "蓋章": ""},
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "李峯甫", "設定數值": 298, "蓋章": ""},
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "羅千金", "設定數值": 298, "蓋章": ""},
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "郭勝隆", "設定數值": 298, "蓋章": ""},
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "吳享運", "設定數值": 232, "蓋章": ""},
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "吳沛軒", "設定數值": 232, "蓋章": ""},
        
        # --- 其他配合 (8%)：只有會計室、秘書室、人事室 ---
        {"分配類別": "其他配合(8%)", "單位": "會計室", "職別": "主計", "姓名": "郭貞彣", "設定數值": 77, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "會計室", "職別": "主計", "姓名": "林玲宜", "設定數值": 78, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "秘書室", "職別": "主任", "姓名": "陳振貴", "設定數值": 78, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "秘書室", "職別": "出納", "姓名": "簡啟峯", "設定數值": 78, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "秘書室", "職別": "巡官", "姓名": "陳鵬翔", "設定數值": 64, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "主任", "姓名": "葉菀容", "設定數值": 78, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "助理員", "姓名": "王韋翔", "設定數值": 77, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "警務佐", "姓名": "李福源", "設定數值": 77, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "警員", "姓名": "陳明祥", "設定數值": 77, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "警員", "姓名": "黃秀吉", "設定數值": 77, "蓋章": ""},

        # --- 勤務督導 (20%)：其餘單位 (含分局長、派出所、各組隊) ---
        {"分配類別": "勤務督導(20%)", "單位": "龍潭分局", "職別": "分局長", "姓名": "施宇峰", "設定數值": 301, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭分局", "職別": "副分局長", "姓名": "何憶雯", "設定數值": 100, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭分局", "職別": "副分局長", "姓名": "蔡志明", "設定數值": 100, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "主任", "姓名": "游新枝", "設定數值": 65, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "巡佐", "姓名": "李文章", "設定數值": 65, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "巡佐", "姓名": "余清富", "設定數值": 65, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "警務佐", "姓名": "陳敬霖", "設定數值": 65, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "警員", "姓名": "黃文興", "設定數值": 65, "蓋章": ""},
        {"分配類別": "勤務
