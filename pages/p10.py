import streamlit as st
# 1. 頁面設定必須在最前面
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

from menu import show_sidebar
show_sidebar() 

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import smtplib
import io
import os
import urllib.parse as _ul
import re
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

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT_TITLE = "桃園市政府警察局龍潭分局"

CHECKIN_POINTS = """1. 中油高原交流道站（龍源路2-20號）
2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）
3. 7-11龍潭佳園門市（中正路三坑段776號）
4. 旭日路三坑自然生態公園停車場
5. 旭日路與大溪區交界處"""

NOTES = """一、各編組執行前由帶班人員在駐地實施勤前教育。
二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。
三、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。
四、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"""

# --- 2. 工具函式 ---
def format_staff_only(val):
    if pd.isna(val) or str(val).strip() in ["None", "nan", ""]: return ""
    s = str(val).replace('\\', '\n').replace('、', '\n').replace('\xa0', ' ')
    s = re.sub(r'(\d{2}[:：]?\d{0,2}\s*-\s*\d{2}[:：]?\d{0,2}[時]?[:：])\s*([^\n\s])', r'\1\n\2', s)
    return '\n'.join([l.strip() for l in s.split('\n') if l.strip()])

@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    try:
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)
    except: return None

def load_from_cloud():
    client = get_client()
    if not client: return None, None, None
    try:
        sh = client.open_by_key(SHEET_ID)
        s = pd.DataFrame(sh.worksheet("危駕_設定").get_all_records())
        c = pd.DataFrame(sh.worksheet("危駕_指揮組").get_all_records())
        p = pd.DataFrame(sh.worksheet("危駕_警力佈署").get_all_records())
        return s, c, p
    except: return None, None, None

def save_to_cloud(p_time, cmdr, df_c, df_p):
    client = get_client()
    if not client: return False
    try:
        sh = client.open_by_key(SHEET_ID)
        sh.worksheet("危駕_設定").clear()
        sh.worksheet("危駕_設定").update(range_name='A1', values=[["Key", "Value"], ["plan_time", p_time], ["commander", cmdr]])
        
        for name, df in [("危駕_指揮組", df_c), ("危駕_警力佈署", df_p)]:
            ws = sh.worksheet(name)
            ws.clear()
            data = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
            ws.update(range_name='A1', values=data)
        return True
    except: return False

# --- 3. 初始化 Session State ---
if 'data_ptl' not in st.session_state:
    s, c, p = load_from_cloud()
    if s is not None:
        sd = dict(zip(s.iloc[:,0].astype(str), s.iloc[:,1].astype(str)))
        st.session_state.p_time = sd.get("plan_time", "115年4月30日22時至翌日6時")
        st.session_state.cmdr = sd.get("commander", "石門所副所長林榮裕")
        st.session_state.data_cmd = c
        st.session_state.data_ptl = p
    else:
        st.session_state.p_time = "115年4月30日22時至翌日6時"
        st.session_state.cmdr = "石門所副所長林榮裕"
        st.session_state.data_cmd = pd.DataFrame(columns=["職稱", "代號", "姓名", "任務"])
        st.session_state.data_ptl = pd.DataFrame(columns=["勤務時段", "代號", "編組", "服勤人員", "任務分工"])

# --- 4. 介面 ---
st.title("🚔 防制危險駕車專案勤務規劃表")

col1, col2 = st.columns(2)
with col1:
    p_time = st.text_input("1. 勤務時間", st.session_state.p_time)
with col2:
    cmdr_input = st.text_input("2. 交通快打指揮官", st.session_state.cmdr)

# 核心邏輯：計算日期
date_match = re.search(r'(?:(\d+)年)?(\d+)月(\d+)日(.*)', p_time)
dedicated_time = ""
normal_time = ""
if date_match:
    y, m, d, part = date_match.group(1), int(date_match.group(2)), int(date_match.group(3)), date_match.group(4).strip()
    y_tw = int(y) if y else (datetime.now().year - 1911)
    base_dt = datetime(y_tw + 1911, m, d)
    next_dt = base_dt + timedelta(days=1)
    dedicated_time = f"{next_dt.month}月{next_dt.day}日\n零時至4時"
    normal_time = f"{m}月{d}日\n{part}"

# 顯示任務編組
st.subheader("3. 任務編組")
res_cmd = st.data_editor(st.session_state.data_cmd, num_rows="dynamic", use_container_width=True).fillna("")

# 顯示警力佈署
st.subheader("4. 警力佈署")
# 自動處理第一列預填
if len(st.session_state.data_ptl) == 0:
    st.session_state.data_ptl = pd.DataFrame([["", "", "", "", ""]], columns=["勤務時段", "代號", "編組", "服勤人員", "任務分工"])

# 在編輯器渲染前，針對「完全空白」的新表單做預填
if str(st.session_state.data_ptl.at[0, '勤務時段']).strip() in ["", "nan"]:
    st.session_state.data_ptl.at[0, '勤務時段'] = dedicated_time
    # 自動帶入代號
    unit_base = "隆安8" if "石門" in cmdr_input else "隆安6" if "龍潭" in cmdr_input else "隆安"
    st.session_state.data_ptl.at[0, '代號'] = unit_base + ("1" if "所長" in cmdr_input and "副" not in cmdr_input else "2")
    st.session_state.data_ptl.at[0, '編組'] = f"專責警力\n（{cmdr_input[:3]}輪值）"

res_ptl = st.data_editor(st.session_state.data_ptl, num_rows="dynamic", use_container_width=True).fillna("")

# --- 後處理：修正新增列的日期 ---
for i in range(len(res_ptl)):
    cur_t = str(res_ptl.at[i, '勤務時段']).strip()
    cur_g = str(res_ptl.at[i, '編組']).strip()
    # 如果是新增列（值為空，或者是從上一列自動複製過來的錯誤跨日時間）
    if i > 0 and ("專責" not in cur_g):
        if cur_t in ["", "nan", "None"] or "零時" in cur_t:
            res_ptl.at[i, '勤務時段'] = normal_time

# 更新到 session_state 以便下次 rerun 保持
st.session_state.data_cmd = res_cmd
st.session_state.data_ptl = res_ptl

# --- 5. 預覽與輸出 (略，維持原本 PDF 與 HTML 邏輯) ---
# ... (此處保留你原本的 get_preview 和 generate_pdf 函式) ...

def get_preview_html(df_c, df_p, cmdr_n, time_s):
    # 簡單預覽邏輯
    return f"<h3>預覽：{time_s}</h3><p>指揮官：{cmdr_n}</p>" # 建議接回你原本的 HTML 渲染代碼

st.markdown("---")
if st.button("💾 同步至雲端並寄送郵件", type="primary"):
    with st.spinner("同步中..."):
        if save_to_cloud(p_time, cmdr_input, res_cmd, res_ptl):
            st.success("✅ 雲端同步成功！")
            # 這裡可以接寄信代碼 send_report_email(...)
        else:
            st.error("❌ 同步失敗。請檢查 Google Sheets 權限或 Secrets 設定。")
