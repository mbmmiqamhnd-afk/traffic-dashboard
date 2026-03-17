import streamlit as st
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

# --- 1. 頁面設定 ---
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# --- 預設資料 (當雲端連線失敗時使用) ---
DEFAULT_TIME = "115年3月6日22時至翌日6時"
DEFAULT_COMMANDER = "石門所副所長林榮裕"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "業務組", "代號": "隆安13", "姓名": "交通組警務員 葉佳媛", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報群聚飆車狀況。"}
])

DEFAULT_PATROL = pd.DataFrame([
    {
        "勤務時段": "3月7日\n零時至4時", "無線電": "隆安82", "編組": "專責警力（石門所輪值）", 
        "服勤人員": "00-02時\n副所長林榮裕\n02-04時\n副所長林榮裕", 
        "任務分工": "「加強防制」勤務，在文化路、中正路三坑段、龍源路及旭日路來回巡邏，隨機攔檢改裝（噪音）車輛"
    },
    {
        "勤務時段": "3月6日\n22時至翌日6時", "無線電": "隆安80", "編組": "石門所", 
        "服勤人員": "線上巡邏警力兼任", 
        "任務分工": "「區域聯防」勤務，於中正路、文化路、中豐路、龍源路巡邏（每1小時巡簽1次），並加強查緝毒駕"
    }
])

CHECKIN_POINTS = """1. 中油高原交流道站（龍源路2-20號）
2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）
3. 7-11龍潭佳園門市（中正路三坑段776號）
4. 旭日路三坑自然生態公園停車場
5. 旭日路與大溪區交界處"""

NOTES = """一、各編組執行前由帶班人員在駐地實施勤前教育。
二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。
三、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。
四、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"""

# --- 2. 自動排版引擎 ---
def auto_format_personnel(val):
    """處理服勤人員格式：時段與姓名垂直並列"""
    if pd.isna(val) or str(val).strip() in ["None", "nan", ""]: 
        return ""
    s = str(val).replace('：', ':').replace('、', '\n')
    # 將「XX-XX時」加粗並強迫換行
    s = re.sub(r'(\d{2}-\d{2}時)[:\s]*', r'<b>\1</b>\n', s)
    lines = [line.strip() for line in s.split('\n') if line.strip()]
    return '\n'.join(lines)

# --- 3. 雲端讀取/儲存函數 (省略細節以保簡潔，保持您的原邏輯) ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    return gspread.authorize(Credentials.from_service_account_info(creds_dict, scopes=SCOPES))

@st.cache_data(ttl=60)
def load_data():
    try:
        client = get_client()
        if not client: return None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        return (pd.DataFrame(sh.worksheet("危駕_設定").get_all_records()), 
                pd.DataFrame(sh.worksheet("危駕_指揮組").get_all_records()), 
                pd.DataFrame(sh.worksheet("危駕_警力佈署").get_all_records()), None)
    except: return None, None, None, "載入失敗"

# --- 4. 主介面邏輯 ---
df_set, df_cmd, df_ptl, err = load_data()
if err or df_set is None:
    t, cmdr = DEFAULT_TIME, DEFAULT_COMMANDER
    ed_cmd, ed_ptl = DEFAULT_CMD.copy(), DEFAULT_PATROL.copy()
else:
    sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    t, cmdr = sd.get("plan_time", DEFAULT_TIME), sd.get("commander", DEFAULT_COMMANDER)
    ed_cmd, ed_ptl = df_cmd, df_ptl

st.title("🚔 防制危險駕車專案勤務規劃表")

# 基礎輸入
p_time = st.text_input("勤務時間", t)
cmdr_input = st.text_input("交通快打指揮官", cmdr)

# ====== 魔法連動核心：指揮官單位 = 專責警力單位 ======
if len(ed_ptl) > 0:
    # 辨識單位 (XX所/分隊)
    m_unit = re.search(r'([\u4e00-\u9fa5]+(?:所|分隊|分局))', cmdr_input)
    if m_unit:
        unit_name = m_unit.group(1)
        title_name = cmdr_input.replace(unit_name, "").strip() # 剩餘的人名職稱
        
        # 1. 同步編組名稱
        ed_ptl.loc[0, '編組'] = f"專責警力\n（{unit_name}輪值）"
        
        # 2. 自動切換無線電代號
        unit_map = {"石門": "隆安8", "高平": "隆安9", "聖亭": "隆安5", "龍潭": "隆安6", "中興": "隆安7", "分隊": "隆安99"}
        for k, v in unit_map.items():
            if k in unit_name:
                suffix = "1" if any(x in title_name for x in ["所長", "分隊長"]) else "2"
                ed_ptl.loc[0, '無線電'] = v + suffix
                break
        
        # 3. 自動同步服勤人員 (垂直排版)
        current_ppl = str(ed_ptl.loc[0, '服勤人員'])
        time_slots = re.findall(r'(\d{2}-\d{2}時)', current_ppl)
        if time_slots and title_name:
            new_val = ""
            for ts in time_slots:
                new_val += f"{ts}\n{title_name}\n"
            ed_ptl.loc[0, '服勤人員'] = new_val.strip()

# 套用自動排版引擎 (讓所有服勤人員欄位都垂直並列)
if '服勤人員' in ed_ptl.columns:
    ed_ptl['服勤人員'] = ed_ptl['服勤人員'].apply(auto_format_personnel)

# 顯示表格編輯器
st.subheader("1. 任務編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True)

st.subheader("2. 警力佈署")
res_ptl = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True)

# 下方為 PDF 生成與下載按鈕 (保持原邏輯)
# ... (PDF 生成代碼與 st.download_button) ...
