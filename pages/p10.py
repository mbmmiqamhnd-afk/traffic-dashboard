import streamlit as st
# 1. 頁面設定必須在最前面
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

from menu import show_sidebar
show_sidebar()

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import io
import os
import re
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

# --- 工具函式 ---
def normalize(s):
    return str(s).replace('\n', '').replace('\r', '').replace(' ', '').strip()

def is_blank(val):
    return normalize(val) in ["", "None", "nan"]

def is_from_first_row(val):
    """判定是否含『零時』，代表可能是從第一列複製過來的錯誤值"""
    return '零時' in str(val)

def format_staff_only(val):
    if pd.isna(val) or str(val).strip() in ["None", "nan", ""]:
        return ""
    s = str(val).replace('\\', '\n').replace('、', '\n').replace('\xa0', ' ')
    s = re.sub(r'(\d{2}[:：]?\d{0,2}\s*-\s*\d{2}[:：]?\d{0,2}[時]?[:：])\s*([^\n\s])', r'\1\n\2', s)
    return '\n'.join([l.strip() for l in s.split('\n') if l.strip()])

@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)
    except:
        return None

def load_from_cloud():
    client = get_client()
    if not client:
        return None, None, None
    try:
        sh = client.open_by_key(SHEET_ID)
        s = pd.DataFrame(sh.worksheet("危駕_設定").get_all_records())
        c = pd.DataFrame(sh.worksheet("危駕_指揮組").get_all_records())
        p = pd.DataFrame(sh.worksheet("危駕_警力佈署").get_all_records())
        return s, c, p
    except:
        return None, None, None

def save_to_cloud(p_time, cmdr, df_c, df_p):
    client = get_client()
    if not client:
        return False
    try:
        sh = client.open_by_key(SHEET_ID)
        sh.worksheet("危駕_設定").clear()
        sh.worksheet("危駕_設定").update(
            range_name='A1',
            values=[["Key", "Value"], ["plan_time", p_time], ["commander", cmdr]]
        )
        for name, df in [("危駕_指揮組", df_c), ("危駕_警力佈署", df_p)]:
            ws = sh.worksheet(name)
            ws.clear()
            data = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
            ws.update(range_name='A1', values=data)
        return True
    except:
        return False

# --- 日期解析與時段字串計算 ---
def calc_time_strings(p_time):
    """
    回傳 (第一列專用, 第二列以後專用)
    """
    date_match = re.search(r'(?:(\d+)年)?(\d+)月(\d+)日', p_time)
    if not date_match:
        return "", ""

    y = date_match.group(1)
    m = int(date_match.group(2))
    d = int(date_match.group(3))

    y_tw = int(y) if y else (datetime.now().year - 1911)
    base_dt = datetime(y_tw + 1911, m, d)
    next_dt = base_dt + timedelta(days=1)

    # 第一列：勤務時間的隔日 零時至4時
    dedicated_time = f"{next_dt.month}月{next_dt.day}日\n零時至4時"
    # 第二列以後：勤務時間的當日 22時至隔日6時
    normal_time = f"{m}月{d}日\n22時至翌日6時"

    return dedicated_time, normal_time

# --- 後處理：校正時段欄位 ---
def fix_time_column(df, dedicated_time, normal_time):
    df = df.copy()
    for i in range(len(df)):
        cur_t = str(df.at[i, '勤務時段']).strip()
        if i == 0:
            # 第一列若為空，則填入隔日零時
            if is_blank(cur_t):
                df.at[i, '勤務時段'] = dedicated_time
        else:
            # 第二列以後：若是空的，或者含有『零時』(代表從第一列誤複製)，則強制填入當日22時
            if is_blank(cur_t) or is_from_first_row(cur_t):
                df.at[i, '勤務時段'] = normal_time
    return df

# --- PDF 字型與產生 (略，維持原本邏輯) ---
def register_font():
    font_paths = ["kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/kaiu.ttf"]
    for fp in font_paths:
        if os.path.exists(fp):
            pdfmetrics.registerFont(TTFont("BiauKai", fp))
            return True
    return False

FONT_AVAILABLE = register_font()
FONT_NAME = "BiauKai" if FONT_AVAILABLE else "Helvetica"

def generate_pdf(df_c, df_p, cmdr_n, time_s):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=12*mm, bottomMargin=12*mm)
    styles = {"title": ParagraphStyle("title", fontName=FONT_NAME, fontSize=14, alignment=1),
              "cell": ParagraphStyle("cell", fontName=FONT_NAME, fontSize=9)}
    story = [Paragraph(f"{UNIT_TITLE} 勤務規劃表", styles["title"]), Spacer(1, 5*mm)]
    
    # 這裡實作 PDF 表格內容 (簡化版)
    # ... 原本的 PDF 產生邏輯 ...
    doc.build(story)
    buf.seek(0)
    return buf

# ============================================================
# 主介面
# ============================================================
st.title("🚔 防制危險駕車專案勤務規劃表")

if 'data_ptl' not in st.session_state:
    s, c, p = load_from_cloud()
    if s is not None:
        sd = dict(zip(s.iloc[:, 0].astype(str), s.iloc[:, 1].astype(str)))
        st.session_state.p_time = sd.get("plan_time", "115年4月30日22時至翌日6時")
        st.session_state.cmdr = sd.get("commander", "石門所副所長林榮裕")
        st.session_state.data_cmd, st.session_state.data_ptl = c, p
    else:
        st.session_state.p_time = "115年4月30日22時至翌日6時"
        st.session_state.cmdr = "石門所副所長林榮裕"
        st.session_state.data_cmd = pd.DataFrame(columns=["職稱", "代號", "姓名", "任務"])
        st.session_state.data_ptl = pd.DataFrame(columns=["勤務時段", "代號", "編組", "服勤人員", "任務分工"])

p_time = st.text_input("1. 勤務時間", st.session_state.p_time)
cmdr_input = st.text_input("2. 交通快打指揮官", st.session_state.cmdr)

# 計算日期字串
dedicated_time, normal_time = calc_time_strings(p_time)

# --- 任務編組 ---
st.subheader("3. 任務編組")
res_cmd = st.data_editor(st.session_state.data_cmd, num_rows="dynamic", use_container_width=True).fillna("")

# --- 警力佈署 ---
st.subheader("4. 警力佈署")
if len(st.session_state.data_ptl) == 0:
    st.session_state.data_ptl = pd.DataFrame([["", "", "", "", ""]], columns=["勤務時段", "代號", "編組", "服勤人員", "任務分工"])

# 預填第一列基本資訊
if is_blank(st.session_state.data_ptl.at[0, '勤務時段']):
    st.session_state.data_ptl.at[0, '勤務時段'] = dedicated_time
    st.session_state.data_ptl.at[0, '編組'] = f"專責警力\n（{cmdr_input[:3]}輪值）"

res_ptl_raw = st.data_editor(st.session_state.data_ptl, num_rows="dynamic", use_container_width=True).fillna("")

# 【核心修正】後處理校正
res_ptl = fix_time_column(res_ptl_raw, dedicated_time, normal_time)

# 更新狀態
st.session_state.p_time, st.session_state.cmdr = p_time, cmdr_input
st.session_state.data_cmd, st.session_state.data_ptl = res_cmd, res_ptl

# --- 按鈕區 ---
if st.button("💾 同步至雲端", type="primary"):
    if save_to_cloud(p_time, cmdr_input, res_cmd, res_ptl):
        st.success("✅ 同步成功！第一列為隔日 00-04，其餘已自動校正為當日 22-06。")
