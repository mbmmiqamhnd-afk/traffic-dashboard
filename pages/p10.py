import streamlit as st
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    pass

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import io, os, smtplib
import urllib.parse as _ul
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

# =========================
# 基本設定
# =========================
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
WS_MAP = {"set": "危駕_設定", "cmd": "危駕_指揮組", "ptl": "危駕_警力佈署"}
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
UNIT_TITLE = "桃園市政府警察局龍潭分局"
CMD_COLS = ["職稱", "代號", "姓名", "任務"]
PTL_COLS = ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]

# =========================
# 預設底稿
# =========================
DEFAULT_PROJECT = "防制危險駕車專案勤務"
DEFAULT_TIME = "115年5月22日22時至翌日6時"
DEFAULT_FAST_CMD = "龍潭所副所長全楚文"

DEFAULT_CMD = pd.DataFrame([ ... ])  # 保持原樣，省略

DEFAULT_PTL = pd.DataFrame([
    {"勤務時段": "5月23日\n零時至4時", "代號": "隆安62", "編組": "專責警力\n（龍潭所輪值）", 
     "服勤人員": "00-02時段：\n警員廖怡惠\n警員劉柏延\n02-04時段：\n警員林軒宇\n警員廖怡惠", 
     "任務分工": "「加強防制」勤務，在文化路、中正路三坑段、龍源路及旭日路來回巡邏，隨機攔檢改裝（噪音）車輛（每2小時至責任區域內指定巡簽地點巡簽1次並守望10分鐘，將守望情形拍照上傳LINE「龍潭分局聯絡平臺」群組）"},
    
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安80", "編組": "石門所", 
     "服勤人員": "線上巡邏組合警力兼任", 
     "任務分工": "「區域聯防」勤務，於中正路、文化路、中豐路、龍源路及旭日路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安90", "編組": "高平所", 
     "服勤人員": "線上巡邏組合警力兼任", 
     "任務分工": "「區域聯防」勤務，於中豐路及龍源路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安990", "編組": "龍潭交通分隊", 
     "服勤人員": "線上巡邏組合警力兼任", 
     "任務分工": "「區域聯防」勤務，於龍源路及旭日路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安50", "編組": "聖亭所", 
     "服勤人員": "線上巡邏組合警力兼任", 
     "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安60", "編組": "龍潭所", 
     "服勤人員": "線上巡邏組合警力兼任", 
     "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安70", "編組": "中興所", 
     "服勤人員": "線上巡邏組合警力兼任", 
     "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
])

DEFAULT_SIGN_POINTS = "巡簽地點：\n1. 中油高原交流道站（龍源路2-20號）\n2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）\n3. 7-11龍潭佳園門市（中正路三坑段776號）\n4. 旭日路三坑自然生態公園停車場\n5. 旭日路與大溪區交界處"
DEFAULT_NOTES = "一、各編組執行前由帶班人員在駐地實施勤前教育。\n二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。\n三、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。\n四、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"

# =========================
# 字體與 Google Sheets 函數（省略，與之前相同）
# =========================
# ... _get_font(), get_client(), init_sheets(), save_data() 保持不變 ...

@st.cache_data(ttl=5)
def load_data():
    try:
        client = get_client()
        if client is None:
            return None, None, None, {}, "授權失敗"
        sh = client.open_by_key(SHEET_ID)
        set_df = pd.DataFrame(sh.worksheet(WS_MAP["set"]).get_all_records()).fillna("")
        cmd_df = pd.DataFrame(sh.worksheet(WS_MAP["cmd"]).get_all_records()).fillna("")
        ptl_df = pd.DataFrame(sh.worksheet(WS_MAP["ptl"]).get_all_records()).fillna("")
        
        if not ptl_df.empty:
            ptl_df = ptl_df.reindex(columns=PTL_COLS, fill_value="")
            # 移除空白列（勤務時段為空）
            ptl_df = ptl_df[ptl_df["勤務時段"].astype(str).str.strip() != ""]
        
        settings = {}
        if not set_df.empty and set_df.shape[1] >= 2:
            settings = dict(zip(set_df.iloc[:,0].astype(str), set_df.iloc[:,1].astype(str)))
        return set_df, cmd_df, ptl_df, settings, None
    except Exception as e:
        return None, None, None, {}, str(e)

# =========================
# PDF 生成（穩定版）
# =========================
def generate_pdf(...):   # 請保留您之前穩定的 generate_pdf 函數
    # ... (內容與上一個版本相同) ...
    pass

# =========================
# UI 主畫面
# =========================
st.title("🚔 防制危險駕車勤務規劃")

if st.sidebar.button("🔧 初始化工作表"):
    init_sheets()
if st.sidebar.button("🔄 強制重新載入"):
    st.cache_data.clear()
    st.rerun()

set_df, cmd_df, ptl_df, settings, err = load_data()
if err:
    st.warning(f"⚠️ 無法連線 Google Sheets（{err}），顯示預設底稿。")

use_cmd = cmd_df if (cmd_df is not None and not cmd_df.empty) else DEFAULT_CMD.copy()
use_ptl = ptl_df if (ptl_df is not None and not ptl_df.empty) else DEFAULT_PTL.copy()

st.subheader("基本設定")
col_a, col_b = st.columns(2)
project_name = col_a.text_input("專案名稱", value=settings.get("project_name", DEFAULT_PROJECT))
time_val = col_b.text_input("勤務時間", value=settings.get("time", DEFAULT_TIME))
fast_cmd = st.text_input("交通快打指揮官", value=settings.get("fast_cmd", DEFAULT_FAST_CMD))

st.subheader("1. 任務編組")
res_cmd = st.data_editor(use_cmd, num_rows="dynamic", use_container_width=True)

st.subheader("2. 警力佈署")
st.caption("💡 相同勤務時段會自動合併 • 已移除空白列")
res_ptl = st.data_editor(use_ptl, num_rows="dynamic", use_container_width=True, height=680)

st.subheader("3. 巡簽地點與備註")
col_c, col_d = st.columns(2)
sign_points = col_c.text_area("巡簽地點", value=settings.get("sign_points", DEFAULT_SIGN_POINTS), height=160)
notes = col_d.text_area("備註", value=settings.get("notes", DEFAULT_NOTES), height=160)

st.markdown("---")
col1, col2, col3 = st.columns(3)

with col1:
    if st.button("💾 儲存至雲端", use_container_width=True):
        s = {"project_name": project_name, "time": time_val, "fast_cmd": fast_cmd, "sign_points": sign_points, "notes": notes}
        if save_data(s, res_cmd, res_ptl):
            st.success("✅ 已儲存")

with col2:
    pdf_buf = generate_pdf(time_val, project_name, fast_cmd, res_cmd, res_ptl, sign_points, notes)
    filename = f"{UNIT_TITLE}執行「{project_name}」規劃表"
    st.download_button("📄 下載 PDF", data=pdf_buf, file_name=f"{filename}.pdf", mime="application/pdf", use_container_width=True)

with col3:
    if st.button("📧 發送 Email", use_container_width=True):
        pdf_buf2 = generate_pdf(time_val, project_name, fast_cmd, res_cmd, res_ptl, sign_points, notes)
        ok, mail_err = send_email(filename, pdf_buf2, filename)
        if ok:
            st.success("✅ 已寄出")
        else:
            st.error(f"❌ 發送失敗：{mail_err}")
