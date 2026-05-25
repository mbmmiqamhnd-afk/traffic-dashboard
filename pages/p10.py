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
# 預設底稿（完整7筆資料）
# =========================
DEFAULT_PROJECT = "防制危險駕車專案勤務"
DEFAULT_TIME = "115年5月22日22時至翌日6時"
DEFAULT_FAST_CMD = "龍潭所副所長全楚文"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "業務組", "代號": "隆安13", "姓名": "交通組巡官 郭勝隆", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報群聚飆車狀況。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青\n執勤官 李文章\n執勤員 黃文興", "任務": "監看群聚告警訊息、指揮、調度及通報本勤務事宜。"},
])

DEFAULT_PTL = pd.DataFrame([
    {"勤務時段": "5月23日\n零時至4時", "代號": "隆安62", "編組": "專責警力\n（龍潭所輪值）", 
     "服勤人員": "00-02時段：\n警員廖怡惠\n警員劉柏延\n02-04時段：\n警員林軒宇\n警員廖怡惠", 
     "任務分工": "「加強防制」勤務，在文化路、中正路三坑段、龍源路及旭日路來回巡邏，隨機攔檢改裝（噪音）車輛（每2小時至責任區域內指定巡簽地點巡簽1次並守望10分鐘，將守望情形拍照上傳LINE「龍潭分局聯絡平臺」群組）"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安80", "編組": "石門所線上巡邏組合警力兼任", "服勤人員": "", 
     "任務分工": "「區域聯防」勤務，於中正路、文化路、中豐路、龍源路及旭日路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安90", "編組": "高平所線上巡邏組合警力兼任", "服勤人員": "", 
     "任務分工": "「區域聯防」勤務，於中豐路及龍源路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安990", "編組": "龍潭交通分隊線上巡邏組合警力兼任", "服勤人員": "", 
     "任務分工": "「區域聯防」勤務，於龍源路及旭日路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安50", "編組": "聖亭所線上巡邏組合警力兼任", "服勤人員": "", 
     "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安60", "編組": "龍潭所線上巡邏組合警力兼任", "服勤人員": "", 
     "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安70", "編組": "中興所線上巡邏組合警力兼任", "服勤人員": "", 
     "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
])

DEFAULT_SIGN_POINTS = "巡簽地點：\n1. 中油高原交流道站（龍源路2-20號）\n2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）\n3. 7-11龍潭佳園門市（中正路三坑段776號）\n4. 旭日路三坑自然生態公園停車場\n5. 旭日路與大溪區交界處"
DEFAULT_NOTES = "一、各編組執行前由帶班人員在駐地實施勤前教育。\n二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。\n三、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。\n四、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"

# =========================
# 字體
# =========================
@st.cache_resource
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames():
        return fname
    for p in ["./kaiu.ttf", "kaiu.ttf", "/usr/share/fonts/truetype/kaiu.ttf", "/app/kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(fname, p))
                return fname
            except:
                pass
    return "Helvetica"

# =========================
# Google Sheets
# =========================
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        st.error("❌ 找不到 gcp_service_account，請確認 Secrets 設定。")
        return None
    try:
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Google 授權失敗：{e}")
        return None

def init_sheets():
    client = get_client()
    if client is None: return
    sh = client.open_by_key(SHEET_ID)
    headers = {WS_MAP["set"]: [["Key", "Value"]], WS_MAP["cmd"]: [CMD_COLS], WS_MAP["ptl"]: [PTL_COLS]}
    for name, header in headers.items():
        try:
            sh.worksheet(name)
        except:
            sh.add_worksheet(title=name, rows="200", cols="20").update(header)
    st.success("初始化完成")
    st.cache_data.clear()
    st.rerun()

@st.cache_data(ttl=5)
def load_data():
    try:
        client = get_client()
        if client is None:
            return None, None, None, {}, "授權失敗"
        sh = client.open_by_key(SHEET_ID)
        
        set_df = pd.DataFrame(sh.worksheet(WS_MAP["set"]).get_all_records()).fillna("")
        cmd_df = pd.DataFrame(sh.worksheet(WS_MAP["cmd"]).get_all_records()).fillna("")
        ptl_raw = sh.worksheet(WS_MAP["ptl"]).get_all_records()
        ptl_df = pd.DataFrame(ptl_raw).fillna("")
        
        # 強制修正欄位
        if not ptl_df.empty:
            ptl_df = ptl_df.reindex(columns=PTL_COLS, fill_value="")
        
        settings = {}
        if not set_df.empty and set_df.shape[1] >= 2:
            settings = dict(zip(set_df.iloc[:,0].astype(str), set_df.iloc[:,1].astype(str)))
            
        return set_df, cmd_df, ptl_df, settings, None
    except Exception as e:
        st.warning(f"載入資料失敗，使用預設底稿。錯誤：{e}")
        return None, None, None, {}, str(e)

# （save_data、generate_pdf、send_email 保持之前穩定版本）

# ... 其餘程式碼請使用我上一個回覆中的完整版本（包含 generate_pdf 和 UI 部分）...

# 主畫面強制顯示預設資料
st.title("🚔 防制危險駕車勤務規劃")

if st.sidebar.button("🔧 初始化工作表"):
    init_sheets()
if st.sidebar.button("🔄 強制重新載入"):
    st.cache_data.clear()
    st.rerun()

set_df, cmd_df, ptl_df, settings, err = load_data()

use_cmd = cmd_df if (cmd_df is not None and not cmd_df.empty) else DEFAULT_CMD.copy()
use_ptl = ptl_df if (ptl_df is not None and not ptl_df.empty and len(ptl_df) > 1) else DEFAULT_PTL.copy()

# 基本設定 + 表格顯示（保持不變）
st.subheader("基本設定")
col_a, col_b = st.columns(2)
project_name = col_a.text_input("專案名稱", value=settings.get("project_name", DEFAULT_PROJECT))
time_val = col_b.text_input("勤務時間", value=settings.get("time", DEFAULT_TIME))
fast_cmd = st.text_input("交通快打指揮官", value=settings.get("fast_cmd", DEFAULT_FAST_CMD))

st.subheader("1. 任務編組")
res_cmd = st.data_editor(use_cmd, num_rows="dynamic", use_container_width=True)

st.subheader("2. 警力佈署")
st.caption("💡 相同勤務時段會自動合併第一欄")
res_ptl = st.data_editor(use_ptl, num_rows="dynamic", use_container_width=True, height=680)

# 其餘 UI 部分（儲存、下載、Email）請接上之前版本
