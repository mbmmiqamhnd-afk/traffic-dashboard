import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib, io, os
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

# --- 1. 頁面設定 ---
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# --- 預設範本資料 ---
DEFAULT_TIME        = "115年3月6日22時至翌日6時"
DEFAULT_BRIEF       = "時間：各編組執行前\n地點：現地勤教"
DEFAULT_COMMANDER   = "石門所副所長林榮裕"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官",   "代號": "隆安1",   "姓名": "分局長 施宇峰",       "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2",   "姓名": "副分局長 何憶雯",     "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3",   "姓名": "副分局長 蔡志明",     "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "業務組",   "代號": "隆安13",  "姓名": "交通組警務員 葉佳媛", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報群聚飆車狀況。"},
    {"職稱": "督導組",   "代號": "隆安681", "姓名": "督察組督察員 黃中彥", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "通訊組",   "代號": "隆安",    "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "監看群聚告警訊息、指揮、調度及通報本勤務事宜。"},
])

DEFAULT_PATROL = pd.DataFrame([
    {"勤務時段": "3月7日零時至4時",       "無線電": "隆安82",  "編組": "專責警力（石門所輪值）", "服勤人員": "00-02時：副所長林榮裕、警員王耀民\n02-04時：副所長林榮裕、警員陳欣妤", "任務分工": "「加強防制」勤務，在文化路、中正路三坑段、龍源路及旭日路來回巡邏，隨機攔檢改裝（噪音）車輛（每2小時至責任區域內指定巡簽地點巡簽1次並守望10分鐘，將守望情形拍照上傳LINE「龍潭分局聯絡平臺」群組）"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安80",  "編組": "石門所",      "服勤人員": "線上巡邏警力兼任",  "任務分工": "「區域聯防」勤務，於中正路、文化路、中豐路、龍源路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次），並加強查緝毒駕"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安90",  "編組": "高平所",      "服勤人員": "線上巡邏警力兼任",  "任務分工": "「區域聯防」勤務，於中豐路及龍源路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安990", "編組": "龍潭交通分隊", "服勤人員": "線上巡邏警力兼任",  "任務分工": "「區域聯防」勤務，於龍源路及溪州橋旁新建道路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安50",  "編組": "聖亭所",      "服勤人員": "線上巡邏警力兼任",  "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安60",  "編組": "龍潭所",      "服勤人員": "線上巡邏警力兼任",  "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安70",  "編組": "中興所",      "服勤人員": "線上巡邏警力兼任",  "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
])

CHECKIN_POINTS = """1. 中油高原交流道站（龍源路2-20號）
2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）
3. 7-11龍潭佳園門市（中正路三坑段776號）
4. 旭日路三坑自然生態公園停車場
5. 旭日路與大溪區交界處"""

NOTES = """一、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。
二、駕駛巡邏車應開啟警示燈，如發現有危險駕車行為「勿追車」，並立即向勤指中心報告，請求以優勢警力執行攔截圍捕。
三、針對下列違法、違規事項加強攔查：
（一）道路交通管理處罰條例第16條（改裝排管）、第18條（改裝車體設備）、第21條（無照駕駛）及43條各款項（蛇行、嚴重超速、逼車、任意減速、拆除消音器、以其他方式造成噪音、兩車以上競速等）及第35條1項2款（毒駕）。
（二）違反刑法185條公共危險罪（以他法致生往來危險者）。
（三）違反社會秩序維護法第72條妨害安寧者，同法第64條聚眾不解散。"""

# --- 2. 建立連線與讀取 ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_data(ttl=60)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("危駕_設定")
        ws_cmd = sh.worksheet("危駕_指揮組")
        ws_ptl = sh.worksheet("危駕_警力佈署")
        return pd.DataFrame(ws_set.get_all_records()), pd.DataFrame(ws_cmd.get_all_records()), pd.DataFrame(ws_ptl.get_all_records()), None
    except Exception as e: return None, None, None, str(e)

def save_data(time_str, briefing, commander, df_cmd, df_patrol):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("危駕_設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["plan_time", time_str], ["briefing", briefing], ["commander", commander]])
        
        for ws_name, df in [("危駕_指揮組", df_cmd), ("危駕_警力佈署", df_patrol)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df =
