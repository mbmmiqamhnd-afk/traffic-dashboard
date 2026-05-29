import streamlit as st

# --- 1. 頁面設定 (必須是全站第一個執行的 Streamlit 指令) ---
st.set_page_config(page_title="二合一專案勤務規劃系統", layout="wide", page_icon="🚓")

# 呼叫側邊欄
try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    st.sidebar.warning("找不到 menu.py，跳過側邊欄載入。")

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib, io, os, traceback, re
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

PTL_COLS = ["組別", "無線電代號", "派遣單位", "職別", "姓名", "任務分工", "攜行裝備", "臨檢目標"]
CP_COLS  = ["組別", "無線電代號", "派遣單位", "職別", "姓名", "任務分工", "臨檢目標場所"]

WS_SET_NAME = "二合一_設定"
WS_CMD_NAME = "二合一_指揮組"
WS_PTL_NAME = "二合一_路檢組"
WS_CP_NAME  = "二合一_擴大臨檢組"

DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年3月25日 18時至22時"
DEFAULT_PROJ    = "0325「雷霆除暴專案」暨自辦擴大臨檢與取締酒後駕車二合一專案"

DEFAULT_BRIEF   = (
    "一、 工作重點任務提示：同仁執行盤查、臨檢及路檢勤務過程中，應強化敵情觀念，提高危機意識，"
    "並特別注意人犯戒護，落實「人犯戒護安全、案件程序安全、執法者及民眾安全」之「三安」要求。\n"
    "二、 行動要領：除法律另有規定外，警察人員執行場所之臨檢，應限於已發生危害或依客觀合理判斷易生危害之場所、"
    "交通工具或公共場所為之。\n"
    "三、 盤查規範：確實依司法院大法官釋字第535號解釋及「警察職權行使法」對於盤查人、車以及實施臨檢之相關規定，"
    "應注意遵守比例原則及考量民眾觀感，不得逾越必要程度。\n"
    "四、 全程蒐證：執行各項干涉、取締、處理糾紛及爭議性勤務，務必全程連續錄音或錄影，以避免因案件招致物議。\n"
    "五、 異議處理：民眾對警察行使職權表示異議，認為無理由者得繼續執行，但經請求時應將異議之理由製作紀錄交付之。"
)

DEFAULT_PTL_FOCUS = (
    "採取全面機動巡邏，針對酒駕熱點攔停盤查；攔獲疑似改裝噪音車，立即引導至「警政大樓廣場」交由環保局檢驗。\n"
    "(雨備方案：於各所轄區易發生刑案地點、金融機構、超商、治安要點及危駕路段加強巡邏盤查人車。)"
)
DEFAULT_CP_FOCUS = (
    "由第一階段之各組機動警力，會合偵查隊專案人員，準時進入目標場所執行威力掃蕩。"
)

DEFAULT_CMD = pd.DataFrame([
    {"項目": "指揮官",     "通訊代號": "隆安 1 號",  "任務目標": "重點機動督導",                                                      "負責人員": "分局長 施宇峰",     "共同執行人員": "秘書 陳鵬翔、警員 張庭溱"},
    {"項目": "副指揮官",   "通訊代號": "隆安 2 號",  "任務目標": "重點機動督導",                                                      "負責人員": "副分局長 何憶雯",   "共同執行人員": "警務佐 曾威仁"},
    {"項目": "副指揮官",   "通訊代號": "隆安 3 號",  "任務目標": "重點機動督導",                                                      "負責人員": "副分局長 蔡志明",   "共同執行人員": "警員 陳明祥"},
    {"項目": "上級督導官", "通訊代號": "建興",        "任務目標": "重點機動督導",                                                      "負責人員": "督察 孫三陽",       "共同執行人員": ""},
    {"項目": "偵查隊",     "通訊代號": "隆安 11號",  "任務目標": "在隊督辦刑案",                                                      "負責人員": "隊長 柯志賢",       "共同執行人員": "偵查員 施明輝"},
    {"項目": "行政組",     "通訊代號": "隆安 5 號",  "任務目標": "督導第一階段臨檢組",                                                "負責人員": "組長 周金柱",       "共同執行人員": "巡官 蕭凱文、警務佐 曾威仁、警員 謝明展"},
    {"項目": "督察組",     "通訊代號": "隆安 6 號",  "任務目標": "機動督導第二階段時檢組",                                            "負責人員": "組長 黃長旗",       "共同執行人員": "警務員 陳冠彰"},
    {"項目": "保安民防組", "通訊代號": "隆安 9 號",  "任務目標": "機動督導第一階段臨檢組；機動督導第二階段路檢組",                               "負責人員": "組長 林良鍾",       "共同執行人員": "巡官 古家杰"},
    {"項目": "交通組",     "通訊代號": "隆安 13號",  "任務目標": "機動督導第一階段路檢組",                                            "負責人員": "組長 楊孟竟",       "共同執行人員": "巡官 郭勝隆"},
    {"項目": "勤務指導",   "通訊代號": "隆安 685號", "任務目標": "指導各路檢點、攔檢點，指導各檢查組勤務執行及狀況處置",        "負責人員": "教官 郭文義",       "共同執行人員": "勤務指導人員"},
    {"項目": "聯絡組",     "通訊代號": "隆安",        "任務目標": "擔任通訊聯絡、指揮管制事宜",                                      "負責人員": "勤指主任 蔡奇青",   "共同執行人員": "執勤官 江文頌、值勤員 曾嘉偉 (18-20時)"},
    {"項目": "偵訊組",     "通訊代號": "隆安 10號",  "任務目標": "負責按捺指紋、照相及移送案件相關事宜",                             "負責人員": "偵查佐 賴享宏、警員 張峻銨", "共同執行人員": "在隊待命受理移送案件"},
    {"項目": "作業組",     "通訊代號": "",            "任務目標": "負責勤務後勤、勤教場地布置相關事宜",                               "負責人員": "警員 葉俊宏、警務員 曾盛鉉", "共同執行人員": "巡官 吳國棟、巡佐 許榮裕、警員 呂紹臺"},
])

DEFAULT_PTL = pd.DataFrame([
    {"組別": "第1路檢組", "無線電代號": "隆安51", "派遣單位": "聖亭所",   "職別": "所長",   "姓名": "鄭榮捷", "任務分工": "帶班兼管制",   "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號前\n（攔檢中興路往龍潭市區方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第1路檢組", "無線電代號": "隆安51", "派遣單位": "聖亭所",   "職別": "警員",   "姓名": "詹宗澤", "任務分工": "指揮管制",     "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號前\n（攔檢中興路往龍潭市區方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第1路檢組", "無線電代號": "隆安51", "派遣單位": "龍潭所",   "職別": "警員",   "姓名": "劉柏延", "任務分工": "攔檢盤查",     "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號前\n（攔檢中興路往龍潭市區方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第1路檢組", "無線電代號": "隆安51", "派遣單位": "龍潭所",   "職別": "警員",   "姓名": "林宸緯", "任務分工": "攔檢盤查",     "攜行裝備": "小電腦、密錄器",               "臨檢目標": "北龍路319號前\n（攔檢中興路往龍潭市區方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第1路檢組", "無線電代號": "隆安51", "派遣單位": "高平所",   "職別": "警員",   "姓名": "黃丞穎", "任務分工": "警戒兼蒐證",   "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號前\n（攔檢中興路往龍潭市區方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第2路檢組", "無線電代號": "隆安82", "派遣單位": "石門所",   "職別": "副所長", "姓名": "林榮裕", "任務分工": "帶班兼管制",   "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號隊面前\n（攔檢龍潭市區往中興路方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第2路檢組", "無線電代號": "隆安82", "派遣單位": "石門所",   "職別": "警員",   "姓名": "陳琦",   "任務分工": "指揮管制",     "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號隊面前\n（攔檢龍潭市區往中興路方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第2路檢組", "無線電代號": "隆安82", "派遣單位": "中興所",   "職別": "巡佐",   "姓名": "蕭漢祥", "任務分工": "攔檢盤查",     "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號隊面前\n（攔檢龍潭市區往中興路方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第2路檢組", "無線電代號": "隆安82", "派遣單位": "中興所",   "職別": "警員",   "姓名": "江益德", "任務分工": "攔檢盤查",     "攜行裝備": "槍彈、無線電、小電腍、密錄器", "臨檢目標": "北龍路319號隊面前\n（攔檢龍潭市區往中興路方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第2路檢組", "無線電代號": "隆安82", "派遣單位": "交通分隊", "職別": "小隊長", "姓名": "林振生", "任務分工": "攔檢盤查",     "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號隊面前\n（攔檢龍潭市區往中興路方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第2路檢組", "無線電代號": "隆安82", "派遣單位": "交通分隊", "職別": "警員",   "姓名": "吳沛軒", "任務分工": "警戒兼蒐證",   "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號隊面前\n（攔檢龍潭市區往中興路方向）\n20時20分分局一樓集合出發臨檢"},
])

DEFAULT_CHECKPOINT = pd.DataFrame([
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "聖亭所", "職別": "所長",   "姓名": "鄭榮捷", "任務分工": "帶班",                             "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nC. 丹陽泰養生館（中豐路281號）IC335\nD. 溫馨汽車旅館（中正路457號）IA337\nE. 凱虹汽車旅館（中正路506號）IA318"},
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "聖亭所", "職別": "警員",   "姓名": "詹宗澤", "任務分工": "製作臨檢紀錄",                     "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nC. 丹陽泰養生館（中豐路281號）IC335\nD. 溫馨汽車旅館（中正路457號）IA337\nE. 凱虹汽車旅館（死中正路506號）IA318"},
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "聖亭所", "職別": "警員",   "姓名": "劉柏延", "任務分工": "盤查兼蒐證",                       "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nC. 丹陽泰養生館（中豐路281號）IC335\nD. 溫馨汽車旅館（中正路457號）IA337\nE. 凱虹汽車旅館（中正路506號）IA318"},
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "龍潭所", "職別": "警員",   "姓名": "林宸緯", "任務分工": "盤查兼蒐證",                       "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nC. 丹陽泰養生館（中豐路281號）IC335\nD. 溫馨汽車旅館（中正路457號）IA337\nE. 凱虹汽車旅館（中正路506號）IA318"},
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "高平所", "職別": "警員",   "姓名": "黃丞穎", "任務分工": "大門警(車)戒兼蒐證",               "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nC. 丹陽泰養生館（中豐路281號）IC335\nD. 溫馨汽車旅館（中正路457號）IA337\nE. 凱虹汽車旅館（中正路506號）IA318"},
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "偵查隊", "職別": "偵查佐", "姓名": "賴享宏", "任務分工": "刑案偵防、社維法案件之處理及移送", "臨檢目標場所": "A. 鉅大撞球館（Play館）（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nC. 丹陽泰養生館（中豐路281號）IC335\nD. 溫馨汽車旅館（中正路457號）IA337\nE. 凱虹汽車旅館（中正路506號）IA318"},
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "偵查隊", "職別": "警員",   "姓名": "張峻銨", "任務分工": "刑案偵防、社維法案件之處理及移送", "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nC. 丹陽泰養生館（中豐路281號）IC335\nD. 溫馨汽車旅館（中正路457號）IA337\nE. 凱虹汽車旅館（心中正路506號）IA318"},
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "石門所",   "職別": "副所長", "姓名": "林榮裕", "任務分工": "帶班",                             "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nF. 憤怒鳥網咖（中興路269號）IB330\nG. 真情男女養生館（中興路387號）IB329\nH. 萬紫千紅舒壓館（中興路491-3號）IB326"},
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "石門所",   "職別": "警員",   "姓名": "陳琦",   "任務分工": "製作臨檢紀錄",                     "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nF. 憤怒鳥網咖（中興路269號）IB330\nG. 真情男女養生館（中興路387號）IB329\nH. 萬紫千紅舒壓館（中興路491-3號）IB326"},
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "中興所",   "職別": "巡佐",   "姓名": "蕭漢祥", "任務分工": "盤查兼蒐證",                       "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nF. 憤怒鳥網咖（中興路269號）IB330\nG. 真情男女養生館（中興路387號）IB329\nH. 萬紫千紅舒壓館（中興路491-3號）IB326"},
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "中興所",   "職別": "警員",   "姓名": "江益德", "任務分工": "盤查兼蒐證",                       "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nF. 憤怒鳥網咖（中興路269號）IB330\nG. 真情男女養生館（中興路387號）IB329\nH. 萬紫千紅舒壓館（中興路491-3號）IB326"},
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "交通分隊", "職別": "小隊長", "姓名": "林振生", "任務分工": "盤查兼蒐證",                       "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nF. 憤怒鳥網咖（中興路269號）IB330\nG. 真情男女養生館（中興路387號）IB329\nH. 萬紫千紅舒壓館（中興路491-3號）IB326"},
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "交通分隊", "職別": "警員",   "姓名": "吳沛軒", "任務分工": "大門警(車)戒兼蒐證",               "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nF. 憤怒鳥網咖（中興路269號）IB330\nG. 真情男女養生館（中興路387號）IB329\nH. 萬紫千紅舒壓館（中興路491-3號）IB326"},
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "偵查隊",   "職別": "警員",   "姓名": "駿宏",   "任務分工": "刑案偵防、社維法案件之處理及移送", "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）IC329\nB. 台灣麻將協會（中豐路558之1號）IC328\nF. 憤怒鳥網咖（中興路269號）IB330\nG. 真情男女養生館（中興路387號）IB329\nH. 萬紫千紅舒壓館（中興路491-3號）IB326"},
])

# ─────────────── 核心輔助函數 ───────────────

def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames():
        return fname
    font_paths = [
        "./kaiu.ttf",
        "kaiu.ttf",
        "/usr/share/fonts/truetype/custom/kaiu.ttf",
        "C:/Windows/Fonts/kaiu.ttf",
    ]
    for p in font_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def safe_str(val):
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s

def clean_df_to_list(df):
    return df.astype(str).values.tolist()

# ─────────────── ★ get_client ───────────────

@st.cache_resource
def get_client():
    try:
        info = dict(st.secrets["gcp_service_account"])
        # 確保 private_key 的 \n 是真實換行而非字面字串
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Google 授權失敗：{e}")
        return None

# ─────────────── 資料載入 ───────────────

@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        if client is None:
            return None, None, None, None, "權限不足或未設定 Secrets"
        sh = client.open_by_key(SHEET_ID)

        try:
            ws_set = sh.worksheet(WS_SET_NAME)
            df_set = pd.DataFrame(ws_set.get_all_records()).fillna("")
        except Exception:
            df_set = None

        try:
            ws_cmd = sh.worksheet(WS_CMD_NAME)
            df_cmd = pd.DataFrame(ws_cmd.get_all_records()).fillna("")
        except Exception:
            df_cmd = pd.DataFrame()

        try:
            ws_ptl = sh.worksheet(WS_PTL_NAME)
            df_ptl = pd.DataFrame(ws_ptl.get_all_records()).fillna("")
        except Exception:
            df_ptl = pd.DataFrame()

        try:
            ws_cp = sh.worksheet(WS_CP_NAME)
            df_cp = pd.DataFrame(ws_cp.get_all_records()).fillna("")
        except Exception:
            df_cp = None

        return df_set, df_cmd, df_ptl, df_cp, None

    except Exception as e:
        return None, None, None, None, str(e)

# ─────────────── 資料儲存 ───────────────

def save_data(unit, time_str, project, briefing, df_cmd, df_ptl, df_cp, stats, ptl_f, cp_f):
    try:
        client = get_client()
        if client is None:
            st.error("❌ 無法取得 Google 授權，請確認 Secrets 設定。")
            return False
        sh = client.open_by_key(SHEET_ID)

        # 1. 二合一_設定
        try:
            ws_set = sh.worksheet(WS_SET_NAME)
        except Exception:
            ws_set = sh.add_worksheet(title=WS_SET_NAME, rows="50", cols="5")
        ws_set.clear()
        ws_set.update(range_name="A1", values=[
            ["Key", "Value"],
            ["unit_name",      unit],
            ["plan_full_time", time_str],
            ["project_name",   project],
            ["briefing_info",  briefing],
            ["stats_cmd",      str(stats["cmd"])],
            ["stats_ptl",      str(stats["ptl_road"])],   # 路檢組
            ["stats_cp",       str(stats["ptl_cp"])],     # 臨檢組
            ["stats_inv",      str(stats["inv"])],
            ["stats_civ",      str(stats["civ"])],
            ["briefing_time",  str(stats["b_time"])],
            ["briefing_loc",   str(stats["b_loc"])],
            ["loc_1",          str(stats["loc_1"])],
            ["loc_2",          str(stats["loc_2"])],
            ["loc_3",          str(stats["loc_3"])],
            ["ptl_focus",      ptl_f],
            ["cp_focus",       cp_f],
        ])

        # 2. 二合一_指揮組
        try:
            ws_c = sh.worksheet(WS_CMD_NAME)
        except Exception:
            ws_c = sh.add_worksheet(title=WS_CMD_NAME, rows="100", cols="20")
        ws_c.clear()
        clean_cmd = df_cmd.dropna(how="all").fillna("")
        if not clean_cmd.empty:
            ws_c.update(range_name="A1", values=[clean_cmd.columns.tolist()] + clean_df_to_list(clean_cmd))

        # 3. 二合一_路檢組
        try:
            ws_p = sh.worksheet(WS_PTL_NAME)
        except Exception:
            ws_p = sh.add_worksheet(title=WS_PTL_NAME, rows="100", cols="20")
        ws_p.clear()
        clean_ptl = df_ptl.dropna(how="all").fillna("")
        if not clean_ptl.empty:
            ws_p.update(range_name="A1", values=[clean_ptl.columns.tolist()] + clean_df_to_list(clean_ptl))

        # 4. 二合一_擴大臨檢組
        try:
            ws_cp = sh.worksheet(WS_CP_NAME)
        except Exception:
            ws_cp = sh.add_worksheet(title=WS_CP_NAME, rows="100", cols="20")
        ws_cp.clear()
        clean_cp = df_cp.dropna(how="all").fillna("")
        if not clean_cp.empty:
            ws_cp.update(range_name="A1", values=[clean_cp.columns.tolist()] + clean_df_to_list(clean_cp))

        load_data.clear()
        return True

    except Exception as e:
        st.error(f"❌ 同步失敗原因：{e}")
        st.code(traceback.format_exc())
        return False

# ─────────────── PDF 生成：規劃表 ───────────────

def generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, stats, ptl_f, cp_f):
    import re

    font = _get_font()
    buf  = io.BytesIO()
    doc  = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=10*mm, rightMargin=10*mm,
        topMargin=12*mm,  bottomMargin=15*mm,
    )
    page_width = A4[0] - 20*mm
    story = []

    style_title      = ParagraphStyle("Title",      fontName=font, fontSize=18, leading=26, alignment=1, spaceAfter=8,    wordWrap="CJK")
    style_section    = ParagraphStyle("Section",    fontName=font, fontSize=14, leading=20, alignment=0, spaceAfter=2*mm, spaceBefore=4*mm, wordWrap="CJK")
    style_text       = ParagraphStyle("Text",       fontName=font, fontSize=14, leading=20, alignment=0, wordWrap="CJK")
    style_cell       = ParagraphStyle("Cell",       fontName=font, fontSize=12, leading=17, alignment=1, wordWrap="CJK")
    style_cell_left  = ParagraphStyle("CellLeft",   fontName=font, fontSize=12, leading=17, alignment=0, wordWrap="CJK")
    style_cp_target  = ParagraphStyle("CpTarget",   fontName=font, fontSize=10, leading=14, alignment=0, wordWrap="CJK")
    style_briefing_hang = ParagraphStyle(
        "BriefingHang",
        fontName=font, fontSize=14, leading=22, alignment=0,
        leftIndent=22, firstLineIndent=-22, spaceAfter=4,
        wordWrap="CJK",
    )

    def clean(t):
        return safe_str(t).replace("\n", "<br/>")

    story.append(Paragraph(f"<b>{unit}執行 {project} 勤務規劃表</b>", style_title))

    # 壹、基本資料
    story.append(Paragraph("<b>壹、 勤務基本資料</b>", style_section))
    date_str      = clean(time_str.split(" ")[0] if " " in time_str else "115年3月25日")
    time_str_only = clean(time_str.split(" ")[1] if " " in time_str else "18時至22時")
    briefing_time_loc_str = f"{stats['b_time']}<br/>{stats['b_loc']}"

    data_basic = [
        [Paragraph(f"<b>{h}</b>", style_cell) for h in ["實施日期", "勤務時間", "指揮官", "勤務編組", "勤前教育時間地點"]],
        [
            Paragraph(date_str, style_cell),
            Paragraph(time_str_only, style_cell),
            Paragraph("分局長 施宇峰", style_cell),
            Paragraph("如任務編組表", style_cell),
            Paragraph(briefing_time_loc_str, style_cell),
        ],
    ]
    t_basic = Table(data_basic, colWidths=[
        page_width*0.19, page_width*0.16, page_width*0.19,
        page_width*0.16, page_width*0.30,
    ])
    t_basic.setStyle(TableStyle([
        ("FONTNAME",   (0,0),(-1,-1), font),
        ("GRID",       (0,0),(-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0),(-1, 0), colors.HexColor("#f2f2f2")),
        ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
    ]))
    story.append(t_basic)

    # 貳、統計表
    story.append(Paragraph("<b>貳、 警力統計及地點統計</b>", style_section))

    style_sub_section = ParagraphStyle("SubSection", fontName=font, fontSize=12, leading=18, alignment=0, spaceAfter=1*mm, spaceBefore=2*mm, wordWrap="CJK")
    story.append(Paragraph("<b>一、 警力統計：</b>", style_sub_section))

    # ── 警力統計表：拆分路檢組 / 臨檢組 ──
    ptl_road = stats.get("ptl_road", 0)
    ptl_cp   = stats.get("ptl_cp",   0)
    total    = stats["cmd"] + ptl_road + ptl_cp + stats["inv"] + stats["civ"]

    data_stats = [
        [Paragraph(f"<b>{h}</b>", style_cell) for h in ["督導組", "路檢組", "臨檢組", "偵訊組", "民力", "總計"]],
        [
            Paragraph(str(stats["cmd"]),  style_cell),
            Paragraph(str(ptl_road),      style_cell),
            Paragraph(str(ptl_cp),        style_cell),
            Paragraph(str(stats["inv"]),  style_cell),
            Paragraph(str(stats["civ"]),  style_cell),
            Paragraph(str(total),         style_cell),
        ],
    ]
    t_stats = Table(data_stats, colWidths=[page_width/6]*6)
    t_stats.setStyle(TableStyle([
        ("FONTNAME",   (0,0),(-1,-1), font),
        ("GRID",       (0,0),(-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0),(-1, 0), colors.HexColor("#f2f2f2")),
        ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
    ]))
    story.append(t_stats)
    story.append(Spacer(1, 2*mm))

    story.append(Paragraph("<b>二、 地點統計：</b>", style_sub_section))

    ptl_loc_count = 0
    if not df_ptl.empty and "組別" in df_ptl.columns:
        ptl_loc_count = df_ptl["組別"].dropna().loc[lambda x: x.astype(str).str.strip() != ""].nunique()

    cp_count = 0
    if df_cp is not None and not df_cp.empty and "臨檢目標場所" in df_cp.columns:
        raw_targets = df_cp["臨檢目標場所"].dropna().unique()
        found_places = set()
        for target in raw_targets:
            target_clean = str(target).strip()
            if target_clean and target_clean.lower() != "nan":
                matches = re.findall(r'(?:^[A-Z0-9熱點]\s*[\.\、\-\：]\s*|[A-Z0-9熱點]\s*[\.\、\-\：]\s*)([^\n]+)', target_clean, re.MULTILINE)
                for item in matches:
                    place_title = item.strip().split("（")[0].split("(")[0][:15]
                    if place_title:
                        found_places.add(place_title)
        cp_count = len(found_places) if found_places else 8

    data_locs = [
        [Paragraph("<b>路檢點</b>", style_cell), Paragraph("<b>臨檢場所</b>", style_cell)],
        [Paragraph(f"{ptl_loc_count} 處", style_cell), Paragraph(f"{cp_count} 處", style_cell)]
    ]
    t_locs = Table(data_locs, colWidths=[page_width*0.5, page_width*0.5])
    t_locs.setStyle(TableStyle([
        ("FONTNAME",   (0,0),(-1,-1), font),
        ("GRID",       (0,0),(-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0),( -1, 0), colors.HexColor("#f2f2f2")),
        ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
        ("TOPPADDING", (0,0),(-1,-1), 6),
    ]))
    story.append(t_locs)
    story.append(Spacer(1, 4*mm))

    # 參、指揮組
    story.append(Paragraph("<b>參、 督導及其他任務編組表</b>", style_section))
    data_cmd = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["項目","通訊代號","任務目標","負責人員","共同人員"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([
            Paragraph(clean(r.get("項目","")),       style_cell),
            Paragraph(clean(r.get("通訊代號","")),    style_cell),
            Paragraph(clean(r.get("任務目標","")),    style_cell_left),
            Paragraph(clean(r.get("負責人員","")),    style_cell),
            Paragraph(clean(r.get("共同執行人員","")),style_cell),
        ])

    t_cmd = Table(data_cmd, colWidths=[
        page_width*0.13, page_width*0.14, page_width*0.26,
        page_width*0.25, page_width*0.22,
    ])
    t_cmd.setStyle(TableStyle([
        ("FONTNAME",   (0,0),(-1,-1), font),
        ("GRID",       (0,0),(-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0),(-1, 0), colors.HexColor("#f2f2f2")),
        ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
    ]))
    story.append(t_cmd)

    # 肆、第一階段定點路檢
    story.append(Paragraph("<b>肆、【第一階段】定點路檢任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務重點：</b><br/>{clean(ptl_f)}", style_text))

    ptl_headers = ["組別","無線電\n代號","派遣\n單位","職別","姓名","任務分工","攜行裝備","臨檢目標"]
    col_w_ptl   = [
        page_width*0.10, page_width*0.09, page_width*0.09, page_width*0.10,
        page_width*0.11, page_width*0.12, page_width*0.15, page_width*0.24,
    ]
    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ptl_headers]]
    rows_ptl = df_ptl.reset_index(drop=True)
    merge_groups = []
    prev_group, grp_start_idx = None, 1
    for i, r in rows_ptl.iterrows():
        grp = safe_str(r.get("組別",""))
        tbl_row = i + 1
        if grp != prev_group:
            if prev_group is not None:
                merge_groups.append((grp_start_idx, tbl_row - 1))
            prev_group = grp
            grp_start_idx = tbl_row
        data_ptl.append([
            Paragraph(clean(r.get("組別","")),      style_cell),
            Paragraph(clean(r.get("無線電代號","")), style_cell),
            Paragraph(clean(r.get("派遣單位","")),   style_cell),
            Paragraph(clean(r.get("職別","")),       style_cell),
            Paragraph(clean(r.get("姓名","")),       style_cell),
            Paragraph(clean(r.get("任務分工","")),   style_cell),
            Paragraph(clean(r.get("攜行裝備","")),   style_cell_left),
            Paragraph(clean(r.get("臨檢目標","")),   style_cell_left),
        ])
    if prev_group is not None:
        merge_groups.append((grp_start_idx, len(rows_ptl)))
    t_ptl = Table(data_ptl, colWidths=col_w_ptl)
    ts_ptl = [
        ("FONTNAME",   (0,0),(-1,-1), font),
        ("GRID",       (0,0),(-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0),(-1, 0), colors.HexColor("#f2f2f2")),
        ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
    ]
    for (rs, re) in merge_groups:
        if re > rs:
            for col in [0, 1, 7]:
                ts_ptl.append(("SPAN", (col, rs), (col, re)))
    t_ptl.setStyle(TableStyle(ts_ptl))
    story.append(t_ptl)

    # 伍、第二階段擴大臨檢
    story.append(Paragraph("<b>伍、【第二階段】場所臨檢任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務重點：</b><br/>{clean(cp_f)}", style_text))

    if df_cp is not None and not df_cp.empty:
        cp_headers = ["組別","無線電\n代號","派遣\n單位","職別","姓名","任務分工","臨檢目標場所"]
        col_w_cp   = [
            page_width*0.10, page_width*0.09, page_width*0.09, page_width*0.10,
            page_width*0.11, page_width*0.16, page_width*0.35,
        ]
        data_cp = [[Paragraph(f"<b>{h}</b>", style_cell) for h in cp_headers]]
        rows_cp = df_cp.reset_index(drop=True)
        cp_merge_groups = []
        cp_prev_group, cp_grp_start = None, 1
        for i, r in rows_cp.iterrows():
            grp = safe_str(r.get("組別",""))
            tbl_row = i + 1
            if grp != cp_prev_group:
                if cp_prev_group is not None:
                    cp_merge_groups.append((cp_grp_start, tbl_row - 1))
                cp_prev_group = grp
                cp_grp_start  = tbl_row
            data_cp.append([
                Paragraph(clean(r.get("組別","")),          style_cell),
                Paragraph(clean(r.get("無線電代號","")),     style_cell),
                Paragraph(clean(r.get("派遣單位","")),       style_cell),
                Paragraph(clean(r.get("職別","")),           style_cell),
                Paragraph(clean(r.get("姓名","")),           style_cell),
                Paragraph(clean(r.get("任務分工","")),       style_cell_left),
                Paragraph(clean(r.get("臨檢目標場所","")),   style_cp_target),
            ])
        if cp_prev_group is not None:
            cp_merge_groups.append((cp_grp_start, len(rows_cp)))
        t_cp = Table(data_cp, colWidths=col_w_cp, splitByRow=True)
        ts_cp = [
            ("FONTNAME",   (0,0),(-1,-1), font),
            ("GRID",       (0,0),(-1,-1), 0.5, colors.black),
            ("BACKGROUND", (0,0),(-1, 0), colors.HexColor("#e6e6e6")),
            ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
        ]
        for (rs, re) in cp_merge_groups:
            if re > rs:
                for col in [0, 1, 6]:
                    ts_cp.append(("SPAN", (col, rs), (col, re)))
        t_cp.setStyle(TableStyle(ts_cp))
        story.append(t_cp)

    # 陸、法令宣導
    story.append(Paragraph("<b>陸、 工作重點與法令宣導</b>", style_section))
    for line in str(briefing).split("\n"):
        if line.strip():
            story.append(Paragraph(clean(line), style_briefing_hang))

    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 10)
        canvas.drawCentredString(A4[0]/2.0, 10*mm, f"-第{canvas.getPageNumber()}頁-")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    return buf.getvalue()

# ─────────────── PDF 生成：簽到表 ───────────────

def generate_attendance_pdf(unit, project, time_str, stats):
    font = _get_font()
    buf  = io.BytesIO()
    doc  = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=10*mm,  bottomMargin=10*mm, # 微調邊界，爭取單頁空間
    )
    page_width = A4[0] - 30*mm
    story = []

    style_title = ParagraphStyle("Title", fontName=font, fontSize=18, leading=26, alignment=1, spaceAfter=8, wordWrap="CJK")
    style_info  = ParagraphStyle("Info",  fontName=font, fontSize=14, leading=22, spaceAfter=1*mm, wordWrap="CJK")
    style_cell  = ParagraphStyle("Cell",  fontName=font, fontSize=14, leading=20, alignment=1, wordWrap="CJK")
    # 新增長官簽核欄位的樣式
    style_sig   = ParagraphStyle("Sig",   fontName=font, fontSize=14, leading=20, alignment=0, wordWrap="CJK") 

    story.append(Paragraph(f"{unit}執行{project}簽到表", style_title))
    date_part = time_str.split(" ")[0] if " " in time_str else "115年3月25日"
    story.append(Paragraph(f"時間：{date_part} {stats['b_time']}", style_info))
    story.append(Paragraph(f"地點：{stats['b_loc']}召開", style_info))
    
    story.append(Spacer(1, 3*mm)) # 縮小間距

    # --- 修改：長官簽核欄位 (分局長與上級督導同一列) ---
    sig_data = [
        [Paragraph("分局長：", style_sig), Paragraph("上級督導：", style_sig)],
        [Paragraph("副分局長：", style_sig), ""]
    ]
    t_sig = Table(sig_data, colWidths=[page_width/2.0]*2)
    t_sig.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2), # 縮小 padding 以節省空間
        # 不設定 GRID，讓表格無外框，僅作排版使用
    ]))
    story.append(t_sig)
    story.append(Spacer(1, 4*mm)) # 縮小間距
    # ------------------------------------

    rows = [
        ("交通組",     "聖亭派出所"),
        ("督察組",     "龍潭派出所"),
        ("行政組",     "中興派出所"),
        ("保安民防組", "石門派出所"),
        ("勤務指揮中心","高平派出所"),
        ("偵查隊",     "三和派出所"),
        ("",           "龍潭交通分隊"),
    ]
    table_data = [[
        Paragraph("單位",     style_cell),
        Paragraph("參加人員", style_cell),
        Paragraph("單位",     style_cell),
        Paragraph("參加人員", style_cell),
    ]]
    for l, r in rows:
        table_data.append([
            Paragraph(l, style_cell) if l else "",
            "",
            Paragraph(r, style_cell) if r else "",
            "",
        ])

    t = Table(
        table_data,
        colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3],
        # ▼ 此處縮小列高 (標題列10mm，資料列20mm)，確保能維持單頁產出
        rowHeights=[10*mm] + [20*mm]*len(rows),
    )
    t.setStyle(TableStyle([
        ("FONTNAME",   (0,0),(-1,-1), font),
        ("GRID",       (0,0),(-1,-1), 0.5, colors.black),
        ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
        ("BACKGROUND", (0,0),(3,  0), colors.whitesmoke),
    ]))
    story.append(t)
    
    # 頁尾維持不變
    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 10)
        canvas.drawCentredString(A4[0]/2.0, 10*mm, f"-第{canvas.getPageNumber()}頁-")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    return buf.getvalue()

# ─────────────── 郵件發送 ───────────────

def send_report_email(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, stats, ptl_f, cp_f):
    try:
        sender = st.secrets["email"]["user"]
        pwd    = st.secrets["email"]["password"]
        msg    = MIMEMultipart()
        msg["From"]    = sender
        msg["To"]      = sender
        msg["Subject"] = f"勤務規劃與簽到表_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為最新版本勤務規劃表。", "plain"))

        pdf1  = generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, stats, ptl_f, cp_f)
        part1 = MIMEBase("application", "pdf")
        part1.set_payload(pdf1)
        encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(f'{unit}規劃表.pdf')}")
        msg.attach(part1)

        pdf2  = generate_attendance_pdf(unit, project, time_str, stats)
        part2 = MIMEBase("application", "pdf")
        part2.set_payload(pdf2)
        encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(f'{unit}簽到表.pdf')}")
        msg.attach(part2)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None

    except Exception as e:
        return False, str(e)

# ─────────────── Streamlit 介面 ───────────────

df_set, df_cmd, df_ptl, df_cp, err = load_data()

default_stats = {
    "cmd":      7,
    "ptl_road": 10,   # 路檢組（第一階段）
    "ptl_cp":   6,    # 臨檢組（第二階段）
    "inv":      3,
    "civ":      0,
    "b_time": "18時30分至19時00分",
    "b_loc":  "本分局2樓會議室",
    "loc_1":  8,
    "loc_2":  2,
    "loc_3":  0,
}

if err or df_set is None:
    u             = DEFAULT_UNIT
    t             = DEFAULT_TIME
    p             = DEFAULT_PROJ
    ed_cmd        = DEFAULT_CMD.copy()
    ed_ptl        = DEFAULT_PTL.copy()
    ed_cp         = DEFAULT_CHECKPOINT.copy()
    p_ptl_focus   = DEFAULT_PTL_FOCUS
    p_cp_focus    = DEFAULT_CP_FOCUS
else:
    d = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    u           = d.get("unit_name",      DEFAULT_UNIT)
    t           = d.get("plan_full_time", DEFAULT_TIME)
    p           = d.get("project_name",   DEFAULT_PROJ)
    p_ptl_focus = d.get("ptl_focus",      DEFAULT_PTL_FOCUS)
    p_cp_focus  = d.get("cp_focus",       DEFAULT_CP_FOCUS)
    default_stats.update({
        "cmd":      int(d.get("stats_cmd",  7)),
        "ptl_road": int(d.get("stats_ptl",  10)),  # 向下相容舊欄位名
        "ptl_cp":   int(d.get("stats_cp",   6)),
        "inv":      int(d.get("stats_inv",  3)),
        "civ":      int(d.get("stats_civ",  0)),
        "b_time": d.get("briefing_time",  "18時30分至19時00分"),
        "b_loc":  d.get("briefing_loc",   "本分局2樓會議室"),
        "loc_1":  int(d.get("loc_1", 8)),
        "loc_2":  int(d.get("loc_2", 2)),
        "loc_3":  int(d.get("loc_3", 0)),
    })
    
    # --- ★ 關鍵修正部分 ---
    ed_cmd = (df_cmd if not df_cmd.empty else DEFAULT_CMD.copy()).astype(str)
    ed_ptl = (df_ptl[PTL_COLS] if (not df_ptl.empty and all(c in df_ptl.columns for c in PTL_COLS)) else DEFAULT_PTL.copy()).astype(str)
    ed_cp  = (df_cp[CP_COLS]   if (df_cp is not None and not df_cp.empty and all(c in df_cp.columns for c in CP_COLS))  else DEFAULT_CHECKPOINT.copy()).astype(str)

# ── 標題
st.title("二合一專案勤務規劃系統 🚓")
if err:
    st.warning(f"⚠️ 無法連線 Google Sheets（{err}），顯示預設資料。")

# ── 基本資訊
p_time = st.text_input("勤務時間", t)

display_project_name = re.sub(r'^\d{4}「?', '', p)
p_input = st.text_input("專案名稱", display_project_name)

date_match = re.search(r'(\d+)年(\d+)月(\d+)日', p_time)
if date_match:
    auto_4_digit = f"{int(date_match.group(2)):02d}{int(date_match.group(3)):02d}"
else:
    auto_4_digit = datetime.now().strftime("%m%d")

if not p_input.startswith("「"):
    p_name = f"{auto_4_digit}「{p_input}"
else:
    p_name = f"{auto_4_digit}{p_input}"

# ── 指揮組及編輯器區塊
st.subheader("參、 指揮編組與重點宣導")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True).dropna(how="all").fillna("")

st.subheader("勤務執行編組 (兩階段)")
tab1, tab2 = st.tabs(["肆、【第一階段】定點路檢", "伍、【第二階段】場所臨檢"])

with tab1:
    res_ptl_focus = st.text_area("【肆】勤務重點", p_ptl_focus, height=80, key="ptl_focus_input")
    st.caption("💡 同一路檢組的多名人員請填寫相同的「組別」與「無線電代號」，PDF 輸出時該欄會自動合併。")
    res_ptl = st.data_editor(
        ed_ptl,
        num_rows="dynamic",
        use_container_width=True,
        key="ptl_ed",
        column_config={
            "組別":      st.column_config.TextColumn("組別",      width="small"),
            "無線電代號": st.column_config.TextColumn("無線電代號", width="small"),
            "派遣單位":   st.column_config.TextColumn("派遣單位",   width="small"),
            "職別":      st.column_config.TextColumn("職別",      width="small"),
            "姓名":      st.column_config.TextColumn("姓名",      width="small"),
            "任務分工":   st.column_config.TextColumn("任務分工",   width="medium"),
            "攜行裝備":   st.column_config.TextColumn("攜行裝備",   width="medium"),
            "臨檢目標":   st.column_config.TextColumn("臨檢目標",   width="large"),
        },
    ).dropna(how="all").fillna("").reset_index(drop=True)

with tab2:
    res_cp_focus = st.text_area("【伍】勤務重點", p_cp_focus, height=80, key="cp_focus_input")
    st.caption("💡 同一臨檢組的多名人員請填寫相同的「組別」與「無線電代號」，PDF 輸出時該欄會自動合併。")
    res_cp = st.data_editor(
        ed_cp,
        num_rows="dynamic",
        use_container_width=True,
        key="cp_ed",
        column_config={
            "組別":        st.column_config.TextColumn("組別",        width="small"),
            "無線電代號":   st.column_config.TextColumn("無線電代號",   width="small"),
            "派遣單位":     st.column_config.TextColumn("派遣單位",     width="small"),
            "職別":        st.column_config.TextColumn("職別",        width="small"),
            "姓名":        st.column_config.TextColumn("姓名",        width="small"),
            "任務分工":     st.column_config.TextColumn("任務分工",     width="medium"),
            "臨檢目標場所": st.column_config.TextColumn("臨檢目標場所", width="large"),
        },
    ).dropna(how="all").fillna("").reset_index(drop=True)

# ── 動態計算警力統計數據 ──

# 1. 督導組：指揮組「負責人員」欄有填寫的列數
if not res_cmd.empty and "負責人員" in res_cmd.columns:
    cmd_series = res_cmd["負責人員"].astype(str).str.strip()
    calc_cmd_count = int(cmd_series[(cmd_series != "") & (cmd_series.str.lower() != "nan")].count())
else:
    calc_cmd_count = 0

# 2. 路檢組：第一階段「姓名」有填寫的列數
ptl_road_count = 0
if not res_ptl.empty and "姓名" in res_ptl.columns:
    ptl_series = res_ptl["姓名"].astype(str).str.strip()
    ptl_road_count = int(ptl_series[(ptl_series != "") & (ptl_series.str.lower() != "nan")].count())

# 3. 臨檢組：第二階段「姓名」有填寫的列數
ptl_cp_count = 0
if not res_cp.empty and "姓名" in res_cp.columns:
    cp_series = res_cp["姓名"].astype(str).str.strip()
    ptl_cp_count = int(cp_series[(cp_series != "") & (cp_series.str.lower() != "nan")].count())

# ── 貳、警力統計與地點統計顯示區塊 ──
st.subheader("貳、 警力統計及地點統計")
col_s1, col_s2, col_s3, col_s4, col_s5 = st.columns(5)

c_cmd  = col_s1.number_input("督導組 (自動計算)", value=calc_cmd_count,  min_value=0, disabled=True)
c_ptl  = col_s2.number_input("路檢組 (自動計算)", value=ptl_road_count,  min_value=0, disabled=True)
c_cp   = col_s3.number_input("臨檢組 (自動計算)", value=ptl_cp_count,    min_value=0, disabled=True)
c_inv  = col_s4.number_input("偵訊組",            value=default_stats["inv"], min_value=0)
c_civ  = col_s5.number_input("民力",              value=default_stats["civ"], min_value=0)

current_stats = {
    "cmd":      c_cmd,
    "ptl_road": c_ptl,   # 路檢組（第一階段）
    "ptl_cp":   c_cp,    # 臨檢組（第二階段）
    "ptl":      c_ptl + c_cp,  # 保留合計欄位供舊邏輯相容
    "inv":      c_inv,
    "civ":      c_civ,
    "total":    c_cmd + c_ptl + c_cp + c_inv + c_civ,
    "b_time":   default_stats["b_time"],
    "b_loc":    default_stats["b_loc"],
    "loc_1":    default_stats["loc_1"],
    "loc_2":    default_stats["loc_2"],
    "loc_3":    default_stats["loc_3"],
}

# ── 操作按鈕
st.markdown("---")

pdf_plan = generate_pdf_from_data(
    u, p_name, p_time, DEFAULT_BRIEF,
    res_cmd, res_ptl, res_cp,
    current_stats, res_ptl_focus, res_cp_focus,
)
st.download_button(
    "📝 下載規劃表 PDF",
    data=pdf_plan,
    file_name=f"{u}規劃表.pdf",
    mime="application/pdf",
    use_container_width=True,
)

if st.button("💾 同步雲端並發送郵件", use_container_width=True):
    with st.spinner("同步中，請稍候…"):
        ok = save_data(
            u, p_time, p_name, DEFAULT_BRIEF,
            res_cmd, res_ptl, res_cp,
            current_stats, res_ptl_focus, res_cp_focus,
        )
    if ok:
        with st.spinner("同步成功，正在寄送郵件…"):
            mail_ok, mail_err = send_report_email(
                u, p_name, p_time, DEFAULT_BRIEF,
                res_cmd, res_ptl, res_cp,
                current_stats, res_ptl_focus, res_cp_focus,
            )
        if mail_ok:
            st.success("✅ 資料已同步至 Google Sheets，郵件發送成功！")
        else:
            st.warning(f"⚠️ 同步成功，但郵件發送失敗：{mail_err}")
