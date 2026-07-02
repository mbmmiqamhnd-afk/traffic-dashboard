import streamlit as st

# ── 必須是第一個 Streamlit 指令 ──────────────────────────────────────────────
st.set_page_config(page_title="專案勤務規劃系統", layout="wide", page_icon="🚓")

try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    pass

import io, os, re, smtplib, urllib.parse as _ul
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

# ══════════════════════════════════════════════════════════════════════════════
# 1. 常數
# ══════════════════════════════════════════════════════════════════════════════
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES   = ["https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"]

DEFAULT_UNIT      = "桃園市政府警察局龍潭分局"
DEFAULT_TIME      = "115年4月10日 19時至23時"
DEFAULT_PROJ_BODY = "「全市取締酒後駕車及防制危險駕車」暨「擴大臨檢」及「取締改裝(噪音)車輛專案監、警、環聯合稽查」"
DEFAULT_BRIEF = (
    "一、 落實三安：同仁執行盤查、臨檢及機動勤務過程中，應強化敵情觀念，提高危機意識，"
    "落實「人犯戒護安全、案件程序安全、執法者及民眾安全」。\n"
    "二、 臨檢合法性：警察人員執行場所之臨檢，應限於已發生危害或依客觀合理判斷易生危害之場所，"
    "進行臨檢前應對當事人告以實施事由，便衣人員並應出示證件(依《警察職權行使法》第6條)。\n"
    "三、 攔停規範：機動攔檢對於已發生危害或易生危害之交通工具，得予以攔停；"
    "若有異常舉動而合理懷疑其將有危害行為時，得要求接受酒精濃度測試(依《警察職權行使法》第8條)。\n"
    "四、 全程蒐證：執行各項干涉、取締、處理糾紛及爭議性勤務(含噪音車引導與酒測)，務必全程連續錄音或錄影。\n"
    "五、 異議處理：民眾對警察行使職權表示異議，認為無理由者得繼續執行，但經請求時應將異議之理由製作紀錄交付之(依《警察職權行使法》第29條)。"
)
DEFAULT_PTL_TIME  = "20時00分至21時30分"
DEFAULT_PTL_FOCUS = "採取全面機動巡邏，針對酒駕熱點攔停盤查；攔獲疑似改裝噪音車，立即引導至「警政大樓廣場」交由環保局檢驗。"
DEFAULT_CP_TIME   = "21時30分至23時00分"
DEFAULT_CP_FOCUS  = "由第一階段之第1至第4組機動警力，會合偵查隊專案人員，於21時20分前集結完畢，21時30分準時進入目標場所執行威力掃蕩。"
DEFAULT_BRIEF_TIME = "19時30分至20時00分"
DEFAULT_BRIEF_LOC  = "分局二樓會議室"

DEFAULT_CMD = pd.DataFrame([
    {"項目": "指揮官",    "通訊代號": "隆安1號",    "任務目標": "勤務核定並重點機動督導",         "負責人員": "分局長 施宇峰",         "共同執行人員": "巡官陳鵬翔、警員張庭溱"},
    {"項目": "副指揮官",  "通訊代號": "隆安2號",    "任務目標": "襄助指揮、重點機動督導",         "負責人員": "副分局長 何憶雯",        "共同執行人員": "警務佐曾威仁"},
    {"項目": "副指揮官",  "通訊代號": "隆安3號",    "任務目標": "襄助指揮、重點機動督導",         "負責人員": "副分局長 蔡志明",        "共同執行人員": "警員陳明祥"},
    {"項目": "行政組",    "通訊代號": "隆安5號",    "任務目標": "督導場所臨檢威力掃蕩第一臨檢組",  "負責人員": "組長 周金柱",            "共同執行人員": "巡官蕭凱文"},
    {"項目": "督察組",    "通訊代號": "隆安6號",    "任務目標": "機動督導各單位勤務紀律",          "負責人員": "組長黃長旗",             "共同執行人員": "警務員 陳冠彰"},
    {"項目": "保安民防組", "通訊代號": "隆安9號",   "任務目標": "督導場所臨檢威力掃蕩第二臨檢組",  "負責人員": "組長林良鍾",             "共同執行人員": "警務員曾盛鉉、警務佐許榮裕、警務佐劉俊德"},
    {"項目": "交通組",    "通訊代號": "隆安13號",   "任務目標": "督導第一階段機動攔查",            "負責人員": "組長 楊孟竟",            "共同執行人員": "巡官郭勝隆、警務員李峯甫、警務員盧冠仁、警員吳享運"},
    {"項目": "聯絡組",    "通訊代號": "隆安",       "任務目標": "擔任通訊聯絡、指揮管制事宜",      "負責人員": "勤務指揮中心 主任蔡奇青", "共同執行人員": "執勤官李文章、執勤員黃文興"},
    {"項目": "偵訊組",    "通訊代號": "隆安10號",   "任務目標": "負責按捺指紋、照相及移送",        "負責人員": "偵查隊隊長 柯志賢",      "共同執行人員": "偵查隊值日小隊"},
    {"項目": "聯合稽查站", "通訊代號": "隆安1382",  "任務目標": "配合環保局及監理站稽查車輛",      "負責人員": "交通組巡官 郭勝隆",      "共同執行人員": "環保局及監理站人員"},
])

DEFAULT_PTL = pd.DataFrame([
    {"單位": "聖亭所", "無線電代號": "", "職別": "副所長", "姓名": "邱品淳", "任務分工": "帶班",      "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "中正路、北龍路周邊及治安要點機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "聖亭所", "無線電代號": "", "職別": "警員",   "姓名": "劉憬霖", "任務分工": "攔檢盤查",  "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "中正路、北龍路周邊及治安要點機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "聖亭所", "無線電代號": "", "職別": "警員",   "姓名": "謝伯昇", "任務分工": "警戒兼蒐證", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "中正路、北龍路周邊及治安要點機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "龍潭所", "無線電代號": "", "職別": "警員",   "姓名": "張家維", "任務分工": "帶班兼蒐證", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "北龍路、中豐路周邊及治安要點機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "龍潭所", "無線電代號": "", "職別": "警員",   "姓名": "王采蘋", "任務分工": "攔檢盤查",  "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "北龍路、中豐路周邊及治安要點機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "中興所", "無線電代號": "", "職別": "所長",   "姓名": "董亦文", "任務分工": "帶班兼蒐證", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "東龍路、中豐路沿線機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "中興所", "無線電代號": "", "職別": "警員",   "姓名": "羅俊傑", "任務分工": "攔檢盤查",  "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "東龍路、中豐路沿線機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "石門所", "無線電代號": "", "職別": "所長",   "姓名": "林育辰", "任務分工": "帶班兼蒐證", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "神龍路、文化路周邊及治安要點機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "三和所", "無線電代號": "", "職別": "警員",   "姓名": "童霈晟", "任務分工": "攔檢盤查",  "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "神龍路、文化路周邊及治安要點機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "石門所", "無線電代號": "", "職別": "巡佐",   "姓名": "林偉政", "任務分工": "帶班兼蒐證", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "中興路、龍新路沿線及治安要點機動攔查。(全程留守機動 20:00-23:00) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "高平所", "無線電代號": "", "職別": "警員",   "姓名": "葉雲翔", "任務分工": "攔檢盤查",  "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "中興路、龍新路沿線及治安要點機動攔查。(全程留守機動 20:00-23:00) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "龍潭交分隊", "無線電代號": "", "職別": "警員", "姓名": "林家豪", "任務分工": "帶班兼蒐證", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "轄內易發生危駕路段、各聯外道路機動攔查。(全程留守機動 20:00-23:00) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "龍潭交分隊", "無線電代號": "", "職別": "警員", "姓名": "吳沛軒", "任務分工": "攔檢盤查",  "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏路段": "轄內易發生危駕路段、各聯外道路機動攔查。(全程留守機動 20:00-23:00) *雨天備案:轄區治安要點巡邏。"},
])

DEFAULT_CHECKPOINT = pd.DataFrame([
    {"單位": "中興所", "無線電代號": "", "職別": "所長",   "姓名": "董亦文", "任務分工": "帶班",           "臨檢目標場所": "A. 鉅大撞球館 (中豐路558號)\nB. 台灣麻將協會 (中豐路558之1號)\nC. 丹陽泰養生館 (中豐路281號)\nD. 溫馨汽車旅館 (中正路457號)\nE. 凱虹汽車旅館 (中正路506號)\n*(各員均需著防彈衣，攜帶槍彈、小電腦、密錄器)*"},
    {"單位": "中興所", "無線電代號": "", "職別": "警員",   "姓名": "羅俊傑", "任務分工": "製作臨檢紀錄",    "臨檢目標場所": "A. 鉅大撞球館、B. 台灣麻將協會、C. 丹陽泰養生館、D. 溫馨汽車旅館、E. 凱虹汽車旅館"},
    {"單位": "龍潭所", "無線電代號": "", "職別": "警員",   "姓名": "張家維", "任務分工": "盤查兼蒐證",     "臨檢目標場所": "A. 鉅大撞球館、B. 台灣麻將協會、C. 丹陽泰養生館、D. 溫馨汽車旅館、E. 凱虹汽車旅館"},
    {"單位": "龍潭所", "無線電代號": "", "職別": "警員",   "姓名": "王采蘋", "任務分工": "盤查兼蒐證",     "臨檢目標場所": "A. 鉅大撞球館、B. 台灣麻將協會、C. 丹陽泰養生館、D. 溫馨汽車旅館、E. 凱虹汽車旅館"},
    {"單位": "偵查隊", "無線電代號": "", "職別": "警員",   "姓名": "許家洋", "任務分工": "刑案偵防、社維法案件查處", "臨檢目標場所": "A. 鉅大撞球館、B. 台灣麻將協會、C. 丹陽泰養生館、D. 溫馨汽車旅館、E. 凱虹汽車旅館"},
    {"單位": "石門所", "無線電代號": "", "職別": "所長",   "姓名": "林育辰", "任務分工": "帶班",           "臨檢目標場所": "A. 鉅大撞球館 (中豐路558號)\nB. 台灣麻將協會 (中豐路558之1號)\nF. 憤怒鳥網咖\nG. 真情男女養生館\nH. 萬紫千紅舒壓館\n*(各員均需著防彈衣，攜帶槍彈、小電腦、密錄器)*"},
    {"單位": "聖亭所", "無線電代號": "", "職別": "副所長", "姓名": "邱品淳", "任務分工": "製作臨檢紀錄",    "臨檢目標場所": "A. 鉅大撞球館、B. 台灣麻將協會、F. 憤怒鳥網咖、G. 真情男女養生館、H. 萬紫千紅舒壓館"},
    {"單位": "聖亭所", "無線電代號": "", "職別": "警員",   "姓名": "劉憬霖", "任務分工": "盤查兼蒐證",     "臨檢目標場所": "A. 鉅大撞球館、B. 台灣麻將協會、F. 憤怒鳥網咖、G. 真情男女養生館、H. 萬紫千紅舒壓館"},
    {"單位": "三和所", "無線電代號": "", "職別": "警員",   "姓名": "謝伯昇", "任務分工": "大門警戒",       "臨檢目標場所": "A. 鉅大撞球館、B. 台灣麻將協會、F. 憤怒鳥網咖、G. 真情男女養生館、H. 萬紫千紅舒壓館"},
    {"單位": "偵查隊", "無線電代號": "", "職別": "小隊長", "姓名": "陳正育", "任務分工": "刑案偵防、社維法案件查處", "臨檢目標場所": "A. 鉅大撞球館、B. 台灣麻將協會、F. 憤怒鳥網咖、G. 真情男女養生館、H. 萬紫千紅舒壓館"},
    {"單位": "偵查隊", "無線電代號": "", "職別": "偵查佐", "姓名": "鄧正斌", "任務分工": "持DV全程蒐證",   "臨檢目標場所": "A. 鉅大撞球館、B. 台灣麻將協會、F. 憤怒鳥網咖、G. 真情男女養生館、H. 萬紫千紅舒壓館"},
])

# ══════════════════════════════════════════════════════════════════════════════
# 2. 工具函數
# ══════════════════════════════════════════════════════════════════════════════
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames():
        return fname
    for p in ["./kaiu.ttf", "kaiu.ttf",
              "/usr/share/fonts/truetype/custom/kaiu.ttf",
              "C:/Windows/Fonts/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def safe_str(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s

def extract_mmdd(time_text: str) -> str:
    m = re.search(r'\d+\s*年\s*(\d+)\s*月\s*(\d+)\s*日', str(time_text))
    if m:
        return f"{int(m.group(1)):02d}{int(m.group(2)):02d}"
    return datetime.now().strftime("%m%d")

UNIT_RADIO_BASE = {
    "聖亭所": "5", "龍潭所": "6", "中興所": "7",
    "石門所": "8", "高平所": "9", "三和所": "3",
    "龍潭交分隊": "99", "偵查隊": "1",
}
SENIOR_RANKS = {"所長", "分隊長", "隊長", "副所長", "小隊長"}

def generate_radio_code(unit: str, rank: str, officer_seq: int = 1) -> str:
    base = UNIT_RADIO_BASE.get(unit.strip(), "")
    if not base:
        return ""
    rank = rank.strip()
    if rank == "所長" or rank == "分隊長" or rank == "隊長":
        return f"{base}1"
    if rank in ("副所長", "小隊長"):
        return f"{base}2"
    return f"{base}{2 + officer_seq}"

# ══════════════════════════════════════════════════════════════════════════════
# 3. 分組邏輯（只保留一份）
# ══════════════════════════════════════════════════════════════════════════════
PTL_UNIT_ORDER = {"聖亭所": 1, "龍潭所": 2, "中興所": 3,
                  "石門所": 4, "三和所": 5, "高平所": 6, "龍潭交分隊": 7}

CP_GROUP1_UNITS = {"中興所", "龍潭所"}

def _normalize_radio_col(res: pd.DataFrame) -> pd.DataFrame:
    """確保無線電代號欄位是純字串型態，防止從 Sheets 讀回 int64 造成寫入失敗。"""
    if "無線電代號" in res.columns:
        res["無線電代號"] = res["無線電代號"].astype(str).str.strip()
        res["無線電代號"] = res["無線電代號"].replace({"nan": "", "None": "", "0": ""})
    else:
        res["無線電代號"] = ""
    return res


def assign_ptl_groups(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    res = df.copy().reset_index(drop=True)
    res = _normalize_radio_col(res)
    res["_ord"] = res["單位"].map(lambda x: PTL_UNIT_ORDER.get(str(x).strip(), 99))
    res = res.sort_values("_ord").drop(columns=["_ord"]).reset_index(drop=True)

    group_ids, radio_codes, unit_officer_count = [], [], {}
    prev_unit, g_idx = None, 0
    for i, row in res.iterrows():
        unit = str(row["單位"]).strip()
        if unit != prev_unit:
            g_idx += 1
            prev_unit = unit
        group_ids.append(f"第{g_idx}巡邏組")

        existing = str(row.get("無線電代號", "")).strip()
        if existing and existing not in ("nan", "None", "0"):
            radio_codes.append(existing)
        elif unit:
            is_officer = row["職別"] in SENIOR_RANKS
            unit_officer_count[unit] = unit_officer_count.get(unit, 0) + (0 if is_officer else 1)
            radio_codes.append(generate_radio_code(unit, row["職別"], unit_officer_count[unit]))
        else:
            radio_codes.append("")

    res["編組"]      = group_ids
    res["無線電代號"] = radio_codes          # 整欄一次賦值，避免逐格寫入型別衝突

    # 同組無線電代號統一用該組第一筆
    for g in res["編組"].unique():
        mask = res["編組"] == g
        first_radio = res.loc[mask, "無線電代號"].iloc[0]
        res.loc[mask, "無線電代號"] = first_radio

    cols = ["編組", "無線電代號", "單位", "職別", "姓名", "任務分工", "攜行裝備", "巡邏路段"]
    return res[[c for c in cols if c in res.columns]]


def assign_cp_groups(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    res = df.copy().reset_index(drop=True)
    res = _normalize_radio_col(res)

    def _group(row):
        existing = str(row.get("編組", "")).strip()
        if existing in ("第1臨檢組", "第1場所臨檢組"):
            return 1
        if existing in ("第2臨檢組", "第2場所臨檢組"):
            return 2
        unit = str(row.get("單位", "")).strip()
        return 1 if unit in CP_GROUP1_UNITS else 2

    res["_g"] = res.apply(_group, axis=1)

    # 組內排序：幹部優先，同單位跟在後面；偵查隊固定排最後
    res["_is_senior"] = res["職別"].apply(lambda x: 0 if str(x).strip() in SENIOR_RANKS else 1)
    res["_is_invest"] = res["單位"].apply(lambda x: 1 if str(x).strip() == "偵查隊" else 0)
    # 組內以「原始列順序」為次排序，確保同單位人員維持原順序
    res["_orig_idx"] = res.index
    res = res.sort_values(["_g", "_is_invest", "_is_senior", "_orig_idx"]).drop(
        columns=["_is_senior", "_is_invest", "_orig_idx"]).reset_index(drop=True)

    # 先指定編組名稱
    group_ids = ["第1臨檢組" if row["_g"] == 1 else "第2臨檢組" for _, row in res.iterrows()]
    res["編組"] = group_ids
    res = res.drop(columns=["_g"])

    # 排序完成後重新計算無線電代號，確保代號對應排序後的第一列（帶班所長）
    unit_officer_count = {}
    radio_codes = []
    for i, row in res.iterrows():
        unit = str(row.get("單位", "")).strip()
        if unit:
            is_officer = row["職別"] in SENIOR_RANKS
            unit_officer_count[unit] = unit_officer_count.get(unit, 0) + (0 if is_officer else 1)
            radio_codes.append(generate_radio_code(unit, row["職別"], unit_officer_count[unit]))
        else:
            radio_codes.append("")

    res["無線電代號"] = radio_codes

    # 整組統一用第一列（帶班所長）的無線電代號
    for g in res["編組"].unique():
        mask = res["編組"] == g
        first_radio = res.loc[mask, "無線電代號"].iloc[0]
        res.loc[mask, "無線電代號"] = first_radio

    cols = ["編組", "無線電代號", "單位", "職別", "姓名", "任務分工", "臨檢目標場所"]
    return res[[c for c in cols if c in res.columns]]

# ══════════════════════════════════════════════════════════════════════════════
# 4. 動態統計
# ══════════════════════════════════════════════════════════════════════════════
def calculate_stats(df_cmd, df_ptl, df_cp):
    supervisors = set()
    for _, row in df_cmd.iterrows():
        aim = str(row.get("任務目標", ""))
        if "督導" in aim or "指導" in aim:
            leader = re.sub(r'(分局長|副分局長|組長|主任|巡官|教官|警務員|警務佐|偵查隊隊長)\s*', '',
                            str(row.get("負責人員", ""))).strip()
            if leader and leader.lower() != "nan":
                supervisors.add(leader)
            for name in re.split(r'[、,\s，]+', str(row.get("共同執行人員", ""))):
                n = re.sub(r'(巡官|警員|警務員|警務佐|巡佐|執勤官|值勤員)\s*', '', name).strip()
                if n and n.lower() != "nan" and "環保局" not in n and "監理站" not in n:
                    supervisors.add(n)
    cmd   = max(len(supervisors), 7)
    ptl_m = len({n for n in df_ptl["姓名"].dropna().astype(str).str.strip() if n}) or 13
    ptl_c = len({n for n in df_cp["姓名"].dropna().astype(str).str.strip() if n})  or 11
    inv   = 2
    return {"cmd": cmd, "ptl_机动": ptl_m, "ptl_场所": ptl_c,
            "inv": inv, "civ": 0, "total": cmd + ptl_m + ptl_c + inv}

# ══════════════════════════════════════════════════════════════════════════════
# 5. PDF 工具
# ══════════════════════════════════════════════════════════════════════════════
def _make_styles(font):
    def S(name, size, align, leading=None, space_after=0, space_before=0):
        return ParagraphStyle(name, fontName=font, fontSize=size,
                              leading=leading or size * 1.4,
                              alignment=align, wordWrap='CJK',
                              spaceAfter=space_after, spaceBefore=space_before)
    return {
        "title":     S("title",   18, TA_CENTER, leading=26, space_after=8),
        "section":   S("section", 14, TA_LEFT,   space_before=8, space_after=4),
        "text":      S("text",    14, TA_LEFT),
        "cell":      S("cell",    14, TA_CENTER),
        "cell_left": S("cl",      14, TA_LEFT),
        "cell_long": S("clong",   11, TA_LEFT,   leading=15),
    }

def _clean(t):
    return safe_str(t).replace("\n", "<br/>")

def _header_row(headers, style):
    return [Paragraph(f"<b>{h}</b>", style) for h in headers]

def _base_table_style(font, header_color="#f2f2f2"):
    return TableStyle([
        ("FONTNAME",    (0, 0), (-1, -1), font),
        ("GRID",        (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND",  (0, 0), (-1,  0), colors.HexColor(header_color)),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
    ])

def _apply_spans(style_cmds, data_list, merge_cols):
    """相鄰儲存格值相同則合併（SPAN）。"""
    if len(data_list) <= 1:
        return
    for col in merge_cols:
        start = 1
        for r in range(2, len(data_list)):
            def text_of(cell):
                return cell.text if hasattr(cell, "text") else str(cell)
            if text_of(data_list[r][col]) == text_of(data_list[r-1][col]) and text_of(data_list[r][col]).strip():
                if r == len(data_list) - 1:
                    style_cmds.append(("SPAN", (col, start), (col, r)))
            else:
                if r - 1 > start:
                    style_cmds.append(("SPAN", (col, start), (col, r - 1)))
                start = r


def generate_main_pdf(unit, project, time_str, briefing,
                      df_cmd, df_ptl, df_cp, stats,
                      ptl_time, ptl_focus, cp_time, cp_focus):
    font   = _get_font()
    buf    = io.BytesIO()
    PW     = A4[0] - 20 * mm
    doc    = SimpleDocTemplate(buf, pagesize=A4,
                               leftMargin=10*mm, rightMargin=10*mm,
                               topMargin=8*mm,   bottomMargin=8*mm)
    S      = _make_styles(font)
    story  = []

    def add_section(title):
        story.append(Paragraph(f"<b>{title}</b>", S["section"]))

    # 標題
    story.append(Paragraph(f"<b>{unit}執行 {project} 勤務規劃表</b>", S["title"]))

    # 壹、基本資料
    add_section("壹、 勤務基本資料")
    date_part = _clean(time_str.split()[0] if time_str.strip() else "115年4月10日")
    time_part = _clean(time_str.split()[1] if " " in time_str else "19時至23時")
    t = Table(
        [_header_row(["實施日期","勤務時間","指揮官","勤務編組","聯合稽查站地點"], S["cell"]),
         [Paragraph(date_part, S["cell"]), Paragraph(time_part, S["cell"]),
          Paragraph("分局長 施宇峰", S["cell"]), Paragraph("如任務編組表", S["cell"]),
          Paragraph("分局廣場", S["cell"])]],
        colWidths=[PW*.14, PW*.18, PW*.32, PW*.14, PW*.22])
    t.setStyle(_base_table_style(font))
    story.append(t)

    # 貳、警力統計
    add_section("貳、 警力統計及地點統計")
    t = Table(
        [_header_row(["督導組","機動攔檢組","場所臨檢組","偵訊組","小計","民力","總計"], S["cell"]),
         [Paragraph(str(v), S["cell"]) for v in [
             stats["cmd"], stats["ptl_机动"], stats["ptl_场所"], stats["inv"],
             stats["cmd"]+stats["ptl_机动"]+stats["ptl_场所"]+stats["inv"],
             stats["civ"], stats["total"]]]],
        colWidths=[PW/7]*7)
    t.setStyle(_base_table_style(font))
    story.append(t)

    # 參、督導編組
    add_section("參、 督導及其他任務編組表")
    data = [_header_row(["項目","通訊代號","任務目標","負責人員","共同人員"], S["cell"])]
    for _, r in df_cmd.iterrows():
        data.append([Paragraph(_clean(r.get("項目","")),      S["cell"]),
                     Paragraph(_clean(r.get("通訊代號","")),  S["cell"]),
                     Paragraph(_clean(r.get("任務目標","")),  S["cell_left"]),
                     Paragraph(_clean(r.get("負責人員","")),  S["cell"]),
                     Paragraph(_clean(r.get("共同執行人員","")), S["cell"])])
    t = Table(data, colWidths=[PW*.12, PW*.14, PW*.28, PW*.26, PW*.20])
    t.setStyle(_base_table_style(font))
    story.append(t)

    # 肆、第一階段
    add_section("肆、【第一階段】機動攔查任務編組")
    story.append(Paragraph(f"<b>勤務時間：</b>{_clean(ptl_time)}", S["text"]))
    story.append(Paragraph(f"<b>勤務重點：</b>{_clean(ptl_focus)}", S["text"]))
    data = [_header_row(["編組","無線電代號","單位","職別","姓名","任務分工","攜行裝備","巡邏路段"], S["cell"])]
    for _, r in df_ptl.iterrows():
        data.append([
            Paragraph(_clean(r.get("編組","")),     S["cell"]),
            Paragraph(_clean(r.get("無線電代號","")), S["cell"]),
            Paragraph(_clean(r.get("單位","")),     S["cell"]),
            Paragraph(_clean(r.get("職別","")),     S["cell"]),
            Paragraph(_clean(r.get("姓名","")),     S["cell"]),
            Paragraph(_clean(r.get("任務分工","")), S["cell_left"]),
            Paragraph(_clean(r.get("攜行裝備","")), S["cell_left"]),
            Paragraph(_clean(r.get("巡邏路段","")), S["cell_long"]),
        ])
    style_cmds = list(_base_table_style(font).getCommands())
    style_cmds[-1] = ("VALIGN", (0, 0), (-1, -1), "TOP")
    _apply_spans(style_cmds, data, [0, 1, 2, 7])
    t = Table(data, colWidths=[PW*.07, PW*.11, PW*.09, PW*.06, PW*.13, PW*.12, PW*.14, PW*.28])
    t.setStyle(TableStyle(style_cmds))
    story.append(t)

    # 伍、第二階段
    add_section("伍、【第二階段】場所臨檢任務編組")
    story.append(Paragraph(f"<b>勤務時間：</b>{_clean(cp_time)}", S["text"]))
    story.append(Paragraph(f"<b>勤務重點：</b>{_clean(cp_focus)}", S["text"]))
    if df_cp is not None and not df_cp.empty:
        data = [_header_row(["編組","無線電代號","單位","職別","姓名","任務分工","臨檢場所"], S["cell"])]
        for _, r in df_cp.iterrows():
            data.append([
                Paragraph(_clean(r.get("編組","")),       S["cell"]),
                Paragraph(_clean(r.get("無線電代號","")),  S["cell"]),
                Paragraph(_clean(r.get("單位","")),       S["cell"]),
                Paragraph(_clean(r.get("職別","")),       S["cell"]),
                Paragraph(_clean(r.get("姓名","")),       S["cell"]),
                Paragraph(_clean(r.get("任務分工","")),   S["cell_left"]),
                Paragraph(_clean(r.get("臨檢目標場所","")), S["cell_long"]),
            ])
        style_cmds = list(_base_table_style(font, "#e6e6e6").getCommands())
        style_cmds[-1] = ("VALIGN", (0, 0), (-1, -1), "TOP")
        _apply_spans(style_cmds, data, [0, 1, 2, 6])
        t = Table(data, colWidths=[PW*.07, PW*.11, PW*.09, PW*.06, PW*.13, PW*.19, PW*.35])
        t.setStyle(TableStyle(style_cmds))
        story.append(t)

    # 陸、法令宣導
    add_section("陸、 工作重點與法令宣導")
    for line in str(briefing).split("\n"):
        if line.strip():
            story.append(Paragraph(_clean(line), S["text"]))

    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 10)
        canvas.drawCentredString(A4[0] / 2, 10*mm, f"-第{canvas.getPageNumber()}頁-")
        canvas.restoreState()

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return buf.getvalue()


def generate_attendance_pdf(unit, project, time_str, brief_time, brief_loc):
    font  = _get_font()
    buf   = io.BytesIO()
    PW    = A4[0] - 30 * mm
    doc   = SimpleDocTemplate(buf, pagesize=A4,
                               leftMargin=15*mm, rightMargin=15*mm,
                               topMargin=10*mm,  bottomMargin=10*mm)
    S     = _make_styles(font)
    story = []

    story.append(Paragraph(f"<b>{unit}執行{project}簽到表</b>", S["title"]))
    date_part = time_str.split()[0] if " " in time_str else time_str
    story.append(Paragraph(f"時間：{date_part} {brief_time}", S["text"]))
    story.append(Paragraph(f"地點：{brief_loc}", S["text"]))
    story.append(Spacer(1, 3*mm))

    t = Table([[Paragraph("<b>分局長：</b>", S["cell_left"]),
                Paragraph("<b>上級督導：</b>", S["cell"]), ""]],
              colWidths=[PW*.3, PW*.4, PW*.3])
    t.setStyle(TableStyle([("FONTNAME", (0,0), (-1,-1), font),
                            ("VALIGN",   (0,0), (-1,-1), "TOP")]))
    story.append(t)
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph("<b>副分局長：</b>", S["text"]))
    story.append(Spacer(1, 6*mm))

    rows = [("交通組","聖亭派出所"), ("督察組","龍潭派出所"), ("行政組","中興派出所"),
            ("保安民防組","石門派出所"), ("勤務指揮中心","高平派出所"),
            ("偵查隊","三和派出所"), ("","龍潭交通分隊")]
    tdata = [_header_row(["單位","參加人員","單位","參加人員"], S["cell"])]
    for l, r in rows:
        tdata.append([Paragraph(l, S["cell"]) if l else "",
                      "", Paragraph(r, S["cell"]) if r else "", ""])

    t = Table(tdata, colWidths=[PW*.2, PW*.3, PW*.2, PW*.3],
              rowHeights=[10*mm] + [26*mm]*len(rows))
    t.setStyle(TableStyle([
        ("FONTNAME",   (0,0), (-1,-1), font),
        ("GRID",       (0,0), (-1,-1), 0.5, colors.black),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("BACKGROUND", (0,0), (3,   0), colors.whitesmoke),
    ]))
    story.append(t)
    doc.build(story)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# 6. Google Sheets
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        d = dict(st.secrets["gcp_service_account"])
        d["private_key"] = d["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(d, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception:
        return None


@st.cache_data(ttl=30)
def load_data():
    """讀取雲端設定，回傳 (設定dict, df_cmd, df_ptl, df_cp, err_str)。"""
    client = get_client()
    if client is None:
        return None, None, None, None, "無法建立 Google Sheets 連線"
    try:
        sh    = client.open_by_key(SHEET_ID)
        cfg   = {r["Key"]: r["Value"]
                 for r in sh.worksheet("三合一_設定").get_all_records()
                 if r.get("Key")}
        df_cmd = pd.DataFrame(sh.worksheet("三合一_指揮組").get_all_records()).fillna("")
        df_ptl = pd.DataFrame(sh.worksheet("三合一_巡邏組").get_all_records()).fillna("")
        df_cp  = pd.DataFrame(sh.worksheet("三合一_擴大臨檢組").get_all_records()).fillna("")
        return cfg, df_cmd, df_ptl, df_cp, None
    except Exception as e:
        return None, None, None, None, str(e)


def save_data(unit, time_str, project, briefing,
              ptl_time, ptl_focus, cp_time, cp_focus, brief_time, brief_loc,
              df_cmd, df_ptl, df_cp, stats):
    client = get_client()
    if client is None:
        return False, "無法建立連線"
    try:
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet("三合一_設定")
        ws.clear()
        ws.update([
            ["Key", "Value"],
            ["unit_name",    unit],
            ["plan_time",    time_str],
            ["project_name", project],
            ["briefing",     briefing],
            ["ptl_time",     ptl_time],
            ["ptl_focus",    ptl_focus],
            ["cp_time",      cp_time],
            ["cp_focus",     cp_focus],
            ["brief_time",   brief_time],
            ["brief_loc",    brief_loc],
            ["stats_cmd",    str(stats["cmd"])],
            ["stats_ptl_机动", str(stats["ptl_机动"])],
            ["stats_ptl_场所", str(stats["ptl_场所"])],
            ["stats_inv",    str(stats["inv"])],
            ["stats_total",  str(stats["total"])],
        ])
        for ws_name, df in [("三合一_指揮組", df_cmd),
                             ("三合一_巡邏組", df_ptl),
                             ("三合一_擴大臨檢組", df_cp)]:
            ws2 = sh.worksheet(ws_name)
            ws2.clear()
            clean = df.dropna(how="all").fillna("")
            if "編組" in clean.columns:
                clean = clean.drop(columns=["編組"])
            if not clean.empty:
                ws2.update([clean.columns.tolist()] + clean.astype(str).values.tolist())
        st.cache_data.clear()
        return True, None
    except Exception as e:
        return False, str(e)

# ══════════════════════════════════════════════════════════════════════════════
# 7. Email
# ══════════════════════════════════════════════════════════════════════════════
def send_email(unit, project, time_str, briefing,
               ptl_time, ptl_focus, cp_time, cp_focus, brief_time, brief_loc,
               df_cmd, df_ptl, df_cp, stats):
    try:
        sender = st.secrets["email"]["user"]
        pwd    = st.secrets["email"]["password"]
        msg    = MIMEMultipart()
        msg["From"]    = sender
        msg["To"]      = sender
        msg["Subject"] = f"勤務規劃與簽到表_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為最新版本勤務規劃表及簽到表。", "plain"))

        for pdf_bytes, filename in [
            (generate_main_pdf(unit, project, time_str, briefing,
                               df_cmd, df_ptl, df_cp, stats,
                               ptl_time, ptl_focus, cp_time, cp_focus), f"{unit}規劃表.pdf"),
            (generate_attendance_pdf(unit, project, time_str,
                                     brief_time, brief_loc), f"{unit}簽到表.pdf"),
        ]:
            part = MIMEBase("application", "pdf")
            part.set_payload(pdf_bytes)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f"attachment; filename*=UTF-8''{_ul.quote(filename)}")
            msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as srv:
            srv.login(sender, pwd)
            srv.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# ══════════════════════════════════════════════════════════════════════════════
# 8. Session State 初始化（含雲端還原）
# ══════════════════════════════════════════════════════════════════════════════
if "initialized" not in st.session_state:
    cfg, df_cmd_cl, df_ptl_cl, df_cp_cl, err = load_data()

    if cfg and not err:
        # ── 雲端資料優先還原 ──
        st.session_state.p_time     = cfg.get("plan_time",    DEFAULT_TIME)
        raw_proj = cfg.get("project_name", DEFAULT_PROJ_BODY)
        # 若雲端儲存的名稱開頭帶有 4 碼數字日期，自動去除
        import re as _re
        st.session_state.proj_body = _re.sub(r'^\d{4}', '', raw_proj) or DEFAULT_PROJ_BODY
        st.session_state.b_info     = cfg.get("briefing",     DEFAULT_BRIEF)
        st.session_state.ptl_time   = cfg.get("ptl_time",     DEFAULT_PTL_TIME)
        st.session_state.ptl_focus  = cfg.get("ptl_focus",    DEFAULT_PTL_FOCUS)
        st.session_state.cp_time    = cfg.get("cp_time",      DEFAULT_CP_TIME)
        st.session_state.cp_focus   = cfg.get("cp_focus",     DEFAULT_CP_FOCUS)
        st.session_state.brief_time = cfg.get("brief_time",   DEFAULT_BRIEF_TIME)
        st.session_state.brief_loc  = cfg.get("brief_loc",    DEFAULT_BRIEF_LOC)
        st.session_state.df_cmd = df_cmd_cl if not df_cmd_cl.empty else DEFAULT_CMD.copy()
        st.session_state.df_ptl = assign_ptl_groups(df_ptl_cl) if not df_ptl_cl.empty else assign_ptl_groups(DEFAULT_PTL.copy())
        st.session_state.df_cp  = assign_cp_groups(df_cp_cl)  if not df_cp_cl.empty  else assign_cp_groups(DEFAULT_CHECKPOINT.copy())
    else:
        # ── 預設值 ──
        st.session_state.p_time     = DEFAULT_TIME
        st.session_state.proj_body  = DEFAULT_PROJ_BODY
        st.session_state.b_info     = DEFAULT_BRIEF
        st.session_state.ptl_time   = DEFAULT_PTL_TIME
        st.session_state.ptl_focus  = DEFAULT_PTL_FOCUS
        st.session_state.cp_time    = DEFAULT_CP_TIME
        st.session_state.cp_focus   = DEFAULT_CP_FOCUS
        st.session_state.brief_time = DEFAULT_BRIEF_TIME
        st.session_state.brief_loc  = DEFAULT_BRIEF_LOC
        st.session_state.df_cmd = DEFAULT_CMD.copy()
        st.session_state.df_ptl = assign_ptl_groups(DEFAULT_PTL.copy())
        st.session_state.df_cp  = assign_cp_groups(DEFAULT_CHECKPOINT.copy())

    st.session_state.initialized = True

# ── 缺漏 key 補丁：版本升級後舊 session 可能缺少新 key，一律補齊預設值 ──
_DEFAULTS = {
    "p_time":      DEFAULT_TIME,
    "proj_body":  DEFAULT_PROJ_BODY,
    "b_info":      DEFAULT_BRIEF,
    "ptl_time":    DEFAULT_PTL_TIME,
    "ptl_focus":  DEFAULT_PTL_FOCUS,
    "cp_time":     DEFAULT_CP_TIME,
    "cp_focus":   DEFAULT_CP_FOCUS,
    "brief_time": DEFAULT_BRIEF_TIME,
    "brief_loc":  DEFAULT_BRIEF_LOC,
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

for _k, _fn in [
    ("df_cmd", lambda: DEFAULT_CMD.copy()),
    ("df_ptl", lambda: assign_ptl_groups(DEFAULT_PTL.copy())),
    ("df_cp",  lambda: assign_cp_groups(DEFAULT_CHECKPOINT.copy())),
]:
    if _k not in st.session_state:
        st.session_state[_k] = _fn()

# ══════════════════════════════════════════════════════════════════════════════
# 9. UI
# ══════════════════════════════════════════════════════════════════════════════
st.title("🚓 專案勤務規劃系統")

# ── 壹、基本資料列 ──
col_time, col_proj = st.columns([1, 2])
with col_time:
    p_time = st.text_input("勤務時間", value=st.session_state.p_time, key="ui_p_time")
    st.session_state.p_time = p_time

mmdd = extract_mmdd(p_time)
with col_proj:
    proj_body = st.text_area(f"專案名稱（日期代碼：{mmdd}，自動加在最前面）",
                             value=st.session_state.proj_body, height=80, key="ui_proj_body")
    st.session_state.proj_body = proj_body

# p_name = 日期代碼(4碼) + 純專案名稱，日期代碼不儲存在 proj_body 裡
p_name = f"{mmdd}{proj_body}"

# ── 貳、警力統計（全自動）──
live_stats = calculate_stats(st.session_state.df_cmd,
                             st.session_state.df_ptl,
                             st.session_state.df_cp)
st.subheader("貳、 警力統計（系統全自動精密統計）")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("督導組",     f"{live_stats['cmd']} 人")
c2.metric("機動攔檢組", f"{live_stats['ptl_机动']} 人")
c3.metric("場所臨檢組", f"{live_stats['ptl_场所']} 人")
c4.metric("偵訊組/民力", f"{live_stats['inv']}人 / {live_stats['civ']}人")
c5.metric("總計服勤警力", f"{live_stats['total']} 人")

# ── 參、指揮編組 ──
st.subheader("參、 督導及指揮編組")
edited_cmd = st.data_editor(st.session_state.df_cmd, num_rows="dynamic",
                            use_container_width=True, key="ed_cmd")
edited_cmd = edited_cmd.dropna(how="all").fillna("")
if not edited_cmd.equals(st.session_state.df_cmd):
    st.session_state.df_cmd = edited_cmd
    st.rerun()

# ── 勤務時間與重點設定（可編輯）──
st.subheader("勤務時間與重點設定")
col_ptl_f, col_cp_f = st.columns(2)
with col_ptl_f:
    ptl_time = st.text_input("【第一階段】機動攔查 勤務時間",
                             value=st.session_state.ptl_time, key="ui_ptl_time")
    st.session_state.ptl_time = ptl_time
    ptl_focus = st.text_area("【第一階段】機動攔查 勤務重點",
                             value=st.session_state.ptl_focus, height=100, key="ui_ptl_focus")
    st.session_state.ptl_focus = ptl_focus
with col_cp_f:
    cp_time = st.text_input("【第二階段】場所臨檢 勤務時間",
                            value=st.session_state.cp_time, key="ui_cp_time")
    st.session_state.cp_time = cp_time
    cp_focus = st.text_area("【第二階段】場所臨檢 勤務重點",
                            value=st.session_state.cp_focus, height=100, key="ui_cp_focus")
    st.session_state.cp_focus = cp_focus

# ── 陸、法令宣導 ──
b_info = st.text_area("陸、 工作重點與法令宣導",
                      value=st.session_state.b_info, height=150, key="ui_b_info")
st.session_state.b_info = b_info

# ── 簽到表集合資訊 ──
st.subheader("簽到表設定")
col_bt, col_bl = st.columns(2)
with col_bt:
    brief_time = st.text_input("簽到集合時間", value=st.session_state.brief_time, key="ui_brief_time")
    st.session_state.brief_time = brief_time
with col_bl:
    brief_loc = st.text_input("簽到集合地點", value=st.session_state.brief_loc, key="ui_brief_loc")
    st.session_state.brief_loc = brief_loc

# ── 任務編組（兩階段）──
st.subheader("勤務執行編組（兩階段）")
tab1, tab2 = st.tabs(["肆、【第一階段】機動攔查", "伍、【第二階段】場所臨檢"])

with tab1:
    edited_ptl = st.data_editor(st.session_state.df_ptl, num_rows="dynamic",
                                 use_container_width=True, key="ed_ptl")
    edited_ptl = edited_ptl.dropna(how="all").fillna("").reset_index(drop=True)
    if not edited_ptl.empty:
        re_ptl = assign_ptl_groups(edited_ptl)
        if not re_ptl.equals(st.session_state.df_ptl):
            st.session_state.df_ptl = re_ptl
            st.rerun()

with tab2:
    edited_cp = st.data_editor(st.session_state.df_cp, num_rows="dynamic",
                                use_container_width=True, key="ed_cp")
    edited_cp = edited_cp.dropna(how="all").fillna("").reset_index(drop=True)
    if not edited_cp.empty:
        re_cp = assign_cp_groups(edited_cp)
        if not re_cp.equals(st.session_state.df_cp):
            st.session_state.df_cp = re_cp
            st.rerun()

# ── 同步＋寄信 ──
st.markdown("---")
if st.button("💾 同步雲端並發送郵件", use_container_width=True, type="primary"):
    with st.spinner("⏳ 正在寫入雲端並寄送郵件，請稍候..."):
        ok, err = save_data(
            DEFAULT_UNIT, p_time, p_name, b_info,
            ptl_time, ptl_focus, cp_time, cp_focus, brief_time, brief_loc,
            st.session_state.df_cmd, st.session_state.df_ptl, st.session_state.df_cp,
            live_stats)
        if not ok:
            st.error(f"❌ 雲端同步失敗：{err}")
            st.stop()

        mail_ok, mail_err = send_email(
            DEFAULT_UNIT, p_name, p_time, b_info,
            ptl_time, ptl_focus, cp_time, cp_focus, brief_time, brief_loc,
            st.session_state.df_cmd, st.session_state.df_ptl, st.session_state.df_cp,
            live_stats)
        if mail_ok:
            st.success(f"✅ 雲端同步完成！專案：「{p_name}」，郵件已發送！")
            st.cache_data.clear()
            st.rerun()
        else:
            st.warning(f"⚠️ 雲端同步成功，但郵件發送失敗：{mail_err}")
