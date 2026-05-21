import streamlit as st

# --- 1. 頁面設定 (必須是全站第一個執行的 Streamlit 指令) ---
st.set_page_config(page_title="專案勤務規劃系統", layout="wide", page_icon="🚓")

# 呼叫側邊欄 (確保在 config 之後)
try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    st.sidebar.warning("找不到 menu.py，跳過側邊欄載入。")

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
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# ★ 預設改為 5 月 29 日全國同步專案
DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年5月29日 20時至24時"
DEFAULT_PROJ    = "0529「全國同步擴大取締酒後駕車及防制危險駕車」暨「擴大臨檢」專案"
DEFAULT_BRIEF   = "一、 落實三安：同仁執行盤查、臨檢及機動勤務過程中，應強化敵情觀念，提高危機意識，落實「人犯戒護安全、案件程序安全、執法者及民眾安全」。\n二、 臨檢合法性：依《警察職權行使法》第6條辦理。\n三、 攔停規範：依《警察職權行使法》第8條辦理。\n四、 全程蒐證：務必全程連續錄音或錄影。\n五、 異議處理：依《警察職權行使法》第29條製作紀錄。"

DEFAULT_PTL_FOCUS = "採取全面機動巡邏，針對酒駕熱點攔停盤查；攔獲疑似改裝噪音車，立即引導至「警政大樓廣場」交由環保局檢驗。\n(註：本階段機動攔查共6組警力。21時30分起，第1至第4組轉入第二階段執行擴大臨檢；第5、第6組全程獨留於路面，持續執行機動攔查至23時。)"
DEFAULT_CP_FOCUS = "由第一階段之第1至第4組機動警力，會合偵查隊專案人員，於21時20分前集結完畢，21時30分準時進入目標場所執行威力掃蕩。"

# ★ 導入 5 月 29 日真實督導組編制
DEFAULT_CMD = pd.DataFrame([
    {"項目": "指揮官", "通訊代號": "隆安 1 號", "任務目標": "勤務核定並重點機動督導", "負責人員": "分局長 施宇峰", "共同執行人員": "巡官 陳鵬翔、警員 張庭溱"},
    {"項目": "副指揮官", "通訊代號": "隆安 2 號", "任務目標": "襄助指揮、重點機動督導", "負責人員": "副分局長 何憶雯", "共同執行人員": "警務佐 曾威仁"},
    {"項目": "副指揮官", "通訊代號": "隆安 3 號", "任務目標": "襄助指揮、重點機動督導", "負責人員": "副分局長 蔡志明", "共同執行人員": "警員 陳明祥"},
    {"項目": "行政組", "通訊代號": "隆安 5 號", "任務目標": "督導擴大臨檢威力掃蕩第一臨檢組", "負責人員": "組長 周金柱", "共同執行人員": "巡官 蕭凱文"},
    {"項目": "督察組", "通訊代號": "隆安 6 號", "任務目標": "機動督導各單位勤務紀律", "負責人員": "組長 黃長旗", "共同執行人員": "警務員 陳冠彰"},
    {"項目": "保安民防組", "通訊代號": "隆安 9 號", "任務目標": "督導擴大臨檢威力掃蕩第二臨檢組", "負責人員": "組長 林良鍾", "共同執行人員": "警務員 曾盛鉉、警務佐 許榮裕、警務佐 劉俊德"},
    {"項目": "交通組", "通訊代號": "隆安 13號", "任務目標": "督導第一階段機動攔查", "負責人員": "組長 楊孟竟", "共同執行人員": "警務員 李峯甫、警務員 盧冠仁、巡官 郭勝隆、警員 吳享運"},
    {"項目": "聯絡組", "通訊代號": "隆安", "任務目標": "擔任通訊聯絡、指揮管制事宜", "負責人員": "勤指主任 蔡奇青", "共同執行人員": "執勤官 李文章、執勤員 黃文興"},
    {"項目": "偵訊組", "通訊代號": "隆安 10號", "任務目標": "負責按捺指紋、照相及移送", "負責人員": "偵查隊長 柯志賢", "共同執行人員": "偵查隊值日小隊"},
    {"項目": "聯合稽查站", "通訊代號": "隆安 1382", "任務目標": "配合環保局及監理站稽查車輛", "負責人員": "交通組巡官 郭勝隆", "共同執行人員": "環保局及監理站人員"}
])

# ★ 導入 5 月 29 日真實定點路檢編制
DEFAULT_PTL = pd.DataFrame([
    # 第1路檢組 (隆安51)
    {"組別": "第1路檢組", "無線電代號": "隆安51", "派遣單位": "聖亭所", "職別": "所長", "姓名": "鄭榮捷", "任務分工": "帶班兼管制", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號前\n（攔檢中興路往龍潭市區方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第1路檢組", "無線電代號": "隆安51", "派遣單位": "聖亭所", "職別": "警員",  "姓名": "詹宗澤", "任務分工": "指揮管制",   "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號前\n（攔檢中興路往龍潭市區方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第1路檢組", "無線電代號": "隆安51", "派遣單位": "龍潭所", "職別": "警員",  "姓名": "劉柏延", "任務分工": "攔檢盤查",   "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號前\n（攔檢中興路往龍潭市區方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第1路檢組", "無線電代號": "隆安51", "派遣單位": "龍潭所", "職別": "警員",  "姓名": "林宸緯", "任務分工": "攔檢盤查",   "攜行裝備": "小電腦、密錄器",                                "臨檢目標": "北龍路319號前\n（攔檢中興路往龍潭市區方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第1路檢組", "無線電代號": "隆安51", "派遣單位": "高平所", "職別": "警員",  "姓名": "黃丞穎", "任務分工": "警戒兼蒐證", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號前\n（攔檢中興路往龍潭市區方向）\n20時20分分局一樓集合出發臨檢"},
    # 第2路檢組 (隆安82)
    {"組別": "第2路檢組", "無線電代號": "隆安82", "派遣單位": "石門所", "職別": "副所長", "姓名": "林榮裕", "任務分工": "帶班兼管制", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號隊面前\n（攔檢龍潭市區往中興路方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第2路檢組", "無線電代號": "隆安82", "派遣單位": "石門所", "職別": "警員",   "姓名": "陳琦",   "任務分工": "指揮管制",   "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號隊面前\n（攔檢龍潭市區往中興路方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第2路檢組", "無線電代號": "隆安82", "派遣單位": "中興所", "職別": "巡佐",   "姓名": "蕭漢祥", "任務分工": "攔檢盤查",   "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號隊面前\n（攔檢龍潭市區往中興路方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第2路檢組", "無線電代號": "隆安82", "派遣單位": "中興所", "職別": "警員",   "姓名": "江益德", "任務分工": "攔檢盤查",   "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號隊面前\n（攔檢龍潭市區往中興路方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第2路檢組", "無線電代號": "隆安82", "派遣單位": "交通分隊", "職別": "小隊長", "姓名": "林振生", "任務分工": "攔檢盤查",  "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號隊面前\n（攔檢龍潭市區往中興路方向）\n20時20分分局一樓集合出發臨檢"},
    {"組別": "第2路檢組", "無線電代號": "隆安82", "派遣單位": "交通分隊", "職別": "警員",   "姓名": "吳沛軒", "任務分工": "警戒兼蒐證", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "臨檢目標": "北龍路319號隊面前\n（攔檢龍潭市區往中興路方向）\n20時20分分局一樓集合出發臨檢"},
])

# ★ 方案 B 核心：手動完美補正 5 月 29 日公文漏掉的「第 2 臨檢組」完整資料
DEFAULT_CHECKPOINT = pd.DataFrame([
    # 第1臨檢組 (隆安51) - 對齊 5/29 專案真實目標
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "聖亭所", "職別": "所長",   "姓名": "鄭榮捷", "任務分工": "帶班",                      "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nC. 丹陽泰養生館（中豐路281號）"},
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "聖亭所", "職別": "警員",   "姓名": "詹宗澤", "任務分工": "製作臨檢紀錄",               "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nC. 丹陽泰養生館（中豐路281號）"},
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "龍潭所", "職別": "警員",   "姓名": "劉柏延", "任務分工": "盤查兼蒐證",                 "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nC. 丹陽泰養生館（中豐路281號）"},
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "龍潭所", "職別": "警員",   "姓名": "林宸緯", "任務分工": "盤查兼蒐證",                 "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nC. 丹陽泰養生館（中豐路281號）"},
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "高平所", "職別": "警員",   "姓名": "黃丞穎", "任務分工": "大門警戒兼蒐證",             "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nC. 丹陽泰養生館（中豐路281號）"},
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "偵查隊", "職別": "偵查佐", "姓名": "賴享宏", "任務分工": "刑案偵防、社維法案件處理及移送", "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nC. 丹陽泰養生館（中豐路281號）"},
    {"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "偵查隊", "職別": "警員",   "姓名": "張峻銨", "任務分工": "刑案偵防、社維法案件處理及移送", "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nC. 丹陽泰養生館（中豐路281號）"},
    
    # 第2臨檢組 (隆安82) - 對齊第一階段路檢人員，手動完美回填
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "石門所", "職別": "副所長", "姓名": "林榮裕", "任務分工": "帶班",                      "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nF. 憤怒鳥網咖（中興路269號）\nG. 真情男女養生館（中興路387號）\nH. 萬紫千紅舒壓館（中興路491-3號）"},
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "石門所", "職別": "警員",   "姓名": "陳琦",   "任務分工": "製作臨檢紀錄",               "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nF. 憤怒鳥網咖（中興路269號）\nG. 真情男女養生館（中興路387號）\nH. 萬紫千紅舒壓館（中興路491-3號）"},
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "中興所", "職別": "巡佐",   "姓名": "蕭漢祥", "任務分工": "盤查兼蒐證",                 "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nF. 憤怒鳥網咖（中興路269號）\nG. 真情男女養生館（中興路387號）\nH. 萬紫千紅舒壓館（中興路491-3號）"},
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "中興所", "職別": "警員",   "姓名": "江益德", "任務分工": "盤查兼蒐證",                 "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nF. 憤怒鳥網咖（中興路269號）\nG. 真情男女養生館（中興路387號）\nH. 萬紫千紅舒壓館（中興路491-3號）"},
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "交通分隊", "職別": "小隊長", "姓名": "林振生", "任務分工": "盤查兼蒐證",                 "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nF. 憤怒鳥網咖（中興路269號）\nG. 真情男女養生館（中興路387號）\nH. 萬紫千紅舒壓館（中興路491-3號）"},
    {"組別": "第2臨檢組", "無線電代號": "隆安82", "派遣單位": "交通分隊", "職別": "警員",   "姓名": "吳沛軒", "任務分工": "大門警戒兼蒐證",             "臨檢目標場所": "A. 鉅大撞球館（中豐路558號）\nB. 台灣麻將協會（中豐路558之1號）\nF. 憤怒鳥網咖（中興路269號）\nG. 真情男女養生館（中興路387號）\nH. 萬紫千紅舒壓館（中興路491-3號）"}
])

# --- 2. 輔助函數 ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    font_paths = ["./kaiu.ttf", "kaiu.ttf", "/usr/share/fonts/truetype/custom/kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]
    for p in font_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def safe_str(val):
    if pd.isna(val) or val is None or str(val).strip().lower() == "nan": return ""
    return str(val)

@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except: return None

@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, None, "權限不足或未設定 Secrets"
        sh = client.open_by_key(SHEET_ID)
        try:
            ws_set = sh.worksheet("三合一_設定")
            df_set = pd.DataFrame(ws_set.get_all_records()).fillna("")
        except: df_set = None
        try:
            ws_cmd = sh.worksheet("三合一_指揮組")
            df_cmd = pd.DataFrame(ws_cmd.get_all_records()).fillna("")
        except: df_cmd = pd.DataFrame()
        try:
            ws_ptl = sh.worksheet("三合一_巡邏組")
            df_ptl = pd.DataFrame(ws_ptl.get_all_records()).fillna("")
        except: df_ptl = pd.DataFrame()
        try:
            ws_cp = sh.worksheet("三合一_擴大臨檢組")
            df_cp = pd.DataFrame(ws_cp.get_all_records()).fillna("")
        except: df_cp = None
        return df_set, df_cmd, df_ptl, df_cp, None
    except Exception as e: return None, None, None, None, str(e)

def save_data(unit, time_str, project, briefing, df_cmd, df_ptl, df_cp, stats, ptl_f, cp_f):
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)

        try: ws_set = sh.worksheet("三合一_設定")
        except: ws_set = sh.add_worksheet(title="三合一_設定", rows="50", cols="5")
        ws_set.clear()
        ws_set.update(range_name='A1', values=[
            ["Key", "Value"], ["unit_name", unit], ["plan_full_time", time_str], ["project_name", project],
            ["briefing_info", briefing], ["stats_cmd", str(stats['cmd'])], ["stats_ptl", str(stats['ptl'])],
            ["stats_inv", str(stats['inv'])], ["stats_civ", str(stats['civ'])], ["briefing_time", str(stats['b_time'])],
            ["briefing_loc", str(stats['b_loc'])], ["loc_1", str(stats['loc_1'])], ["loc_2", str(stats['loc_2'])],
            ["loc_3", str(stats['loc_3'])], ["ptl_focus", ptl_f], ["cp_focus", cp_f]
        ])

        for ws_name, df in [("三合一_指揮組", df_cmd), ("三合一_巡邏組", df_ptl), ("三合一_擴大臨檢組", df_cp)]:
            if df is None: continue
            try: ws = sh.worksheet(ws_name)
            except: ws = sh.add_worksheet(title=ws_name, rows="100", cols="20")
            ws.clear()
            clean_df = df.dropna(how='all').fillna("")
            if not clean_df.empty:
                ws.update(range_name='A1', values=[clean_df.columns.tolist()] + clean_df.astype(str).values.tolist())

        load_data.clear()
        return True
    except: return False

# --- 3. PDF 生成功能 ---
def generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, stats, ptl_f, cp_f):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=10*mm, rightMargin=10*mm, topMargin=12*mm, bottomMargin=15*mm)
    page_width = A4[0] - 20*mm
    story = []

    style_title    = ParagraphStyle('Title',    fontName=font, fontSize=18, leading=26, alignment=1, spaceAfter=8,    wordWrap='CJK')
    style_section  = ParagraphStyle('Section',  fontName=font, fontSize=14, leading=20, alignment=0, spaceAfter=2*mm, spaceBefore=4*mm, wordWrap='CJK')
    style_text     = ParagraphStyle('Text',     fontName=font, fontSize=14, leading=20, alignment=0, wordWrap='CJK')
    style_cell     = ParagraphStyle('Cell',     fontName=font, fontSize=12, leading=17, alignment=1, wordWrap='CJK')
    style_cell_left= ParagraphStyle('CellLeft', fontName=font, fontSize=12, leading=17, alignment=0, wordWrap='CJK')
    
    # 擴大臨檢長文字目標場所專用精細樣式（防截斷崩潰）
    style_cp_target = ParagraphStyle('CpTarget', fontName=font, fontSize=10, leading=14, alignment=0, wordWrap='CJK')

    def clean(t): return safe_str(t).replace("\n", "<br/>")

    story.append(Paragraph(f"<b>{unit}執行 {project} 勤務規劃表</b>", style_title))

    # 1. 基本資料
    story.append(Paragraph("<b>壹、 勤務基本資料</b>", style_section))
    date_str      = clean(time_str.split(" ")[0] if " " in time_str else "115年5月29日")
    time_str_only = clean(time_str.split(" ")[1] if " " in time_str else "20時至24時")
    data_basic = [
        [Paragraph("<b>實施日期</b>", style_cell), Paragraph("<b>勤務時間</b>", style_cell), Paragraph("<b>指揮官</b>", style_cell), Paragraph("<b>勤務編組</b>", style_cell), Paragraph("<b>聯合稽查站地點</b>", style_cell)],
        [Paragraph(date_str, style_cell), Paragraph(time_str_only, style_cell), Paragraph("分局長 施宇峰", style_cell), Paragraph("如任務編組表", style_cell), Paragraph("本局廣場", style_cell)]
    ]
    t_basic = Table(data_basic, colWidths=[page_width*0.18, page_width*0.2, page_width*0.18, page_width*0.18, page_width*0.26])
    t_basic.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_basic)

    # 2. 統計表
    story.append(Paragraph("<b>貳、 警力統計及地點統計</b>", style_section))
    data_stats = [
        [Paragraph("督導組", style_cell), Paragraph("攔臨組", style_cell), Paragraph("偵訊組", style_cell), Paragraph("小計", style_cell), Paragraph("民力", style_cell), Paragraph("總計", style_cell)],
        [Paragraph(str(stats['cmd']), style_cell), Paragraph(str(stats['ptl']), style_cell), Paragraph(str(stats['inv']), style_cell), Paragraph(str(stats['cmd']+stats['ptl']+stats['inv']), style_cell), Paragraph(str(stats['civ']), style_cell), Paragraph(str(stats['total']), style_cell)]
    ]
    t_stats = Table(data_stats, colWidths=[page_width*0.16]*6)
    t_stats.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_stats)

    # 3. 指揮組
    story.append(Paragraph("<b>參、 督導及其他任務編組表</b>", style_section))
    data_cmd = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["項目", "通訊代號", "任務目標", "負責人員", "共同人員"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([
            Paragraph(clean(r.get('項目','')), style_cell),
            Paragraph(clean(r.get('通訊代號','')), style_cell),
            Paragraph(clean(r.get('任務目標','')), style_cell_left),
            Paragraph(clean(r.get('負責人員','')), style_cell),
            Paragraph(clean(r.get('共同執行人員','')), style_cell)
        ])
    t_cmd = Table(data_cmd, colWidths=[page_width*0.14, page_width*0.16, page_width*0.3, page_width*0.2, page_width*0.2])
    t_cmd.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_cmd)

    # 4. 第一階段（定點路檢）
    story.append(Paragraph("<b>肆、【第一階段】定點路檢任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務重點：</b>{clean(ptl_f)}", style_text))

    ptl_headers = ["組別", "無線電\n代號", "派遣\n單位", "職別", "姓名", "任務分工", "攜行裝備", "臨檢目標"]
    col_w_ptl   = [page_width*0.10, page_width*0.10, page_width*0.09, page_width*0.07,
                   page_width*0.08, page_width*0.12, page_width*0.17, page_width*0.27]

    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ptl_headers]]
    rows_ptl = df_ptl.reset_index(drop=True)
    
    merge_groups = []
    prev_group = None
    grp_start_idx = 1

    for i, r in rows_ptl.iterrows():
        grp = safe_str(r.get('組別', ''))
        table_row_idx = i + 1
        if grp != prev_group:
            if prev_group is not None:
                merge_groups.append((grp_start_idx, table_row_idx - 1))
            prev_group = grp
            grp_start_idx = table_row_idx

        data_ptl.append([
            Paragraph(clean(r.get('組別','')),      style_cell),
            Paragraph(clean(r.get('無線電代號','')), style_cell),
            Paragraph(clean(r.get('派遣單位','')),   style_cell),
            Paragraph(clean(r.get('職別','')),       style_cell),
            Paragraph(clean(r.get('姓名','')),       style_cell),
            Paragraph(clean(r.get('任務分工','')),   style_cell),
            Paragraph(clean(r.get('攜行裝備','')),   style_cell_left),
            Paragraph(clean(r.get('臨檢目標','')),   style_cell_left),
        ])

    if prev_group is not None:
        merge_groups.append((grp_start_idx, len(rows_ptl)))

    t_ptl = Table(data_ptl, colWidths=col_w_ptl)
    ts_ptl = [
        ('FONTNAME',   (0,0), (-1,-1), font),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#f2f2f2')),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
    ]
    for (rs, re) in merge_groups:
        if re > rs:
            for col in [0, 1, 7]:
                ts_ptl.append(('SPAN', (col, rs), (col, re)))
    t_ptl.setStyle(TableStyle(ts_ptl))
    story.append(t_ptl)

    # 5. 第二階段（擴大臨檢）
    story.append(Paragraph("<b>伍、【第二階段】擴大臨檢任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務重點：</b>{clean(cp_f)}", style_text))
    
    if df_cp is not None and not df_cp.empty:
        cp_headers = ["組別", "無線電\n代號", "派遣\n單位", "職別", "姓名", "任務分工", "臨檢目標場所"]
        col_w_cp   = [page_width*0.10, page_width*0.10, page_width*0.09, page_width*0.07,
                      page_width*0.08, page_width*0.16, page_width*0.40]

        data_cp = [[Paragraph(f"<b>{h}</b>", style_cell) for h in cp_headers]]
        rows_cp = df_cp.reset_index(drop=True)
        
        cp_merge_groups = []
        cp_prev_group = None
        cp_grp_start = 1

        for i, r in rows_cp.iterrows():
            grp = safe_str(r.get('組別', ''))
            tbl_row = i + 1
            if grp != cp_prev_group:
                if cp_prev_group is not None:
                    cp_merge_groups.append((cp_grp_start, tbl_row - 1))
                cp_prev_group = grp
                cp_grp_start = tbl_row
                
            data_cp.append([
                Paragraph(clean(r.get('組別','')),        style_cell),
                Paragraph(clean(r.get('無線電代號','')),   style_cell),
                Paragraph(clean(r.get('派遣單位','')),     style_cell),
                Paragraph(clean(r.get('職別','')),         style_cell),
                Paragraph(clean(r.get('姓名','')),         style_cell),
                Paragraph(clean(r.get('任務分工','')),     style_cell_left),
                Paragraph(clean(r.get('臨檢目標場所','')), style_cp_target), 
            ])
            
        if cp_prev_group is not None:
            cp_merge_groups.append((cp_grp_start, len(rows_cp)))

        t_cp = Table(data_cp, colWidths=col_w_cp, splitByRow=True)
        ts_cp = [
            ('FONTNAME',   (0,0), (-1,-1), font),
            ('GRID',       (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#e6e6e6')),
            ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ]
        
        for (rs, re) in cp_merge_groups:
            if re > rs:
                for col in [0, 1, 6]:
                    ts_cp.append(('SPAN', (col, rs), (col, re)))
        t_cp.setStyle(TableStyle(ts_cp))
        story.append(t_cp)

    # 6. 宣導
    story.append(Paragraph("<b>陸、 工作重點與法令宣導</b>", style_section))
    for line in str(briefing).split('\n'):
        if line.strip(): story.append(Paragraph(clean(line), style_text))

    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 10)
        canvas.drawCentredString(A4[0]/2.0, 10*mm, f"-第{canvas.getPageNumber()}頁-")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    return buf.getvalue()

def generate_attendance_pdf(unit, project, time_str, stats):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 30*mm
    story = []

    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=26, alignment=1, spaceAfter=12, wordWrap='CJK')
    style_info  = ParagraphStyle('Info',  fontName=font, fontSize=14, leading=22, spaceAfter=1*mm, wordWrap='CJK')
    style_cell  = ParagraphStyle('Cell',  fontName=font, fontSize=14, leading=20, alignment=1, wordWrap='CJK')

    story.append(Paragraph(f"{unit}執行{project}簽到表", style_title))
    date_part = time_str.split(' ')[0] if ' ' in time_str else "115年5月29日"
    story.append(Paragraph(f"時間:{date_part}{stats['b_time']}", style_info))
    story.append(Paragraph(f"地點:{stats['b_loc']}召開", style_info))

    table_data = [[Paragraph("單位", style_cell), Paragraph("參加人員", style_cell), Paragraph("單位", style_cell), Paragraph("參加人員", style_cell)]]
    rows = [("交通組", "聖亭派出所"), ("督察組", "龍潭派出所"), ("行政組", "中興派出所"), ("保安民防組", "石門派出所"), ("勤務指揮中心", "高平派出所"), ("偵查隊", "三和派出所"), ("", "龍潭交通分隊")]
    for l, r in rows:
        table_data.append([Paragraph(l, style_cell) if l else "", "", Paragraph(r, style_cell) if r else "", ""])

    t = Table(table_data, colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3], rowHeights=[12*mm] + [24*mm]*len(rows))
    t.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('BACKGROUND',(0,0),(3,0),colors.whitesmoke)]))
    story.append(Spacer(1, 10*mm))
    story.append(t)

    doc.build(story)
    return buf.getvalue()

def send_report_email(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, stats, ptl_f, cp_f):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"], msg["Subject"] = sender, sender, f"勤務規劃與簽到表_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為最新版本勤務規劃表。", "plain"))

        pdf1 = generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, stats, ptl_f, cp_f)
        part1 = MIMEBase("application", "pdf")
        part1.set_payload(pdf1)
        encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(f'{unit}規劃表.pdf')}")
        msg.attach(part1)

        pdf2 = generate_attendance_pdf(unit, project, time_str, stats)
        part2 = MIMEBase("application", "pdf")
        part2.set_payload(pdf2)
        encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(f'{unit}簽到表.pdf')}")
        msg.attach(part2)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# ─────────────── Streamlit 介面 ───────────────
df_set, df_cmd, df_ptl, df_cp, err = load_data()

# 調整 5/29 專案真實警力統計初始數據
default_stats = {
    'cmd': 7, 'ptl': 16, 'inv': 3, 'civ': 0,
    'b_time': '18時30分至19時00分', 'b_loc': '分局二樓會議室',
    'loc_1': 8, 'loc_2': 2, 'loc_3': 0
}

PTL_COLS = ["組別", "無線電代號", "派遣單位", "職別", "姓名", "任務分工", "攜行裝備", "臨檢目標"]

if err or df_set is None:
    u, t, p, b = DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_PROJ, DEFAULT_BRIEF
    ed_cmd, ed_ptl, ed_cp = DEFAULT_CMD.copy(), DEFAULT_PTL.copy(), DEFAULT_CHECKPOINT.copy()
    p_ptl_focus, p_cp_focus = DEFAULT_PTL_FOCUS, DEFAULT_CP_FOCUS
else:
    d = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    u  = d.get("unit_name", DEFAULT_UNIT)
    t  = d.get("plan_full_time", DEFAULT_TIME)
    p  = d.get("project_name", DEFAULT_PROJ)
    b  = d.get("briefing_info", DEFAULT_BRIEF)
    p_ptl_focus = d.get("ptl_focus", DEFAULT_PTL_FOCUS)
    p_cp_focus  = d.get("cp_focus",  DEFAULT_CP_FOCUS)
    default_stats.update({
        'cmd':    int(d.get("stats_cmd",  7)),
        'ptl':    int(d.get("stats_ptl",  16)),
        'inv':    int(d.get("stats_inv",  3)),
        'civ':    int(d.get("stats_civ",  0)),
        'b_time': d.get("briefing_time",  "18時30分至19時00分"),
        'b_loc':  d.get("briefing_loc",   "分局二樓會議室"),
        'loc_1':  int(d.get("loc_1", 8)),
        'loc_2':  int(d.get("loc_2", 2)),
        'loc_3':  int(d.get("loc_3", 0)),
    })
    ed_cmd = df_cmd if not df_cmd.empty else DEFAULT_CMD.copy()
    if not df_ptl.empty and all(c in df_ptl.columns for c in PTL_COLS):
        ed_ptl = df_ptl[PTL_COLS]
    else:
        ed_ptl = DEFAULT_PTL.copy()
    if df_cp is not None and not df_cp.empty and all(c in df_cp.columns for c in CP_COLS):
        ed_cp = df_cp[CP_COLS]
    else:
        ed_cp = DEFAULT_CHECKPOINT.copy()

st.title("🚓 專案勤務規劃系統")
p_time = st.text_input("勤務時間", t)
p_name = st.text_input("專案名稱", p)

st.subheader("貳、 警力統計及地點統計")
col_s1, col_s2, col_s3, col_s4 = st.columns(4)
c_cmd = col_s1.number_input("督導組", value=default_stats['cmd'])
c_ptl = col_s2.number_input("攔臨組", value=default_stats['ptl'])
c_inv = col_s3.number_input("偵訊組", value=default_stats['inv'])
c_civ = col_s4.number_input("民力",   value=default_stats['civ'])
current_stats = {
    'cmd': c_cmd, 'ptl': c_ptl, 'inv': c_inv, 'civ': c_civ,
    'total': c_cmd+c_ptl+c_inv+c_civ,
    'b_time': default_stats['b_time'], 'b_loc': default_stats['b_loc'],
    'loc_1': default_stats['loc_1'], 'loc_2': default_stats['loc_2'], 'loc_3': default_stats['loc_3']
}

st.subheader("參、 指導編組與重點宣導")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True).dropna(how='all').fillna("")
b_info  = st.text_area("陸、 工作重點與法令宣導", b, height=150)

st.subheader("勤務執行編組 (兩階段)")
tab1, tab2 = st.tabs(["肆、【第一階段】定點路檢", "伍、【第二階段】擴大臨檢"])

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
        }
    ).dropna(how='all').fillna("").reset_index(drop=True)

with tab2:
    res_cp_focus = st.text_area("【伍】勤務重點", p_cp_focus, height=80, key="cp_focus_input")
    st.caption("💡 同一臨檢組的多名人員請填寫相同的「組別」與「無線電代號」，PDF 輸出時該欄會自動合併。")
    res_cp = st.data_editor(
        ed_cp,
        num_rows="dynamic",
        use_container_width=True,
        key="cp_ed",
        column_config={
            "組別":      st.column_config.TextColumn("組別",      width="small"),
            "無線電代號": st.column_config.TextColumn("無線電代號", width="small"),
            "派遣單位":   st.column_config.TextColumn("派遣單位",   width="small"),
            "職別":      st.column_config.TextColumn("職別",      width="small"),
            "姓名":      st.column_config.TextColumn("姓名",      width="small"),
            "任務分工":   st.column_config.TextColumn("任務分工",   width="medium"),
            "臨檢目標場所": st.column_config.TextColumn("臨檢目標場所", width="large"),
        }
    ).dropna(how='all').fillna("").reset_index(drop=True)

st.markdown("---")
pdf_plan = generate_pdf_from_data(u, p_name, p_time, b_info, res_cmd, res_ptl, res_cp, current_stats, res_ptl_focus, res_cp_focus)
st.download_button("📝 下載規劃表", data=pdf_plan, file_name=f"{u}規劃表.pdf", use_container_width=True)

if st.button("💾 同步雲端並發送郵件", use_container_width=True):
    if save_data(u, p_time, p_name, b_info, res_cmd, res_ptl, res_cp, current_stats, res_ptl_focus, res_cp_focus):
        ok, mail_err = send_report_email(u, p_name, p_time, b_info, res_cmd, res_ptl, res_cp, current_stats, res_ptl_focus, res_cp_focus)
        if ok: st.success("✅ 已同步並寄出！")
        else:  st.warning(f"⚠️ 同步成功但郵件失敗: {mail_err}")
    else:
        st.error("❌ 同步失敗，請確認 Google Sheets 權限設定。")
