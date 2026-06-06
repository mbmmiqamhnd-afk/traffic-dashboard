import streamlit as st

# --- 1. 頁面設定 (必須是全站第一個執行的 Streamlit 指令) ---
st.set_page_config(page_title="專案勤務規劃系統", layout="wide", page_icon="🚓")

# 呼叫側邊欄 (確保在 config 之後)
try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    pass

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
import re

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

DEFAULT_UNIT = "桃園市政府警察局龍潭分局"
# 【正本還原】對齊 0410 勤務時間與專案名稱
DEFAULT_TIME = "115年4月10日 19時至23時"
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

DEFAULT_PTL_FOCUS = "採取全面機動巡邏，針對酒駕熱點攔停盤查；攔獲疑似改裝噪音車，立即引導至「警政大樓廣場」交由環保局檢驗。"
DEFAULT_CP_FOCUS = "由第一階段之第1至第4組機動警力，會合偵查隊專案人員，於21時20分前集結完畢，21時30分準時進入目標場所執行威力掃蕩。"

# 參、督導及其他任務編組表（【正本還原】對齊 0410 正式公文同仁名單）
DEFAULT_CMD = pd.DataFrame([
    {"項目": "指揮官", "通訊代號": "隆安1號", "任務目標": "勤務核定並重點機動督導", "負責人員": "分局長 施宇峰", "共同執行人員": "巡官陳鵬翔、警員張庭溱"},
    {"項目": "副指揮官", "通訊代號": "隆安2號", "任務目標": "襄助指揮、重點機動督導", "負責人員": "副分局長 何憶雯", "共同執行人員": "警務佐曾威仁"},
    {"項目": "副指揮官", "通訊代號": "隆安3號", "任務目標": "襄助指揮、重點機動督導", "負責人員": "副分局長 蔡志明", "共同執行人員": "警員陳明祥"},
    {"項目": "行政組", "通訊代號": "隆安5號", "任務目標": "督導擴大臨檢威力掃蕩第一臨檢組", "負責人員": "組長 周金柱", "共同執行人員": "巡官蕭凱文"},
    {"項目": "督察組", "通訊代號": "隆安6號", "任務目標": "機動督導各單位勤務紀律", "負責人員": "組長黃長旗", "共同執行人員": "警務員 陳冠彰"},
    {"項目": "保安民防組", "通訊代號": "隆安9號", "任務目標": "督導擴大臨檢威力掃蕩第二臨檢組", "負責人員": "組長林良鍾", "共同執行人員": "警務員曾盛鉉、警務佐許榮裕、警務佐劉俊德"},
    {"項目": "交通組\n聯絡組", "通訊代號": "隆安 13號\n隆安", "任務目標": "督導第一階段機動欄查\n擔任通訊聯絡、指揮管制事宜", "負責人員": "組長 楊孟竟\n勤指中心主任蔡奇青", "共同執行人員": "巡官郭勝隆、警員吳享運、執勤官李文章、執勤員黃文興、警務員李峯甫、警務員盧冠仁"},
    {"項目": "偵訊組", "通訊代號": "隆安10號", "任務目標": "負責按捺指紋、照相及移送", "負責人員": "偵查隊隊長 柯志賢", "共同執行人員": "偵查隊值日小隊"},
    {"項目": "聯合稽查站", "通訊代號": "隆安1382", "任務目標": "配合環保局及監理站稽查車輛", "負責人員": "交通組巡官 郭勝隆", "共同執行人員": "環保局及監理站人員"}
])

# 肆、【正本還原】第一階段機動攔查（共6組實際服勤同仁、裝備、責任區與雨天備案）
DEFAULT_PTL = pd.DataFrame([
    {"單位": "聖亭所", "服勤人員": "副所長邱品淳\n警員劉憬霖\n警員謝伯昇", "任務分工": "帶班\n攔檢盤查\n警戒兼蒐證", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "中正路、北龍路周邊及治安要點機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "龍潭所", "服勤人員": "警員張家維\n警員王采蘋", "任務分工": "帶班兼蒐證\n攔檢盤查", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "北龍路、中豐路周邊及治安要點機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "中興所", "服勤人員": "所長董亦文\n警員羅俊傑", "任務分工": "帶班兼蒐證\n攔檢盤查", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "東龍路、中豐路沿線機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "石門所\n三和所", "服勤人員": "所長林育辰\n警員童霈晟", "任務分工": "帶班兼蒐證\n攔檢盤查", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "神龍路、文化路周邊及治安要點機動攔查。(20:00-21:30機動，後轉臨檢) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "石門所\n高平所", "服勤人員": "巡佐林偉政\n警員葉雲翔", "任務分工": "帶班兼蒐證\n攔檢盤查", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "中興路、龍新路沿線及治安要點機動攔查。(全程留守機動 20:00-23:00) *雨天備案:轄區治安要點巡邏。"},
    {"單位": "龍潭交警\n通分隊", "服勤人員": "警員林家豪\n警員吳沛軒", "任務分工": "帶班兼蒐證\n攔檢盤查", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "轄內易發生危駕路段、各聯外道路機動攔查。(全程留守機動 20:00-23:00) *雨天備案:轄區治安要點巡邏。"}
])

# 伍、【正本還原】第二階段擴大臨檢（精準對齊 0410 實際 2 組核心威力掃蕩編組）
DEFAULT_CHECKPOINT = pd.DataFrame([
    {"單位": "中興所\n龍潭所\n偵查隊", "服勤人員": "所長董亦文\n警員羅俊傑\n警員張家維\n警員王采蘋\n警員許家洋", "任務分工": "帶班\n製作臨檢紀錄\n盤查兼蒐證\n盤查兼蒐證\n刑案偵防、社維法案件查處", "臨檢目標場所": "A. 鉅大撞球館 (中豐路558號)\nB. 台灣麻將協會 (中豐路558之1號)\nC. 丹陽泰養生館 (中豐路281號)\nD. 溫馨汽車旅館 (中正路457號)\nE. 凱虹汽車旅館 (中正路506號)\n*(各員均需著防彈衣，攜帶槍彈、小電腦、密錄器)*"},
    {"單位": "石門所\n聖亭所\n三和所\n偵查隊", "服勤人員": "所長林育辰\n副所長邱品淳\n警員劉憬霖\n警員謝伯昇\n小隊長陳正育\n偵查佐鄧正斌", "任務分工": "帶班\n製作臨檢紀錄\n盤查兼蒐證\n大門警戒\n刑案偵防、社維法案件查處\n持DV全程蒐證", "臨檢目標場所": "A. 鉅大撞球館 (中豐路558號)\nB. 台灣麻將協會 (中豐路558之1號)\nF. 憤怒鳥網咖 (中興路269號)\nG. 真情男女養生館 (中興路387號)\nH. 萬紫千紅舒壓館 (中興路491-3號)\n*(各員均需著防彈衣，攜帶槍彈、小電腦、密錄器)*"}
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

def extract_mmdd(time_text):
    try:
        match = re.search(r'(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日', time_text)
        if match:
            month = int(match.group(2))
            day = int(match.group(3))
            return f"{month:02d}{day:02d}"
    except:
        pass
    return datetime.now().strftime("%m%d")

@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: 
        st.error("❌ 找不到 Secrets 中的 GCP 金鑰！")
        return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e: 
        st.error(f"❌ Google 授權失敗：{e}")
        return None

@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, None, "無法取得 Google 授權"
        sh = client.open_by_key(SHEET_ID)
        
        try:
            ws_set = sh.worksheet("三合一_設定")
            df_set = pd.DataFrame(ws_set.get_all_records()).fillna("")
        except: df_set = pd.DataFrame()
        
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
        except: df_cp = pd.DataFrame()
        
        return df_set, df_cmd, df_ptl, df_cp, None
    except Exception as e: 
        return None, None, None, None, str(e)

def save_data(unit, time_str, project, briefing, df_cmd, df_ptl, df_cp, stats, ptl_f, cp_f):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        
        try: ws_set = sh.worksheet("三合一_設定")
        except: ws_set = sh.add_worksheet(title="三合一_設定", rows="50", cols="5")
        ws_set.clear()
        
        ws_set.update([
            ["Key", "Value"], ["unit_name", unit], ["plan_full_time", time_str], ["project_name", project],
            ["briefing_info", briefing], ["stats_cmd", str(stats['cmd'])], ["stats_ptl", str(stats['ptl'])],
            ["stats_inv", str(stats['inv'])], ["stats_civ", str(stats['civ'])], ["briefing_time", str(stats['b_time'])],
            ["briefing_loc", str(stats['b_loc'])], ["ptl_focus", ptl_f], ["cp_focus", cp_f]
        ])
        
        for ws_name, df in [("三合一_指揮組", df_cmd), ("三合一_巡邏組", df_ptl), ("三合一_擴大臨檢組", df_cp)]:
            if df is None: continue
            try: ws = sh.worksheet(ws_name)
            except: ws = sh.add_worksheet(title=ws_name, rows="100", cols="20")
            ws.clear()
            
            clean_df = df.dropna(how='all').fillna("")
            if "組別" in clean_df.columns:
                clean_df = clean_df.drop(columns=["組別"])
                
            if not clean_df.empty:
                ws.update([clean_df.columns.tolist()] + clean_df.astype(str).values.tolist())
        
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"❌ 儲存至雲端時發生錯誤：{e}")
        return False

# --- 3. PDF 生成功能 ---
def generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, stats, ptl_f, cp_f):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=10*mm, rightMargin=10*mm, topMargin=12*mm, bottomMargin=15*mm)
    page_width = A4[0] - 20*mm
    story = []
    
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=26, alignment=1, spaceAfter=8, wordWrap='CJK')
    style_section = ParagraphStyle('Section', fontName=font, fontSize=14, leading=20, alignment=0, spaceAfter=2*mm, spaceBefore=4*mm, wordWrap='CJK')
    style_text = ParagraphStyle('Text', fontName=font, fontSize=14, leading=20, alignment=0, wordWrap='CJK')
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=20, alignment=1, wordWrap='CJK')
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=20, alignment=0, wordWrap='CJK')
    
    def clean(t): return safe_str(t).replace("\n", "<br/>")

    story.append(Paragraph(f"<b>{unit}執行 {project} 勤務規劃表</b>", style_title))
    
    story.append(Paragraph("<b>壹、 勤務基本資料</b>", style_section))
    date_str = clean(time_str.split(" ")[0] if " " in time_str else "115年4月10日")
    time_str_only = clean(time_str.split(" ")[1] if " " in time_str else "19時至23時")
    data_basic = [[Paragraph("<b>實施日期</b>", style_cell), Paragraph("<b>勤務時間</b>", style_cell), Paragraph("<b>指揮官</b>", style_cell), Paragraph("<b>勤務編組</b>", style_cell), Paragraph("<b>聯合稽查站地點</b>", style_cell)], 
                  [Paragraph(date_str, style_cell), Paragraph(time_str_only, style_cell), Paragraph("分局長 施宇峰", style_cell), Paragraph("如任務編組表", style_cell), Paragraph("分局廣場", style_cell)]]
    t_basic = Table(data_basic, colWidths=[page_width*0.18, page_width*0.2, page_width*0.18, page_width*0.18, page_width*0.26])
    t_basic.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_basic)
    
    story.append(Paragraph("<b>貳、 警力統計及地點統計</b>", style_section))
    data_stats = [[Paragraph("督導組", style_cell), Paragraph("攔臨組", style_cell), Paragraph("偵訊組", style_cell), Paragraph("小計", style_cell), Paragraph("民力", style_cell), Paragraph("總計", style_cell)], 
                  [Paragraph(str(stats['cmd']), style_cell), Paragraph(str(stats['ptl']), style_cell), Paragraph(str(stats['inv']), style_cell), Paragraph(str(stats['cmd']+stats['ptl']+stats['inv']), style_cell), Paragraph(str(stats['civ']), style_cell), Paragraph(str(stats['total']), style_cell)]]
    t_stats = Table(data_stats, colWidths=[page_width*0.16]*6)
    t_stats.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_stats)

    story.append(Paragraph("<b>參、 督導及其他任務編組表</b>", style_section))
    data_cmd = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["項目", "通訊代號", "任務目標", "負責人員", "共同人員"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([Paragraph(clean(r.get('項目')), style_cell), Paragraph(clean(r.get('通訊代號')), style_cell), Paragraph(clean(r.get('任務目標')), style_cell_left), Paragraph(clean(r.get('負責人員')), style_cell), Paragraph(clean(r.get('共同執行人員')), style_cell)])
    
    t_cmd = Table(data_cmd, colWidths=[page_width*0.12, page_width*0.14, page_width*0.28, page_width*0.26, page_width*0.2])
    t_cmd.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_cmd)
    
    story.append(Paragraph("<b>肆、【第一階段】機動攔查任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務重點：</b>{clean(ptl_f)}", style_text)) 
    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["組別", "單位", "服勤人員", "任務分工", "攜行裝備", "責任區"]]]
    for _, r in df_ptl.iterrows():
        data_ptl.append([Paragraph(clean(r.get('組別')), style_cell), Paragraph(clean(r.get('單位')), style_cell), Paragraph(clean(r.get('服勤人員')), style_cell_left), Paragraph(clean(r.get('任務分工')), style_cell_left), Paragraph(clean(r.get('攜行裝備')), style_cell_left), Paragraph(clean(r.get('巡邏與攔查責任區')), style_cell_left)])
    t_ptl = Table(data_ptl, colWidths=[page_width*0.11, page_width*0.12, page_width*0.20, page_width*0.21, page_width*0.16, page_width*0.20])
    t_ptl.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_ptl)

    story.append(Paragraph("<b>伍、【第二階段】擴大臨檢任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務重點：</b>{clean(cp_f)}", style_text))
    if df_cp is not None and not df_cp.empty:
        data_cp = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["組別", "單位", "服勤人員", "任務分工", "臨檢場所"]]]
        for _, r in df_cp.iterrows():
            data_cp.append([Paragraph(clean(r.get('組別')), style_cell), Paragraph(clean(r.get('單位')), style_cell), Paragraph(clean(r.get('服勤人員')), style_cell_left), Paragraph(clean(r.get('任務分工')), style_cell_left), Paragraph(clean(r.get('臨檢目標場所')), style_cell_left)])
        t_cp = Table(data_cp, colWidths=[page_width*0.11, page_width*0.12, page_width*0.24, page_width*0.24, page_width*0.29])
        t_cp.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e6e6e6')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
        story.append(t_cp)
    
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
    style_info = ParagraphStyle('Info', fontName=font, fontSize=14, leading=22, spaceAfter=1*mm, wordWrap='CJK')
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=20, alignment=1, wordWrap='CJK')
    
    story.append(Paragraph(f"{unit}執行{project}簽到表", style_title))
    date_part = time_str.split(' ')[0] if ' ' in time_str else "115年4月10日"
    story.append(Paragraph(f"時間:{date_part}{stats['b_time']}", style_info))
    story.append(Paragraph(f"地點:{stats['b_loc']}召開", style_info))
    
    table_data = [[Paragraph("單位", style_cell), Paragraph("參加人員", style_cell), Paragraph("單位", style_cell), Paragraph("參加人員", style_cell)]]
    rows = [("交通組", "聖亭派出所"), ("督察組", "龍潭派出所"), ("行政組", "中興派出所"), ("保安民防組", "石門派出所"), ("勤務指揮中心", "高平派出所"), ("偵查隊", "三和派出所"), ("", "龍潭交通分隊")]
    for l, r in rows: table_data.append([Paragraph(l, style_cell) if l else "", "", Paragraph(r, style_cell) if r else "", ""])
    
    t = Table(table_data, colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3], rowHeights=[12*mm] + [24*mm]*len(rows))
    t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), font), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BACKGROUND', (0,0), (3,0), colors.whitesmoke)]))
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
    except Exception as e:
        return False, str(e)

# --- 4. Streamlit 介面與 Session State 狀態記憶防護機制 ---
# 【核心修正】檢查並初始化記憶區，若為首次載入或雲端為空，100% 採用 0410 完整正本底稿
if "initialized" not in st.session_state:
    st.session_state.p_time = DEFAULT_TIME
    st.session_state.proj_body = DEFAULT_PROJ_BODY
    st.session_state.b_info = DEFAULT_BRIEF
    st.session_state.stats_data = {'cmd': 7, 'ptl': 16, 'inv': 3, 'civ': 0, 'b_time': '19時30分至20時00分', 'b_loc': '分局二樓會議室'}
    st.session_state.df_cmd = DEFAULT_CMD.copy()
    st.session_state.df_ptl = DEFAULT_PTL.copy()
    st.session_state.df_cp = DEFAULT_CHECKPOINT.copy()
    st.session_state.initialized = True

st.title("🚓 專案勤務規劃系統")

col_time, col_proj = st.columns([1, 2])
with col_time:
    p_time = st.text_input("勤務時間", value=st.session_state.p_time, key="input_p_time")
    st.session_state.p_time = p_time

# 自動連動擷取 4 碼數字 (此例會正確抓出 0410)
mmdd_code = extract_mmdd(p_time)

with col_proj:
    input_proj_body = st.text_input(f"專案名稱 (目前連動代碼: {mmdd_code})", value=st.session_state.proj_body, key="input_proj_body")
    st.session_state.proj_body = input_proj_body

p_name = f"{mmdd_code}{input_proj_body}"

st.subheader("貳、 警力統計及地點統計")
col_s1, col_s2, col_s3, col_s4 = st.columns(4)
c_cmd = col_s1.number_input("督導組", value=st.session_state.stats_data['cmd'])
c_ptl = col_s2.number_input("攔臨組", value=st.session_state.stats_data['ptl'])
c_inv = col_s3.number_input("偵訊組", value=st.session_state.stats_data['inv'])
c_civ = col_s4.number_input("民力", value=st.session_state.stats_data['civ'])

current_stats = {
    'cmd': c_cmd, 'ptl': c_ptl, 'inv': c_inv, 'civ': c_civ, 
    'total': c_cmd+c_ptl+c_inv+c_civ, 
    'b_time': st.session_state.stats_data['b_time'], 
    'b_loc': st.session_state.stats_data['b_loc']
}
st.session_state.stats_data.update(current_stats)

st.subheader("參、 指導編組與重點宣導")
# 透過記憶緩衝回寫，確保在網頁上編輯的名字不會因重繪而平白消失
res_cmd = st.data_editor(st.session_state.df_cmd, num_rows="dynamic", use_container_width=True, key="cmd_editor").dropna(how='all').fillna("")
st.session_state.df_cmd = res_cmd

b_info = st.text_area("陸、 工作重點與法令宣導", value=st.session_state.b_info, height=150, key="input_b_info")
st.session_state.b_info = b_info

st.subheader("勤務執行編組 (兩階段)")
tab1, tab2 = st.tabs(["肆、【第一階段】機動攔查", "伍、【第二階段】擴大臨檢"])

with tab1:
    res_ptl_raw = st.data_editor(st.session_state.df_ptl, num_rows="dynamic", use_container_width=True, key="ptl_ed").dropna(how='all').fillna("").reset_index(drop=True)
    st.session_state.df_ptl = res_ptl_raw
    if not res_ptl_raw.empty:
        res_ptl = res_ptl_raw.copy()
        res_ptl["組別"] = [f"第{i+1}巡邏組" for i in range(len(res_ptl))]
        res_ptl = res_ptl[["組別"] + [col for col in res_ptl.columns if col != "組別"]]
    else:
        res_ptl = res_ptl_raw

with tab2:
    res_cp_raw = st.data_editor(st.session_state.df_cp, num_rows="dynamic", use_container_width=True, key="cp_ed").dropna(how='all').fillna("").reset_index(drop=True)
    st.session_state.df_cp = res_cp_raw
    if not res_cp_raw.empty:
        res_cp = res_cp_raw.copy()
        res_cp["組別"] = [f"第{i+1}臨檢組" for i in range(len(res_cp))]
        res_cp = res_cp[["組別"] + [col for col in res_cp.columns if col != "組別"]]
    else:
        res_cp = res_cp_raw

st.markdown("---")

if st.button("💾 同步雲端並發送郵件", use_container_width=True):
    with st.spinner("⏳ 正在寫入雲端並寄送郵件，請稍候..."):
        if save_data(DEFAULT_UNIT, st.session_state.p_time, p_name, st.session_state.b_info, st.session_state.df_cmd, st.session_state.df_ptl, st.session_state.df_cp, st.session_state.stats_data, DEFAULT_PTL_FOCUS, DEFAULT_CP_FOCUS):
            ok, mail_err = send_report_email(DEFAULT_UNIT, p_name, st.session_state.p_time, st.session_state.b_info, st.session_state.df_cmd, res_ptl, res_cp, st.session_state.stats_data, DEFAULT_PTL_FOCUS, DEFAULT_CP_FOCUS)
            if ok: 
                st.success(f"✅ 資料已成功同步至雲端！專案名稱：「{p_name}」，公文 PDF 郵件已發送完成！")
                st.cache_data.clear()
                st.rerun()
            else: 
                st.error(f"❌ 雲端同步成功，但寄送郵件失敗！錯誤細節：{mail_err}")
