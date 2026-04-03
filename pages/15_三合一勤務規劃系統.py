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
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
import re

# --- 1. 頁面設定 ---
st.set_page_config(page_title="三合一專案勤務規劃系統", layout="wide", page_icon="🚓")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年4月10日 19時至23時"
DEFAULT_PROJ    = "0410取締酒後駕車暨監警環聯合稽查及擴大臨檢 三合一專案"
DEFAULT_BRIEF   = "一、 落實三安：同仁執行盤查、臨檢及機動勤務過程中，應強化敵情觀念，提高危機意識，落實「人犯戒護安全、案件程序安全、執法者及民眾安全」。\n二、 臨檢合法性：警察人員執行場所之臨檢，應限於已發生危害或依客觀合理判斷易生危害之場所，進行臨檢前應對當事人告以實施事由，便衣人員並應出示證件（依《警察職權行使法》第6條）。\n三、 攔停規範：機動攔檢對於已發生危害或易生危害之交通工具，得予以攔停；若有異常舉動而合理懷疑其將有危害行為時，得要求接受酒精濃度測試（依《警察職權行使法》第8條）。\n四、 全程蒐證：執行各項干涉、取締、處理糾紛及爭議性勤務（含噪音車引導與酒測），務必全程連續錄音或錄影。\n五、 異議處理：民眾對警察行使職權表示異議，認為無理由者得繼續執行，但經請求時應將異議之理由製作紀錄交付之（依《警察職權行使法》第29條）。"

# 根據專案更新指揮編組
DEFAULT_CMD = pd.DataFrame([
    {"項目": "指揮官", "通訊代號": "隆安 1 號", "任務目標": "勤務核定並重點機動督導", "負責人員": "分局長 施宇峰", "共同執行人員": "巡官 郭勝隆"},
    {"項目": "副指揮官", "通訊代號": "隆安 2 號", "任務目標": "襄助指揮、重點機動督導", "負責人員": "副分局長 何憶雯", "共同執行人員": "警務佐 曾威仁"},
    {"項目": "副指揮官", "通訊代號": "隆安 3 號", "任務目標": "襄助指揮、重點機動督導", "負責人員": "副分局長 蔡志明", "共同執行人員": "警員 陳明祥"},
    {"項目": "行政組", "通訊代號": "隆安 5 號", "任務目標": "督導第二階段臨檢勤務", "負責人員": "組長 周金柱", "共同執行人員": "巡官 蕭凱文"},
    {"項目": "督察組", "通訊代號": "隆安 6 號", "任務目標": "機動督導各單位勤務紀律", "負責人員": "督察組長 黃長旗", "共同執行人員": "警務員 陳冠彰"},
    {"項目": "交通組", "通訊代號": "隆安 13號", "任務目標": "機動督導第一階段攔檢組", "負責人員": "交通組長 楊孟竟", "共同執行人員": "警務員 盧冠仁"},
    {"項目": "聯絡組", "通訊代號": "隆安", "任務目標": "擔任通訊聯絡、指揮管制事宜", "負責人員": "勤指主任 蔡奇青", "共同執行人員": "執勤官、值勤員"},
    {"項目": "偵訊組", "通訊代號": "隆安 10號", "任務目標": "負責按捺指紋、照相及移送", "負責人員": "偵查隊 偵查佐", "共同執行人員": "在隊待命受理"},
    {"項目": "稽查站", "通訊代號": "聯合站", "任務目標": "警政大樓廣場聯合稽查警戒", "負責人員": "交通組派遣 2 名", "共同執行人員": "配合環保、監理"},
])

# 第一階段機動巡邏編組
DEFAULT_PTL = pd.DataFrame([
    {"單位": "聖亭所", "職別/姓名": "所長 鄭榮捷\n警員 詹宗澤", "任務分工": "帶班\n盤查兼警戒", "攜行裝備": "槍彈、無線電\n小電腦、密錄器", "巡邏與攔查責任區": "中正路、北龍路周邊及治安要點機動攔查。\n(20:00-21:30機動，後轉臨檢)"},
    {"單位": "龍潭所", "職別/姓名": "所長 孫祥愷\n警員 沈庭禾", "任務分工": "盤查兼警戒", "攜行裝備": "槍彈、無線電\n小電腦、密錄器", "巡邏與攔查責任區": "北龍路、中豐路周邊及治安要點機動攔查。\n(20:00-21:30機動，後轉臨檢)"},
    {"單位": "高平所", "職別/姓名": "警員 邱春松\n警員 唐銘聰", "任務分工": "盤查兼警戒", "攜行裝備": "槍彈、無線電\n小電腦、密錄器", "巡邏與攔查責任區": "東龍路、中豐路沿線機動攔查。\n(20:00-21:30機動，後轉臨檢)"},
    {"單位": "石門所", "職別/姓名": "巡佐 林偉政\n警員 鄒詠如", "任務分工": "帶班\n盤查兼警戒", "攜行裝備": "槍彈、無線電\n小電腦、密錄器", "巡邏與攔查責任區": "神龍路、文化路周邊及治安要點機動攔查。\n(20:00-21:30機動，後轉臨檢)"},
    {"單位": "中興所", "職別/姓名": "所長 董亦文\n警員 徐毓汶", "任務分工": "盤查兼警戒", "攜行裝備": "槍彈、無線電\n小電腦、密錄器", "巡邏與攔查責任區": "中興路、龍新路沿線及治安要點機動攔查。\n(全程留守機動 20:00-23:00)"},
    {"單位": "交分隊", "職別/姓名": "小隊長 林振生\n警員 吳沛軒", "任務分工": "盤查兼警戒", "攜行裝備": "槍彈、無線電\n小電腦、密錄器", "巡邏與攔查責任區": "轄內易發生危駕路段、各聯外道路機動攔查。\n(全程留守機動 20:00-23:00)"},
])

# 第二階段擴大臨檢編組
DEFAULT_CHECKPOINT = pd.DataFrame([
    {"單位": "聖亭所\n\n龍潭所\n\n偵查隊", "職別/姓名": "所長 鄭榮捷 警員 詹宗澤\n\n所長 孫祥愷 警員 沈庭禾\n\n偵查佐 賴享宏 警員 張峻銨", "任務分工": "帶班 製作臨檢紀錄\n\n盤查兼蒐證\n\n刑案偵防、社維法 著刑事背心、DV", "臨檢目標場所": "A. 鉅大撞球館 (中豐路558號)\nB. 台灣麻將協會 (中豐路558之1號)\nC. 丹陽泰養生館 (中豐路281號)\nD. 溫馨汽車旅館 (中正路457號)\nE. 凱虹汽車旅館 (中正路506號)\n\n*(各員均需著防彈衣，攜帶槍彈、小電腦、密錄器)*"},
    {"單位": "石門所\n\n高平所\n\n偵查隊", "職別/姓名": "巡佐 林偉政 警員 鄒詠如\n\n警員 邱春松 警員 唐銘聰\n\n偵查隊警員 2名", "任務分工": "帶班 製作臨檢紀錄\n\n大門警戒兼盤查\n\n刑案偵防、社維法", "臨檢目標場所": "F. 憤怒鳥網咖 (中興路269號)\nG. 真情男女養生館 (中興路387號)\nH. 萬紫千紅舒壓館 (中興路491-3號)\n\n*(各員均需著防彈衣，攜帶槍彈、小電腦、密錄器)*"},
])

# --- 2. 輔助函數 ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    font_paths = ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/custom/kaiu.ttf"]
    for p in font_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def parse_meeting_time(time_str):
    try:
        match = re.search(r"(\d+)至", time_str)
        if match:
            start_hour = int(match.group(1))
            end_hour = start_hour + 1
            return f"{start_hour}時30分至{end_hour}時00分"
    except:
        pass
    return "19時30分至20時00分"

def safe_str(val):
    if pd.isna(val) or val is None or str(val).strip().lower() == "nan":
        return ""
    return str(val)

@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        
        # 改為讀取 三合一 專屬工作表，並加入例外處理以防工作表尚未建立
        try:
            ws_set = sh.worksheet("三合一_設定")
            df_set = pd.DataFrame(ws_set.get_all_records()).fillna("")
        except:
            df_set = None
            
        try:
            ws_cmd = sh.worksheet("三合一_指揮組")
            df_cmd = pd.DataFrame(ws_cmd.get_all_records()).fillna("")
        except:
            df_cmd = pd.DataFrame()
            
        try:
            ws_ptl = sh.worksheet("三合一_巡邏組")
            df_ptl = pd.DataFrame(ws_ptl.get_all_records()).fillna("")
        except:
            df_ptl = pd.DataFrame()
            
        try:
            ws_cp = sh.worksheet("三合一_擴大臨檢組")
            df_cp = pd.DataFrame(ws_cp.get_all_records()).fillna("")
        except:
            df_cp = None
            
        return df_set, df_cmd, df_ptl, df_cp, None
    except Exception as e: return None, None, None, None, str(e)

def save_data(unit, time_str, project, briefing, df_cmd, df_ptl, df_cp, stats):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        
        # 儲存至 三合一 專屬工作表
        try:
            ws_set = sh.worksheet("三合一_設定")
        except:
            ws_set = sh.add_worksheet(title="三合一_設定", rows="50", cols="5")
            
        ws_set.clear()
        ws_set.update([["Key", "Value"], 
                       ["unit_name", unit], 
                       ["plan_full_time", time_str], 
                       ["project_name", project], 
                       ["briefing_info", briefing],
                       ["stats_cmd", str(stats['cmd'])],
                       ["stats_ptl", str(stats['ptl'])],
                       ["stats_inv", str(stats['inv'])],
                       ["stats_civ", str(stats['civ'])],
                       ["briefing_time", str(stats['b_time'])],
                       ["briefing_loc", str(stats['b_loc'])],
                       ["loc_1", str(stats['loc_1'])],
                       ["loc_2", str(stats['loc_2'])],
                       ["loc_3", str(stats['loc_3'])]])
        
        # 建立或更新其餘專屬工作表
        for ws_name, df in [("三合一_指揮組", df_cmd), ("三合一_巡邏組", df_ptl), ("三合一_擴大臨檢組", df_cp)]:
            try:
                ws = sh.worksheet(ws_name)
            except:
                ws = sh.add_worksheet(title=ws_name, rows="100", cols="20")
            ws.clear()
            df_cleaned = df.dropna(how='all').fillna("")
            if not df_cleaned.empty:
                ws.update([df_cleaned.columns.tolist()] + df_cleaned.values.tolist())
        load_data.clear()
        return True
    except: return False

# --- 3. PDF 生成功能 ---
def generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, stats):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []
    
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=24, alignment=1, spaceAfter=10)
    style_section = ParagraphStyle('Section', fontName=font, fontSize=15, leading=20, alignment=0, spaceAfter=3*mm, spaceBefore=4*mm)
    style_text = ParagraphStyle('Text', fontName=font, fontSize=12, leading=18, alignment=0)
    
    # 凸排段落樣式
    style_briefing = ParagraphStyle('Briefing', fontName=font, fontSize=12, leading=18, alignment=0, leftIndent=32, firstLineIndent=-32, spaceAfter=2*mm)
    
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=18, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=18, alignment=0)
    
    def clean(t): return safe_str(t).replace("\n", "<br/>")

    # 大標題
    story.append(Paragraph(f"<b>{unit}執行 {project} 勤務規劃表</b>", style_title))
    
    # 壹、 勤務基本資料
    story.append(Paragraph("<b>壹、 勤務基本資料</b>", style_section))
    
    date_str = clean(time_str.split(" ")[0] if " " in time_str else "115年4月10日")
    time_str_only = clean(time_str.split(" ")[1] if " " in time_str else "19時至23時")
    
    data_basic = [
        [Paragraph("<b>實施日期</b>", style_cell), Paragraph("<b>勤務時間</b>", style_cell), Paragraph("<b>指揮官</b>", style_cell), Paragraph("<b>勤務編組</b>", style_cell), Paragraph("<b>聯合稽查站地點</b>", style_cell)],
        # 使用 <nobr> 標籤強制 PDF 不換行
        [Paragraph(f"<nobr>{date_str}</nobr>", style_cell),
         Paragraph(f"<nobr>{time_str_only}</nobr>", style_cell),
         Paragraph("分局長 施宇峰", style_cell),
         Paragraph("如各階段任務編組表", style_cell),
         Paragraph("龍潭區警政聯合辦公大樓廣場", style_cell)]
    ]
    # 微調欄寬比例 (把日期與時間欄寬從 0.15 加大到 0.2，避免字體被擠壓)
    t_basic = Table(data_basic, colWidths=[page_width*0.2, page_width*0.2, page_width*0.18, page_width*0.18, page_width*0.24])
    t_basic.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_basic)
    
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph("<b>勤務時程分配：</b>", style_text))
    story.append(Paragraph("19:00 - 19:30：各單位由駐地往分局移動路程。<br/>19:30 - 20:00：勤前教育（地點：本分局2樓會議室）。<br/>20:00 - 23:00：第一階段（機動攔查與聯合稽查）。<br/>21:30 - 23:00：第二階段（擴大臨檢威力掃蕩）。", style_text))
    
    # 貳、 警力使用統計表與地點統計
    story.append(Paragraph("<b>貳、 警力使用統計表及勤前教育、地點統計</b>", style_section))
    data_stats = [
        [Paragraph("<b>單位</b>", style_cell), Paragraph("<b>業務及督導組</b>", style_cell), Paragraph("<b>攔檢與臨檢組</b>", style_cell), Paragraph("<b>偵訊組</b>", style_cell), Paragraph("<b>小計</b>", style_cell), Paragraph("<b>民力</b>", style_cell), Paragraph("<b>總計</b>", style_cell)],
        [Paragraph("龍潭分局", style_cell), Paragraph(str(stats['cmd']), style_cell), Paragraph(str(stats['ptl']), style_cell), Paragraph(str(stats['inv']), style_cell), Paragraph(str(stats['cmd']+stats['ptl']+stats['inv']), style_cell), Paragraph(str(stats['civ']), style_cell), Paragraph(str(stats['total']), style_cell)]
    ]
    t_stats = Table(data_stats, colWidths=[page_width*0.2, page_width*0.16, page_width*0.16, page_width*0.12, page_width*0.12, page_width*0.12, page_width*0.12])
    t_stats.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_stats)
    
    story.append(Spacer(1, 3*mm))
    
    data_loc = [
        [Paragraph("<b>勤前教育時間</b>", style_cell), Paragraph("<b>勤前教育地點</b>", style_cell), Paragraph("<b>臨檢點</b>", style_cell), Paragraph("<b>盤查點</b>", style_cell), Paragraph("<b>聯外道路</b>", style_cell)],
        [Paragraph(clean(stats['b_time']), style_cell), Paragraph(clean(stats['b_loc']), style_cell), Paragraph(f"{stats['loc_1']}處", style_cell), Paragraph(f"{stats['loc_2']}處", style_cell), Paragraph(f"{stats['loc_3']}處", style_cell)]
    ]
    t_loc = Table(data_loc, colWidths=[page_width*0.25, page_width*0.27, page_width*0.16, page_width*0.16, page_width*0.16])
    t_loc.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_loc)

    # 參、 督導及其他任務編組表
    story.append(Paragraph("<b>參、 督導及其他任務編組表 (19:00 - 23:00)</b>", style_section))
    data_cmd = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["項目", "通訊代號", "任務目標", "負責人員", "共同執行人員"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([
            Paragraph(clean(r.get('項目')), style_cell), 
            Paragraph(clean(r.get('通訊代號')), style_cell),
            Paragraph(clean(r.get('任務目標')), style_cell_left), 
            Paragraph(clean(r.get('負責人員')), style_cell),
            Paragraph(clean(r.get('共同執行人員')), style_cell)
        ])
    t_cmd = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.15, page_width*0.3, page_width*0.2, page_width*0.2])
    t_cmd.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_cmd)
    
    # 肆、 第一階段
    story.append(Paragraph("<b>肆、【第一階段 20:00 - 23:00】機動攔查任務編組</b>", style_section))
    story.append(Paragraph("<b>勤務重點：</b>取消定點路檢，採取全面機動巡邏。針對酒駕熱點攔停盤查；攔獲疑似改裝噪音車，立即引導至「警政大樓廣場」交由環保局檢驗。<br/>（註：本階段機動攔查共6組警力。21時30分起，第1至第4組轉入第二階段執行擴大臨檢；第5、第6組全程獨留於路面，持續執行機動攔查至23時。）", style_text))
    story.append(Spacer(1, 2*mm))
    
    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["組別", "單位", "職別/姓名", "任務分工", "攜行裝備", "巡邏與攔查責任區"]]]
    for _, r in df_ptl.iterrows():
        data_ptl.append([
            Paragraph(clean(r.get('組別')), style_cell), 
            Paragraph(clean(r.get('單位')), style_cell), 
            Paragraph(clean(r.get('職別/姓名')), style_cell_left), 
            Paragraph(clean(r.get('任務分工')), style_cell_left), 
            Paragraph(clean(r.get('攜行裝備')), style_cell_left), 
            Paragraph(clean(r.get('巡邏與攔查責任區')), style_cell_left)
        ])
    t_ptl = Table(data_ptl, colWidths=[page_width*0.12, page_width*0.1, page_width*0.2, page_width*0.15, page_width*0.18, page_width*0.25])
    t_ptl.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_ptl)

    # 伍、 第二階段
    story.append(Paragraph("<b>伍、【第二階段 21:30 - 23:00】擴大臨檢任務編組</b>", style_section))
    story.append(Paragraph("<b>勤務重點：</b>由第一階段之第1至第4組機動警力，會合偵查隊專案人員，於21時20分前集結完畢，21時30分準時進入目標場所執行威力掃蕩。", style_text))
    story.append(Spacer(1, 2*mm))

    data_cp = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["組別", "單位", "職別/姓名", "任務分工", "臨檢目標場所"]]]
    for _, r in df_cp.iterrows():
        data_cp.append([
            Paragraph(clean(r.get('組別')), style_cell), 
            Paragraph(clean(r.get('單位')), style_cell), 
            Paragraph(clean(r.get('職別/姓名')), style_cell_left), 
            Paragraph(clean(r.get('任務分工')), style_cell_left), 
            Paragraph(clean(r.get('臨檢目標場所')), style_cell_left)
        ])
    t_cp = Table(data_cp, colWidths=[page_width*0.12, page_width*0.12, page_width*0.25, page_width*0.2, page_width*0.31])
    t_cp.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e6e6e6')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_cp)
    
    story.append(Spacer(1, 1*mm))
    story.append(Paragraph("備註：臨檢完畢後若有剩餘時間，於各所轄內治安熱點、涉毒區段加強巡守，以防制刑案發生。", style_text))

    # 陸、 工作重點與法令宣導
    story.append(Paragraph("<b>陸、 工作重點與法令宣導</b>", style_section))
    for line in str(briefing).split('\n'):
        if line.strip():
            story.append(Paragraph(f"{clean(line)}", style_briefing))

    doc.build(story)
    return buf.getvalue()

# 簽到表
def generate_attendance_pdf(unit, project, time_str, briefing):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 30*mm
    story = []
    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, leading=22, alignment=1, spaceAfter=8)
    style_top_info = ParagraphStyle('TopInfo', fontName=font, fontSize=12, leading=18, alignment=0)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=24, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=24, alignment=0) 
    style_note = ParagraphStyle('Note', fontName=font, fontSize=12, leading=15, alignment=0)
    
    story.append(Paragraph(f"{unit}執行{project}勤前教育會議人員簽到表", style_title))
    meeting_range = parse_meeting_time(time_str)
    date_part = time_str.split(' ')[0] if ' ' in time_str else "115年4月10日"
    story.append(Paragraph(f"時間：{date_part} {meeting_range}", style_top_info))
    story.append(Paragraph(f"地點：本分局2樓會議室", style_top_info))
    story.append(Spacer(1, 3*mm))
    table_data = [[Paragraph("分局長：", style_cell_left), "", Paragraph("上級督導：", style_cell_left), ""],
                  [Paragraph("副分局長：", style_cell_left), "", "", ""],
                  [Paragraph("單位", style_cell), Paragraph("參加人員", style_cell), Paragraph("單位", style_cell), Paragraph("參加人員", style_cell)]]
    rows = [("交通組", "中興派出所"), ("勤務指揮中心", "石門派出所"), ("督察組", "高平派出所"), ("偵查隊", "三和派出所"), ("龍潭派出所", "龍潭交通分隊"), ("聖亭派出所", "")]
    for l, r in rows: table_data.append([Paragraph(l, style_cell), "", Paragraph(r, style_cell), ""])
    t = Table(table_data, colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3], rowHeights=[18*mm, 18*mm, 10*mm] + [26*mm]*len(rows))
    t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), font), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                           ('ALIGN', (0,0), (0,0), 'LEFT'), ('ALIGN', (2,0), (2,0), 'LEFT'), ('ALIGN', (0,1), (0,1), 'LEFT'), ('SPAN', (0,1), (3,1)),
                           ('BACKGROUND', (0,2), (0,2), colors.whitesmoke), ('BACKGROUND', (2,2), (2,2), colors.whitesmoke)]))
    story.append(t)
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph("備註：請將行動電話調整為靜音。", style_note))
    doc.build(story)
    return buf.getvalue()

# 寄信功能
def send_report_email(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, stats):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        msg["Subject"] = f"{unit}執行{project}勤務規劃與簽到表_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為勤務規劃表與人員簽到表 PDF。", "plain", "utf-8"))
        
        pdf_plan_name = f"{unit}執行{project}勤務規劃表.pdf"
        pdf_attendance_name = f"{unit}執行{project}勤前教育會議人員簽到表.pdf"

        pdf1 = generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, stats)
        part1 = MIMEBase("application", "pdf")
        part1.set_payload(pdf1)
        encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(pdf_plan_name)}")
        msg.attach(part1)
        
        pdf2 = generate_attendance_pdf(unit, project, time_str, briefing)
        part2 = MIMEBase("application", "pdf")
        part2.set_payload(pdf2)
        encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(pdf_attendance_name)}")
        msg.attach(part2)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)


# --- 主程式介面 ---
df_set, df_cmd, df_ptl, df_cp, err = load_data()

# 讀取預設警力數值及勤教/地點統計
default_stats = {'cmd': 6, 'ptl': 31, 'inv': 2, 'civ': 0, 'b_time': '19時30分', 'b_loc': '本分局2樓會議室', 'loc_1': 8, 'loc_2': 6, 'loc_3': 0}

if err or df_set is None:
    u, t, p, b = DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_PROJ, DEFAULT_BRIEF
    ed_cmd, ed_ptl, ed_cp = DEFAULT_CMD.copy(), DEFAULT_PTL.copy(), DEFAULT_CHECKPOINT.copy()
else:
    d = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    u, t, p = d.get("unit_name", DEFAULT_UNIT), d.get("plan_full_time", DEFAULT_TIME), d.get("project_name", DEFAULT_PROJ)
    
    # 強制更新宣導內容
    b = d.get("briefing_info", DEFAULT_BRIEF)
    if "落實三安" not in str(b):
        b = DEFAULT_BRIEF
        
    default_stats['cmd'] = int(d.get("stats_cmd", 6))
    default_stats['ptl'] = int(d.get("stats_ptl", 31))
    default_stats['inv'] = int(d.get("stats_inv", 2))
    default_stats['civ'] = int(d.get("stats_civ", 0))
    default_stats['b_time'] = d.get("briefing_time", "19時30分")
    default_stats['b_loc'] = d.get("briefing_loc", "本分局2樓會議室")
    default_stats['loc_1'] = int(d.get("loc_1", 8))
    default_stats['loc_2'] = int(d.get("loc_2", 6))
    default_stats['loc_3'] = int(d.get("loc_3", 0))
    
    # 格式檢查
    ed_cmd = df_cmd if not df_cmd.empty and "項目" in df_cmd.columns else DEFAULT_CMD.copy()
    
    if not df_ptl.empty and "單位" in df_ptl.columns:
        ed_ptl = df_ptl.drop(columns=["組別"]) if "組別" in df_ptl.columns else df_ptl
    else:
        ed_ptl = DEFAULT_PTL.copy()
        
    if df_cp is not None and not df_cp.empty and "臨檢目標場所" in df_cp.columns:
        ed_cp = df_cp.drop(columns=["組別"]) if "組別" in df_cp.columns else df_cp
    else:
        ed_cp = DEFAULT_CHECKPOINT.copy()

st.title("🚓 三合一專案勤務規劃系統")
c1, c2 = st.columns(2)

# 先渲染右側的勤務時間輸入框，獲取最新的時間數值
p_time = c2.text_input("勤務時間", t)

# 擷取月份與日期 (忽略年)
date_match = re.search(r"(\d+)月(\d+)日", p_time)
if date_match:
    m = date_match.group(1).zfill(2) # 自動補零至2碼
    d = date_match.group(2).zfill(2) # 自動補零至2碼
    new_prefix = f"{m}{d}"
    
    # 將專案名稱最前面的數字替換為 4 碼 (MMDD)
    p = re.sub(r"^\d+", new_prefix, p)

# 再渲染左側的專案名稱輸入框，帶入更新後的預設值 p
p_name = c1.text_input("專案名稱", p)

# 警力加總與地點統計區塊
st.subheader("貳、 警力使用與地點統計 (手動微調區)")
col_s1, col_s2, col_s3, col_s4 = st.columns(4)
c_cmd = col_s1.number_input("業務及督導組 (人)", value=default_stats['cmd'], min_value=0)
c_ptl = col_s2.number_input("攔檢與臨檢組 (人)", value=default_stats['ptl'], min_value=0)
c_inv = col_s3.number_input("偵訊組 (人)", value=default_stats['inv'], min_value=0)
c_civ = col_s4.number_input("民力 (人)", value=default_stats['civ'], min_value=0)
c_total = c_cmd + c_ptl + c_inv + c_civ

st.write("📍 **勤前教育與地點統計**")
col_b1, col_b2 = st.columns(2)
b_time = col_b1.text_input("勤前教育時間", default_stats['b_time'])
b_loc = col_b2.text_input("勤前教育地點", default_stats['b_loc'])

col_l1, col_l2, col_l3 = st.columns(3)
loc_1 = col_l1.number_input("臨檢點 (處)", value=default_stats['loc_1'], min_value=0)
loc_2 = col_l2.number_input("盤查點 (處)", value=default_stats['loc_2'], min_value=0)
loc_3 = col_l3.number_input("聯外道路 (處)", value=default_stats['loc_3'], min_value=0)

current_stats = {
    'cmd': c_cmd, 'ptl': c_ptl, 'inv': c_inv, 'civ': c_civ, 'total': c_total,
    'b_time': b_time, 'b_loc': b_loc, 'loc_1': loc_1, 'loc_2': loc_2, 'loc_3': loc_3
}

st.subheader("參、 督導及其他任務編組表")
res_cmd_raw = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True)
res_cmd = res_cmd_raw.dropna(how='all').fillna("")

# 工作重點與法令宣導
b_info = st.text_area("陸、 工作重點與法令宣導", b, height=200)

st.subheader("勤務執行編組 (兩階段)")
tab1, tab2 = st.tabs(["肆、【第一階段】機動攔查", "伍、【第二階段】擴大臨檢威力掃蕩"])

with tab1:
    st.caption("💡 取消定點路檢，採取全面機動巡邏。（系統會自動編號為「第1巡邏組」...）")
    res_ptl_raw = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True, key="ptl_editor")
    res_ptl = res_ptl_raw.dropna(how='all').fillna("").reset_index(drop=True)
    if not res_ptl.empty:
        res_ptl.insert(0, "組別", [f"第{i+1}巡邏組" for i in range(len(res_ptl))])

with tab2:
    st.caption("💡 針對治安場所執行威力掃蕩。（系統會自動編號為「第1臨檢組」...）")
    res_cp_raw = st.data_editor(ed_cp, num_rows="dynamic", use_container_width=True, key="cp_editor")
    res_cp = res_cp_raw.dropna(how='all').fillna("").reset_index(drop=True)
    if not res_cp.empty:
        res_cp.insert(0, "組別", [f"第{i+1}臨檢組" for i in range(len(res_cp))])

def get_html():
    style = "<style>body{font-family:'標楷體';padding:10px;line-height:1.5;} th,td{border:1px solid black;padding:6px;font-size:14pt;text-align:center;} .middle-block{font-size:12pt;margin:15px 0 15px 0;text-align:left;} h3, h4 {margin-top: 25px;}</style>"
    
    html = f"<html>{style}<body><h2 style='text-align:center'>{u}執行<br>{p_name}<br>勤務規劃表</h2>"
    
    html += "<h4>壹、 勤務基本資料</h4><table><tr><th>實施日期</th><th>勤務時間</th><th>指揮官</th><th>勤務編組</th><th>聯合稽查站地點</th></tr>"
    # 加入 style='white-space: nowrap;' 強制網頁不換行
    html += f"<tr><td style='white-space: nowrap;'>{p_time.split(' ')[0]}</td><td style='white-space: nowrap;'>{p_time.split(' ')[1] if ' ' in p_time else '19時至23時'}</td><td>分局長 施宇峰</td><td>如各階段任務編組表</td><td>龍潭區警政聯合辦公大樓廣場</td></tr></table>"
    html += "<div class='middle-block'><b>勤務時程分配：</b><br>19:00 - 19:30：各單位由駐地往分局移動路程。<br>19:30 - 20:00：勤前教育（地點：本分局2樓會議室）。<br>20:00 - 23:00：第一階段（機動攔查與聯合稽查）。<br>21:30 - 23:00：第二階段（擴大臨檢威力掃蕩）。</div>"
    
    html += "<h4>貳、 警力使用統計表及勤前教育、地點統計</h4><table><tr><th>單位</th><th>業務及督導組</th><th>攔檢與臨檢組</th><th>偵訊組</th><th>小計</th><th>民力</th><th>總計</th></tr>"
    html += f"<tr><td>龍潭分局</td><td>{c_cmd}</td><td>{c_ptl}</td><td>{c_inv}</td><td>{c_cmd+c_ptl+c_inv}</td><td>{c_civ}</td><td>{c_total}</td></tr></table>"
    
    html += "<table style='margin-top: 15px;'><tr><th>勤前教育時間</th><th>勤前教育地點</th><th>臨檢點</th><th>盤查點</th><th>聯外道路</th></tr>"
    html += f"<tr><td>{b_time}</td><td>{b_loc}</td><td>{loc_1}處</td><td>{loc_2}處</td><td>{loc_3}處</td></tr></table>"
    
    html += "<h4>參、 督導及其他任務編組表 (19:00 - 23:00)</h4><table><tr><th>項目</th><th>通訊代號</th><th>任務目標</th><th>負責人員</th><th>共同執行人員</th></tr>"
    for _, r in res_cmd.iterrows():
        html += f"<tr><td>{safe_str(r.get('項目')).replace('\n', '<br>')}</td><td>{safe_str(r.get('通訊代號')).replace('\n', '<br>')}</td><td style='text-align:left'>{safe_str(r.get('任務目標')).replace('\n','<br>')}</td><td>{safe_str(r.get('負責人員')).replace('\n', '<br>')}</td><td>{safe_str(r.get('共同執行人員')).replace('\n', '<br>')}</td></tr>"
    html += "</table>"
    
    html += "<h4>肆、【第一階段 20:00 - 23:00】機動攔查任務編組</h4><div class='middle-block'><b>勤務重點：</b>取消定點路檢，採取全面機動巡邏。針對酒駕熱點攔停盤查；攔獲疑似改裝噪音車，立即引導至「警政大樓廣場」交由環保局檢驗。<br>（註：本階段機動攔查共6組警力。21時30分起，第1至第4組轉入第二階段執行擴大臨檢；第5、第6組全程獨留於路面，持續執行機動攔查至23時。）</div>"
    html += "<table><tr><th>組別</th><th>單位</th><th>職別/姓名</th><th>任務分工</th><th>攜行裝備</th><th>巡邏與攔查責任區</th></tr>"
    for _, r in res_ptl.iterrows():
        html += f"<tr><td>{safe_str(r.get('組別')).replace('\n', '<br>')}</td><td>{safe_str(r.get('單位')).replace('\n','<br>')}</td><td style='text-align:left'>{safe_str(r.get('職別/姓名')).replace('\n','<br>')}</td><td style='text-align:left'>{safe_str(r.get('任務分工')).replace('\n', '<br>')}</td><td style='text-align:left'>{safe_str(r.get('攜行裝備')).replace('\n', '<br>')}</td><td style='text-align:left'>{safe_str(r.get('巡邏與攔查責任區')).replace('\n', '<br>')}</td></tr>"
    html += "</table>"
    
    html += "<h4>伍、【第二階段 21:30 - 23:00】擴大臨檢任務編組</h4><div class='middle-block'><b>勤務重點：</b>由第一階段之第1至第4組機動警力，會合偵查隊專案人員，於21時20分前集結完畢，21時30分準時進入目標場所執行威力掃蕩。</div>"
    html += "<table><tr><th>組別</th><th>單位</th><th>職別/姓名</th><th>任務分工</th><th>臨檢目標場所</th></tr>"
    for _, r in res_cp.iterrows():
        html += f"<tr><td>{safe_str(r.get('組別')).replace('\n', '<br>')}</td><td>{safe_str(r.get('單位')).replace('\n','<br>')}</td><td style='text-align:left'>{safe_str(r.get('職別/姓名')).replace('\n','<br>')}</td><td style='text-align:left'>{safe_str(r.get('任務分工')).replace('\n', '<br>')}</td><td style='text-align:left'>{safe_str(r.get('臨檢目標場所')).replace('\n', '<br>')}</td></tr>"
    html += "</table><div class='middle-block'>備註：臨檢完畢後若有剩餘時間，於各所轄內治安熱點、涉毒區段加強巡守，以防制刑案發生。</div>"

    html += f"<h4>陸、 工作重點與法令宣導</h4><div class='middle-block'>"
    for line in str(b_info).split('\n'):
        if line.strip():
            html += f"<div style='padding-left: 3em; text-indent: -3em; margin-bottom: 8px;'>{safe_str(line).replace('\n', '<br>')}</div>"
    html += "</div>"
    
    return html + "</body></html>"

st.markdown("---")
with st.expander("點擊展開即時預覽 (包含完整六大項段落與地點統計)"):
    st.components.v1.html(get_html(), height=800, scrolling=True)

col_dl1, col_dl2 = st.columns(2)

download_plan_name = f"{u}執行{p_name}勤務規劃表.pdf"
download_attendance_name = f"{u}執行{p_name}勤前教育會議人員簽到表.pdf"

pdf_plan = generate_pdf_from_data(u, p_name, p_time, b_info, res_cmd, res_ptl, res_cp, current_stats)
col_dl1.download_button("📝 下載 1.勤務規劃表", data=pdf_plan, file_name=download_plan_name, use_container_width=True)

pdf_attendance = generate_attendance_pdf(u, p_name, p_time, b_info)
col_dl2.download_button("🖋️ 下載 2.人員簽到表", data=pdf_attendance, file_name=download_attendance_name, use_container_width=True)

if st.button("💾 同步雲端並發送備份郵件", use_container_width=True):
    with st.spinner("同步中..."):
        if save_data(u, p_time, p_name, b_info, res_cmd, res_ptl, res_cp, current_stats):
            ok, mail_err = send_report_email(u, p_name, p_time, b_info, res_cmd, res_ptl, res_cp, current_stats)
            if ok: st.success("✅ 同步成功並已寄出郵件！")
            else: st.warning(f"⚠️ 雲端已同步，但郵件失敗: {mail_err}")
