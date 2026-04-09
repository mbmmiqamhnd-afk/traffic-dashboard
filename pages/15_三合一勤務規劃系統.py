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
st.set_page_config(page_title="專案勤務規劃系統", layout="wide", page_icon="🚓")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年4月11日 19時至23時"
DEFAULT_PROJ    = "0411取締酒後駕車與防制危險駕車及噪音車輛專案"
DEFAULT_BRIEF   = "一、 落實三安：同仁執行盤查、臨檢及機動勤務過程中，應強化敵情觀念，提高危機意識，落實「人犯戒護安全、案件程序安全、執法者及民眾安全」。\n二、 臨檢合法性：警察人員執行場所之臨檢，應限於已發生危害或依客觀合理判斷易生危害之場所，進行臨檢前應對當事人告以實施事由，便衣人員並應出示證件（依《警察職權行使法》第6條）。\n三、 攔停規範：機動攔檢對於已發生危害或易生危害之交通工具，得予以攔停；若有異常舉動而合理懷疑其將有危害行為時，得要求接受酒精濃度測試（依《警察職權行使法》第8條）。\n四、 全程蒐證：執行各項干涉、取締、處理糾紛及爭議性勤務（含噪音車引導與酒測），務必全程連續錄音或錄影。\n五、 異議處理：民眾對警察行使職權表示異議，認為無理由者得繼續執行，但經請求時應將異議之理由製作紀錄交付之（依《警察職權行使法》第29條）。"

DEFAULT_PTL_FOCUS = "取消定點路檢，採取全面機動巡邏。針對酒駕熱點攔停盤查；攔獲疑似改裝噪音車，立即引導至「警政大樓廣場」交由環保局檢驗。\n（註：本階段機動攔查共6組警力。21時30分起，第1至第4組轉入第二階段執行擴大臨檢；第5、第6組全程獨留於路面，持續執行機動攔查至23時。）"
DEFAULT_CP_FOCUS = "由第一階段之第1至第4組機動警力，會合偵查隊專案人員，於21時20分前集結完畢，21時30分準時進入目標場所執行威力掃蕩。"

DEFAULT_CMD = pd.DataFrame([
    {"項目": "指揮官", "通訊代號": "隆安 1 號", "任務目標": "勤務核定並重點機動督導", "負責人員": "分局長 施宇峰", "共同執行人員": "巡官 郭勝隆"},
    {"項目": "副指揮官", "通訊代號": "隆安 2 號", "任務目標": "襄助指揮、重點機動督導", "負責人員": "副分局長 何憶雯", "共同執行人員": "警務佐 曾威仁"},
    {"項目": "副指揮官", "通訊代號": "隆安 3 號", "任務目標": "襄助指揮、重點機動督導", "負責人員": "副分局長 蔡志明", "共同執行人員": "警員 陳明祥"},
    {"項目": "行政組", "通訊代號": "隆安 5 號", "任務目標": "督導第二階段臨檢勤務", "負責人員": "組長 周金柱", "共同執行人員": "巡官 蕭凱文"},
    {"項目": "督察組", "通訊代號": "隆安 6 號", "任務目標": "機動督導各單位勤務紀律", "負責人員": "督察組長 黃長旗", "共同執行人員": "警務員 陳冠彰"},
    {"項目": "交通組", "通訊代號": "隆安 13號", "任務目標": "機動督導第一階段攔檢組", "負責人員": "交通組長 楊孟竟", "共同執行人員": "警務員 盧冠仁"},
    {"項目": "聯絡組", "通訊代號": "隆安", "任務目標": "擔任通訊聯絡、指揮管制事宜", "負責人員": "勤指主任 蔡奇青", "共同執行人員": "執勤官、值勤員"},
    {"項目": "偵訊組", "通訊代號": "隆安 10號", "任務目標": "負責按捺指紋、照相及移送", "負責人員": "偵查隊 偵查佐", "共同執行人員": "在隊待命受理"},
])

DEFAULT_PTL = pd.DataFrame([
    {"單位": "聖亭所", "服勤人員": "所長 鄭榮捷\n警員 詹宗澤", "任務分工": "帶班\n盤查兼警戒", "攜行裝備": "槍彈、無線電\n小電腦、密錄器", "巡邏與攔查責任區": "中正路、北龍路周邊及治安要點機動攔查。\n(20:00-21:30機動，後轉臨檢)"},
    {"單位": "龍潭所", "服勤人員": "所長 孫祥愷\n警員 沈庭禾", "任務分工": "盤查兼警戒", "攜行裝備": "槍彈、無線電\n小電腦、密錄器", "巡邏與攔查責任區": "北龍路、中豐路周邊及治安要點機動攔查。\n(20:00-21:30機動，後轉臨檢)"},
    {"單位": "高平所", "服勤人員": "警員 邱春松\n警員 唐銘聰", "任務分工": "盤查兼警戒", "攜行裝備": "槍彈、無線電\n小電腦、密錄器", "巡邏與攔查責任區": "東龍路、中豐路沿線機動攔查。\n(20:00-21:30機動，後轉臨檢)"},
    {"單位": "石門所", "服勤人員": "巡佐 林偉政\n警員 鄒詠如", "任務分工": "帶班\n盤查兼警戒", "攜行裝備": "槍彈、無線電\n小電腦、密錄器", "巡邏與攔查責任區": "神龍路、文化路周邊及治安要點機動攔查。\n(20:00-21:30機動，後轉臨檢)"},
    {"單位": "中興所", "服勤人員": "所長 董亦文\n警員 徐毓汶", "任務分工": "盤查兼警戒", "攜行裝備": "槍彈、無線電\n小電腦、密錄器", "巡邏與攔查責任區": "中興路、龍新路沿線及治安要點機動攔查。\n(全程留守機動 20:00-23:00)"},
    {"單位": "交分隊", "服勤人員": "小隊長 林振生\n警員 吳沛軒", "任務分工": "盤查兼警戒", "攜行裝備": "槍彈、無線電\n小電腦、密錄器", "巡邏與攔查責任區": "轄內易發生危駕路段、各聯外道路機動攔查。\n(全程留守機動 20:00-23:00)"},
])

DEFAULT_CHECKPOINT = pd.DataFrame([
    {"單位": "聖亭所\n龍潭所\n偵查隊", "服勤人員": "所長 鄭榮捷 警員 詹宗澤\n所長 孫祥愷 警員 沈庭禾\n偵查佐 賴享宏 警員 張峻銨", "任務分工": "帶班 製作臨檢紀錄\n盤查兼蒐證\n刑案偵防、社維法", "臨檢目標場所": "A. 鉅大撞球館 (中豐路558號)\nB. 台灣麻將協會 (中豐路558之1號)\nC. 丹陽泰養生館 (中豐路281號)\nD. 溫馨汽車旅館 (中正路457號)\nE. 凱虹汽車旅館 (中正路506號)"},
    {"單位": "石門所\n高平所\n偵查隊", "服勤人員": "巡佐 林偉政 警員 鄒詠如\n警員 邱春松 警員 唐銘聰\n偵查隊警員 2名", "任務分工": "帶班 製作臨檢紀錄\n大門警戒兼盤查\n刑案偵防、社維法", "臨檢目標場所": "F. 憤怒鳥網咖 (中興路269號)\nG. 真情男女養生館 (中興路387號)\nH. 萬紫千紅舒壓館 (中興路491-3號)"},
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

def safe_str(val):
    if pd.isna(val) or val is None or str(val).strip().lower() == "nan": return ""
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
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        try: ws_set = sh.worksheet("三合一_設定")
        except: ws_set = sh.add_worksheet(title="三合一_設定", rows="50", cols="5")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["unit_name", unit], ["plan_full_time", time_str], ["project_name", project], ["briefing_info", briefing], ["stats_cmd", str(stats['cmd'])], ["stats_ptl", str(stats['ptl'])], ["stats_inv", str(stats['inv'])], ["stats_civ", str(stats['civ'])], ["briefing_time", str(stats['b_time'])], ["briefing_loc", str(stats['b_loc'])], ["loc_1", str(stats['loc_1'])], ["loc_2", str(stats['loc_2'])], ["loc_3", str(stats['loc_3'])], ["ptl_focus", ptl_f], ["cp_focus", cp_f]])
        for ws_name, df in [("三合一_指揮組", df_cmd), ("三合一_巡邏組", df_ptl), ("三合一_擴大臨檢組", df_cp)]:
            try: ws = sh.worksheet(ws_name)
            except: ws = sh.add_worksheet(title=ws_name, rows="100", cols="20")
            ws.clear()
            df_cleaned = df.dropna(how='all').fillna("")
            if not df_cleaned.empty: ws.update([df_cleaned.columns.tolist()] + df_cleaned.values.tolist())
        load_data.clear()
        return True
    except: return False

# --- 3. 勤務規劃表 PDF 生成功能 ---
def generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, stats, ptl_f, cp_f):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=10*mm, rightMargin=10*mm, topMargin=12*mm, bottomMargin=15*mm)
    page_width = A4[0] - 20*mm
    story = []
    
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=26, alignment=1, spaceAfter=8)
    style_section = ParagraphStyle('Section', fontName=font, fontSize=14, leading=20, alignment=0, spaceAfter=2*mm, spaceBefore=4*mm)
    style_text = ParagraphStyle('Text', fontName=font, fontSize=14, leading=20, alignment=0)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=20, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=20, alignment=0)
    
    def clean(t): return safe_str(t).replace("\n", "<br/>")

    story.append(Paragraph(f"<b>{unit}執行 {project} 勤務規劃表</b>", style_title))
    story.append(Paragraph("<b>壹、 勤務基本資料</b>", style_section))
    date_str = clean(time_str.split(" ")[0] if " " in time_str else "115年4月11日")
    time_str_only = clean(time_str.split(" ")[1] if " " in time_str else "19時至23時")
    data_basic = [[Paragraph("<b>實施日期</b>", style_cell), Paragraph("<b>勤務時間</b>", style_cell), Paragraph("<b>指揮官</b>", style_cell), Paragraph("<b>勤務編組</b>", style_cell), Paragraph("<b>聯合稽查站地點</b>", style_cell)], [Paragraph(f"<nobr>{date_str}</nobr>", style_cell), Paragraph(f"<nobr>{time_str_only}</nobr>", style_cell), Paragraph("分局長 施宇峰", style_cell), Paragraph("如各階段任務編組表", style_cell), Paragraph("龍潭區警政聯合辦公大樓廣場", style_cell)]]
    t_basic = Table(data_basic, colWidths=[page_width*0.18, page_width*0.2, page_width*0.18, page_width*0.18, page_width*0.26])
    t_basic.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_basic)
    
    story.append(Paragraph("<b>貳、 警力統計及地點統計</b>", style_section))
    data_stats = [[Paragraph("<b>單位</b>", style_cell), Paragraph("<b>督導組</b>", style_cell), Paragraph("<b>攔臨組</b>", style_cell), Paragraph("<b>偵訊組</b>", style_cell), Paragraph("<b>小計</b>", style_cell), Paragraph("<b>民力</b>", style_cell), Paragraph("<b>總計</b>", style_cell)], [Paragraph("龍潭分局", style_cell), Paragraph(str(stats['cmd']), style_cell), Paragraph(str(stats['ptl']), style_cell), Paragraph(str(stats['inv']), style_cell), Paragraph(str(stats['cmd']+stats['ptl']+stats['inv']), style_cell), Paragraph(str(stats['civ']), style_cell), Paragraph(str(stats['total']), style_cell)]]
    t_stats = Table(data_stats, colWidths=[page_width*0.16, page_width*0.14, page_width*0.14, page_width*0.14, page_width*0.14, page_width*0.14, page_width*0.14])
    t_stats.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_stats)
    
    story.append(Spacer(1, 2*mm))
    data_loc = [[Paragraph("<b>勤前時間</b>", style_cell), Paragraph("<b>勤前地點</b>", style_cell), Paragraph("<b>臨檢點</b>", style_cell), Paragraph("<b>盤查點</b>", style_cell), Paragraph("<b>聯外道路</b>", style_cell)], [Paragraph(clean(stats['b_time']), style_cell), Paragraph(clean(stats['b_loc']), style_cell), Paragraph(f"{stats['loc_1']}處", style_cell), Paragraph(f"{stats['loc_2']}處", style_cell), Paragraph(f"{stats['loc_3']}處", style_cell)]]
    t_loc = Table(data_loc, colWidths=[page_width*0.25, page_width*0.27, page_width*0.16, page_width*0.16, page_width*0.16])
    t_loc.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_loc)

    story.append(Paragraph("<b>參、 督導及其他任務編組表 (19:00 - 23:00)</b>", style_section))
    data_cmd = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["項目", "通訊代號", "任務目標", "負責人員", "共同人員"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([Paragraph(clean(r.get('項目')), style_cell), Paragraph(clean(r.get('通訊代號')), style_cell), Paragraph(clean(r.get('任務目標')), style_cell_left), Paragraph(clean(r.get('負責人員')), style_cell), Paragraph(clean(r.get('共同執行人員')), style_cell)])
    t_cmd = Table(data_cmd, colWidths=[page_width*0.14, page_width*0.16, page_width*0.3, page_width*0.2, page_width*0.2])
    t_cmd.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_cmd)
    
    story.append(Paragraph("<b>肆、【第一階段】機動攔查任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務重點：</b>{clean(ptl_f)}", style_text)) 
    story.append(Spacer(1, 1*mm))
    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["組別", "單位", "服勤人員", "任務分工", "攜行裝備", "責任區"]]]
    for _, r in df_ptl.iterrows():
        data_ptl.append([Paragraph(clean(r.get('組別')), style_cell), Paragraph(clean(r.get('單位')), style_cell), Paragraph(clean(r.get('服勤人員')), style_cell_left), Paragraph(clean(r.get('任務分工')), style_cell_left), Paragraph(clean(r.get('攜行裝備')), style_cell_left), Paragraph(clean(r.get('巡邏與攔查責任區')), style_cell_left)])
    
    t_ptl = Table(data_ptl, colWidths=[page_width*0.11, page_width*0.12, page_width*0.20, page_width*0.21, page_width*0.16, page_width*0.20])
    t_ptl.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_ptl)

    story.append(Paragraph("<b>伍、【第二階段】擴大臨檢任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務重點：</b>{clean(cp_f)}", style_text))
    story.append(Spacer(1, 1*mm))
    data_cp = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["組別", "單位", "服勤人員", "任務分工", "臨檢場所"]]]
    for _, r in df_cp.iterrows():
        data_cp.append([Paragraph(clean(r.get('組別')), style_cell), Paragraph(clean(r.get('單位')), style_cell), Paragraph(clean(r.get('服勤人員')), style_cell_left), Paragraph(clean(r.get('任務分工')), style_cell_left), Paragraph(clean(r.get('臨檢目標場所')), style_cell_left)])
    
    t_cp = Table(data_cp, colWidths=[page_width*0.11, page_width*0.12, page_width*0.24, page_width*0.24, page_width*0.29])
    t_cp.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e6e6e6')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_cp)
    
    story.append(Paragraph("<b>陸、 工作重點與法令宣導</b>", style_section))
    for line in str(briefing).split('\n'):
        if line.strip(): story.append(Paragraph(f"{clean(line)}", style_text))
    
    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 12)
        page_num_text = f"-第{canvas.getPageNumber()}頁-"
        canvas.drawCentredString(A4[0]/2.0, 10*mm, page_num_text)
        canvas.restoreState()

    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    return buf.getvalue()


# =====================================================================
# --- 專案簽到表 (更新：縮減表格行高以容納於一頁) ---
# =====================================================================
def generate_attendance_pdf(unit, project, time_str, stats):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 30*mm
    story = []
    
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=26, alignment=1, spaceAfter=12)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=14, leading=22, spaceAfter=1*mm)
    style_sig = ParagraphStyle('Sig', fontName=font, fontSize=14, leading=22)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=20, alignment=1)
    
    story.append(Paragraph(f"{unit}執行{project}簽到表", style_title))
    
    date_part = time_str.split(' ')[0] if ' ' in time_str else "115年4月11日"
    story.append(Paragraph(f"時間:{date_part}{stats['b_time']}", style_info))
    story.append(Paragraph(f"地點:{stats['b_loc']}召開", style_info))
    story.append(Spacer(1, 6*mm))
    
    # 頂部簽核區
    sig_data = [
        [Paragraph("分局長:", style_sig), Paragraph("上級督導:", style_sig)],
        [Paragraph("副分局長:", style_sig), ""]
    ]
    t_sig = Table(sig_data, colWidths=[page_width*0.5, page_width*0.5])
    t_sig.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(t_sig)
    story.append(Spacer(1, 4*mm))
    
    # 簽到單位區
    table_data = [
        [Paragraph("單位", style_cell), Paragraph("參加人員", style_cell), Paragraph("單位", style_cell), Paragraph("參加人員", style_cell)]
    ]
    
    rows = [
        ("交通組", "聖亭派出所"),
        ("督察組", "龍潭派出所"),
        ("行政組", "中興派出所"),
        ("保安民防組", "石門派出所"),
        ("勤務指揮中心", "高平派出所"),
        ("偵查隊", "三和派出所"),
        ("", "龍潭交通分隊")
    ]
    
    for l, r in rows: 
        table_data.append([Paragraph(l, style_cell) if l else "", "", Paragraph(r, style_cell) if r else "", ""])
        
    # ⚠️ 這裡將每列高度從 25*mm 縮減為 18*mm，保證表格不會被擠到第二頁
    t = Table(table_data, colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3], rowHeights=[12*mm] + [18*mm]*len(rows))
    t.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font), 
        ('GRID', (0,0), (-1,-1), 0.5, colors.black), 
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (3,0), colors.whitesmoke)
    ]))
    story.append(t)
    
    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 12)
        page_num_text = f"-第{canvas.getPageNumber()}頁-"
        canvas.drawCentredString(A4[0]/2.0, 10*mm, page_num_text)
        canvas.restoreState()

    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    return buf.getvalue()


def send_report_email(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, stats, ptl_f, cp_f):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        msg["Subject"] = f"{unit}執行{project}勤務規劃與簽到表_{datetime.now().strftime('%m%d')}"
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

# --- Streamlit 介面 ---
df_set, df_cmd, df_ptl, df_cp, err = load_data()
default_stats = {'cmd': 6, 'ptl': 31, 'inv': 2, 'civ': 0, 'b_time': '18時30分至19時00分', 'b_loc': '分局二樓會議室', 'loc_1': 8, 'loc_2': 6, 'loc_3': 0}
p_ptl_focus, p_cp_focus = DEFAULT_PTL_FOCUS, DEFAULT_CP_FOCUS

if err or df_set is None:
    u, t, p, b = DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_PROJ, DEFAULT_BRIEF
    ed_cmd, ed_ptl, ed_cp = DEFAULT_CMD.copy(), DEFAULT_PTL.copy(), DEFAULT_CHECKPOINT.copy()
else:
    d = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    u, t, p, b = d.get("unit_name", DEFAULT_UNIT), d.get("plan_full_time", DEFAULT_TIME), d.get("project_name", DEFAULT_PROJ), d.get("briefing_info", DEFAULT_BRIEF)
    p_ptl_focus, p_cp_focus = d.get("ptl_focus", DEFAULT_PTL_FOCUS), d.get("cp_focus", DEFAULT_CP_FOCUS)
    default_stats.update({'cmd': int(d.get("stats_cmd", 6)), 'ptl': int(d.get("stats_ptl", 31)), 'inv': int(d.get("stats_inv", 2)), 'civ': int(d.get("stats_civ", 0)), 'b_time': d.get("briefing_time", "18時30分至19時00分"), 'b_loc': d.get("briefing_loc", "分局二樓會議室"), 'loc_1': int(d.get("loc_1", 8)), 'loc_2': int(d.get("loc_2", 6)), 'loc_3': int(d.get("loc_3", 0))})
    ed_cmd = df_cmd if not df_cmd.empty else DEFAULT_CMD.copy()
    ed_ptl = df_ptl.drop(columns=["組別"]) if not df_ptl.empty and "組別" in df_ptl.columns else (df_ptl if not df_ptl.empty else DEFAULT_PTL.copy())
    if "職別/姓名" in ed_ptl.columns: ed_ptl = ed_ptl.rename(columns={"職別/姓名": "服勤人員"})
    ed_cp = df_cp.drop(columns=["組別"]) if df_cp is not None and not df_cp.empty and "組別" in df_cp.columns else (df_cp if df_cp is not None and not df_cp.empty else DEFAULT_CHECKPOINT.copy())
    if ed_cp is not None and "職別/姓名" in ed_cp.columns: ed_cp = ed_cp.rename(columns={"職別/姓名": "服勤人員"})

st.title("🚓 專案勤務規劃系統")
c1, c2 = st.columns(2)
p_time = c2.text_input("勤務時間", t)
date_match = re.search(r"(\d+)月(\d+)日", p_time)
if date_match:
    m, d_str = date_match.group(1).zfill(2), date_match.group(2).zfill(2)
    p = re.sub(r"^\d+", f"{m}{d_str}", p)
p_name = c1.text_input("專案名稱 (注意雲端快取)", p)

st.subheader("貳、 警力統計及地點統計")
col_s1, col_s2, col_s3, col_s4 = st.columns(4)
c_cmd, c_ptl, c_inv, c_civ = col_s1.number_input("督導組", value=default_stats['cmd']), col_s2.number_input("攔臨組", value=default_stats['ptl']), col_s3.number_input("偵訊組", value=default_stats['inv']), col_s4.number_input("民力", value=default_stats['civ'])
c_total = c_cmd + c_ptl + c_inv + c_civ
col_b1, col_b2 = st.columns(2)
b_time, b_loc = col_b1.text_input("勤前時間", default_stats['b_time']), col_b2.text_input("勤前地點", default_stats['b_loc'])
col_l1, col_l2, col_l3 = st.columns(3)
loc_1, loc_2, loc_3 = col_l1.number_input("臨檢點", value=default_stats['loc_1']), col_l2.number_input("盤查點", value=default_stats['loc_2']), col_l3.number_input("聯外道路", value=default_stats['loc_3'])
current_stats = {'cmd': c_cmd, 'ptl': c_ptl, 'inv': c_inv, 'civ': c_civ, 'total': c_total, 'b_time': b_time, 'b_loc': b_loc, 'loc_1': loc_1, 'loc_2': loc_2, 'loc_3': loc_3}

st.subheader("參、 指導編組與重點宣導")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True).dropna(how='all').fillna("")
b_info = st.text_area("陸、 工作重點與法令宣導", b, height=150)
st.subheader("💡 階段勤務重點編輯")
col_f1, col_f2 = st.columns(2)
cur_ptl_focus, cur_cp_focus = col_f1.text_area("機動攔查重點", value=p_ptl_focus, height=100), col_f2.text_area("擴大臨檢重點", value=p_cp_focus, height=100)

st.subheader("勤務執行編組 (兩階段)")
tab1, tab2 = st.tabs(["肆、【第一階段】機動攔查", "伍、【第二階段】擴大臨檢"])
with tab1:
    res_ptl = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True, key="ptl_editor").dropna(how='all').fillna("").reset_index(drop=True)
    if not res_ptl.empty: res_ptl.insert(0, "組別", [f"第{i+1}巡邏組" for i in range(len(res_ptl))])
with tab2:
    res_cp = st.data_editor(ed_cp, num_rows="dynamic", use_container_width=True, key="cp_editor").dropna(how='all').fillna("").reset_index(drop=True)
    if not res_cp.empty: res_cp.insert(0, "組別", [f"第{i+1}臨檢組" for i in range(len(res_cp))])

st.markdown("---")
col_dl1, col_dl2 = st.columns(2)

pdf_plan = generate_pdf_from_data(u, p_name, p_time, b_info, res_cmd, res_ptl, res_cp, current_stats, cur_ptl_focus, cur_cp_focus)
col_dl1.download_button("📝 下載勤務規劃表", data=pdf_plan, file_name=f"{u}規劃表.pdf", use_container_width=True)

pdf_attendance = generate_attendance_pdf(u, p_name, p_time, current_stats)
col_dl2.download_button("🖋️ 下載簽到表 (已調整行高至一頁)", data=pdf_attendance, file_name=f"{u}簽到表.pdf", use_container_width=True)

if st.button("💾 同步雲端並發送郵件", use_container_width=True):
    if save_data(u, p_time, p_name, b_info, res_cmd, res_ptl, res_cp, current_stats, cur_ptl_focus, cur_cp_focus):
        ok, mail_err = send_report_email(u, p_name, p_time, b_info, res_cmd, res_ptl, res_cp, current_stats, cur_ptl_focus, cur_cp_focus)
        if ok: st.success("✅ 已同步並寄出！")
        else: st.warning(f"⚠️ 同步成功但郵件失敗: {mail_err}")
