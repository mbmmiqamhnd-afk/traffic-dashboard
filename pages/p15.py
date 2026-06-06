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

DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年5月29日 20時至24時"
DEFAULT_PROJ_BODY = "「全國同步擴大取締酒後駕車及防制危險駕車」暨「擴大臨檢」"
DEFAULT_BRIEF   = "一、 工作重點任務提示：同仁執行盤查、臨檢及路檢勤務過程中，應強化敵情觀念，提高危機意識，並特別注意人犯戒護，落實「人犯戒護安全、案件程序安全、執法者及民眾安全」之「三安」要求。\n二、 行動要領：除法律另有規定外，局察人員執行場所之臨檢，應限於已發生危害或依客觀合理判斷易生危害之場所、交通工具或公共場所為之。\n三、 盤查規範：確實依司法院大法官釋字第535號解釋及「警察職權行使法」對於盤查人、車以及實施臨檢之相關規定，應注意遵守比例原則及考量民眾觀感，不得逾越必要程度。\n四、 全程蒐證：執行各項干涉、取締、處理糾紛及爭議性勤務，務必全程連續錄音或錄影，以避免因案件招致物議。\n五、 異議處理：民眾對警察行使職權表示異議，認為無理由者得繼續執行，但經請求時應將異議之理由製作紀錄交付之。"

DEFAULT_PTL_FOCUS = "採取定點路檢，針對酒駕熱點攔停盤查。"
DEFAULT_CP_FOCUS = "由第一階段之第1至第2組定點路檢警力，會合偵查隊專案人員，於22時20分前集結完畢，22時30分準時進入目標場所執行威力掃蕩。"

# 參、督導及其他任務編組表底稿資料
DEFAULT_CMD = pd.DataFrame([
    {"項目": "指揮官", "通訊代號": "隆安1號", "任務目標": "重點機動督導", "負責人員": "分局長 施宇峰", "共同執行人員": "秘書陳鵬翔、警員張庭溱"},
    {"項目": "副指揮官", "通訊代號": "隆安2號", "任務目標": "重點機動督導", "負責人員": "副分局長何憶雯", "共同執行人員": "警務佐曾威仁"},
    {"項目": "副指揮官", "通訊代號": "隆安3號", "任務目標": "重點機動督導", "負責人員": "副分局長蔡志明", "共同執行人員": "警員陳明祥"},
    {"項目": "上級督導官", "通訊代號": "建興", "任務目標": "重點機動督導", "負責人員": "督察孫三陽", "共同執行人員": ""},
    {"項目": "偵查隊", "通訊代號": "隆安11號", "任務目標": "在隊督辦刑案", "負責人員": "隊長柯志賢", "共同執行人員": "偵查員 施明輝"},
    {"項目": "行政組", "通訊代號": "隆安5號", "任務目標": "督導第一階段臨檢組", "負責人員": "組長 周金柱", "共同執行人員": "巡官蕭凱文、警務佐曾威仁、警員謝明展"},
    {"項目": "督察組", "通訊代號": "隆安6號", "任務目標": "機動督導第二階段臨檢組", "負責人員": "組長 黃長旗", "共同執行人員": "警務員 陳冠彰"},
    {"項目": "保安民防組", "通訊代號": "隆安9號", "任務目標": "機動督導第一階段臨檢組；機動督導第二階段路檢組", "負責人員": "組長林良鍾", "共同執行人員": "巡官古家杰"},
    {"項目": "交通組", "通訊代號": "隆安13號", "任務目標": "機動督導第一階段路檢組", "負責人員": "組長 楊孟竟", "共同執行人員": "巡官郭勝隆"},
    {"項目": "勤務指導", "通訊代號": "隆安 685號", "任務目標": "指導各路檢點、攔檢點，指導各檢查組勤務執行及狀況處置", "負責人員": "教官郭文義", "共同執行人員": "勤務指導人員"},
    {"項目": "聯絡組", "通訊代號": "隆安", "任務目標": "擔任通訊聯絡、指揮管制事宜", "負責人員": "勤指主任蔡奇青", "共同執行人員": "執勤官江文頌、值勤員曾嘉偉 (18-20時)"},
    {"項目": "偵訊組", "通訊代號": "隆安10號", "任務目標": "負責按捺指紋、照相及移送案件相關事宜", "負責人員": "偵查佐賴享宏", "共同執行人員": "警員張峻銨 (在隊待命受理移送案件)"},
    {"項目": "作業組", "通訊代號": "", "任務目標": "負責勤務後勤、勤教場地布置相關事宜", "負責人員": "警員葉俊宏", "共同執行人員": "警務員曾盛鉉、巡官吳國棟、巡佐許榮裕、呂紹臺、警員"}
])

# 肆、第一階段機動攔查底稿資料
DEFAULT_PTL = pd.DataFrame([
    {"單位": "分局規劃", "服勤人員": "待派同仁", "任務分工": "帶班兼管制", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "各指定路段/熱點"},
    {"單位": "分局規劃", "服勤人員": "待派同仁", "任務分工": "指揮管制", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "各指定路段/熱點"},
    {"單位": "分局規劃", "服勤人員": "待派同仁", "任務分工": "攔檢盤查", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "各指定路段/熱點"},
    {"單位": "分局規劃", "服勤人員": "待派同仁", "任務分工": "攔檢盤查", "攜行裝備": "小電腦、密錄器", "巡邏與攔查責任區": "各指定路段/熱點"},
    {"單位": "分局規劃", "服勤人員": "待派同仁", "任務分工": "警戒兼蒐證", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "各指定路段/熱點"},
    {"單位": "分局規劃", "服勤人員": "待派同仁", "任務分工": "帶班兼管制", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "各指定路段/熱點"},
    {"單位": "分局規劃", "服勤人員": "待派同仁", "任務分工": "指揮管制", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "各指定路段/熱點"},
    {"單位": "分局規劃", "服勤人員": "待派同仁", "任務分工": "攔檢盤查", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "各指定路段/熱點"},
    {"單位": "分局規劃", "服勤人員": "待派同仁", "任務分工": "攔檢盤查", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "各指定路段/熱點"},
    {"單位": "分局規劃", "服勤人員": "待派同仁", "任務分工": "攔檢盤查", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "各指定路段/熱點"},
    {"單位": "分局規劃", "服勤人員": "待派同仁", "任務分工": "警戒兼蒐證", "攜行裝備": "槍彈、無線電、小電腦、密錄器", "巡邏與攔查責任區": "各指定路段/熱點"}
])

# 伍、第二階段擴大臨檢底稿資料
DEFAULT_CHECKPOINT = pd.DataFrame([
    {"單位": "臨檢編組", "服勤人員": "待派同仁", "任務分工": "帶班", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、C.丹陽泰養生館、D.溫馨汽車旅館、E.凱虹汽車旅館"},
    {"單位": "臨檢編組", "服勤人員": "待派同仁", "任務分工": "製作臨檢紀錄", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、C.丹陽泰養生館、D.溫馨汽車旅館、E.凱虹汽車旅館"},
    {"單位": "臨檢編組", "服勤人員": "待派同仁", "任務分工": "盤查兼蒐證", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、C.丹陽泰養生館、D.溫馨汽車旅館、E.凱虹汽車旅館"},
    {"單位": "臨檢編組", "服勤人員": "待派同仁", "任務分工": "盤查兼蒐證", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、C.丹陽泰養生館、D.溫馨汽車旅館、E.凱虹汽車旅館"},
    {"單位": "臨檢編組", "服勤人員": "待派同仁", "任務分工": "大門警(車)戒兼蒐證", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、C.丹陽泰養生館、D.溫馨汽車旅館、E.凱虹汽車旅館"},
    {"單位": "偵查隊", "服勤人員": "專案同仁", "任務分工": "刑案偵防、社維法案件處理及移送", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、C.丹陽泰養生館、D.溫馨汽車旅館、E.凱虹汽車旅館"},
    {"單位": "偵查隊", "服勤人員": "專案同仁", "任務分工": "刑案偵防、社維法案件處理及移送", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、C.丹陽泰養生館、D.溫馨汽車旅館、E.凱虹汽車旅館"},
    {"單位": "臨檢編組", "服勤人員": "待派同仁", "任務分工": "帶班", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、F.憤怒鳥網咖、G.真情男女養生館、H.萬紫千紅舒壓館"},
    {"單位": "臨檢編組", "服勤人員": "待派同仁", "任務分工": "製作臨檢紀錄", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、F.憤怒鳥網咖、G.真情男女養生館、H.萬紫千紅舒壓館"},
    {"單位": "臨檢編組", "服勤人員": "待派同仁", "任務分工": "盤查兼蒐證", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、F.憤怒鳥網咖、G.真情男女養生館、H.萬紫千紅舒壓館"},
    {"單位": "臨檢編組", "服勤人員": "待派同仁", "任務分工": "盤查兼蒐證", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、F.憤怒鳥網咖、G.真情男女養生館、H.萬紫千紅舒壓館"},
    {"單位": "臨檢編組", "服勤人員": "待派同仁", "任務分工": "盤查兼蒐證", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、F.憤怒鳥網咖、G.真情男女養生館、H.萬紫千紅舒壓館"},
    {"單位": "臨檢編組", "服勤人員": "待派同仁", "任務分工": "大門警(車)戒兼蒐證", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、F.憤怒鳥網咖、G.真情男女養生館、H.萬紫千紅舒壓館"},
    {"單位": "偵查隊", "服勤人員": "專案同仁", "任務分工": "刑案偵防、社維法案件處理及移送", "臨檢目標場所": "A.鉅大撞球館、B.台灣麻將協會、F.憤怒鳥網咖、G.真情男女養生館、H.萬紫千紅舒壓館"}
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
    """從時間文字自動精準擷取4碼數字"""
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
        
        # 設定頁
        try: ws_set = sh.worksheet("三合一_設定")
        except: ws_set = sh.add_worksheet(title="三合一_設定", rows="50", cols="5")
        ws_set.clear()
        
        # 【徹底解決】此處已完全移除了對 stats['loc_1'] ~ stats['loc_3'] 的調用，與 current_stats 完美對齊
        ws_set.update([
            ["Key", "Value"], ["unit_name", unit], ["plan_full_time", time_str], ["project_name", project],
            ["briefing_info", briefing], ["stats_cmd", str(stats['cmd'])], ["stats_ptl", str(stats['ptl'])],
            ["stats_inv", str(stats['inv'])], ["stats_civ", str(stats['civ'])], ["briefing_time", str(stats['b_time'])],
            ["briefing_loc", str(stats['b_loc'])], ["ptl_focus", ptl_f], ["cp_focus", cp_f]
        ])
        
        # 各勤務編組工作表
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
    
    # 壹、基本資料
    story.append(Paragraph("<b>壹、 勤務基本資料</b>", style_section))
    date_str = clean(time_str.split(" ")[0] if " " in time_str else "115年5月29日")
    time_str_only = clean(time_str.split(" ")[1] if " " in time_str else "20時至24時")
    data_basic = [[Paragraph("<b>實施日期</b>", style_cell), Paragraph("<b>勤務時間</b>", style_cell), Paragraph("<b>指揮官</b>", style_cell), Paragraph("<b>勤務編組</b>", style_cell), Paragraph("<b>聯合稽查站地點</b>", style_cell)], 
                  [Paragraph(date_str, style_cell), Paragraph(time_str_only, style_cell), Paragraph("分局長 施宇峰", style_cell), Paragraph("如任務編組表", style_cell), Paragraph("本局廣場", style_cell)]]
    t_basic = Table(data_basic, colWidths=[page_width*0.18, page_width*0.2, page_width*0.18, page_width*0.18, page_width*0.26])
    t_basic.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_basic)
    
    # 貳、統計表
    story.append(Paragraph("<b>貳、 警力統計及地點統計</b>", style_section))
    data_stats = [[Paragraph("督導組", style_cell), Paragraph("攔臨組", style_cell), Paragraph("偵訊組", style_cell), Paragraph("小計", style_cell), Paragraph("民力", style_cell), Paragraph("總計", style_cell)], 
                  [Paragraph(str(stats['cmd']), style_cell), Paragraph(str(stats['ptl']), style_cell), Paragraph(str(stats['inv']), style_cell), Paragraph(str(stats['cmd']+stats['ptl']+stats['inv']), style_cell), Paragraph(str(stats['civ']), style_cell), Paragraph(str(stats['total']), style_cell)]]
    t_stats = Table(data_stats, colWidths=[page_width*0.16]*6)
    t_stats.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_stats)

    # 參、督導組
    story.append(Paragraph("<b>參、 督導及其他任務編組表</b>", style_section))
    data_cmd = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["項目", "通訊代號", "任務目標", "負責人員", "共同人員"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([Paragraph(clean(r.get('項目')), style_cell), Paragraph(clean(r.get('通訊代號')), style_cell), Paragraph(clean(r.get('任務目標')), style_cell_left), Paragraph(clean(r.get('負責人員')), style_cell), Paragraph(clean(r.get('共同執行人員')), style_cell)])
    
    # 【負責人員寬度擴大】設為 0.26，提供充足空間以完美容納 8 個中文字寬度不變形
    t_cmd = Table(data_cmd, colWidths=[page_width*0.12, page_width*0.14, page_width*0.28, page_width*0.26, page_width*0.2])
    t_cmd.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_cmd)
    
    # 肆、第一階段
    story.append(Paragraph("<b>肆、【第一階段】機動攔查任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務重點：</b>{clean(ptl_f)}", style_text)) 
    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["組別", "單位", "服勤人員", "任務分工", "攜行裝備", "責任區"]]]
    for _, r in df_ptl.iterrows():
        data_ptl.append([Paragraph(clean(r.get('組別')), style_cell), Paragraph(clean(r.get('單位')), style_cell), Paragraph(clean(r.get('服勤人員')), style_cell_left), Paragraph(clean(r.get('任務分工')), style_cell_left), Paragraph(clean(r.get('攜行裝備')), style_cell_left), Paragraph(clean(r.get('巡邏與攔查責任區')), style_cell_left)])
    t_ptl = Table(data_ptl, colWidths=[page_width*0.11, page_width*0.12, page_width*0.20, page_width*0.21, page_width*0.16, page_width*0.20])
    t_ptl.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t_ptl)

    # 伍、第二階段
    story.append(Paragraph("<b>伍、【第二階段】擴大臨檢任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務重點：</b>{clean(cp_f)}", style_text))
    if df_cp is not None and not df_cp.empty:
        data_cp = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["組別", "單位", "服勤人員", "任務分工", "臨檢場所"]]]
        for _, r in df_cp.iterrows():
            data_cp.append([Paragraph(clean(r.get('組別')), style_cell), Paragraph(clean(r.get('單位')), style_cell), Paragraph(clean(r.get('服勤人員')), style_cell_left), Paragraph(clean(r.get('任務分工')), style_cell_left), Paragraph(clean(r.get('臨檢目標場所')), style_cell_left)])
        t_cp = Table(data_cp, colWidths=[page_width*0.11, page_width*0.12, page_width*0.24, page_width*0.24, page_width*0.29])
        t_cp.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e6e6e6')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
        story.append(t_cp)
    
    # 陸、工作重點
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
    date_part = time_str.split(' ')[0] if ' ' in time_str else "115年5月29日"
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

# --- Streamlit 介面 ---
df_set, df_cmd, df_ptl, df_cp, err = load_data()
default_stats = {'cmd': 7, 'ptl': 16, 'inv': 3, 'civ': 0, 'b_time': '19時30分至20時00分', 'b_loc': '分局二樓會議室'}

if err:
    st.error(f"⚠️ 雲端資料載入異常：{err}")

if df_set is None or df_set.empty:
    u, t, b = DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_BRIEF
    proj_body = DEFAULT_PROJ_BODY
    ed_cmd, ed_ptl, ed_cp = DEFAULT_CMD.copy(), DEFAULT_PTL.copy(), DEFAULT_CHECKPOINT.copy()
    p_ptl_focus, p_cp_focus = DEFAULT_PTL_FOCUS, DEFAULT_CP_FOCUS
else:
    d = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    u, t, b = d.get("unit_name", DEFAULT_UNIT), d.get("plan_full_time", DEFAULT_TIME), d.get("briefing_info", DEFAULT_BRIEF)
    p_ptl_focus, p_cp_focus = d.get("ptl_focus", DEFAULT_PTL_FOCUS), d.get("cp_focus", DEFAULT_CP_FOCUS)
    
    raw_proj = d.get("project_name", DEFAULT_PROJ_BODY)
    proj_body = re.sub(r'^\d{4}', '', raw_proj)
    
    default_stats.update({'cmd': int(d.get("stats_cmd", 7) or 7), 'ptl': int(d.get("stats_ptl", 16) or 16), 'inv': int(d.get("stats_inv", 3) or 3), 'civ': int(d.get("stats_civ", 0) or 0), 'b_time': d.get("briefing_time", "19時30分至20時00分"), 'b_loc': d.get("briefing_loc", "分局二樓會議室")})
    ed_cmd = df_cmd if not df_cmd.empty else DEFAULT_CMD.copy()
    ed_ptl = df_ptl.drop(columns=["組別"]) if not df_ptl.empty and "組別" in df_ptl.columns else (df_ptl if not df_ptl.empty else DEFAULT_PTL.copy())
    ed_cp = df_cp.drop(columns=["組別"]) if df_cp is not None and not df_cp.empty and "組別" in df_cp.columns else (df_cp if df_cp is not None and not df_cp.empty else DEFAULT_CHECKPOINT.copy())

st.title("🚓 專案勤務規劃系統")

col_time, col_proj = st.columns([1, 2])
with col_time:
    p_time = st.text_input("勤務時間", t)

# 動態從時間分析出4碼 (例如 0529)
mmdd_code = extract_mmdd(p_time)

with col_proj:
    # 網頁端可直接編輯，前方提示當下連動的4碼數字
    input_proj_body = st.text_input(f"專案名稱 (目前連動代碼: {mmdd_code})", proj_body)

# 自動拼裝為最終完整的專案全名
p_name = f"{mmdd_code}{input_proj_body}"

st.subheader("貳、 警力統計及地點統計")
col_s1, col_s2, col_s3, col_s4 = st.columns(4)
c_cmd, c_ptl, c_inv, c_civ = col_s1.number_input("督導組", value=default_stats['cmd']), col_s2.number_input("攔臨組", value=default_stats['ptl']), col_s3.number_input("偵訊組", value=default_stats['inv']), col_s4.number_input("民力", value=default_stats['civ'])

# 【已對齊】current_stats 移除 loc 稽查點
current_stats = {'cmd': c_cmd, 'ptl': c_ptl, 'inv': c_inv, 'civ': c_civ, 'total': c_cmd+c_ptl+c_inv+c_civ, 'b_time': default_stats['b_time'], 'b_loc': default_stats['b_loc']}

st.subheader("參、 指導編組與重點宣導")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True).dropna(how='all').fillna("")
b_info = st.text_area("陸、 工作重點與法令宣導", b, height=150)

st.subheader("勤務執行編組 (兩階段)")
tab1, tab2 = st.tabs(["肆、【第一階段】機動攔查", "伍、【第二階段】擴大臨檢"])

with tab1:
    res_ptl_raw = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True, key="ptl_ed").dropna(how='all').fillna("").reset_index(drop=True)
    if not res_ptl_raw.empty:
        res_ptl = res_ptl_raw.copy()
        res_ptl["組別"] = [f"第{i+1}巡邏組" for i in range(len(res_ptl))]
        res_ptl = res_ptl[["組別"] + [col for col in res_ptl.columns if col != "組別"]]
    else:
        res_ptl = res_ptl_raw

with tab2:
    res_cp_raw = st.data_editor(ed_cp, num_rows="dynamic", use_container_width=True, key="cp_ed").dropna(how='all').fillna("").reset_index(drop=True)
    if not res_cp_raw.empty:
        res_cp = res_cp_raw.copy()
        res_cp["組別"] = [f"第{i+1}臨檢組" for i in range(len(res_cp))]
        res_cp = res_cp[["組別"] + [col for col in res_cp.columns if col != "組別"]]
    else:
        res_cp = res_cp_raw

st.markdown("---")
# 已依需求移除下載規劃表按鈕 (st.download_button)

if st.button("💾 同步雲端並發送郵件", use_container_width=True):
    with st.spinner("⏳ 正在寫入雲端並寄送郵件，請稍候..."):
        if save_data(u, p_time, p_name, b_info, res_cmd, res_ptl, res_cp, current_stats, p_ptl_focus, p_cp_focus):
            ok, mail_err = send_report_email(u, p_name, p_time, b_info, res_cmd, res_ptl, res_cp, current_stats, p_ptl_focus, p_cp_focus)
            if ok: 
                st.success(f"✅ 資料已成功同步至雲端！專案完整名稱為「{p_name}」，規劃表及簽到表 PDF 郵件已發送完成！")
                st.rerun()
            else: 
                st.error(f"❌ 雲端同步成功，但寄送郵件失敗！錯誤細節：{mail_err}")
