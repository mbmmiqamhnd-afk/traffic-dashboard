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
st.set_page_config(page_title="二階段勤務規劃系統", layout="wide", page_icon="🚓")

# --- 常數與專屬工作表設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

WS_MAP = {
    "set": "二階段_設定",
    "cmd": "二階段_指揮組",
    "ptl": "二階段_巡邏組",
    "cp":  "二階段_路檢組"
}

# 根據 0408 專案更新預設值
DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年4月8日20至24時"
DEFAULT_PROJ    = "0408全國同步擴大取締酒後駕車與防制危險駕車及噪音車輛專案勤務"
DEFAULT_BRIEF   = "20時30分於分局二樓會議室召開" 

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "建興", "姓名": "駐區督察 孫三陽", "任務": "重點機動督導"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長黃長旗\n督察組督察員 黃中彥\n督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官郭文義", "任務": "指導各編組勤務執行及狀況處置"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟\n交通組警務員盧冠仁\n交通組警務員李峯甫\n交通組巡官郭勝隆\n交通組巡官羅千金\n交通組警員吳享運\n勤指中心警員張庭溱\n(代理人:巡官陳鵬翔)", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "行政組警務佐曾威仁\n人事室警員陳明祥\n主任蔡奇青\n執勤官李文章\n執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜"},
])

DEFAULT_PTL = pd.DataFrame([
    {"編組": "第一巡邏組", "無線電": "隆安50", "單位": "聖亭所", "服勤人員": "", "任務分工": "巡邏:\n於易發生酒駕、危險駕車路段(中豐路、中豐路中山段、中豐路上林段、大昌路一段、中正路)加強攔檢。"},
    {"編組": "第二巡邏組", "無線電": "隆安60", "單位": "龍潭所", "服勤人員": "", "任務分工": "巡邏:\n於易發生酒駕、危險駕車路段(聖亭路、聖亭路八德段、干城路、龍平路、湧光路、工五路)加強攔檢。"},
    {"編組": "第三巡邏組", "無線電": "隆安70", "單位": "中興所", "服勤人員": "", "任務分工": "巡邏:\n於易發生酒駕、危險駕車路段(北龍路、中興路、福龍路、武漢路、自由街、民主街)加強攔檢"},
    {"編組": "第四巡邏組", "無線電": "隆安80", "單位": "石門所", "服勤人員": "", "任務分工": "巡邏:\n於易發生酒駕、危險駕車路段(大昌路二段、中正路上華段、中正路三林段、中正路佳安段、中正路三坑段、龍源路大平段、文化路)加強攔檢"},
    {"編組": "第五巡邏組", "無線電": "隆安90", "單位": "高平所、三和所", "服勤人員": "", "任務分工": "巡邏:\n於易發生酒駕、危險駕車路段(中豐路高平段、中原路、龍源路、楊銅路反及龍新路三和段至三水段)加強攔檢。"},
    {"編組": "第六巡邏組", "無線電": "隆安990", "單位": "龍潭交通分隊", "服勤人員": "", "任務分工": "巡邏:\n於易發生酒駕、危險駕車路段(中豐路高平段、中原路、龍源路、工五路)加強攔檢。"},
])

DEFAULT_CHECKPOINT = pd.DataFrame([
    {"編組": "第一路檢組", "無線電": "隆安70", "單位": "中興所、高平所", "服勤人員": "", "任務分工": "路檢:中正路三坑段與美國路口(攔檢往龍潭市區方向車輛)\n*雨天備案:轄區治安要點巡邏"},
    {"編組": "第二路檢組", "無線電": "隆安50", "單位": "聖亭所、三和所", "服勤人員": "", "任務分工": "路檢:中正路三坑段與美國路口(攔檢往龍源路方向車輛)。\n*雨天備案:轄區治安要點巡邏。"},
    {"編組": "第一機動攔檢組", "無線電": "隆安990", "單位": "交通分隊", "服勤人員": "", "任務分工": "攔截圍捕:於中正路三坑段適當地點,機動攔檢迴轉規避車輛。\n*雨天備案:轄區治安要點巡邏。"},
    {"編組": "第二機動攔檢組", "無線電": "隆安60", "單位": "龍潭所", "服勤人員": "", "任務分工": "路檢:於中正路三坑段適當地點,機動攔檢迴轉規避車輛。\n*雨天備案:轄區治安要點巡邏。"},
    {"編組": "第三機動攔檢組", "無線電": "隆安80", "單位": "石門所", "服勤人員": "", "任務分工": "攔截圍捕:於美國路文化路來回梭巡,機動攔檢迴轉規避車輛\n*雨天備案:轄區治安要點巡邏"},
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
    except: pass
    return "19時30分至20時00分"

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

def init_sheets():
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        headers = {
            WS_MAP["set"]: [["Key", "Value"]],
            WS_MAP["cmd"]: [["職稱", "代號", "姓名", "任務"]],
            WS_MAP["ptl"]: [["編組", "無線電", "單位", "服勤人員", "任務分工"]],
            WS_MAP["cp"]:  [["編組", "無線電", "單位", "服勤人員", "任務分工"]]
        }
        for ws_name, head in headers.items():
            try:
                sh.worksheet(ws_name)
                st.sidebar.info(f"✔ {ws_name} 已存在")
            except:
                sh.add_worksheet(title=ws_name, rows="100", cols="20").update(head)
                st.sidebar.success(f"➕ 已建立 {ws_name}")
        load_data.clear()
        st.rerun()
    except Exception as e:
        st.error(f"初始化失敗：{e}")

@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, None, "權限不足"
        sh = client.open_by_key(SHEET_ID)
        
        ws_set = sh.worksheet(WS_MAP["set"])
        ws_cmd = sh.worksheet(WS_MAP["cmd"])
        ws_ptl = sh.worksheet(WS_MAP["ptl"])
        ws_cp  = sh.worksheet(WS_MAP["cp"])
            
        return (pd.DataFrame(ws_set.get_all_records()).fillna(""), 
                pd.DataFrame(ws_cmd.get_all_records()).fillna(""), 
                pd.DataFrame(ws_ptl.get_all_records()).fillna(""), 
                pd.DataFrame(ws_cp.get_all_records()).fillna(""), None)
    except Exception as e: return None, None, None, None, str(e)

def save_data(unit, time_str, project, briefing, df_cmd, df_ptl, df_cp):
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        
        ws_set = sh.worksheet(WS_MAP["set"])
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["unit_name", unit], ["plan_full_time", time_str], ["project_name", project], ["briefing_info", briefing]])
        
        for ws_name, df in [(WS_MAP["cmd"], df_cmd), (WS_MAP["ptl"], df_ptl), (WS_MAP["cp"], df_cp)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df_cleaned = df.dropna(how='all').fillna("")
            if not df_cleaned.empty:
                ws.update([df_cleaned.columns.tolist()] + df_cleaned.values.tolist())
        load_data.clear()
        return True
    except: return False

def generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=24, alignment=1, spaceAfter=8)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=10)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=18, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=18, alignment=0)
    style_middle_block = ParagraphStyle('MiddleBlock', fontName=font, fontSize=14, leading=22, spaceAfter=2*mm, alignment=TA_LEFT, leftIndent=5*mm)
    style_table_title = ParagraphStyle('TTitle', fontName=font, fontSize=16, alignment=1, leading=22)

    story.append(Paragraph(f"{unit}執行{project}勤務規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def clean(t): return safe_str(t).replace("\n", "<br/>").replace("、", "<br/>")
    def clean_text_only(t): return safe_str(t).replace("\n", "<br/>")

    data_cmd = [[Paragraph("<b>任 務 編 組</b>", style_table_title), '', '', ''],
                [Paragraph(f"<b>{h}</b>", style_cell) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([Paragraph(f"<b>{clean_text_only(r.get('職稱'))}</b>", style_cell), Paragraph(clean_text_only(r.get('代號')), style_cell), Paragraph(clean(r.get('姓名')), style_cell), Paragraph(clean_text_only(r.get('任務')), style_cell_left)])
    
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.12, page_width*0.28, page_width*0.45])
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('SPAN',(0,0),(-1,0)),('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t1)
    
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph("<b>📢 勤前教育：</b>", style_middle_block))
    story.append(Paragraph(f"{clean_text_only(briefing)}", style_middle_block))
    story.append(Spacer(1, 6*mm))

    story.append(Paragraph("<b>第一階段：21時至22時30分，機動巡邏</b>", style_middle_block))
    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "代號", "單位", "服勤人員", "任務分工"]]]
    for _, r in df_ptl.iterrows():
        data_ptl.append([Paragraph(clean_text_only(r.get('編組')), style_cell), Paragraph(clean_text_only(r.get('無線電')), style_cell), Paragraph(clean(r.get('單位')), style_cell), Paragraph(clean(r.get('服勤人員')), style_cell), Paragraph(clean_text_only(r.get('任務分工')), style_cell_left)])
    t2 = Table(data_ptl, colWidths=[page_width*0.15, page_width*0.12, page_width*0.13, page_width*0.20, page_width*0.40])
    t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t2)

    story.append(Spacer(1, 8*mm))
    story.append(Paragraph("<b>第二階段：22時30分至24時，定點路檢及機動攔檢</b>", style_middle_block))
    data_cp = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "代號", "單位", "服勤人員", "任務分工"]]]
    for _, r in df_cp.iterrows():
        task_text = f"{clean_text_only(r.get('任務分工'))}"
        data_cp.append([Paragraph(clean_text_only(r.get('編組')), style_cell), Paragraph(clean_text_only(r.get('無線電')), style_cell), Paragraph(clean(r.get('單位')), style_cell), Paragraph(clean(r.get('服勤人員')), style_cell), Paragraph(task_text, style_cell_left)])
    t3 = Table(data_cp, colWidths=[page_width*0.15, page_width*0.12, page_width*0.13, page_width*0.20, page_width*0.40])
    t3.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e6e6e6')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t3)
    doc.build(story)
    return buf.getvalue()

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
    story.append(Paragraph(f"{unit}執行{project}簽到表", style_title))
    meeting_range = parse_meeting_time(time_str)
    date_part = time_str.split('日')[0] + '日' if '日' in time_str else ""
    story.append(Paragraph(f"時間：{date_part}{meeting_range}", style_top_info))
    loc = str(briefing).strip() if "於" not in str(briefing) else str(briefing).strip().split("於")[1]
    story.append(Paragraph(f"地點：{loc}", style_top_info))
    story.append(Spacer(1, 3*mm))
    table_data = [[Paragraph("分局長：", style_cell_left), "", Paragraph("上級督導：", style_cell_left), ""],
                  [Paragraph("副分局長：", style_cell_left), "", "", ""],
                  [Paragraph("單位", style_cell), Paragraph("參加人員", style_cell), Paragraph("單位", style_cell), Paragraph("參加人員", style_cell)]]
    rows = [("交通組", "中興派出所"), ("勤務指揮中心", "石門派出所"), ("督察組", "高平派出所"), ("聖亭派出所", "三和派出所"), ("龍潭派出所", "龍潭交通分隊")]
    for l, r in rows: table_data.append([Paragraph(l, style_cell), "", Paragraph(r, style_cell), ""])
    t = Table(table_data, colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3], rowHeights=[18*mm, 18*mm, 10*mm] + [26*mm]*len(rows))
    t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), font), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('SPAN', (0,1), (3,1))]))
    story.append(t)
    doc.build(story)
    return buf.getvalue()

def send_report_email(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        msg["Subject"] = f"{unit}執行{project}規劃與簽到表_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為 PDF。", "plain", "utf-8"))
        p1 = generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp)
        part1 = MIMEBase("application", "pdf")
        part1.set_payload(p1)
        encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(unit+'規劃表.pdf')}")
        msg.attach(part1)
        p2 = generate_attendance_pdf(unit, project, time_str, briefing)
        part2 = MIMEBase("application", "pdf")
        part2.set_payload(p2)
        encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(unit+'簽到表.pdf')}")
        msg.attach(part2)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

def auto_assign_radio_code(df):
    if df.empty: return df
    base_prefixes = {"交通分隊": "99", "聖亭": "5", "龍潭": "6", "中興": "7", "石門": "8", "高平": "9", "三和": "3"}
    for idx, row in df.iterrows():
        unit, person = safe_str(row.get('單位')), safe_str(row.get('服勤人員'))
        if not unit: continue
        first_unit = re.split(r'[\n、 ]', unit.strip())[0]
        base_pfx = next((v for k, v in base_prefixes.items() if k in first_unit), "")
        if base_pfx:
            if "副所長" in person: df.at[idx, '無線電'] = f"隆安{base_pfx}2"
            elif "所長" in person: df.at[idx, '無線電'] = f"隆安{base_pfx}1"
            elif not safe_str(row.get('無線電')).startswith(f"隆安{base_pfx}"): df.at[idx, '無線電'] = f"隆安{base_pfx}0"
    return df

# 🌟 新增：智慧連動對齊演算法 (完美修正版)
def sync_personnel_data(df_ptl, df_cp):
    """拆解第一階段單位，並將人員精準映射至第二階段（支援換行與空白，並自動去除重複）"""
    if df_ptl.empty or df_cp.empty: return df_cp
    
    # 強化的分隔符號：支援 頓號、逗號、全半形空白、換行(\n)、斜線
    split_pattern = r'[、,，\s/]+'
    
    # 1. 拆解第一階段的單位名稱，建立 mapping 字典
    p_dict = {}
    for _, row in df_ptl.iterrows():
        unit_str = str(row.get('單位', '')).replace('龍潭交通分隊', '交通分隊')
        units = re.split(split_pattern, unit_str.strip())
        
        persons_str = str(row.get('服勤人員', '')).strip()
        # 提前把人員也獨立拆開
        current_persons = []
        for p in re.split(split_pattern, persons_str):
            p = p.strip()
            if p and p not in current_persons:
                current_persons.append(p)
        
        for u in units:
            u = u.strip()
            if u:
                # 若字典裡沒有這個單位，先建立空陣列
                if u not in p_dict:
                    p_dict[u] = []
                # 把人員「累加」進去，避免被覆蓋
                for p in current_persons:
                    if p not in p_dict[u]:
                        p_dict[u].append(p)
            
    # 2. 依照第二階段的單位進行組合
    df_cp_new = df_cp.copy()
    for idx, row in df_cp_new.iterrows():
        unit_str = str(row.get('單位', '')).replace('龍潭交通分隊', '交通分隊')
        units_cp = re.split(split_pattern, unit_str.strip())
        
        combined = []
        for u in units_cp:
            u = u.strip()
            # 把第一階段累積的人員名單全部倒出來
            if u in p_dict:
                for p in p_dict[u]:
                    if p not in combined: 
                        combined.append(p)
        
        # 若有組合出人員，則複寫進欄位中 (統一用頓號組合)
        if combined:
            df_cp_new.at[idx, '服勤人員'] = "、".join(combined)
            
    return df_cp_new

# --- 3. 主程式介面 ---
st.sidebar.title("🛠️ 雲端設定")
if st.sidebar.button("初始化/檢查雲端分頁"):
    init_sheets()

df_set, df_cmd, df_ptl, df_cp, err = load_data()

if err: 
    st.sidebar.warning(f"雲端尚未就緒: {err}")

d = {}
if df_set is not None and isinstance(df_set, pd.DataFrame) and not df_set.empty and df_set.shape[1] >= 2:
    try:
        keys = df_set.iloc[:, 0].astype(str).tolist()
        vals = df_set.iloc[:, 1].astype(str).tolist()
        d = dict(zip(keys, vals))
    except Exception as e:
        st.sidebar.error(f"讀取設定檔發生錯誤: {e}")

u = d.get("unit_name", DEFAULT_UNIT) if d.get("unit_name") else DEFAULT_UNIT
t = d.get("plan_full_time", DEFAULT_TIME) if d.get("plan_full_time") else DEFAULT_TIME
p = d.get("project_name", DEFAULT_PROJ) if d.get("project_name") else DEFAULT_PROJ
b = d.get("briefing_info", DEFAULT_BRIEF) if d.get("briefing_info") else DEFAULT_BRIEF

ed_cmd = df_cmd if (df_cmd is not None and not df_cmd.empty) else DEFAULT_CMD.copy()
ed_ptl = df_ptl if (df_ptl is not None and not df_ptl.empty) else DEFAULT_PTL.copy()
ed_cp  = df_cp  if (df_cp is not None and not df_cp.empty) else DEFAULT_CHECKPOINT.copy()

st.title("🚓 二階段勤務規劃系統 (專屬分頁版)")
c1, c2 = st.columns(2)
p_name = c1.text_input("專案名稱", p)
p_time = c2.text_input("勤務時間", t)

st.subheader("1. 指揮編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True).dropna(how='all').fillna("")
b_info = st.text_area("📢 勤前教育", b, height=70)

st.subheader("2. 勤務編組")
tab1, tab2 = st.tabs(["📍 第一階段：機動巡邏", "🚧 第二階段：定點路檢"])

# --- 第一階段 ---
with tab1:
    res_ptl = auto_assign_radio_code(st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True, key="ptl_v2").dropna(how='all').fillna(""))

# --- 第二階段 (加入連動機制) ---
with tab2:
    st.info("💡 點擊下方按鈕，系統會自動拆解比對「單位」名稱，將第一階段的人員精準帶入第二階段！(帶入後仍可手動微調)")
    if st.button("🔄 一鍵自動帶入第一階段人員"):
        # 進行計算並將結果存入 session_state
        st.session_state["synced_cp"] = sync_personnel_data(res_ptl, ed_cp)
        st.rerun() # 重新整理頁面以顯示最新資料
        
    # 如果有被同步過的新資料，就使用新資料；否則使用預設/雲端讀取的資料
    current_cp = st.session_state.get("synced_cp", ed_cp)
    res_cp = auto_assign_radio_code(st.data_editor(current_cp, num_rows="dynamic", use_container_width=True, key="cp_v2").dropna(how='all').fillna(""))
    
    # 使用者若在表格上手動打字修改，我們也更新回暫存，確保狀態一致
    st.session_state["synced_cp"] = res_cp

st.markdown("---")
col_dl1, col_dl2 = st.columns(2)
pdf_plan = generate_pdf_from_data(u, p_name, p_time, b_info, res_cmd, res_ptl, res_cp)
col_dl1.download_button("📝 下載勤務規劃表", data=pdf_plan, file_name=f"{u}_{p_name}.pdf", use_container_width=True)
pdf_attendance = generate_attendance_pdf(u, p_name, p_time, b_info)
col_dl2.download_button("🖋️ 下載人員簽到表", data=pdf_attendance, file_name=f"{u}_簽到表.pdf", use_container_width=True)

if st.button("💾 同步雲端專屬分頁並發送備份郵件", use_container_width=True):
    with st.spinner("同步中..."):
        if save_data(u, p_time, p_name, b_info, res_cmd, res_ptl, res_cp):
            ok, mail_err = send_report_email(u, p_name, p_time, b_info, res_cmd, res_ptl, res_cp)
            if ok: st.success("✅ 同步成功！")
            else: st.warning(f"⚠️ 雲端已同步，但郵件失敗: {mail_err}")
