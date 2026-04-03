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
st.set_page_config(page_title="雲端勤務規劃系統", layout="wide", page_icon="🚓")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

WS_MAP = {
    "set": "專案_設定",
    "cmd": "專案_指揮組",
    "ptl": "專案_巡邏組"
}

DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年3月28日19至23時"
DEFAULT_PROJ    = "0328「全市取締酒後駕車與防制危險駕車及噪音車輛」合併「取締改裝(噪音)車輛專案監、警、環聯合稽查」"
DEFAULT_BRIEF   = "19時30分於分局二樓會議室召開" 
DEFAULT_STATION = "時間:20時至23時\n地點:桃園市龍潭區中正路269號(龍星國民小學)大門口"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長施宇峰", "任務": "核定本勤務執行並重點機動督導"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "建興", "姓名": "駐區督察 孫三陽", "任務": "重點機動督導"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長黃長旗\n督察組督察員 黃中彥\n督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官郭文義", "任務": "指導各編組勤務執行及狀況處置"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟\n交通組警務員盧冠仁\n交通組警務員李峯甫\n交通組巡官郭勝隆\n交通組巡官羅千金\n交通組警員吳享運\n勤指中心警員張庭溱\n(代理人:巡官陳鵬翔)", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "行政組警務佐曾威仁\n人事室警員陳明祥\n主任蔡奇青\n執勤官李文章\n執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜"},
])

DEFAULT_PTL = pd.DataFrame([
    {"編組": "第一巡邏組", "無線電": "隆安52", "單位": "聖亭所", "服勤人員": "副所長邱品淳\n警員傅維強", "任務分工": "於中正路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第二巡邏組", "無線電": "隆安61", "單位": "龍潭所", "服勤人員": "所長孫祥愷\n警員沈庭禾\n警員周浚豪\n警員黃子軒", "任務分工": "於北龍路周邊易有噪音車輛滋擾聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第三巡邏組", "無線電": "隆安71", "單位": "中興所", "服勤人員": "所長董亦文\n警員徐毓汶\n警員蔡震東", "任務分工": "於龍新路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第四巡邏組", "無線電": "隆安83", "單位": "石門所", "服勤人員": "巡佐林偉政\n警員鄒詠如", "任務分工": "於神龍路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第五巡邏組", "無線電": "隆安93", "單位": "高平所\n三和所", "服勤人員": "警員邱春松\n警員唐銘聰", "任務分工": "於東龍路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第六巡邏組", "無線電": "隆安993", "單位": "龍潭交通分隊", "服勤人員": "小隊長林振生\n警員張登冠", "任務分工": "於中豐路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
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
            WS_MAP["ptl"]: [["編組", "無線電", "單位", "服勤人員", "任務分工"]]
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
        if client is None: return None, None, None, "權限不足"
        sh = client.open_by_key(SHEET_ID)
        
        ws_set = sh.worksheet(WS_MAP["set"])
        ws_cmd = sh.worksheet(WS_MAP["cmd"])
        ws_ptl = sh.worksheet(WS_MAP["ptl"])
        return pd.DataFrame(ws_set.get_all_records()).fillna(""), pd.DataFrame(ws_cmd.get_all_records()).fillna(""), pd.DataFrame(ws_ptl.get_all_records()).fillna(""), None
    except Exception as e: return None, None, None, str(e)

def save_data(unit, time_str, project, briefing, station, df_cmd, df_ptl):
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        
        ws_set = sh.worksheet(WS_MAP["set"])
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["unit_name", unit], ["plan_full_time", time_str], ["project_name", project], ["briefing_info", briefing], ["check_station", station]])
        
        for ws_name, df in [(WS_MAP["cmd"], df_cmd), (WS_MAP["ptl"], df_ptl)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df_cleaned = df.dropna(how='all').fillna("")
            if not df_cleaned.empty:
                ws.update([df_cleaned.columns.tolist()] + df_cleaned.values.tolist())
        load_data.clear()
        return True
    except: return False

# --- 3. PDF 生成功能 ---
A4_SIZE = (float(595.275), float(841.890))

def generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    font = _get_font()
    buf = io.BytesIO()
    
    margin_lr = float(12 * mm)
    margin_tb = float(15 * mm)
    
    doc = SimpleDocTemplate(
        buf, 
        pagesize=A4_SIZE, 
        leftMargin=margin_lr, 
        rightMargin=margin_lr, 
        topMargin=margin_tb, 
        bottomMargin=margin_tb
    )
    page_width = A4_SIZE[0] - (2 * margin_lr)
    story = []
    
    # 🌟 核心修復：將表格內文的 fontSize 統一調降為 12，給予中英文字符充足的喘息空間
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=24, alignment=1, spaceAfter=8, wordWrap='CJK')
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=10, wordWrap='CJK')
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=12, leading=16, alignment=1, wordWrap='CJK')
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=12, leading=16, alignment=0, wordWrap='CJK')
    style_middle_block = ParagraphStyle('MiddleBlock', fontName=font, fontSize=14, leading=22, spaceAfter=2*mm, alignment=TA_LEFT, leftIndent=5*mm, firstLineIndent=0, wordWrap='CJK')
    style_table_title = ParagraphStyle('TTitle', fontName=font, fontSize=16, alignment=1, leading=22, wordWrap='CJK')

    story.append(Paragraph(f"{unit}{project}勤務規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def clean(t): return str(t).replace("\n", "<br/>").replace("、", "<br/>")

    data_cmd = [[Paragraph("<b>任 務 編 組</b>", style_table_title), '', '', ''],
                [Paragraph(f"<b>{h}</b>", style_cell) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        # 🌟 將所有欄位（包含任務與代號）全面套用 clean 處理
        data_cmd.append([
            Paragraph(f"<b>{r.get('職稱','')}</b>", style_cell), 
            Paragraph(clean(r.get('代號','')), style_cell),
            Paragraph(clean(r.get('姓名','')), style_cell), 
            Paragraph(clean(r.get('任務','')), style_cell_left)
        ])
    
    # 🌟 重新調配比例，讓前面字數較少的欄位稍微加寬一點，並加上 repeatRows=2 允許跨頁表頭
    t1 = Table(data_cmd, colWidths=[page_width*0.14, page_width*0.14, page_width*0.27, page_width*0.45], repeatRows=2)
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('SPAN',(0,0),(-1,0)),
                            ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t1)
    
    briefing_clean = str(briefing).strip().replace('\n', '<br/>')
    station_clean = str(station).strip().replace('\n', '<br/>')

    story.append(Spacer(1, 6*mm))
    story.append(Paragraph("<b>📢 勤前教育：</b>", style_middle_block))
    story.append(Paragraph(f"{briefing_clean}", style_middle_block))
    story.append(Spacer(1, 2*mm)) 
    story.append(Paragraph("<b>🚧 環保局臨時檢驗站開設：</b>", style_middle_block))
    story.append(Paragraph(f"{station_clean}", style_middle_block))
    story.append(Spacer(1, 6*mm))

    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "代號", "單位", "服勤人員", "任務分工"]]]
    for _, r in df_ptl.iterrows():
        task = f"{r.get('任務分工','')}<br/><font color='blue' size='11'>*雨備方案：各治安要點巡邏。</font>"
        
        data_ptl.append([
            Paragraph(clean(r.get('編組','')), style_cell), 
            Paragraph(clean(r.get('無線電','')), style_cell), 
            Paragraph(clean(r.get('單位','')), style_cell), 
            Paragraph(clean(r.get('服勤人員','')), style_cell), 
            Paragraph(task, style_cell_left)
        ])
    
    # 🌟 同樣調整巡邏組的比例與加上 repeatRows=1
    t2 = Table(data_ptl, colWidths=[page_width*0.14, page_width*0.14, page_width*0.16, page_width*0.20, page_width*0.36], repeatRows=1)
    t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('FONTSIZE',(0,0),(-1,-1),12),('ALIGN',(0,1),(1,-1),'CENTER'),
                            ('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t2)
    doc.build(story)
    return buf.getvalue()

def generate_attendance_pdf(unit, project, time_str, briefing):
    font = _get_font()
    buf = io.BytesIO()
    
    margin_lr = float(15 * mm)
    margin_tb = float(15 * mm)
    
    doc = SimpleDocTemplate(
        buf, 
        pagesize=A4_SIZE, 
        leftMargin=margin_lr, 
        rightMargin=margin_lr, 
        topMargin=margin_tb, 
        bottomMargin=margin_tb
    )
    page_width = A4_SIZE[0] - (2 * margin_lr)
    story = []

    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, leading=22, alignment=1, spaceAfter=8, wordWrap='CJK')
    style_top_info = ParagraphStyle('TopInfo', fontName=font, fontSize=12, leading=18, alignment=0, wordWrap='CJK')
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=24, alignment=1, wordWrap='CJK')
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=24, alignment=0, wordWrap='CJK') 
    style_note = ParagraphStyle('Note', fontName=font, fontSize=11, leading=15, alignment=0, wordWrap='CJK')

    story.append(Paragraph(f"{unit}執行{project}勤前教育會議人員簽到表", style_title))
    
    meeting_range = parse_meeting_time(time_str)
    date_part = time_str.split('日')[0] + '日' if '日' in time_str else ""
    story.append(Paragraph(f"時間：{date_part}{meeting_range}", style_top_info))
    
    loc = str(briefing).strip() if "於" not in str(briefing) else str(briefing).strip().split("於")[1]
    story.append(Paragraph(f"地點：{loc}", style_top_info))
    story.append(Spacer(1, 3*mm))

    table_data = []
    table_data.append([Paragraph("分局長：", style_cell_left), "", Paragraph("上級督導：", style_cell_left), ""])
    table_data.append([Paragraph("副分局長：", style_cell_left), "", "", ""])
    table_data.append([Paragraph("單位", style_cell), Paragraph("參加人員", style_cell), 
                        Paragraph("單位", style_cell), Paragraph("參加人員", style_cell)])
    
    rows = [
        ("交通組", "中興派出所"),
        ("勤務指揮中心", "石門派出所"),
        ("督察組", "高平派出所"),
        ("聖亭派出所", "三和派出所"),
        ("龍潭派出所", "龍潭交通分隊")
    ]
    for l, r in rows:
        table_data.append([Paragraph(l, style_cell), "", Paragraph(r, style_cell), ""])
    
    data_row_height = 20 * mm * (4/3)
    row_heights = [18*mm, 18*mm, 10*mm] + [data_row_height] * len(rows)
    
    t = Table(table_data, colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3], rowHeights=row_heights)
    t.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('ALIGN', (2,0), (2,0), 'LEFT'),
        ('ALIGN', (0,1), (0,1), 'LEFT'),
        ('SPAN', (0,1), (3,1)), 
        ('BACKGROUND', (0,2), (0,2), colors.whitesmoke),
        ('BACKGROUND', (2,2), (2,2), colors.whitesmoke),
    ]))
    story.append(t)

    story.append(Spacer(1, 5*mm))
    story.append(Paragraph("備註：請將行動電話調整為靜音。", style_note))

    doc.build(story)
    return buf.getvalue()

# --- 4. 寄信功能 ---
def send_report_email(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    try:
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        msg["Subject"] = f"勤務規劃與簽到表_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為最新的勤務規劃表與人員簽到表 PDF。", "plain", "utf-8"))
        
        pdf1 = generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl)
        part1 = MIMEBase("application", "pdf")
        part1.set_payload(pdf1)
        encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote('規劃表.pdf')}")
        msg.attach(part1)
        
        pdf2 = generate_attendance_pdf(unit, project, time_str, briefing)
        part2 = MIMEBase("application", "pdf")
        part2.set_payload(pdf2)
        encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote('簽到表.pdf')}")
        msg.attach(part2)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# --- 5. 主介面與智慧代號計算邏輯 ---
st.sidebar.title("🛠️ 雲端設定")
if st.sidebar.button("初始化/檢查雲端分頁"):
    init_sheets()

if st.sidebar.button("⚠️ 強制重置為最新專案資料 (覆蓋雲端)"):
    with st.spinner("重置中..."):
        save_data(DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_PROJ, DEFAULT_BRIEF, DEFAULT_STATION, DEFAULT_CMD, DEFAULT_PTL)
        st.cache_data.clear()
        st.rerun()

df_set, df_cmd, df_ptl, err = load_data()

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
s = d.get("check_station", DEFAULT_STATION) if d.get("check_station") else DEFAULT_STATION

ed_cmd = df_cmd if (df_cmd is not None and not df_cmd.empty) else DEFAULT_CMD.copy()
ed_ptl = df_ptl if (df_ptl is not None and not df_ptl.empty) else DEFAULT_PTL.copy()

st.title("🚓 專案勤務規劃管理系統")
c1, c2 = st.columns(2)

p_name_raw = c1.text_input("專案名稱", p)
p_time = c2.text_input("勤務時間", t)

# 🌟 連動邏輯：擷取月份與日期，覆蓋專案名稱前 4 碼
match = re.search(r"(\d+)月(\d+)日", p_time)
if match:
    mm = str(match.group(1)).zfill(2)
    dd = str(match.group(2)).zfill(2)
    date_prefix = f"{mm}{dd}"
    
    if re.match(r"^\d{4}", p_name_raw):
        p_name = f"{date_prefix}{p_name_raw[4:]}"
    else:
        p_name = f"{date_prefix}{p_name_raw}"
else:
    p_name = p_name_raw

st.subheader("1. 指揮編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True).dropna(how='all').fillna("")

b_info = st.text_area("📢 勤前教育", b, height=70)
s_info = st.text_area("🚧 環保局臨時檢驗站開設", s, height=70)

st.subheader("2. 巡邏編組")
st.info("💡 系統會自動依據「單位」（優先看上方單位）及「服勤人員」計算無線電代號（所長為1、副所長為2）。未標示主副官時將保留您輸入的代號，或自動帶入單位基礎代號。")
res_ptl_raw = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True).dropna(how='all').fillna("")

def auto_assign_radio_code(df):
    base_prefixes = {"交通分隊": "99", "聖亭": "5", "龍潭": "6", "中興": "7", "石門": "8", "高平": "9", "三和": "3"}
    for idx, row in df.iterrows():
        unit = str(row.get('單位', ''))
        person = str(row.get('服勤人員', ''))
        current_radio = str(row.get('無線電', '')).strip()
        if not unit: continue
        first_unit = re.split(r'[\n、 ]', unit.strip())[0]
        base_pfx = ""
        for k, v in base_prefixes.items():
            if k in first_unit:
                base_pfx = v
                break
        if base_pfx:
            if "副所長" in person: df.at[idx, '無線電'] = f"隆安{base_pfx}2"
            elif "所長" in person: df.at[idx, '無線電'] = f"隆安{base_pfx}1"
            else:
                if not current_radio.startswith(f"隆安{base_pfx}"): df.at[idx, '無線電'] = f"隆安{base_pfx}0"
    return df

res_ptl = auto_assign_radio_code(res_ptl_raw.copy())

def get_html():
    style = "<style>body{font-family:'標楷體';padding:10px;} th,td{border:1px solid black;padding:6px;font-size:12pt;text-align:center;} .middle-block{font-size:12pt;margin:15px 0 15px 20px;line-height:1.6; text-align:left;}</style>"
    
    html = f"<html>{style}<body><h3 style='text-align:center'>{u}<br>{p_name}勤務規劃表</h3><div style='text-align:right'><b>時間：{p_time}</b></div><table><tr><th colspan='4'>任 務 編 組</th></tr>"
    
    for _, r in res_cmd.iterrows():
        html += f"<tr><td><b>{r.get('職稱','')}</b></td><td>{r.get('代號','')}</td><td>{str(r.get('姓名','')).replace('、','<br>')}</td><td style='text-align:left'>{r.get('任務','')}</td></tr>"
    b_html = str(b_info).strip().replace('\n', '<br>')
    s_html = str(s_info).strip().replace('\n', '<br>')
    html += f"</table><div class='middle-block'><b>📢 勤前教育：</b><br>{b_html}<br><br><b>🚧 環保局臨時檢驗站開設：</b><br>{s_html}</div>"
    html += "<table><tr><th>編組</th><th>代號</th><th>單位</th><th>人員</th><th>任務</th></tr>"
    for _, r in res_ptl.iterrows():
        html += f"<tr><td style='white-space:nowrap;'>{r.get('編組','')}</td><td style='white-space:nowrap;'>{r.get('無線電','')}</td><td>{str(r.get('單位','')).replace('、','<br>')}</td><td>{str(r.get('服勤人員','')).replace('、','<br>')}</td><td style='text-align:left'>{r.get('任務分工','')}</td></tr>"
    return html + "</table></body></html>"

st.markdown("---")
st.subheader("📄 報表下載與同步")

with st.expander("點擊展開即時預覽 (代號已依據所長/副所長自動校正)"):
    st.components.v1.html(get_html(), height=500, scrolling=True)

col_dl1, col_dl2 = st.columns(2)

pdf_plan = generate_pdf_from_data(u, p_name, p_time, b_info, s_info, res_cmd, res_ptl)
col_dl1.download_button("📝 下載 1.勤務規劃表", data=pdf_plan, file_name=f"規劃表_{datetime.now().strftime('%m%d')}.pdf", use_container_width=True)

pdf_attendance = generate_attendance_pdf(u, p_name, p_time, b_info)
col_dl2.download_button("🖋️ 下載 2.人員簽到表", data=pdf_attendance, file_name=f"簽到表_{datetime.now().strftime('%m%d')}.pdf", use_container_width=True)

if st.button("💾 同步雲端並發送備份郵件 (含兩份 PDF)", use_container_width=True):
    with st.spinner("同步中..."):
        save_data(u, p_time, p_name, b_info, s_info, res_cmd, res_ptl)
        ok, mail_err = send_report_email(u, p_name, p_time, b_info, s_info, res_cmd, res_ptl)
        if ok: st.success("✅ 同步成功，規劃表與簽到表已寄至信箱！")
        else: st.warning(f"⚠️ 雲端已同步，但郵件失敗: {mail_err}")
