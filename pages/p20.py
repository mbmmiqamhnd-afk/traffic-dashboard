import streamlit as st
import pandas as pd
import gspread
from gspread.exceptions import WorksheetNotFound, APIError
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib, io, os, traceback
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

# --- 1. 頁面設定 (必須是第一個 Streamlit 指令) ---
st.set_page_config(page_title="聯合稽查(二階段)勤務規劃系統", layout="wide", page_icon="🚓")

# 呼叫側邊欄
try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    pass

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

WS_MAP = {
    "set": "專案_設定",
    "cmd": "專案_指揮組",
    "ptl": "專案_巡邏組",
    "cp":  "專案_路檢組"  # 新增路檢組工作表
}

DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年3月28日19至23時"
DEFAULT_PROJ    = "0328「全市取締酒後駕車與防制危險駕車及噪音車輛」合併「取締改裝(噪音)車輛專案監、警、環聯合稽查」"
DEFAULT_BRIEF   = "19時30分於分局二樓會議室召開"
DEFAULT_STATION = "時間:20時至23時\n地點:桃園市龍潭區大昌路一段277號(龍潭區警政聯合辦公大樓)廣場"
DEFAULT_P1_DESC = "第一階段：21時至22時30分，機動巡邏"
DEFAULT_P2_DESC = "第二階段：22時30分至24時，定點路檢及機動攔檢"

EXPECTED_PTL_COLS = ["編組", "無線電代號", "單位", "職別", "姓名", "任務分工", "巡邏路段"]
EXPECTED_CP_COLS  = ["編組", "無線電代號", "單位", "職別", "姓名", "任務分工", "路檢地點"]

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "無線電代號": "隆安1", "負責人員": "分局長施宇峰", "任務": "核定本勤務執行並重點機動督導"},
    {"職稱": "副指揮官", "無線電代號": "隆安2", "負責人員": "副分局長何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "無線電代號": "隆安3", "負責人員": "副分局長蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "作業及督巡組", "無線電代號": "隆安13", "負責人員": "交通組組長 楊孟竟", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
])

DEFAULT_PTL = pd.DataFrame([
    {"編組": "第一巡邏組", "無線電代號": "隆安52", "單位": "聖亭所", "職別": "副所長", "姓名": "曹培翔", "任務分工": "帶班兼蒐證", "巡邏路段": "於中正路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第一巡邏組", "無線電代號": "隆安52", "單位": "聖亭所", "職別": "警員", "姓名": "詹宗澤", "任務分工": "攔檢盤查", "巡邏路段": "於中正路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"}
])

DEFAULT_CP = pd.DataFrame([
    {"編組": "第一路檢組", "無線電代號": "隆安52", "單位": "聖亭所", "職別": "副所長", "姓名": "曹培翔", "任務分工": "帶班兼蒐證", "路檢地點": "中正路與五福街口"},
    {"編組": "第一路檢組", "無線電代號": "隆安52", "單位": "聖亭所", "職別": "警員", "姓名": "詹宗澤", "任務分工": "攔檢盤查", "路檢地點": "中正路與五福街口"}
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

def clean_df(df):
    if df is None or df.empty: return pd.DataFrame()
    cleaned = df.replace(r'^\s*$', pd.NA, regex=True).dropna(how='all').fillna("")
    return cleaned

def parse_meeting_time(time_str):
    try:
        match = re.search(r"(\d+)至", time_str)
        if match:
            start_hour = int(match.group(1))
            end_hour = start_hour + 1
            return f"{start_hour}時30分至{end_hour}時00分"
    except: pass
    return "19時30分至20時00分"

def parse_briefing_time_range(briefing_str):
    try:
        match = re.search(r"(\d+)時(?:(\d+)分)?", briefing_str)
        if match:
            hour = int(match.group(1))
            minute = int(match.group(2)) if match.group(2) else 0
            end_minute = minute + 30
            end_hour = hour
            if end_minute >= 60:
                end_minute -= 60
                end_hour += 1
            return f"{hour}時{minute:02d}分至{end_hour}時{end_minute:02d}分"
    except: pass
    return "19時30分至20時00分"

def safe_str(val):
    if pd.isna(val) or val is None or str(val).strip().lower() == "nan": return ""
    return str(val)

def clean_df_to_list(df):
    return df.astype(str).values.tolist()

def get_merge_styles(df, merge_cols):
    span_styles = []
    if df.empty: return span_styles
    cols_list = df.columns.tolist()
    for col_name in merge_cols:
        if col_name not in cols_list: continue
        c_idx = cols_list.index(col_name)
        start_idx = 0
        while start_idx < len(df):
            val = str(df.iloc[start_idx][col_name]).strip()
            if not val: 
                start_idx += 1
                continue
            end_idx = start_idx
            while end_idx + 1 < len(df):
                next_val = str(df.iloc[end_idx + 1][col_name]).strip()
                if next_val == val:
                    end_idx += 1
                else: break
            if end_idx > start_idx:
                span_styles.append(('SPAN', (c_idx, start_idx + 1), (c_idx, end_idx + 1)))
            start_idx = end_idx + 1
    return span_styles

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
            WS_MAP["cmd"]: [["職稱", "無線電代號", "負責人員", "任務"]],
            WS_MAP["ptl"]: [EXPECTED_PTL_COLS],
            WS_MAP["cp"]:  [EXPECTED_CP_COLS]
        }
        for ws_name, head in headers.items():
            try:
                sh.worksheet(ws_name)
                st.sidebar.info(f"✔ {ws_name} 已存在")
            except:
                sh.add_worksheet(title=ws_name, rows="100", cols="20").update(range_name='A1', values=head)
                st.sidebar.success(f"➕ 已建立 {ws_name}")
        st.cache_data.clear()
        st.rerun()
    except Exception as e:
        st.error(f"初始化失敗：{e}")

@st.cache_data(ttl=600)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, None, "權限不足"
        sh = client.open_by_key(SHEET_ID)
        
        try: df_set = pd.DataFrame(sh.worksheet(WS_MAP["set"]).get_all_records()).fillna("")
        except: df_set = pd.DataFrame()
        try: df_cmd = pd.DataFrame(sh.worksheet(WS_MAP["cmd"]).get_all_records()).fillna("")
        except: df_cmd = pd.DataFrame()
        try: df_ptl = pd.DataFrame(sh.worksheet(WS_MAP["ptl"]).get_all_records()).fillna("")
        except: df_ptl = pd.DataFrame()
        try: df_cp  = pd.DataFrame(sh.worksheet(WS_MAP["cp"]).get_all_records()).fillna("")
        except: df_cp  = pd.DataFrame()

        return df_set, df_cmd, df_ptl, df_cp, None
    except Exception as e: return None, None, None, None, str(e)

def save_data(unit, time_str, project, briefing, station, p1_desc, p2_desc, df_cmd, df_ptl, df_cp):
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        
        try: ws_set = sh.worksheet(WS_MAP["set"])
        except WorksheetNotFound: ws_set = sh.add_worksheet(title=WS_MAP["set"], rows="50", cols="5")
        ws_set.clear()
        ws_set.update(range_name='A1', values=[
            ["Key", "Value"], ["unit_name", unit], ["plan_full_time", time_str], ["project_name", project], 
            ["briefing_info", briefing], ["check_station", station], ["phase1_desc", p1_desc], ["phase2_desc", p2_desc]
        ])
        
        for ws_name, df, expected_cols in [(WS_MAP["cmd"], df_cmd, ["職稱", "無線電代號", "負責人員", "任務"]), 
                                           (WS_MAP["ptl"], df_ptl, EXPECTED_PTL_COLS), 
                                           (WS_MAP["cp"], df_cp, EXPECTED_CP_COLS)]:
            try: ws = sh.worksheet(ws_name)
            except WorksheetNotFound: ws = sh.add_worksheet(title=ws_name, rows="100", cols="20")
            ws.clear()
            df_cleaned = df.dropna(how='all').fillna("")
            if not df_cleaned.empty:
                cols = df_cleaned.columns.tolist()
                ws.update(range_name='A1', values=[cols] + clean_df_to_list(df_cleaned))
            else:
                ws.update(range_name='A1', values=[expected_cols])
                
        st.cache_data.clear()
        return True
    except Exception as e: 
        st.error(f"儲存失敗：{e}")
        return False

# --- 3. PDF 生成功能 ---
A4_SIZE = (float(595.275), float(841.890))

def add_page_number(canvas, doc):
    canvas.saveState()
    canvas.setFont(_get_font(), 11)
    page_num = canvas.getPageNumber()
    text = f"- 第 {page_num} 頁 -"
    canvas.drawCentredString(A4_SIZE[0] / 2.0, 10 * mm, text)
    canvas.restoreState()

def generate_pdf_from_data(unit, project, time_str, briefing, station, p1_desc, p2_desc, df_cmd, df_ptl, df_cp):
    font = _get_font()
    buf = io.BytesIO()
    margin_lr, margin_tb = float(12 * mm), float(15 * mm)
    doc = SimpleDocTemplate(buf, pagesize=A4_SIZE, leftMargin=margin_lr, rightMargin=margin_lr, topMargin=margin_tb, bottomMargin=margin_tb)
    page_width = A4_SIZE[0] - (2 * margin_lr)
    story = []

    style_title        = ParagraphStyle('Title',       fontName=font, fontSize=18, leading=24, alignment=1, spaceAfter=8,   wordWrap='CJK')
    style_info         = ParagraphStyle('Info',        fontName=font, fontSize=12, alignment=2, spaceAfter=10,              wordWrap='CJK')
    style_cell         = ParagraphStyle('Cell',        fontName=font, fontSize=13, leading=18, alignment=1,                 wordWrap='CJK')
    style_cell_left    = ParagraphStyle('CellLeft',    fontName=font, fontSize=13, leading=18, alignment=0,                 wordWrap='CJK')
    style_middle_block = ParagraphStyle('MiddleBlock', fontName=font, fontSize=14, leading=22, spaceAfter=2*mm, alignment=TA_LEFT, leftIndent=5*mm, wordWrap='CJK')
    style_table_title  = ParagraphStyle('TTitle',      fontName=font, fontSize=16, alignment=1, leading=22,                 wordWrap='CJK')

    story.append(Paragraph(f"{unit}{project}勤務規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))

    def clean(t): return safe_str(t).replace("\n", "<br/>").replace("、", "<br/>")

    # -- 指揮組 --
    df_cmd = clean_df(df_cmd)
    data_cmd = [[Paragraph("<b>任 務 編 組</b>", style_table_title), '', '', ''],
                [Paragraph(f"<b>{h}</b>", style_cell) for h in ["職稱", "無線電代號", "負責人員", "任務"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([
            Paragraph(f"<b>{r.get('職稱','')}</b>", style_cell),
            Paragraph(clean(r.get('無線電代號','')), style_cell),
            Paragraph(clean(r.get('負責人員','')), style_cell),
            Paragraph(clean(r.get('任務','')), style_cell_left)
        ])

    t1 = Table(data_cmd, colWidths=[page_width*0.14, page_width*0.14, page_width*0.35, page_width*0.37], repeatRows=2)
    t1.setStyle(TableStyle([
        ('FONTNAME',   (0,0), (-1,-1), font),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.black),
        ('SPAN',       (0,0), (-1,0)),
        ('BACKGROUND', (0,0), (-1,1),  colors.HexColor('#f2f2f2')),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t1)

    story.append(Spacer(1, 6*mm))
    story.append(Paragraph("<b>📢 勤前教育：</b>", style_middle_block))
    story.append(Paragraph(str(briefing).strip().replace('\n', '<br/>'), style_middle_block))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph("<b>🚧 環保局臨時檢驗站開設：</b>", style_middle_block))
    story.append(Paragraph(str(station).strip().replace('\n', '<br/>'), style_middle_block))
    story.append(Spacer(1, 6*mm))

    # -- 第一階段 (巡邏組) --
    df_ptl = clean_df(df_ptl)
    if not df_ptl.empty:
        story.append(Paragraph(f"<b>{p1_desc}</b>", style_middle_block))
        span_styles_ptl = get_merge_styles(df_ptl, ["編組", "無線電代號", "單位", "巡邏路段"])
        data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in EXPECTED_PTL_COLS]]
        
        for _, r in df_ptl.iterrows():
            task_route = f"{r.get('巡邏路段','')}<br/><font color='blue' size='11'>*雨備方案：各治安要點巡邏。</font>"
            data_ptl.append([
                Paragraph(clean(r.get('編組','')), style_cell),
                Paragraph(clean(r.get('無線電代號','')), style_cell),
                Paragraph(clean(r.get('單位','')), style_cell),
                Paragraph(clean(r.get('職別','')), style_cell),
                Paragraph(clean(r.get('姓名','')), style_cell),
                Paragraph(clean(r.get('任務分工','')), style_cell),
                Paragraph(task_route, style_cell_left)
            ])
            
        t2 = Table(data_ptl, colWidths=[page_width*0.11, page_width*0.11, page_width*0.12, page_width*0.10, page_width*0.12, page_width*0.13, page_width*0.31], repeatRows=1)
        base_style_ptl = [
            ('FONTNAME',   (0,0), (-1,-1), font),
            ('FONTSIZE',   (0,0), (-1,-1), 13),
            ('ALIGN',      (0,1), (5,-1),  'CENTER'),
            ('GRID',       (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0),  colors.HexColor('#f2f2f2')),
            ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ]
        t2.setStyle(TableStyle(base_style_ptl + span_styles_ptl))
        story.append(t2)
        story.append(Spacer(1, 6*mm))

    # -- 第二階段 (路檢組) --
    df_cp = clean_df(df_cp)
    if not df_cp.empty:
        story.append(Paragraph(f"<b>{p2_desc}</b>", style_middle_block))
        span_styles_cp = get_merge_styles(df_cp, ["編組", "無線電代號", "單位", "路檢地點"])
        data_cp = [[Paragraph(f"<b>{h}</b>", style_cell) for h in EXPECTED_CP_COLS]]
        
        for _, r in df_cp.iterrows():
            data_cp.append([
                Paragraph(clean(r.get('編組','')), style_cell),
                Paragraph(clean(r.get('無線電代號','')), style_cell),
                Paragraph(clean(r.get('單位','')), style_cell),
                Paragraph(clean(r.get('職別','')), style_cell),
                Paragraph(clean(r.get('姓名','')), style_cell),
                Paragraph(clean(r.get('任務分工','')), style_cell),
                Paragraph(clean(r.get('路檢地點','')), style_cell_left)
            ])
            
        t3 = Table(data_cp, colWidths=[page_width*0.11, page_width*0.11, page_width*0.12, page_width*0.10, page_width*0.12, page_width*0.13, page_width*0.31], repeatRows=1)
        base_style_cp = [
            ('FONTNAME',   (0,0), (-1,-1), font),
            ('FONTSIZE',   (0,0), (-1,-1), 13),
            ('ALIGN',      (0,1), (5,-1),  'CENTER'),
            ('GRID',       (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0),  colors.HexColor('#e6e6e6')),
            ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ]
        t3.setStyle(TableStyle(base_style_cp + span_styles_cp))
        story.append(t3)
        
    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    return buf.getvalue()

def generate_attendance_pdf(unit, project, time_str, briefing):
    font = _get_font()
    buf = io.BytesIO()
    margin_lr, margin_tb = float(15 * mm), float(15 * mm)
    doc = SimpleDocTemplate(buf, pagesize=A4_SIZE, leftMargin=margin_lr, rightMargin=margin_lr, topMargin=margin_tb, bottomMargin=margin_tb)
    page_width = A4_SIZE[0] - (2 * margin_lr)
    story = []

    style_title    = ParagraphStyle('Title',    fontName=font, fontSize=16, leading=22, alignment=1, spaceAfter=8, wordWrap='CJK')
    style_top_info = ParagraphStyle('TopInfo',  fontName=font, fontSize=12, leading=18, alignment=0,               wordWrap='CJK')
    style_cell     = ParagraphStyle('Cell',     fontName=font, fontSize=14, leading=24, alignment=1,               wordWrap='CJK')
    style_cell_left= ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=24, alignment=0,               wordWrap='CJK')
    style_note     = ParagraphStyle('Note',     fontName=font, fontSize=11, leading=15, alignment=0,               wordWrap='CJK')

    story.append(Paragraph(f"{unit}執行{project}勤前教育會議人員簽到表", style_title))
    meeting_range = parse_briefing_time_range(str(briefing))
    date_part = time_str.split('日')[0] + '日' if '日' in time_str else ""
    story.append(Paragraph(f"時間：{date_part}{meeting_range}", style_top_info))
    loc = str(briefing).strip() if "於" not in str(briefing) else str(briefing).strip().split("於")[1]
    story.append(Paragraph(f"地點：{loc}", style_top_info))
    story.append(Spacer(1, 3*mm))

    table_data = [
        [Paragraph("分局長：", style_cell_left), "", Paragraph("上級督導：", style_cell_left), ""],
        [Paragraph("副分局長：", style_cell_left), "", "", ""],
        [Paragraph("<b>單位</b>", style_cell), Paragraph("<b>參加人員</b>", style_cell), Paragraph("<b>單位</b>", style_cell), Paragraph("<b>參加人員</b>", style_cell)]
    ]
    
    rows = [
        ("交通組", "中興派出所"), 
        ("督察組", "石門派出所"), 
        ("勤務指揮中心", "高平派出所"), 
        ("聖亭派出所", "三和派出所"), 
        ("龍潭派出所", "龍潭交通分隊")
    ]
    
    for l, r in rows:
        table_data.append([Paragraph(l, style_cell), "", Paragraph(r, style_cell), ""])

    row_heights = [18*mm, 18*mm, 10*mm] + [25*mm] * len(rows)
    t = Table(table_data, colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3], rowHeights=row_heights)
    t.setStyle(TableStyle([
        ('FONTNAME',   (0,0), (-1,-1), font),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
        ('ALIGN',      (0,0), (0,0),   'LEFT'),
        ('ALIGN',      (2,0), (2,0),   'LEFT'),
        ('ALIGN',      (0,1), (0,1),   'LEFT'),
        ('SPAN',       (0,1), (3,1)),
        ('BACKGROUND', (0,2), (3,2),   colors.whitesmoke),
    ]))
    story.append(t)
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph("備註：請將行動電話調整為靜音。", style_note))
    doc.build(story)
    return buf.getvalue()

# --- 4. 寄信功能 ---
def send_report_email(unit, project, time_str, briefing, station, p1_desc, p2_desc, df_cmd, df_ptl, df_cp):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        max_msg = MIMEMultipart()
        max_msg["From"], max_msg["To"], max_msg["Subject"] = sender, sender, f"{unit}勤務規劃與簽到表_{datetime.now().strftime('%m%d')}"
        max_msg.attach(MIMEText("附件為最新的勤務規劃表與人員簽到表 PDF。", "plain", "utf-8"))
        
        plan_filename = f"{unit}{project}勤務規劃表.pdf"
        p1 = generate_pdf_from_data(unit, project, time_str, briefing, station, p1_desc, p2_desc, df_cmd, df_ptl, df_cp)
        part1 = MIMEBase("application", "pdf"); part1.set_payload(p1); encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(plan_filename)}"); max_msg.attach(part1)
        
        attendance_filename = f"{unit}勤務簽到表.pdf"
        p2 = generate_attendance_pdf(unit, project, time_str, briefing)
        part2 = MIMEBase("application", "pdf"); part2.set_payload(p2); encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(attendance_filename)}"); max_msg.attach(part2)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, max_msg.as_string())
        return True, None
    except Exception as e: return False, str(e)


# --- 自動指派無線電代號 ---
def auto_assign_radio_code(df):
    if df is None or df.empty: return df
    df_copy = df.copy()
    base_prefixes = {"交通分隊": "99", "聖亭": "5", "龍潭": "6", "中興": "7", "石門": "8", "高平": "9", "三和": "3"}
    
    active_group_name = ""
    active_group_radio = ""
    
    for idx, row in df_copy.iterrows():
        group_name = safe_str(row.get('編組')).strip()
        unit = safe_str(row.get('單位')).strip()
        person = safe_str(row.get('姓名')).strip()
        rank = safe_str(row.get('職別')).strip()
        current_radio = safe_str(row.get('無線電代號')).strip()
        
        clean_unit = re.sub(r'\s+', '', unit)
        clean_rank = re.sub(r'\s+', '', rank)
        clean_person = re.sub(r'\s+', '', person)
        
        if idx == 0 or (group_name != "" and group_name != active_group_name):
            if group_name != "":
                active_group_name = group_name
            
            base_pfx = next((v for k, v in base_prefixes.items() if k in clean_unit), "")
            expected_radio = ""
            
            if base_pfx:
                if "副所長" in clean_rank or "小隊長" in clean_rank or "副所長" in clean_person or "小隊長" in clean_person: 
                    expected_radio = f"隆安{base_pfx}2"
                elif "所長" in clean_rank or "分隊長" in clean_rank or "所長" in clean_person or "分隊長" in clean_person: 
                    expected_radio = f"隆安{base_pfx}1"
                else: 
                    expected_radio = f"隆安{base_pfx}0"
            
            if current_radio != "":
                active_group_radio = current_radio
            else:
                active_group_radio = expected_radio
        
        if '無線電代號' in df_copy.columns:
            df_copy.at[idx, '無線電代號'] = active_group_radio
            
    return df_copy


# --- 5. 主介面 ---
st.sidebar.title("🛠️ 雲端設定")
if st.sidebar.button("初始化/檢查雲端分頁"): init_sheets()
if st.sidebar.button("⚠️ 強制重置為預設資料 (覆蓋雲端)"):
    with st.spinner("重置中..."):
        save_data(DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_PROJ, DEFAULT_BRIEF, DEFAULT_STATION, DEFAULT_P1_DESC, DEFAULT_P2_DESC, DEFAULT_CMD, DEFAULT_PTL, DEFAULT_CP)
        if "ptl_editable_df" in st.session_state: del st.session_state.ptl_editable_df
        st.cache_data.clear()
        st.rerun()

df_set, df_cmd, df_ptl, df_cp, err = load_data()

# 舊資料相容處理 (無線電 -> 無線電代號)
if isinstance(df_ptl, pd.DataFrame) and not df_ptl.empty:
    if '無線電' in df_ptl.columns and '無線電代號' not in df_ptl.columns: df_ptl['無線電代號'] = df_ptl['無線電']
    for c in EXPECTED_PTL_COLS:
        if c not in df_ptl.columns: df_ptl[c] = ""
    df_ptl = df_ptl[EXPECTED_PTL_COLS]
else:
    df_ptl = pd.DataFrame(columns=EXPECTED_PTL_COLS)

if isinstance(df_cp, pd.DataFrame) and not df_cp.empty:
    if '無線電' in df_cp.columns and '無線電代號' not in df_cp.columns: df_cp['無線電代號'] = df_cp['無線電']
    for c in EXPECTED_CP_COLS:
        if c not in df_cp.columns: df_cp[c] = ""
    df_cp = df_cp[EXPECTED_CP_COLS]
else:
    df_cp = pd.DataFrame(columns=EXPECTED_CP_COLS)


d = dict(zip(df_set.iloc[:, 0].astype(str), df_set.iloc[:, 1].astype(str))) if df_set is not None and not df_set.empty else {}
u = d.get("unit_name", DEFAULT_UNIT)
t = d.get("plan_full_time", DEFAULT_TIME)
p = d.get("project_name", DEFAULT_PROJ)
b = d.get("briefing_info", DEFAULT_BRIEF)
s = d.get("check_station", DEFAULT_STATION)
p1_d = d.get("phase1_desc", DEFAULT_P1_DESC)
p2_d = d.get("phase2_desc", DEFAULT_P2_DESC)

st.title("🚓 聯合稽查(二階段)勤務規劃系統")
c1, c2 = st.columns(2)

p_time = c2.text_input("勤務時間", value=t)
match = re.search(r"(\d+)月(\d+)日", p_time)
new_prefix = f"{str(match.group(1)).zfill(2)}{str(match.group(2)).zfill(2)}" if match else ""

if p and re.match(r"^\d{4}", p): display_p_name = p[4:]
else: display_p_name = p

p_name_display_input = c1.text_input("專案名稱", value=display_p_name)
p_name = f"{new_prefix}{p_name_display_input}" if new_prefix else p_name_display_input

st.subheader("⚙️ 階段標題設定")
cc1, cc2 = st.columns(2)
phase1_desc = cc1.text_input("第一階段標題", p1_d)
phase2_desc = cc2.text_input("第二階段標題", p2_d)

st.subheader("1. 指揮編組")
res_cmd = st.data_editor(df_cmd if df_cmd is not None and not df_cmd.empty else DEFAULT_CMD.copy(), num_rows="dynamic", use_container_width=True).dropna(how='all').fillna("")
b_info, s_info = st.text_area("📢 勤前教育", b, height=70), st.text_area("🚧 環保局臨時檢驗站開設", s, height=70)

st.subheader("2. 勤務編組")

with st.expander("📋 點此打開【今日出勤名冊快速貼上區】(針對巡邏組)", expanded=False):
    st.markdown("""
    **💡 智慧辨識貼上說明（3欄、4欄皆通用）：**
    * **【模式 A：一般同單位模式】** 直接貼 **3 個資料** 👉 `單位 職別 姓名`
    * **【模式 B：跨單位聯合模式】** 貼上 **4 個資料** 👉 `編組名稱 單位 職別 姓名`
    """)
    
    paste_placeholder = "聖亭所 副所長 曹培翔\n聖亭所 警員 詹宗澤\n第三巡邏組 中興所 警員 林國仁"
    raw_paste = st.text_area("請在此貼上名冊文字：", value="", placeholder=paste_placeholder, height=200)
    
    if st.button("⚡ 解析名冊並匯入下方【第一階段】表格", use_container_width=True):
        if raw_paste.strip():
            lines = raw_paste.strip().split("\n")
            parsed_ptl = []
            
            route_map = {
                "聖亭": "於中正路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。",
                "龍潭": "於北龍路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。",
                "中興": "於中興路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。",
                "石門": "於神龍路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。",
                "高平": "於東龍路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。",
                "三和": "於東龍路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。",
                "交通": "於中豐路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"
            }
            
            for line in lines:
                if not line.strip(): continue
                tokens = re.split(r'[\s,\t]+', line.strip())
                
                if len(tokens) == 3:
                    u_name = tokens[0].strip()
                    title  = tokens[1].strip()
                    name   = tokens[2].strip()
                    g_name = u_name.replace("派出所", "組").replace("所", "組").replace("分隊", "組")
                elif len(tokens) >= 4:
                    g_name = tokens[0].strip()
                    u_name = tokens[1].strip()
                    title  = tokens[2].strip()
                    name   = tokens[3].strip()
                else:
                    continue 
                    
                default_route = next((v for k, v in route_map.items() if k in u_name), "於轄區內易有噪音車輛滋擾路段巡邏。")
                
                parsed_ptl.append({
                    "編組": g_name,
                    "無線電代號": "", 
                    "單位": u_name,
                    "職別": title,
                    "姓名": name,
                    "任務分工": "機動巡查" if "警員" in title else "帶班兼蒐證",
                    "巡邏路段": default_route
                })
            
            if parsed_ptl:
                st.session_state.ptl_editable_df = pd.DataFrame(parsed_ptl)
                st.success("🎉 名冊智慧解析成功！已載入下方第一階段表格。")
                st.rerun()
            else:
                st.error("❌ 無法解析文字，請確認每行輸入是否包含最少 3 個或 4 個空白隔開的資料。")

auto_sync_radio = st.checkbox("✨ 啟用自動推算與統一同編組「無線電代號」 (若需完全手動自訂每列代號，請取消勾選)", value=True)

tab1, tab2 = st.tabs(["📍 第一階段 (巡邏)", "🚧 第二階段 (路檢)"])

with tab1:
    st.info(f"當前標題：{phase1_desc}")
    if "ptl_editable_df" not in st.session_state:
        st.session_state.ptl_editable_df = df_ptl if not df_ptl.empty else DEFAULT_PTL.copy()

    raw_ptl = st.data_editor(st.session_state.ptl_editable_df, num_rows="dynamic", use_container_width=True, key="ptl_editor")
    res_ptl = auto_assign_radio_code(raw_ptl).dropna(how='all').fillna("") if auto_sync_radio else raw_ptl.dropna(how='all').fillna("")
    
    if not res_ptl.equals(st.session_state.ptl_editable_df):
        st.session_state.ptl_editable_df = res_ptl.copy()
        st.rerun()

with tab2:
    st.info(f"當前標題：{phase2_desc}")
    raw_cp = st.data_editor(df_cp if not df_cp.empty else DEFAULT_CP.copy(), num_rows="dynamic", use_container_width=True, key="cp_editor")
    res_cp = auto_assign_radio_code(raw_cp).dropna(how='all').fillna("") if auto_sync_radio else raw_cp.dropna(how='all').fillna("")

st.markdown("---")

if st.button("💾 同步雲端並發送備份郵件", use_container_width=True):
    with st.spinner("處理中..."):
        save_data(u, p_time, p_name, b_info, s_info, phase1_desc, phase2_desc, res_cmd, res_ptl, res_cp)
        ok, max_err = send_report_email(u, p_name, p_time, b_info, s_info, phase1_desc, phase2_desc, res_cmd, res_ptl, res_cp)
        if ok: 
            st.success("✅ 同步與郵件發送成功！")
            if "ptl_editable_df" in st.session_state: del st.session_state.ptl_editable_df
            st.rerun()
        else: st.warning(f"⚠️ 雲端已同步，但郵件失敗: {max_err}")
