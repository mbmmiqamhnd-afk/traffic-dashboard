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
from menu import show_sidebar

# --- 1. 頁面設定 (必須是第一個 Streamlit 指令) ---
st.set_page_config(page_title="聯合稽查勤務規劃系統", layout="wide", page_icon="🚓")

# 呼叫側邊欄
show_sidebar()

# --- 常幕與設定 ---
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
    {"職稱": "指揮官", "無線電代號": "隆安1", "負責人員": "分局長施宇峰", "任務": "核定本勤務執行並重點機動督導"},
    {"職稱": "副指揮官", "無線電代號": "隆安2", "負責人員": "副分局長何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "無線電代號": "隆安3", "負責人員": "副分局長蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "無線電代號": "建興", "負責人員": "駐區督察 孫三陽", "任務": "重點機動督導"},
    {"職稱": "督導組", "無線電代號": "隆安6", "負責人員": "督察組組長黃長旗\n督察組督察員 黃中彥\n督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律"},
    {"職稱": "指導組", "無線電代號": "隆安684", "負責人員": "督察組教官郭文義", "任務": "指導各編組勤務執行及狀況處置"},
    {"職稱": "作業及督巡組", "無線電代號": "隆安13", "負責人員": "交通組組長 楊孟竟\n交通組警務員盧冠仁\n交通組警務員李峯甫\n交通組警務員葉佳媛\n交通組巡官郭勝隆\n交通組警員吳享運\n勤指中心警員張庭溱\n(代理人:巡官陳鵬翔)", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組", "無線電代號": "隆安", "負責人員": "行政組警務佐曾威仁\n人事室警員陳明祥\n主任蔡奇青\n執勤官李文章\n執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜"},
])

DEFAULT_PTL = pd.DataFrame([
    {"編組": "聖亭組1", "無線電": "隆安52", "單位": "聖亭所", "職別": "副所長", "姓名": "邱品淳", "任務分工": "機動巡查", "巡邏路段": "於中正路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "聖亭組1", "無線電": "隆安52", "單位": "聖亭所", "職別": "警員", "姓名": "傅維強", "任務分工": "安全維護", "巡邏路段": "於中正路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "聯合組1", "無線電": "隆安61", "單位": "龍潭所", "職別": "所長", "姓名": "孫祥愷", "任務分工": "改裝車查緝", "巡邏路段": "於北龍路周邊易有噪音車輛滋擾聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "聯合組1", "無線電": "隆安61", "單位": "中興所", "職別": "警員", "姓名": "蔡震東", "任務分工": "跨區聯合稽查", "巡邏路段": "於北龍路周邊易有噪音車輛滋擾聚集路段機動巡查改裝噪音車輛。"},
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
            WS_MAP["cmd"]: [["職稱", "無線電代號", "負責人員", "任務"]], # 欄位修改
            WS_MAP["ptl"]: [["編組", "無線電", "單位", "職別", "姓名", "任務分工", "巡邏路段"]]
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

def add_page_number(canvas, doc):
    canvas.saveState()
    font_name = _get_font()
    canvas.setFont(font_name, 11)
    page_num = canvas.getPageNumber()
    text = f"- 第 {page_num} 頁 -"
    canvas.drawCentredString(A4_SIZE[0] / 2.0, 10 * mm, text)
    canvas.restoreState()

def generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    font = _get_font()
    buf = io.BytesIO()
    margin_lr = float(12 * mm)
    margin_tb = float(15 * mm)
    doc = SimpleDocTemplate(buf, pagesize=A4_SIZE, leftMargin=margin_lr, rightMargin=margin_lr, topMargin=margin_tb, bottomMargin=margin_tb)
    page_width = A4_SIZE[0] - (2 * margin_lr)
    story = []

    style_title        = ParagraphStyle('Title',       fontName=font, fontSize=18, leading=24, alignment=1, spaceAfter=8,   wordWrap='CJK')
    style_info         = ParagraphStyle('Info',        fontName=font, fontSize=12, alignment=2, spaceAfter=10,              wordWrap='CJK')
    style_cell         = ParagraphStyle('Cell',        fontName=font, fontSize=13, leading=18, alignment=1,                 wordWrap='CJK')
    style_cell_left    = ParagraphStyle('CellLeft',    fontName=font, fontSize=13, leading=18, alignment=0,                 wordWrap='CJK')
    style_middle_block = ParagraphStyle('MiddleBlock', fontName=font, fontSize=14, leading=22, spaceAfter=2*mm,
                                        alignment=TA_LEFT, leftIndent=5*mm, firstLineIndent=0,                              wordWrap='CJK')
    style_table_title  = ParagraphStyle('TTitle',      fontName=font, fontSize=16, alignment=1, leading=22,                 wordWrap='CJK')

    story.append(Paragraph(f"{unit}{project}勤務規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))

    def clean(t): return str(t).replace("\n", "<br/>").replace("、", "<br/>")

    data_cmd = [[Paragraph("<b>任 務 編 組</b>", style_table_title), '', '', ''],
                [Paragraph(f"<b>{h}</b>", style_cell) for h in ["職稱", "無線電代號", "負責人員", "任務"]]] # PDF表頭修改
    for _, r in df_cmd.iterrows():
        data_cmd.append([
            Paragraph(f"<b>{r.get('職稱','')}</b>", style_cell),
            Paragraph(clean(r.get('無線電代號','')), style_cell), # 取值修改
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

    # PDF巡邏編組抬頭修改為「無線電代號」
    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "無線電代號", "單位", "職別", "姓名", "任務分工", "巡邏路段"]]]
    
    if not df_ptl.empty:
        t2_styles = [
            ('FONTNAME',   (0,0), (-1,-1), font),
            ('FONTSIZE',   (0,0), (-1,-1), 13),
            ('ALIGN',      (0,1), (5,-1),  'CENTER'),
            ('GRID',       (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0),  colors.HexColor('#f2f2f2')),
            ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ]
        
        current_row = 1
        for g_name, g_df in df_ptl.groupby("編組", sort=False):
            start_row = current_row
            
            for _, r in g_df.iterrows():
                task_route = f"{r.get('巡邏路段','')}<br/><font color='blue' size='11'>*雨備方案：各治安要點巡邏。</font>"
                data_ptl.append([
                    Paragraph(clean(r.get('編組','')), style_cell),
                    Paragraph(clean(r.get('無線電','')), style_cell),
                    Paragraph(clean(r.get('單位','')), style_cell),
                    Paragraph(clean(r.get('職別','')), style_cell),
                    Paragraph(clean(r.get('姓名','')), style_cell),
                    Paragraph(clean(r.get('任務分工','')), style_cell),
                    Paragraph(task_route, style_cell_left)
                ])
                current_row += 1
            end_row = current_row - 1
            
            if start_row < end_row:
                t2_styles.append(('SPAN', (0, start_row), (0, end_row)))
                t2_styles.append(('SPAN', (6, start_row), (6, end_row)))
                
                sub_start_rad = start_row
                for r_idx in range(start_row + 1, end_row + 1):
                    prev_val = str(g_df.iloc[sub_start_rad - start_row]['無線電']).strip()
                    curr_val = str(g_df.iloc[r_idx - start_row]['無線電']).strip()
                    if prev_val != curr_val:
                        if sub_start_rad < r_idx - 1:
                            t2_styles.append(('SPAN', (1, sub_start_rad), (1, r_idx - 1)))
                        sub_start_rad = r_idx
                if sub_start_rad < end_row:
                    t2_styles.append(('SPAN', (1, sub_start_rad), (1, end_row)))
                    
                sub_start_uni = start_row
                for r_idx in range(start_row + 1, end_row + 1):
                    prev_val = str(g_df.iloc[sub_start_uni - start_row]['單位']).strip()
                    curr_val = str(g_df.iloc[r_idx - start_row]['單位']).strip()
                    if prev_val != curr_val:
                        if sub_start_uni < r_idx - 1:
                            t2_styles.append(('SPAN', (2, sub_start_uni), (2, r_idx - 1)))
                        sub_start_uni = r_idx
                if sub_start_uni < end_row:
                    t2_styles.append(('SPAN', (2, sub_start_uni), (2, end_row)))

        t2 = Table(data_ptl, colWidths=[page_width*0.11, page_width*0.11, page_width*0.12, page_width*0.10, page_width*0.12, page_width*0.13, page_width*0.31], repeatRows=1)
        t2.setStyle(TableStyle(t2_styles))
        story.append(t2)
        
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
    meeting_range = parse_meeting_time(time_str)
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
def send_report_email(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        max_msg = MIMEMultipart()
        max_msg["From"], max_msg["To"], max_msg["Subject"] = sender, sender, f"勤務規劃與簽到表_{datetime.now().strftime('%m%d')}"
        max_msg.attach(MIMEText("附件為最新的勤務規劃表與人員簽到表 PDF。", "plain", "utf-8"))
        
        plan_filename = f"{unit}{project}勤務規劃表.pdf"
        
        for pdf_func, name in [(generate_pdf_from_data, plan_filename), (generate_attendance_pdf, '簽到表.pdf')]:
            args = (unit, project, time_str, briefing, station, df_cmd, df_ptl) if pdf_func == generate_pdf_from_data else (unit, project, time_str, briefing)
            data = pdf_func(*args)
            part = MIMEBase("application", "pdf")
            part.set_payload(data)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(name)}")
            max_msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, max_msg.as_string())
        return True, None
    except Exception as e: return False, str(e)


# --- 自動指派無線電代號 ---
def auto_assign_radio_code(df):
    if df.empty: return df
    
    prefixes = {"交通分隊": "99", "聖亭": "5", "龍潭": "6", "中興": "7", "石門": "8", "高平": "9", "三和": "3"}
    
    first_row_indices = df.groupby("編組", sort=False).head(1).index
    
    for idx in first_row_indices:
        unit = str(df.at[idx, '單位'])
        title = str(df.at[idx, '職別']).strip()
        current_radio = str(df.at[idx, '無線電']).strip()
        
        base_pfx = next((v for k, v in prefixes.items() if k in unit), "")
        
        if base_pfx:
            if (current_radio == "" or 
                not current_radio.startswith(f"隆安{base_pfx}") or 
                current_radio in [f"隆安{base_pfx}0", f"隆安{base_pfx}1", f"隆安{base_pfx}2"]):
                
                if "副所長" in title or "小隊長" in title:
                    df.at[idx, '無線電'] = f"隆安{base_pfx}2"
                elif "所長" in title or "分隊長" in title:
                    df.at[idx, '無線電'] = f"隆安{base_pfx}1"
                else:
                    df.at[idx, '無線電'] = f"隆安{base_pfx}0"

    radio_map = dict(zip(df.loc[first_row_indices, "編組"], df.loc[first_row_indices, "無線電"]))
    
    for idx, row in df.iterrows():
        g_name = row.get("編組")
        if g_name in radio_map:
            df.at[idx, '無線電'] = radio_map[g_name]
            
    return df


# --- 5. 主介面 ---
st.sidebar.title("🛠️ 雲端設定")
if st.sidebar.button("初始化/檢查雲端分頁"): init_sheets()
if st.sidebar.button("⚠️ 強制重置為最新專案資料 (覆蓋雲端)"):
    with st.spinner("重置中..."):
        save_data(DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_PROJ, DEFAULT_BRIEF, DEFAULT_STATION, DEFAULT_CMD, DEFAULT_PTL)
        if "ptl_editable_df" in st.session_state: del st.session_state.ptl_editable_df
        st.cache_data.clear()
        st.rerun()

df_set, df_cmd, df_ptl, err = load_data()
d = dict(zip(df_set.iloc[:, 0].astype(str), df_set.iloc[:, 1].astype(str))) if df_set is not None and not df_set.empty else {}
u = d.get("unit_name", DEFAULT_UNIT)
t = d.get("plan_full_time", DEFAULT_TIME)
p = d.get("project_name", DEFAULT_PROJ)
b = d.get("briefing_info", DEFAULT_BRIEF)
s = d.get("check_station", DEFAULT_STATION)

st.title("🚓 聯合稽查勤務規劃 management系統")
c1, c2 = st.columns(2)

p_time = c2.text_input("勤務時間", value=t)

match = re.search(r"(\d+)月(\d+)日", p_time)
new_prefix = f"{str(match.group(1)).zfill(2)}{str(match.group(2)).zfill(2)}" if match else ""

if "p_name_input" not in st.session_state:
    st.session_state.p_name_input = p

current_p = st.session_state.p_name_input
if new_prefix:
    if re.match(r"^\d{4}", current_p):
        if current_p[:4] != new_prefix:
            st.session_state.p_name_input = f"{new_prefix}{current_p[4:]}"
    else:
        st.session_state.p_name_input = f"{new_prefix}{current_p}"

p_name = c1.text_input("專案名稱", key="p_name_input")


st.subheader("1. 指揮編組")
res_cmd = st.data_editor(df_cmd if df_cmd is not None and not df_cmd.empty else DEFAULT_CMD.copy(), num_rows="dynamic", use_container_width=True).dropna(how='all').fillna("")
b_info, s_info = st.text_area("📢 勤前教育", b, height=70), st.text_area("🚧 環保局臨時檢驗站開設", s, height=70)


st.subheader("2. 巡邏編組")

with st.expander("📋 點此打開【今日出勤名冊快速貼上區】", expanded=False):
    st.markdown("""
    **💡 智慧辨識貼上說明（3欄、4欄皆通用）：**
    * **【模式 A：一般同單位模式】** 直接貼 **3 個資料** 👉 `單位 職別 姓名`
    * **【模式 B：跨單位聯合模式】** 貼上 **4 個資料** 👉 `編組名稱 單位 職別 姓名`
    """)
    
    paste_placeholder = "聖亭所 副所長 邱品淳\n聖亭所 警員 傅維強\n聯合組1 龍潭所 所長 孫祥愷\n聯合組1 中興所 警員 蔡震東"
    raw_paste = st.text_area("請在此貼上名冊文字：", value="", placeholder=paste_placeholder, height=200)
    
    if st.button("⚡ 解析名冊並匯入下方表格", use_container_width=True):
        if raw_paste.strip():
            lines = raw_paste.strip().split("\n")
            parsed_ptl = []
            
            route_map = {
                "聖亭": "於中正路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。",
                "龍潭": "於北龍路周邊易有噪音車輛滋擾聚集路段機動巡查改裝噪音車輛。",
                "中興": "於龍新路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。",
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
                    "無線電": "", 
                    "單位": u_name,
                    "職別": title,
                    "姓名": name,
                    "任務分工": "機動巡查" if "警員" in title else "安全維護及督導",
                    "巡邏路段": default_route
                })
            
            if parsed_ptl:
                st.session_state.ptl_editable_df = pd.DataFrame(parsed_ptl)
                st.success("🎉 名冊智慧解析成功！已載入下方表格。")
                st.rerun()
            else:
                st.error("❌ 無法解析文字，請確認每行輸入是否包含最少 3 個或 4 個空白隔開的資料。")

if "ptl_editable_df" not in st.session_state:
    st.session_state.ptl_editable_df = df_ptl if df_ptl is not None and not df_ptl.empty else DEFAULT_PTL.copy()

res_ptl_raw = st.data_editor(st.session_state.ptl_editable_df, num_rows="dynamic", use_container_width=True).dropna(how='all').fillna("")

res_ptl = auto_assign_radio_code(res_ptl_raw.copy())

if not res_ptl.equals(st.session_state.ptl_editable_df):
    st.session_state.ptl_editable_df = res_ptl.copy()
    st.rerun()

st.markdown("---")
col_dl1, col_dl2 = st.columns(2)

plan_filename = f"{u}{p_name}勤務規劃表.pdf"

col_dl1.download_button("📝 下載 1.勤務規劃表", data=generate_pdf_from_data(u, p_name, p_time, b_info, s_info, res_cmd, res_ptl), file_name=plan_filename, use_container_width=True)
col_dl2.download_button("🖋️ 下載 2.人員簽到表", data=generate_attendance_pdf(u, p_name, p_time, b_info), file_name=f"簽到表_{datetime.now().strftime('%m%d')}.pdf", use_container_width=True)

if st.button("💾 同步雲端並發送備份郵件", use_container_width=True):
    with st.spinner("處理中..."):
        save_data(u, p_time, p_name, b_info, s_info, res_cmd, res_ptl)
        ok, max_err = send_report_email(u, p_name, p_time, b_info, s_info, res_cmd, res_ptl)
        if ok: 
            st.success("✅ 同步與郵件發送成功！")
            if "ptl_editable_df" in st.session_state: del st.session_state.ptl_editable_df
            st.rerun()
        else: st.warning(f"⚠️ 雲端已同步，但郵件失敗: {max_err}")
