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
    style_cell         = ParagraphStyle('Cell',        fontName=font, fontSize=14, leading=20, alignment=1,                 wordWrap='CJK')
    style_cell_left    = ParagraphStyle('CellLeft',    fontName=font, fontSize=14, leading=20, alignment=0,                 wordWrap='CJK')
    style_middle_block = ParagraphStyle('MiddleBlock', fontName=font, fontSize=14, leading=22, spaceAfter=2*mm,
                                        alignment=TA_LEFT, leftIndent=5*mm, firstLineIndent=0,                              wordWrap='CJK')
    style_table_title  = ParagraphStyle('TTitle',      fontName=font, fontSize=16, alignment=1, leading=22,                 wordWrap='CJK')

    story.append(Paragraph(f"{unit}{project}勤務規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))

    def clean(t): return str(t).replace("\n", "<br/>").replace("、", "<br/>")

    data_cmd = [[Paragraph("<b>任 務 編 組</b>", style_table_title), '', '', ''],
                [Paragraph(f"<b>{h}</b>", style_cell) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([
            Paragraph(f"<b>{r.get('職稱','')}</b>", style_cell),
            Paragraph(clean(r.get('代號','')), style_cell),
            Paragraph(clean(r.get('姓名','')), style_cell),
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

    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "代號", "單位", "服勤人員", "任務分工"]]]
    for _, r in df_ptl.iterrows():
        task = f"{r.get('任務分工','')}<br/><font color='blue' size='12'>*雨備方案：各治安要點巡邏。</font>"
        data_ptl.append([
            Paragraph(clean(r.get('編組','')), style_cell),
            Paragraph(clean(r.get('無線電','')), style_cell),
            Paragraph(clean(r.get('單位','')), style_cell),
            Paragraph(clean(r.get('服勤人員','')), style_cell),
            Paragraph(task, style_cell_left)
        ])

    t2 = Table(data_ptl, colWidths=[page_width*0.14, page_width*0.14, page_width*0.16, page_width*0.20, page_width*0.36], repeatRows=1)
    t2.setStyle(TableStyle([
        ('FONTNAME',   (0,0), (-1,-1), font),
        ('FONTSIZE',   (0,0), (-1,-1), 14),
        ('ALIGN',      (0,1), (1,-1),  'CENTER'),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0),  colors.HexColor('#f2f2f2')),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
    ]))
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
    
    # --- 調整後的單位順序：督察組在上，勤指中心在下 ---
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
        msg = MIMEMultipart()
        msg["From"], msg["To"], msg["Subject"] = sender, sender, f"勤務規劃與簽到表_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為最新的勤務規劃表與人員簽到表 PDF。", "plain", "utf-8"))
        
        plan_filename = f"{unit}{project}勤務規劃表.pdf"
        
        for pdf_func, name in [(generate_pdf_from_data, plan_filename), (generate_attendance_pdf, '簽到表.pdf')]:
            args = (unit, project, time_str, briefing, station, df_cmd, df_ptl) if pdf_func == generate_pdf_from_data else (unit, project, time_str, briefing)
            data = pdf_func(*args)
            part = MIMEBase("application", "pdf")
            part.set_payload(data)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(name)}")
            msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# --- 5. 主介面 ---
st.sidebar.title("🛠️ 雲端設定")
if st.sidebar.button("初始化/檢查雲端分頁"): init_sheets()
if st.sidebar.button("⚠️ 強制重置為最新專案資料 (覆蓋雲端)"):
    with st.spinner("重置中..."):
        save_data(DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_PROJ, DEFAULT_BRIEF, DEFAULT_STATION, DEFAULT_CMD, DEFAULT_PTL)
        st.cache_data.clear()
        st.rerun()

df_set, df_cmd, df_ptl, err = load_data()
d = dict(zip(df_set.iloc[:, 0].astype(str), df_set.iloc[:, 1].astype(str))) if df_set is not None and not df_set.empty else {}
u = d.get("unit_name", DEFAULT_UNIT)
t = d.get("plan_full_time", DEFAULT_TIME)
p = d.get("project_name", DEFAULT_PROJ)
b = d.get("briefing_info", DEFAULT_BRIEF)
s = d.get("check_station", DEFAULT_STATION)

# 將標題修改為「聯合稽查」
st.title("🚓 聯合稽查勤務規劃管理系統")
c1, c2 = st.columns(2)

# 1. 先渲染勤務時間，取得最新的時間輸入值
p_time = c2.text_input("勤務時間", value=t)

# 2. 自動擷取勤務時間的月、日，產生 4 碼前綴
match = re.search(r"(\d+)月(\d+)日", p_time)
new_prefix = f"{str(match.group(1)).zfill(2)}{str(match.group(2)).zfill(2)}" if match else ""

# 3. 初始化專案名稱的 session_state (若剛載入則套用預設值 p)
if "p_name_input" not in st.session_state:
    st.session_state.p_name_input = p

# 4. 判斷並連動更新狀態：如果時間改變了，立刻抽換專案名稱前 4 碼
current_p = st.session_state.p_name_input
if new_prefix:
    if re.match(r"^\d{4}", current_p):
        if current_p[:4] != new_prefix:
            st.session_state.p_name_input = f"{new_prefix}{current_p[4:]}"
    else:
        st.session_state.p_name_input = f"{new_prefix}{current_p}"

# 5. 渲染專案名稱於左側欄，直接綁定 key 來連動畫面的文字框
p_name = c1.text_input("專案名稱", key="p_name_input")


st.subheader("1. 指揮編組")
res_cmd = st.data_editor(df_cmd if df_cmd is not None and not df_cmd.empty else DEFAULT_CMD.copy(), num_rows="dynamic", use_container_width=True).dropna(how='all').fillna("")
b_info, s_info = st.text_area("📢 勤前教育", b, height=70), st.text_area("🚧 環保局臨時檢驗站開設", s, height=70)

st.subheader("2. 巡邏編組")
res_ptl_raw = st.data_editor(df_ptl if df_ptl is not None and not df_ptl.empty else DEFAULT_PTL.copy(), num_rows="dynamic", use_container_width=True).dropna(how='all').fillna("")

def auto_assign_radio_code(df):
    prefixes = {"交通分隊": "99", "聖亭": "5", "龍潭": "6", "中興": "7", "石門": "8", "高平": "9", "三和": "3"}
    for idx, row in df.iterrows():
        unit, person = str(row.get('單位', '')), str(row.get('服勤人員', ''))
        first_unit = re.split(r'[\n、 ]', unit.strip())[0]
        base_pfx = next((v for k, v in prefixes.items() if k in first_unit), "")
        if base_pfx:
            if "副所長" in person: df.at[idx, '無線電'] = f"隆安{base_pfx}2"
            elif "所長" in person: df.at[idx, '無線電'] = f"隆安{base_pfx}1"
            elif not str(row.get('無線電','')).startswith(f"隆安{base_pfx}"): df.at[idx, '無線電'] = f"隆安{base_pfx}0"
    return df

res_ptl = auto_assign_radio_code(res_ptl_raw.copy())

st.markdown("---")
col_dl1, col_dl2 = st.columns(2)

plan_filename = f"{u}{p_name}勤務規劃表.pdf"

col_dl1.download_button("📝 下載 1.勤務規劃表", data=generate_pdf_from_data(u, p_name, p_time, b_info, s_info, res_cmd, res_ptl), file_name=plan_filename, use_container_width=True)
col_dl2.download_button("🖋️ 下載 2.人員簽到表", data=generate_attendance_pdf(u, p_name, p_time, b_info), file_name=f"簽到表_{datetime.now().strftime('%m%d')}.pdf", use_container_width=True)

if st.button("💾 同步雲端並發送備份郵件", use_container_width=True):
    with st.spinner("處理中..."):
        save_data(u, p_time, p_name, b_info, s_info, res_cmd, res_ptl)
        ok, m_err = send_report_email(u, p_name, p_time, b_info, s_info, res_cmd, res_ptl)
        if ok: st.success("✅ 同步與郵件發送成功！")
        else: st.warning(f"⚠️ 雲端已同步，但郵件失敗: {m_err}")
