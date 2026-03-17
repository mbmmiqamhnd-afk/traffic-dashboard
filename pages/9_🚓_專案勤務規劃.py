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
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
import re

# --- 1. 頁面設定 ---
st.set_page_config(page_title="雲端勤務規劃系統", layout="wide", page_icon="🚓")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年3月20日19至23時"
DEFAULT_PROJ    = "0320「取締改裝(噪音)車輛專案監、警、環聯合稽查」"
DEFAULT_BRIEF   = "於分局二樓會議室召開" 
DEFAULT_STATION = "環保局臨時檢驗站開設時間：20時至23時\n地點：桃園市龍潭區大昌路一段277號（龍潭區警政聯合辦公大樓）廣場"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "駐區督察", "姓名": "孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫", "任務": "規劃本勤務、重點機動督導及回報績效。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_PTL = pd.DataFrame([
    {"編組": "第一巡邏組", "無線電": "隆安54", "單位": "聖亭所", "服勤人員": "巡佐傅錫城、警員曾建凱", "任務分工": "於大昌路一段周邊巡查。"},
    {"編組": "第二巡邏組", "無線電": "隆安62", "單位": "龍潭所", "服勤人員": "副所長全楚文、警員龔品璇", "任務分工": "於大昌路二段周邊巡查。"},
    {"編組": "第三巡邏組", "無線電": "隆安72", "單位": "中興所", "服勤人員": "副所長薛德祥、警員冷柔萱", "任務分工": "於中興路周邊巡查。"},
    {"編組": "第四巡邏組", "無線電": "隆安83", "單位": "石門所", "服勤人員": "巡佐林偉政、警員盧瑾瑤", "任務分工": "於北龍路周邊巡查。"},
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
    """提取勤務開始時間並計算第一小時的後半小時"""
    try:
        match = re.search(r"(\d+)至", time_str)
        if match:
            start_hour = int(match.group(1))
            end_hour = start_hour + 1
            return f"{start_hour}時30分至{end_hour}時00分"
    except:
        pass
    return "19時30分至20時00分"

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
        if client is None: return None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("設定")
        ws_cmd = sh.worksheet("指揮組")
        ws_ptl = sh.worksheet("巡邏組")
        return pd.DataFrame(ws_set.get_all_records()), pd.DataFrame(ws_cmd.get_all_records()), pd.DataFrame(ws_ptl.get_all_records()), None
    except Exception as e: return None, None, None, str(e)

def save_data(unit, time_str, project, briefing, station, df_cmd, df_ptl):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["unit_name", unit], ["plan_full_time", time_str], ["project_name", project], ["briefing_info", briefing], ["check_station", station]])
        for ws_name, df in [("指揮組", df_cmd), ("巡邏組", df_ptl)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df = df.fillna("")
            ws.update([df.columns.tolist()] + df.values.tolist())
        load_data.clear()
        return True
    except: return False

# --- 3. PDF 生成功能 ---

def generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []
    
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=24, alignment=1, spaceAfter=8)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=10)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=18, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=18, alignment=0)
    style_note = ParagraphStyle('Note', fontName=font, fontSize=14, leading=20, spaceAfter=5)
    style_table_title = ParagraphStyle('TTitle', fontName=font, fontSize=16, alignment=1, leading=22)

    story.append(Paragraph(f"{unit}執行{project}規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def clean(t): return str(t).replace("\n", "<br/>").replace("、", "<br/>")

    data_cmd = [[Paragraph("<b>任 務 編 組</b>", style_table_title), '', '', ''],
                [Paragraph(f"<b>{h}</b>", style_cell) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([Paragraph(f"<b>{r.get('職稱','')}</b>", style_cell), Paragraph(str(r.get('代號','')), style_cell),
                         Paragraph(clean(r.get('姓名','')), style_cell), Paragraph(str(r.get('任務','')), style_cell_left)])
    
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.12, page_width*0.28, page_width*0.45])
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('SPAN',(0,0),(-1,0)),
                            ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph(f"<b>📢 勤前教育：</b>{briefing}", style_note))
    story.append(Paragraph(f"<b>🚧 檢驗站資訊：</b><br/>{station.replace(chr(10), '<br/>')}", style_note))
    story.append(Spacer(1, 6*mm))

    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "代號", "單位", "服勤人員", "任務分工"]]]
    for _, r in df_ptl.iterrows():
        task = f"{r.get('任務分工','')}<br/><font color='blue' size='11'>*雨備方案：各治安要點巡邏。</font>"
        data_ptl.append([str(r.get('編組','')), str(r.get('無線電','')), Paragraph(clean(r.get('單位','')), style_cell), 
                         Paragraph(clean(r.get('服勤人員','')), style_cell), Paragraph(task, style_cell_left)])
    
    t2 = Table(data_ptl, colWidths=[page_width*0.15, page_width*0.12, page_width*0.13, page_width*0.20, page_width*0.40])
    t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('FONTSIZE',(0,0),(-1,-1),14),('ALIGN',(0,1),(1,-1),'CENTER'),
                            ('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t2)
    doc.build(story)
    return buf.getvalue()

# (B) 簽到表
def generate_attendance_pdf(unit, project, time_str, briefing):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 30*mm
    story = []

    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, leading=22, alignment=1, spaceAfter=8)
    style_top_info = ParagraphStyle('TopInfo', fontName=font, fontSize=12, leading=18, alignment=0)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=12, leading=22, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=12, leading=22, alignment=0) #
    style_note = ParagraphStyle('Note', fontName=font, fontSize=11, leading=15, alignment=0)

    # 1. 標題
    story.append(Paragraph(f"{unit}執行{project}勤前教育會議人員簽到表", style_title))
    
    # 2. 自動計算時間
    meeting_range = parse_meeting_time(time_str)
    date_part = time_str.split('日')[0] + '日' if '日' in time_str else ""
    story.append(Paragraph(f"時間：{date_part}{meeting_range}", style_top_info))
    
    # 3. 地點
    loc = briefing if "於" not in briefing else briefing.split("於")[1]
    story.append(Paragraph(f"地點：{loc}", style_top_info))
    story.append(Spacer(1, 3*mm))

    # 4. 核心表格
    table_data = []
    
    # 第一列：分局長 與 上級督導 (文字靠左)
    table_data.append([Paragraph("分局長：", style_cell_left), "", Paragraph("上級督導：", style_cell_left), ""])
    
    # 第二列：副分局長 (跨欄合併、文字靠左)
    table_data.append([Paragraph("副分局長：", style_cell_left), "", "", ""])

    # 第三列：標題
    table_data.append([Paragraph("單位", style_cell), Paragraph("參加人員", style_cell), 
                        Paragraph("單位", style_cell), Paragraph("參加人員", style_cell)])
    
    # 單位內容 (交通組與勤指互換位置)
    rows = [
        ("交通組", "中興派出所"),
        ("勤務指揮中心", "石門派出所"),
        ("督察組", "高平派出所"),
        ("聖亭派出所", "三和派出所"),
        ("龍潭派出所", "龍潭交通分隊")
    ]
    for l, r in rows:
        table_data.append([Paragraph(l, style_cell), "", Paragraph(r, style_cell), ""])
    
    row_heights = [18*mm, 18*mm, 10*mm] + [20*mm] * len(rows)
    
    t = Table(table_data, colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3], rowHeights=row_heights)
    t.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        # 特定列靠左對齊
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('ALIGN', (2,0), (2,0), 'LEFT'),
        ('ALIGN', (0,1), (0,1), 'LEFT'),
        # 合併欄位
        ('SPAN', (0,1), (3,1)), 
        ('BACKGROUND', (0,2), (0,2), colors.whitesmoke),
        ('BACKGROUND', (2,2), (2,2), colors.whitesmoke),
    ]))
    story.append(t)

    # 5. 備註
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph("備註：請將行動電話調整為靜音。", style_note))

    doc.build(story)
    return buf.getvalue()

# --- 4. 寄信功能 (夾帶兩份 PDF) ---
def send_report_email(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    try:
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        msg["Subject"] = f"勤務規劃與簽到表_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為最新的勤務規劃表與人員簽到表 PDF。", "plain", "utf-8"))
        
        # 附件 1: 規劃表
        pdf1 = generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl)
        part1 = MIMEBase("application", "pdf")
        part1.set_payload(pdf1)
        encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote('規劃表.pdf')}")
        msg.attach(part1)
        
        # 附件 2: 簽到表
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

# --- 5. 主介面 ---
df_set, df_cmd, df_ptl, err = load_data()
if err or df_set is None:
    u, t, p, b, s = DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_PROJ, DEFAULT_BRIEF, DEFAULT_STATION
    ed_cmd, ed_ptl = DEFAULT_CMD.copy(), DEFAULT_PTL.copy()
else:
    d = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    u, t, p, b, s = d.get("unit_name", DEFAULT_UNIT), d.get("plan_full_time", DEFAULT_TIME), d.get("project_name", DEFAULT_PROJ), d.get("briefing_info", DEFAULT_BRIEF), d.get("check_station", DEFAULT_STATION)
    ed_cmd, ed_ptl = df_cmd, df_ptl

st.title("🚓 專案勤務規劃管理系統")
c1, c2 = st.columns(2)
p_name = c1.text_input("專案名稱", p)
p_time = c2.text_input("勤務時間", t)

st.subheader("1. 指揮編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True)

c3, c4 = st.columns(2)
b_info = c3.text_area("📢 勤前教育地點", b, height=70)
s_info = c4.text_area("🚧 檢驗站資訊", s, height=70)

st.subheader("2. 巡邏編組")
res_ptl = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True)

st.markdown("---")
st.subheader("📄 報表下載與同步")
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
