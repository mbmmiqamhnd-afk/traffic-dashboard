import streamlit as st

# 呼叫側邊欄 (確保在 config 之後)
try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    st.sidebar.warning("找不到 menu.py，跳過側邊欄載入。")

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

# --- 常數與雲端工作表設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# 使用與二階段系統一致的工作表對應
WS_MAP = {
    "set": "二階段_設定",
    "cmd": "二階段_指揮組",
    "p1_cp": "二階段_巡邏組",  # 第一階段定點路檢
    "p2_cp": "二階段_路檢組"   # 第二階段擴大臨檢
}

DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年4月11日 19時至23時"
DEFAULT_PROJ    = "0411取締酒後駕車與防制危險駕車及噪音車輛專案"
DEFAULT_BRIEF   = "一、 落實三安：同仁執行盤查、臨檢及機動勤務過程中，應強化敵情觀念，提高危機意識，落實「人犯戒護安全、案件程序安全、執法者及民眾安全」。\n二、 臨檢合法性：依《警察職權行使法》第6條辦理。\n三、 攔停規範：依《警察職權行使法》第8條辦理。\n四、 全程蒐證：務必全程連續錄音或錄影。\n五、 異議處理：依《警察職權行使法》第29條製作紀錄。"

DEFAULT_P1_DESC = "第一階段：19時至21時30分，定點路檢站部署勤務"
DEFAULT_P2_DESC = "第二階段：21時30分至23時，擴大臨檢與威力掃蕩"

# 比照二階段系統定點路檢的標準資料架構
DEFAULT_P1_DF = pd.DataFrame([
    {"編組": "第一組", "無線電": "隆安51", "單位": "聖亭所", "服勤人員": "所長 鄭榮捷\n警員 詹宗澤", "任務分工": "中正路與大昌路口\n攔停兼警戒"},
    {"編組": "第二組", "無線電": "隆安61", "單位": "龍潭所", "服勤人員": "所長 孫祥愷\n警員 沈庭禾", "任務分工": "北龍路與金龍路口\n攔停兼警戒"},
])

DEFAULT_P2_DF = pd.DataFrame([
    {"編組": "臨檢一組", "無線電": "", "單位": "聖亭所\n龍潭所", "服勤人員": "所長 鄭榮捷\n所長 孫祥愷", "任務分工": "帶班 製作臨檢紀錄\nA. 鉅大撞球館\nB. 台灣麻將協會"},
])

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安 1 號", "姓名": "分局長 施宇峰", "任務": "勤務核定並重點機動督導"},
    {"職稱": "副指揮官", "代號": "隆安 2 號", "姓名": "副分局長 何憶雯", "任務": "襄助指揮、重點機動督導"},
    {"職稱": "行政組", "代號": "隆安 5 號", "姓名": "組長 周金柱", "任務": "督導第二階段臨檢勤務"},
    {"職稱": "督察組", "代號": "隆安 6 號", "姓名": "督察組長 黃長旗", "任務": "機動督導各單位勤務紀律"},
    {"職稱": "交通組", "代號": "隆安 13號", "姓名": "交通組長 楊孟竟", "任務": "機動督導第一階段攔檢組"},
])

# --- 2. 核心輔助函數 ---
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

def clean_df(df):
    """核心清洗機制：過濾 Data Editor 產生的純空白列"""
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
    return "18時30分至19時00分"

def draw_page_number(canvas, doc):
    page_num = canvas.getPageNumber()
    canvas.setFont(_get_font(), 10)
    canvas.drawCentredString(105 * mm, 10 * mm, f"- 第 {page_num} 頁 -")

@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except: return None

@st.cache_data(ttl=5)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, None, "權限不足"
        sh = client.open_by_key(SHEET_ID)
        return (pd.DataFrame(sh.worksheet(WS_MAP["set"]).get_all_records()).fillna(""), 
                pd.DataFrame(sh.worksheet(WS_MAP["cmd"]).get_all_records()).fillna(""), 
                pd.DataFrame(sh.worksheet(WS_MAP["p1_cp"]).get_all_records()).fillna(""), 
                pd.DataFrame(sh.worksheet(WS_MAP["p2_cp"]).get_all_records()).fillna(""), None)
    except Exception as e: return None, None, None, None, str(e)

def save_data(unit, time_str, project, briefing, df_cmd, df_p1, df_p2, p1_desc, p2_desc):
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet(WS_MAP["set"])
        ws_set.clear()
        ws_set.update(range_name='A1', values=[
            ["Key", "Value"], 
            ["unit_name", unit], 
            ["plan_full_time", time_str], 
            ["project_name", project], 
            ["briefing_info", briefing],
            ["phase1_desc", p1_desc],
            ["phase2_desc", p2_desc]
        ])
        for ws_name, df in [(WS_MAP["cmd"], df_cmd), (WS_MAP["p1_cp"], df_p1), (WS_MAP["p2_cp"], df_p2)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df_cleaned = clean_df(df)
            if not df_cleaned.empty:
                ws.update(range_name='A1', values=[df_cleaned.columns.tolist()] + df_cleaned.astype(str).values.tolist())
        st.cache_data.clear()
        return True
    except: return False

# --- 3. 無線電代號自動填入邏輯 (完全比照二階段定點路檢系統) ---
def auto_assign_radio_code(df):
    if df is None or df.empty: return df
    base_prefixes = {"交通分隊": "99", "聖亭": "5", "龍潭": "6", "中興": "7", "石門": "8", "高平": "9", "三和": "3"}
    for idx, row in df.iterrows():
        unit, person, current_radio = safe_str(row.get('單位')), safe_str(row.get('服勤人員')), safe_str(row.get('無線電'))
        if current_radio != "": continue
        if not unit: continue
        first_unit = re.split(r'[\n、 ]', unit.strip())[0]
        base_pfx = next((v for k, v in base_prefixes.items() if k in first_unit), "")
        if base_pfx:
            if "副所長" in person: df.at[idx, '無線電'] = f"隆安{base_pfx}2"
            elif "所長" in person: df.at[idx, '無線電'] = f"隆安{base_pfx}1"
            else: df.at[idx, '無線電'] = f"隆安{base_pfx}0"
    return df

# 一鍵自動人員交叉帶入邏輯
def sync_personnel_data(df_p1, df_p2):
    df_p1 = clean_df(df_p1)
    if df_p1.empty or df_p2.empty: return df_p2
    split_pattern = r'[、,，\s/]+'
    p_dict = {}
    for _, row in df_p1.iterrows():
        unit_str = str(row.get('單位', '')).replace('龍潭交通分隊', '交通分隊')
        units = [u.strip() for u in re.split(split_pattern, unit_str) if u.strip()]
        persons_str = str(row.get('服勤人員', '')).strip()
        current_persons = [p.strip() for p in re.split(split_pattern, persons_str) if p.strip()]
        for u in units:
            if u not in p_dict: p_dict[u] = []
            if not current_persons: continue
            for p in current_persons:
                if p not in p_dict[u]: p_dict[u].append(p)
                
    df_p2_new = df_p2.copy()
    for idx, row in df_p2_new.iterrows():
        u_str = str(row.get('單位', '')).replace('龍潭交通分隊', '交通分隊')
        u_list = [u.strip() for u in re.split(split_pattern, u_str) if u.strip()]
        combined = []
        for u in u_list:
            for p in p_dict.get(u, []):
                if p not in combined: combined.append(p)
        if combined: df_p2_new.at[idx, '服勤人員'] = "、".join(combined)
    return df_p2_new

# --- 4. PDF 報表生成 (比照二階段路檢雙定點格式) ---
def generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_p1, df_p2, p1_desc, p2_desc):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []
    
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=24, alignment=1, spaceAfter=8, wordWrap='CJK')
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=10, wordWrap='CJK')
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=18, alignment=1, wordWrap='CJK')
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=18, alignment=0, wordWrap='CJK')
    style_middle_block = ParagraphStyle('MiddleBlock', fontName=font, fontSize=14, leading=22, spaceAfter=2*mm, alignment=TA_LEFT, leftIndent=5*mm, wordWrap='CJK')
    style_table_title = ParagraphStyle('TTitle', fontName=font, fontSize=16, alignment=1, leading=22, wordWrap='CJK')
    
    def clean_p(t): return safe_str(t).replace("\n", "<br/>").replace("、", "<br/>")
    def clean_text_only(t): return safe_str(t).replace("\n", "<br/>")

    story.append(Paragraph(f"{unit}執行{project}勤務規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    # 督導指揮編組表
    df_cmd = clean_df(df_cmd)
    data_cmd = [[Paragraph("<b>任 務 編 組</b>", style_table_title), '', '', ''],
                [Paragraph(f"<b>{h}</b>", style_cell) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([Paragraph(f"<b>{clean_text_only(r.get('職稱'))}</b>", style_cell), Paragraph(clean_text_only(r.get('代號')), style_cell), Paragraph(clean_p(r.get('姓名')), style_cell), Paragraph(clean_text_only(r.get('任務')), style_cell_left)])
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.12, page_width*0.28, page_width*0.45])
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('SPAN',(0,0),(-1,0)),('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t1)
    
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph("<b>📢 勤前教育：</b>", style_middle_block))
    story.append(Paragraph(f"{clean_text_only(briefing)}", style_middle_block))
    story.append(Spacer(1, 5*mm))
    
    # 第一階段表格 (完全沿用定點路檢的標準5欄位格式)
    df_p1 = clean_df(df_p1)
    story.append(Paragraph(f"<b>{p1_desc}</b>", style_middle_block))
    data_p1 = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "無線電", "單位", "服勤人員", "任務分工"]]]
    for _, r in df_p1.iterrows():
        data_p1.append([Paragraph(clean_text_only(r.get('編組')), style_cell), Paragraph(clean_text_only(r.get('無線電')), style_cell), Paragraph(clean_p(r.get('單位')), style_cell), Paragraph(clean_p(r.get('服勤人員')), style_cell), Paragraph(clean_text_only(r.get('任務分工')), style_cell_left)])
    t2 = Table(data_p1, colWidths=[page_width*0.15, page_width*0.12, page_width*0.13, page_width*0.20, page_width*0.40])
    t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t2)
    
    story.append(Spacer(1, 6*mm))
    
    # 第二階段表格 (定點路檢 / 擴大臨檢)
    df_p2 = clean_df(df_p2)
    story.append(Paragraph(f"<b>{p2_desc}</b>", style_middle_block))
    data_p2 = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "無線電", "單位", "服勤人員", "任務分工"]]]
    for _, r in df_p2.iterrows():
        data_p2.append([Paragraph(clean_text_only(r.get('編組')), style_cell), Paragraph(clean_text_only(r.get('無線電')), style_cell), Paragraph(clean_p(r.get('單位')), style_cell), Paragraph(clean_p(r.get('服勤人員')), style_cell), Paragraph(clean_text_only(r.get('任務分工')), style_cell_left)])
    t3 = Table(data_p2, colWidths=[page_width*0.15, page_width*0.12, page_width*0.13, page_width*0.20, page_width*0.40])
    t3.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e6e6e6')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t3)
    
    doc.build(story, onFirstPage=draw_page_number, onLaterPages=draw_page_number)
    return buf.getvalue()

def generate_attendance_pdf(unit, project, time_str, briefing):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 30*mm
    story = []
    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, leading=22, alignment=1, spaceAfter=8, wordWrap='CJK')
    style_top_info = ParagraphStyle('TopInfo', fontName=font, fontSize=12, leading=18, alignment=0, wordWrap='CJK')
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=24, alignment=1, wordWrap='CJK')
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=24, alignment=0, wordWrap='CJK') 
    
    story.append(Paragraph(f"{unit}執行{project}勤務簽到表", style_title))
    meeting_range = parse_meeting_time(time_str)
    date_part = time_str.split('日')[0] + '日' if '日' in time_str else ""
    story.append(Paragraph(f"時間：{date_part}{meeting_range}", style_top_info))
    loc = str(briefing).strip() if "於" not in str(briefing) else str(briefing).strip().split("於")[1]
    story.append(Paragraph(f"地點：{loc if loc else '分局二樓會議室'}", style_top_info))
    story.append(Spacer(1, 3*mm))
    
    table_data = [[Paragraph("分局長：", style_cell_left), "", Paragraph("上級督導：", style_cell_left), ""],
                  [Paragraph("副分局長：", style_cell_left), "", "", ""],
                  [Paragraph("單位", style_cell), Paragraph("參加人員", style_cell), Paragraph("單位", style_cell), Paragraph("參加人員", style_cell)]]
    
    rows = [("交通組", "中興派出所"), ("督察組", "石門派出所"), ("勤務指揮中心", "高平派出所"), ("聖亭派出所", "三和派出所"), ("龍潭派出所", "龍潭交通分隊")]
    for l, r in rows: table_data.append([Paragraph(l, style_cell), "", Paragraph(r, style_cell), ""])
    t = Table(table_data, colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3], rowHeights=[18*mm, 18*mm, 10*mm] + [26*mm]*len(rows))
    t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), font), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('SPAN', (0,1), (3,1))]))
    story.append(t)
    doc.build(story, onFirstPage=draw_page_number, onLaterPages=draw_page_number)
    return buf.getvalue()

def send_report_email(unit, project, time_str, briefing, df_cmd, df_p1, df_p2, p1_desc, p2_desc):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        msg["Subject"] = f"{unit}執行{project}二合一規劃與簽到表_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為最新版二合一專案勤務規劃文件與簽到表。", "plain", "utf-8"))
        
        p1 = generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_p1, df_p2, p1_desc, p2_desc)
        part1 = MIMEBase("application", "pdf"); part1.set_payload(p1); encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(f'{unit}二合一規劃表.pdf')}"); msg.attach(part1)
        
        p2 = generate_attendance_pdf(unit, project, time_str, briefing)
        part2 = MIMEBase("application", "pdf"); part2.set_payload(p2); encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(f'{unit}二合一簽到表.pdf')}"); msg.attach(part2)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd); server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# --- 5. Streamlit 主系統介面 ---
df_set, df_cmd, df_p1, df_p2, err = load_data()
d = dict(zip(df_set.iloc[:, 0].astype(str), df_set.iloc[:, 1].astype(str))) if df_set is not None else {}

u = d.get("unit_name", DEFAULT_UNIT)
t = d.get("plan_full_time", DEFAULT_TIME)
p = d.get("project_name", DEFAULT_PROJ)
b = d.get("briefing_info", DEFAULT_BRIEF)
p1_d = d.get("phase1_desc", DEFAULT_P1_DESC)
p2_d = d.get("phase2_desc", DEFAULT_P2_DESC)

st.title("🚓 二合一勤務規劃系統")
c1, c2 = st.columns(2)
p_name = c1.text_input("專案名稱", p)
p_time = c2.text_input("勤務時間", t)

st.subheader("⚙️ 階段標題與說明")
cc1, cc2 = st.columns(2)
phase1_desc = cc1.text_input("第一階段標題說明", p1_d)
phase2_desc = cc2.text_input("第二階段標題說明", p2_d)

st.subheader("1. 指揮編組")
res_cmd = st.data_editor(df_cmd if not df_cmd.empty else DEFAULT_CMD.copy(), num_rows="dynamic", use_container_width=True)
b_info = st.text_area("📢 勤前教育", b, height=70)

st.subheader("2. 勤務執行編組")
tab1, tab2 = st.tabs(["📍 第一階段（定點路檢）", "🚧 第二階段（擴大臨檢）"])

with tab1:
    st.info(f"當前標題：{phase1_desc}")
    # 第一階段採用完全一致的「定點路檢五欄位架構」，並內建自動隆安無線電代碼生成邏輯
    p1_source = df_p1 if not df_p1.empty else DEFAULT_P1_DF.copy()
    if "無線電" not in p1_source.columns and "編組" in p1_source.columns:
        p1_source = DEFAULT_P1_DF.copy() # 舊欄位不相容時強制重置為新架構
    res_p1 = auto_assign_radio_code(st.data_editor(p1_source, num_rows="dynamic", use_container_width=True, key="p1_editor"))

with tab2:
    st.info(f"當前標題：{phase2_desc}")
    p2_source = df_p2 if not df_p2.empty else DEFAULT_P2_DF.copy()
    
    if st.button("🔄 一鍵自動帶入第一階段人員"):
        st.session_state["synced_p2"] = sync_personnel_data(res_p1, p2_source)
        st.rerun()
        
    current_p2 = st.session_state.get("synced_p2", p2_source)
    res_p2 = auto_assign_radio_code(st.data_editor(current_p2, num_rows="dynamic", use_container_width=True, key="p2_editor"))

st.markdown("---")
pdf_plan = generate_pdf_from_data(u, p_name, p_time, b_info, clean_df(res_cmd), clean_df(res_p1), clean_df(res_p2), phase1_desc, phase2_desc)
pdf_attendance = generate_attendance_pdf(u, p_name, p_time, b_info)

col_dl1, col_dl2 = st.columns(2)
col_dl1.download_button("📝 下載規劃表", data=pdf_plan, file_name=f"{u}執行{p_name}勤務規劃表.pdf", use_container_width=True)
col_dl2.download_button("🖋️ 下載簽到表", data=pdf_attendance, file_name=f"{u}執行{p_name}勤務簽到表.pdf", use_container_width=True)

if st.button("💾 同步雲端並發送 Email 備份", use_container_width=True):
    with st.spinner("同步雲端與發信處理中..."):
        if save_data(u, p_time, p_name, b_info, res_cmd, res_p1, res_p2, phase1_desc, phase2_desc):
            ok, mail_err = send_report_email(u, p_name, p_time, b_info, res_cmd, res_p1, res_p2, phase1_desc, phase2_desc)
            if ok: st.success("✅ 同步與發信成功！二合一階段設定、標準路檢編組已完美同步至 Google 雲端。")
            else: st.error(f"❌ 發信失敗: {mail_err}")
