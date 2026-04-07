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

# --- 常數與雲端設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

WS_MAP = {
    "set": "二階段_設定",
    "cmd": "二階段_指揮組",
    "ptl": "二階段_巡邏組",
    "cp":  "二階段_路檢組"
}

DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年4月8日20至24時"
DEFAULT_PROJ    = "0408全國同步擴大取締酒後駕車與防制危險駕車及噪音車輛專案勤務"
DEFAULT_BRIEF   = "20時30分於分局二樓會議室召開" 

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

@st.cache_data(ttl=5)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, None, "權限不足"
        sh = client.open_by_key(SHEET_ID)
        return (pd.DataFrame(sh.worksheet(WS_MAP["set"]).get_all_records()).fillna(""), 
                pd.DataFrame(sh.worksheet(WS_MAP["cmd"]).get_all_records()).fillna(""), 
                pd.DataFrame(sh.worksheet(WS_MAP["ptl"]).get_all_records()).fillna(""), 
                pd.DataFrame(sh.worksheet(WS_MAP["cp"]).get_all_records()).fillna(""), None)
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
        st.cache_data.clear()
        return True
    except: return False

# --- PDF 生成：修正標題與重複字詞 ---
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

    # 🌟 修正標題：移除重複的「勤務」二字
    display_title = f"{unit}執行{project}規劃表"
    story.append(Paragraph(display_title, style_title))
    story.append(Paragraph(f"時間：{time_str}", style_info))
    
    def clean(t): return safe_str(t).replace("\n", "<br/>").replace("、", "<br/>")
    def clean_text_only(t): return safe_str(t).replace("\n", "<br/>")

    # 1. 指揮組
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

    # 2. 勤務執行組 (合併版)
    story.append(Paragraph("<b>任務執行編組</b>", style_middle_block))
    data_task = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "代號", "單位", "服勤人員", "任務分工"]]]
    for _, r in df_cp.iterrows():
        data_task.append([Paragraph(clean_text_only(r.get('編組')), style_cell), Paragraph(clean_text_only(r.get('無線電')), style_cell), Paragraph(clean(r.get('單位')), style_cell), Paragraph(clean(r.get('服勤人員')), style_cell), Paragraph(clean_text_only(r.get('任務分工')), style_cell_left)])
    t2 = Table(data_task, colWidths=[page_width*0.15, page_width*0.12, page_width*0.13, page_width*0.20, page_width*0.40])
    t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t2)

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
    # 🌟 簽到表同步標題格式
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

# 🌟 發信函數：附件檔名對齊「規劃表」
def send_report_email(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        msg["Subject"] = f"{unit}執行{project}_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為最新 PDF 勤務規劃文件。", "plain", "utf-8"))
        
        # 🌟 關鍵修正：檔名對齊規劃表 (移除重複勤務二字)
        p1_name = f"{unit}執行{project}規劃表.pdf"
        p2_name = f"{unit}執行{project}簽到表.pdf"

        p1 = generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp)
        part1 = MIMEBase("application", "pdf"); part1.set_payload(p1); encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(p1_name)}"); msg.attach(part1)
        
        p2 = generate_attendance_pdf(unit, project, time_str, briefing)
        part2 = MIMEBase("application", "pdf"); part2.set_payload(p2); encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(p2_name)}"); msg.attach(part2)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd); server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# --- 邏輯功能 (略) ---
def auto_assign_radio_code(df):
    if df.empty: return df
    base_prefixes = {"交通分隊": "99", "聖亭": "5", "龍潭": "6", "中興": "7", "石門": "8", "高平": "9", "三和": "3"}
    for idx, row in df.iterrows():
        unit, person, cur_radio = safe_str(row.get('單位')), safe_str(row.get('服勤人員')), safe_str(row.get('無線電'))
        if cur_radio != "": continue
        if not unit: continue
        first_unit = re.split(r'[\n、 ]', unit.strip())[0]
        base_pfx = next((v for k, v in base_prefixes.items() if k in first_unit), "")
        if base_pfx:
            if "副所長" in person: df.at[idx, '無線電'] = f"隆安{base_pfx}2"
            elif "所長" in person: df.at[idx, '無線電'] = f"隆安{base_pfx}1"
            else: df.at[idx, '無線電'] = f"隆安{base_pfx}0"
    return df

def sync_personnel_data(df_ptl, df_cp):
    if df_ptl.empty or df_cp.empty: return df_cp
    split_pattern = r'[、,，\s/]+'
    p_dict = {}
    for _, row in df_ptl.iterrows():
        u_str = str(row.get('單位', '')).replace('龍潭交通分隊', '交通分隊')
        units = [u.strip() for u in re.split(split_pattern, u_str) if u.strip()]
        p_str = str(row.get('服勤人員', '')).strip()
        current_persons = [p.strip() for p in re.split(split_pattern, p_str) if p.strip()]
        for u in units:
            if u not in p_dict: p_dict[u] = []
            if not current_persons: continue
            for p in current_persons:
                if p not in p_dict[u]: p_dict[u].append(p)
    df_cp_new = df_cp.copy()
    for idx, row in df_cp_new.iterrows():
        u_str = str(row.get('單位', '')).replace('龍潭交通分隊', '交通分隊')
        u_list = [u.strip() for u in re.split(split_pattern, u_str) if u.strip()]
        combined = []
        for u in u_list:
            for p in p_dict.get(u, []):
                if p not in combined: combined.append(p)
        if combined: df_cp_new.at[idx, '服勤人員'] = "、".join(combined)
    return df_cp_new

# --- 3. 主程式介面 ---
df_set, df_cmd, df_ptl, df_cp, err = load_data()
d = dict(zip(df_set.iloc[:, 0].astype(str), df_set.iloc[:, 1].astype(str))) if df_set is not None else {}

u_val = d.get("unit_name", DEFAULT_UNIT)
t_val = d.get("plan_full_time", DEFAULT_TIME)
p_val = d.get("project_name", DEFAULT_PROJ)
b_val = d.get("briefing_info", DEFAULT_BRIEF)

st.title("🚓 勤務規劃管理系統")
c1, c2 = st.columns(2)
p_name = c1.text_input("專案名稱", p_val)
p_time = c2.text_input("時間範圍", t_val)

st.subheader("1. 指揮編組")
res_cmd = st.data_editor(df_cmd, num_rows="dynamic", use_container_width=True)
b_info = st.text_area("📢 勤前教育", b_val, height=70)

st.subheader("2. 執行編組編排")
if st.button("🔄 同步第一階段人員到最終編組"):
    st.session_state["synced_cp"] = sync_personnel_data(df_ptl, df_cp)
    st.rerun()

tab1, tab2 = st.tabs(["巡邏組原始資料", "規劃表顯示編組"])
with tab1:
    res_ptl = auto_assign_radio_code(st.data_editor(df_ptl, num_rows="dynamic", use_container_width=True, key="ptl_editor"))
with tab2:
    cur_cp = st.session_state.get("synced_cp", df_cp)
    res_cp = auto_assign_radio_code(st.data_editor(cur_cp, num_rows="dynamic", use_container_width=True, key="cp_editor"))

st.markdown("---")
# 生成 PDF
pdf_plan = generate_pdf_from_data(u_val, p_name, p_time, b_info, res_cmd, res_ptl, res_cp)
pdf_attendance = generate_attendance_pdf(u_val, p_name, p_time, b_info)

# 🌟 下載按鈕檔名修正
col_dl1, col_dl2 = st.columns(2)
col_dl1.download_button(f"📝 下載：{u_val}執行{p_name}規劃表.pdf", data=pdf_plan, file_name=f"{u_val}執行{p_name}規劃表.pdf", use_container_width=True)
col_dl2.download_button(f"🖋️ 下載：{u_val}執行{p_name}簽到表.pdf", data=pdf_attendance, file_name=f"{u_val}執行{p_name}簽到表.pdf", use_container_width=True)

if st.button("💾 儲存並發送 Email 備份", use_container_width=True):
    if save_data(u_val, p_time, p_name, b_info, res_cmd, res_ptl, res_cp):
        ok, m_err = send_report_email(u_val, p_name, p_time, b_info, res_cmd, res_ptl, res_cp)
        if ok: st.success("✅ 存檔發信成功！標題與檔名已修正為「規劃表」。")
        else: st.error(f"❌ 存檔成功但發信失敗: {m_err}")
