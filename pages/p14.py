import streamlit as st

# --- 1. 頁面設定 (必須是全站第一個執行的 Streamlit 指令) ---
st.set_page_config(page_title="二階段勤務規劃系統", layout="wide", page_icon="🚓")

# 呼叫側邊欄 (確保在 config 之後)
try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    pass

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

# --- 常數與工作表設定 ---
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
DEFAULT_PROJ    = "全國同步擴大取締酒後駕車與防制危險駕車及噪音車輛專案勤務"
DEFAULT_BRIEF   = "20時30分於分局二樓會議室召開"
DEFAULT_P1_DESC = "第一階段：21時至22時30分，機動巡邏"
DEFAULT_P2_DESC = "第二階段：22時30分至24時，定點路檢及機動攔檢"

EXPECTED_PTL_COLS = ["編組", "無線電", "單位", "職別", "姓名", "任務分工", "巡邏路段"]
EXPECTED_CP_COLS  = ["編組", "無線電", "單位", "職別", "姓名", "任務分工", "路檢地點"]

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

def extract_4_digit_date(time_str):
    try:
        match = re.search(r"(\d+)月(\d+)日", time_str)
        if match:
            month = match.group(1).zfill(2)
            day = match.group(2).zfill(2)
            return f"{month}{day}"
    except: pass
    return ""

def safe_str(val):
    if pd.isna(val) or val is None or str(val).strip().lower() == "nan": return ""
    return str(val)

def clean_df_to_list(df):
    return df.astype(str).values.tolist()

def draw_page_number(canvas, doc):
    page_num = canvas.getPageNumber()
    text = f"- 第 {page_num} 頁 -"
    canvas.setFont(_get_font(), 10)
    canvas.drawCentredString(105 * mm, 10 * mm, text)

# 計算 PDF 表格需要合併的儲存格 (SPAN)
def get_merge_styles(df, merge_cols):
    span_styles = []
    if df.empty: return span_styles
    
    cols_list = df.columns.tolist()
    for col_name in merge_cols:
        if col_name not in cols_list:
            continue
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
                else:
                    break
                    
            if end_idx > start_idx:
                span_styles.append(('SPAN', (c_idx, start_idx + 1), (c_idx, end_idx + 1)))
            
            start_idx = end_idx + 1
            
    return span_styles

# --- Google 授權 ---
@st.cache_resource
def get_client():
    try:
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Google 授權失敗：{e}")
        return None

@st.cache_data(ttl=600)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, None, "權限不足或未設定密鑰"
        sh = client.open_by_key(SHEET_ID)

        try:
            df_set = pd.DataFrame(sh.worksheet(WS_MAP["set"]).get_all_records()).fillna("")
        except:
            df_set = None

        try:
            df_cmd = pd.DataFrame(sh.worksheet(WS_MAP["cmd"]).get_all_records()).fillna("")
        except:
            df_cmd = pd.DataFrame()

        try:
            df_ptl = pd.DataFrame(sh.worksheet(WS_MAP["ptl"]).get_all_records()).fillna("")
        except:
            df_ptl = pd.DataFrame()

        try:
            df_cp = pd.DataFrame(sh.worksheet(WS_MAP["cp"]).get_all_records()).fillna("")
        except:
            df_cp = pd.DataFrame()

        return df_set, df_cmd, df_ptl, df_cp, None
    except Exception as e:
        return None, None, None, None, str(e)

def save_data(unit, time_str, project, briefing, df_cmd, df_ptl, df_cp, p1_desc, p2_desc):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)

        try:
            ws_set = sh.worksheet(WS_MAP["set"])
        except WorksheetNotFound:
            ws_set = sh.add_worksheet(title=WS_MAP["set"], rows="50", cols="5")
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

        try:
            ws_cmd = sh.worksheet(WS_MAP["cmd"])
        except WorksheetNotFound:
            ws_cmd = sh.add_worksheet(title=WS_MAP["cmd"], rows="100", cols="20")
        ws_cmd.clear()
        clean_cmd = df_cmd.dropna(how="all").fillna("")
        if not clean_cmd.empty:
            ws_cmd.update(range_name='A1', values=[clean_cmd.columns.tolist()] + clean_df_to_list(clean_cmd))

        try:
            ws_ptl = sh.worksheet(WS_MAP["ptl"])
        except WorksheetNotFound:
            ws_ptl = sh.add_worksheet(title=WS_MAP["ptl"], rows="100", cols="20")
        ws_ptl.clear()
        clean_ptl = df_ptl.dropna(how="all").fillna("")
        if not clean_ptl.empty:
            ws_ptl.update(range_name='A1', values=[clean_ptl.columns.tolist()] + clean_df_to_list(clean_ptl))

        try:
            ws_cp = sh.worksheet(WS_MAP["cp"])
        except WorksheetNotFound:
            ws_cp = sh.add_worksheet(title=WS_MAP["cp"], rows="100", cols="20")
        ws_cp.clear()
        clean_cp = df_cp.dropna(how="all").fillna("")
        if not clean_cp.empty:
            ws_cp.update(range_name='A1', values=[clean_cp.columns.tolist()] + clean_df_to_list(clean_cp))

        st.cache_data.clear()
        return True
    except APIError as e:
        st.error(f"❌ Google API 流量限制或連線錯誤：{e}")
        return False
    except Exception as e:
        st.error(f"❌ 同步失敗原因：{e}")
        st.code(traceback.format_exc())
        return False

# --- PDF 相關函數 ---
def generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, p1_desc, p2_desc):
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

    story.append(Paragraph(f"{unit}執行{project}勤務規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))

    def clean_p(t): return safe_str(t).replace("\n", "<br/>").replace("、", "<br/>")
    def clean_text_only(t): return safe_str(t).replace("\n", "<br/>")

    df_cmd = clean_df(df_cmd)
    data_cmd = [[Paragraph("<b>任 務 編 組</b>", style_table_title), '', '', ''],
                [Paragraph(f"<b>{h}</b>", style_cell) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([
            Paragraph(f"<b>{clean_text_only(r.get('職稱'))}</b>", style_cell),
            Paragraph(clean_text_only(r.get('代號')), style_cell),
            Paragraph(clean_p(r.get('姓名')), style_cell),
            Paragraph(clean_text_only(r.get('任務')), style_cell_left)
        ])
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.12, page_width*0.28, page_width*0.45])
    t1.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),font),
        ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('SPAN',(0,0),(-1,0)),
        ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE')
    ]))
    story.append(t1)

    story.append(Spacer(1, 6*mm))
    story.append(Paragraph("<b>📢 勤前教育：</b>", style_middle_block))
    story.append(Paragraph(f"{clean_text_only(briefing)}", style_middle_block))
    story.append(Spacer(1, 6*mm))

    # --- 第一階段：巡邏組 ---
    df_ptl = clean_df(df_ptl)
    story.append(Paragraph(f"<b>{p1_desc}</b>", style_middle_block))
    
    span_styles_ptl = get_merge_styles(df_ptl, ["編組", "無線電", "單位", "巡邏路段"])
    
    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in EXPECTED_PTL_COLS]]
    for _, r in df_ptl.iterrows():
        data_ptl.append([
            Paragraph(clean_text_only(r.get('編組')), style_cell),
            Paragraph(clean_text_only(r.get('無線電')), style_cell),
            Paragraph(clean_p(r.get('單位')), style_cell),
            Paragraph(clean_p(r.get('職別')), style_cell),
            Paragraph(clean_p(r.get('姓名')), style_cell),
            Paragraph(clean_p(r.get('任務分工')), style_cell),
            Paragraph(clean_text_only(r.get('巡邏路段')), style_cell_left)
        ])
    t2 = Table(data_ptl, colWidths=[page_width*0.10, page_width*0.12, page_width*0.12, page_width*0.10, page_width*0.14, page_width*0.14, page_width*0.28])
    base_style_ptl = [
        ('FONTNAME',(0,0),(-1,-1),font),
        ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE')
    ]
    t2.setStyle(TableStyle(base_style_ptl + span_styles_ptl))
    story.append(t2)

    story.append(Spacer(1, 8*mm))

    # --- 第二階段：路檢組 ---
    df_cp = clean_df(df_cp)
    story.append(Paragraph(f"<b>{p2_desc}</b>", style_middle_block))
    
    span_styles_cp = get_merge_styles(df_cp, ["編組", "無線電", "單位", "路檢地點"])
    
    data_cp = [[Paragraph(f"<b>{h}</b>", style_cell) for h in EXPECTED_CP_COLS]]
    for _, r in df_cp.iterrows():
        data_cp.append([
            Paragraph(clean_text_only(r.get('編組')), style_cell),
            Paragraph(clean_text_only(r.get('無線電')), style_cell),
            Paragraph(clean_p(r.get('單位')), style_cell),
            Paragraph(clean_p(r.get('職別')), style_cell),
            Paragraph(clean_p(r.get('姓名')), style_cell),
            Paragraph(clean_p(r.get('任務分工')), style_cell),
            Paragraph(clean_text_only(r.get('路檢地點')), style_cell_left)
        ])
    t3 = Table(data_cp, colWidths=[page_width*0.10, page_width*0.12, page_width*0.12, page_width*0.10, page_width*0.14, page_width*0.14, page_width*0.28])
    base_style_cp = [
        ('FONTNAME',(0,0),(-1,-1),font),
        ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e6e6e6')),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE')
    ]
    t3.setStyle(TableStyle(base_style_cp + span_styles_cp))
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
    story.append(Paragraph(f"地點：{loc}", style_top_info))
    story.append(Spacer(1, 3*mm))

    table_data = [
        [Paragraph("分局長：", style_cell_left), "", Paragraph("上級督導：", style_cell_left), ""],
        [Paragraph("副分局長：", style_cell_left), "", "", ""],
        [Paragraph("單位", style_cell), Paragraph("參加人員", style_cell), Paragraph("單位", style_cell), Paragraph("參加人員", style_cell)]
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
    t = Table(
        table_data,
        colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3],
        rowHeights=[18*mm, 18*mm, 10*mm] + [26*mm]*len(rows)
    )
    t.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('SPAN', (0,1), (3,1))
    ]))
    story.append(t)
    doc.build(story, onFirstPage=draw_page_number, onLaterPages=draw_page_number)
    return buf.getvalue()

def send_report_email(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, p1_desc, p2_desc):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        msg["Subject"] = f"{unit}執行{project}規劃與簽到表_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為最新 PDF 規劃文件。", "plain", "utf-8"))

        plan_filename = f"{unit}執行{project}勤務規劃表.pdf"
        attendance_filename = f"{unit}執行{project}勤務簽到表.pdf"

        p1 = generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp, p1_desc, p2_desc)
        part1 = MIMEBase("application", "pdf"); part1.set_payload(p1); encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(plan_filename)}"); msg.attach(part1)

        p2 = generate_attendance_pdf(unit, project, time_str, briefing)
        part2 = MIMEBase("application", "pdf"); part2.set_payload(p2); encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(attendance_filename)}"); msg.attach(part2)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd); server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# --- 核心邏輯區 ---
def auto_assign_radio_code(df):
    if df is None or df.empty: return df
    df_copy = df.copy()
    base_prefixes = {"交通分隊": "99", "聖亭": "5", "龍潭": "6", "中興": "7", "石門": "8", "高平": "9", "三和": "3"}
    for idx, row in df_copy.iterrows():
        unit, person, rank, current_radio = safe_str(row.get('單位')), safe_str(row.get('姓名')), safe_str(row.get('職別')), safe_str(row.get('無線電'))
        if current_radio != "": continue
        if not unit: continue
        first_unit = re.split(r'[\n、 ]', unit.strip())[0]
        base_pfx = next((v for k, v in base_prefixes.items() if k in first_unit), "")
        if base_pfx:
            if "副所長" in rank or "副所長" in person: df_copy.at[idx, '無線電'] = f"隆安{base_pfx}2"
            elif "所長" in rank or "所長" in person: df_copy.at[idx, '無線電'] = f"隆安{base_pfx}1"
            else: df_copy.at[idx, '無線電'] = f"隆安{base_pfx}0"
    return df_copy

# --- 3. 主程式介面 ---

if st.sidebar.button("🔄 強制從雲端更新資料"):
    st.cache_data.clear()
    st.rerun()

st.title("🚓 二階段勤務規劃系統")

df_set, df_cmd, df_ptl, df_cp, err = load_data()

if err:
    if "429" in str(err) or "Quota exceeded" in str(err):
        st.error("⚠️ Google 雲端連線過於頻繁（API 額度暫時用光），請等待 1 分鐘後再重新整理網頁。")
    else:
        st.warning(f"⚠️ 無法連線 Google Sheets ({err})，顯示預設或已載入資料。")

df_set = df_set if isinstance(df_set, pd.DataFrame) else pd.DataFrame()
df_cmd = df_cmd if (isinstance(df_cmd, pd.DataFrame) and not df_cmd.empty) else pd.DataFrame(columns=["職稱", "代號", "姓名", "任務"])

# --- 舊版本試算表欄位轉換與相容機制 ---
if isinstance(df_ptl, pd.DataFrame) and not df_ptl.empty:
    if '服勤人員' in df_ptl.columns and '姓名' not in df_ptl.columns:
        df_ptl['姓名'] = df_ptl['服勤人員']
    if '任務分工' in df_ptl.columns and '巡邏路段' not in df_ptl.columns:
        df_ptl['巡邏路段'] = df_ptl['任務分工']
        df_ptl['任務分工'] = ""
    for c in EXPECTED_PTL_COLS:
        if c not in df_ptl.columns: df_ptl[c] = ""
    df_ptl = df_ptl[EXPECTED_PTL_COLS]
else:
    df_ptl = pd.DataFrame(columns=EXPECTED_PTL_COLS)

if isinstance(df_cp, pd.DataFrame) and not df_cp.empty:
    if '服勤人員' in df_cp.columns and '姓名' not in df_cp.columns:
        df_cp['姓名'] = df_cp['服勤人員']
    if '任務分工' in df_cp.columns and '路檢地點' not in df_cp.columns:
        df_cp['路檢地點'] = df_cp['任務分工']
        df_cp['任務分工'] = ""
    for c in EXPECTED_CP_COLS:
        if c not in df_cp.columns: df_cp[c] = ""
    df_cp = df_cp[EXPECTED_CP_COLS]
else:
    df_cp = pd.DataFrame(columns=EXPECTED_CP_COLS)

d = dict(zip(df_set.iloc[:, 0].astype(str), df_set.iloc[:, 1].astype(str))) if not df_set.empty else {}

u    = d.get("unit_name",      DEFAULT_UNIT)
t    = d.get("plan_full_time", DEFAULT_TIME)
p    = d.get("project_name",   DEFAULT_PROJ)
b    = d.get("briefing_info",  DEFAULT_BRIEF)
p1_d = d.get("phase1_desc",    DEFAULT_P1_DESC)
p2_d = d.get("phase2_desc",    DEFAULT_P2_DESC)

clean_p_name = re.sub(r"^\d{4}", "", p)

c1, c2 = st.columns(2)
p_time  = c2.text_input("勤務時間", t)
p_input = c1.text_input("專案名稱", clean_p_name)

date_code = extract_4_digit_date(p_time)
p_name = f"{date_code}{p_input}" if date_code else p_input

st.subheader("⚙️ 階段標題與說明")
cc1, cc2 = st.columns(2)
phase1_desc = cc1.text_input("第一階段標題說明", p1_d)
phase2_desc = cc2.text_input("第二階段標題說明", p2_d)

st.subheader("1. 指揮編組")
res_cmd = st.data_editor(df_cmd, num_rows="dynamic", use_container_width=True).dropna(how="all").fillna("")
b_info  = st.text_area("📢 勤前教育", b, height=70)

st.subheader("2. 勤務編組")
tab1, tab2 = st.tabs(["📍 第一階段 (巡邏)", "🚧 第二階段 (路檢)"])

with tab1:
    st.info(f"當前標題：{phase1_desc}")
    raw_ptl = st.data_editor(df_ptl, num_rows="dynamic", use_container_width=True, key="ptl_editor")
    res_ptl = auto_assign_radio_code(raw_ptl).dropna(how="all").fillna("")

with tab2:
    st.info(f"當前標題：{phase2_desc}")
    raw_cp = st.data_editor(df_cp, num_rows="dynamic", use_container_width=True, key="cp_editor")
    res_cp = auto_assign_radio_code(raw_cp).dropna(how="all").fillna("")

st.markdown("---")

# --- 移除原本的下載按鈕欄位，直接保留雲端備份與發信功能 ---
if st.button("💾 同步雲端並發送 Email 備份", use_container_width=True):
    with st.spinner("同步中，請稍候…"):
        if save_data(u, p_time, p_name, b_info, res_cmd, res_ptl, res_cp, phase1_desc, phase2_desc):
            with st.spinner("同步成功，正在寄送郵件…"):
                ok, mail_err = send_report_email(u, p_name, p_time, b_info, res_cmd, res_ptl, res_cp, phase1_desc, phase2_desc)
            if ok:
                st.success(f"✅ 同步與發信成功！已在後台為專案自動補上「{date_code}」代碼。")
                st.rerun()  # 同步成功後重整前端，確保畫面立即更新最新配發的呼叫代碼
            else:
                st.error(f"❌ 發信失敗: {mail_err}")
