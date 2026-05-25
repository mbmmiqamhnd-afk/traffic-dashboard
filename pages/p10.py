import streamlit as st

# --- 1. 頁面設定 (必須是全站第一個執行的 Streamlit 指令) ---
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

from menu import show_sidebar
show_sidebar()

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import io, os, re, smtplib
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

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT_TITLE = "桃園市政府警察局龍潭分局"

WS_MAP = {
    "set": "危駕_設定",
    "cmd": "危駕_指揮組",
    "ptl": "危駕_警力佈署"
}

CHECKIN_POINTS = """1. 中油高原交流道站（龍源路2-20號）
2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）
3. 7-11龍潭佳園門市（中正路三坑段776號）
4. 旭日路三坑自然生態公園停車場
5. 旭日路與大溪區交界處"""

NOTES = """一、各編組執行前由帶班人員在駐地實施勤前教育。
二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。
三、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。
四、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"""

DEFAULT_TIME_VAL = "115年5月22日22時至翌日6時"
DEFAULT_CMDR_VAL = "龍潭所副所長全楚文"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "業務組", "代號": "隆安13", "姓名": "交通組巡官郭勝隆", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報群聚飆車狀況。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗", "任務": "督導各編組服儀裝備及勤務紀律"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任蔡奇青\n執勤官李文章\n執勤員黃文興", "任務": "監看群聚告警訊息、指揮、調度及通報本勤務事宜"}
])

DEFAULT_SCHEDULE = pd.DataFrame([
    {
        "勤務時段": "5月23日\n零時至4時", 
        "代號": "隆安62", 
        "編組": "專責警力\n（龍潭輪值）", 
        "服勤人員": "00-02時段:\n警員廖怡惠\n警員劉柏延\n\n02-04時段:\n警員林軒宇\n警員廖怡惠", 
        "任務分工": "「加強防制」勤務，在文化路、中正路三坑段、龍源路及旭日路來回巡邏，隨機攔檢改裝(噪音)車輛（每2小時至責任區域內指定巡簽地點巡簽1次並守望10分鐘，將守望情形拍照上傳LINE「龍潭分局聯絡平臺」群組）"
    },
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安80", "編組": "石門", "服勤人員": "線上巡邏組合警力兼任", "任務分工": "「區域聯防」勤務，於中正路、文化路、中豐路、龍源路及旭日路巡邏(每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次)"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安90", "編組": "高平", "服勤人員": "線上巡邏組合警力兼任", "任務分工": "「區域聯防」勤務，於中豐路及龍源路巡邏(每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次)"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安990", "編組": "交通分隊", "服勤人員": "線上巡邏組合警力兼任", "任務分工": "「區域聯防」勤務，於龍源路及及旭日路巡邏(每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次)"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安50", "編組": "聖亭", "服勤人員": "線上巡邏組合警力兼任", "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安60", "編組": "龍潭", "服勤人員": "線上巡邏組合警力兼任", "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安70", "編組": "中興", "服勤人員": "線上巡邏組合警力兼任", "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"}
])

def normalize(s): return str(s).replace('\n', '').replace('\r', '').replace(' ', '').strip()
def is_blank(val): return normalize(val) in ["", "None", "nan"]

# --- 2. Google Sheets 連線 (終極金鑰清洗版) ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    
    # 終極防禦：針對 InvalidByte(1623, 61) 的格式清洗
    if "private_key" in creds_dict:
        pk = str(creds_dict["private_key"])
        # 防呆：如果使用者不小心把其他欄位(如client_email=...)貼進了引號內，用正則切斷
        if "client_email" in pk:
            pk = pk.split("client_email")[0]
        
        pk = pk.replace("\\n", "\n")
        
        # 強制重新拼裝標準 PEM 格式，根絕一切隱形字元與錯誤等號
        if "-----BEGIN PRIVATE KEY-----" in pk and "-----END PRIVATE KEY-----" in pk:
            body = pk.split("-----BEGIN PRIVATE KEY-----")[-1].split("-----END PRIVATE KEY-----")[0]
            # 清除所有的換行、空格、引號
            body = re.sub(r'[\n\r\s\"\']', '', body)
            # 每 64 個字元強制換行 (PEM 標準)
            clean_body = "\n".join([body[i:i+64] for i in range(0, len(body), 64)])
            creds_dict["private_key"] = f"-----BEGIN PRIVATE KEY-----\n{clean_body}\n-----END PRIVATE KEY-----\n"

    try:
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Google 授權失敗：{e}")
        return None

def init_sheets():
    try:
        client = get_client()
        if not client: return
        sh = client.open_by_key(SHEET_ID)
        headers = {
            WS_MAP["set"]: [["Key", "Value"]],
            WS_MAP["cmd"]: [["職稱", "代號", "姓名", "任務"]],
            WS_MAP["ptl"]: [["勤務時段", "代號", "編組", "服勤人員", "任務分工"]]
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
    except Exception as e: 
        return None, None, None, str(e)

def save_data(p_time, cmdr, df_c, df_p):
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet(WS_MAP["set"])
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["plan_time", p_time], ["commander", cmdr]])
        for ws_name, df in [(WS_MAP["cmd"], df_c), (WS_MAP["ptl"], df_p)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df_cleaned = df.dropna(how='all').fillna("")
            if not df_cleaned.empty:
                ws.update([df_cleaned.columns.tolist()] + df_cleaned.values.tolist())
        load_data.clear()
        return True
    except Exception as e: 
        st.error(f"儲存失敗：{e}")
        return False

# --- 3. 寄信功能 ---
def send_report_email(time_str, commander, df_cmd, df_patrol, custom_filename):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"], msg["Subject"] = sender, sender, custom_filename
        msg.attach(MIMEText(f"附件為「{custom_filename}」PDF 規劃表。", "plain", "utf-8"))
        
        pdf_buf = generate_pdf(time_str, commander, df_cmd, df_patrol)
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_buf.getvalue())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(custom_filename+'.pdf')}")
        msg.attach(part)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# --- 4. 邏輯運算與 PDF ---
def calc_time_strings(p_time):
    date_match = re.search(r'(?:(\d+)年)?(\d+)月(\d+)日(.*)', p_time)
    if not date_match: return "", ""
    y, m, d = date_match.group(1), int(date_match.group(2)), int(date_match.group(3))
    time_part = date_match.group(4).strip() or "22時至翌日6時"
    y_tw = int(y) if y else (datetime.now().year - 1911)
    base_dt = datetime(y_tw + 1911, m, d)
    next_dt = base_dt + timedelta(days=1)
    return f"{next_dt.month}月{next_dt.day}日\n零時至4時", f"{m}月{d}日\n{time_part}"

def get_unit_details(cmdr_input):
    unit_base = "隆安99" if "分隊" in cmdr_input else "隆安8" if "石門" in cmdr_input else "隆安9" if "高平" in cmdr_input else "隆安5" if "聖亭" in cmdr_input else "隆安6" if "龍潭" in cmdr_input else "隆安7" if "中興" in cmdr_input else "隆安"
    suffix = "1" if ("所長" in cmdr_input and "副" not in cmdr_input) else "2"
    # 依修正紀錄：移除「所」等單位頭銜
    unit_name = "石門" if "石門" in cmdr_input else "高平" if "高平" in cmdr_input else "聖亭" if "聖亭" in cmdr_input else "龍潭" if "龍潭" in cmdr_input else "中興" if "中興" in cmdr_input else "交通分隊" if "分隊" in cmdr_input else cmdr_input[:2]
    return unit_base + suffix, f"專責警力\n（{unit_name}輪值）"

def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    for p in ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/custom/kaiu.ttf"]:
        if os.path.exists(p): pdfmetrics.registerFont(TTFont(fname, p)); return fname
    return "Helvetica"

def generate_pdf(time_str, commander, df_cmd, df_patrol):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []
    
    def draw_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 10)
        canvas.drawCentredString(A4[0] / 2, 10 * mm, f"- {doc.page} -")
        canvas.restoreState()
    
    style_title = ParagraphStyle('T', fontName=font, fontSize=16, alignment=1, spaceAfter=8)
    style_info = ParagraphStyle('I', fontName=font, fontSize=12, alignment=2, spaceAfter=10)
    style_th = ParagraphStyle('H', fontName=font, fontSize=16, alignment=1, leading=22)
    style_cell = ParagraphStyle('C', fontName=font, fontSize=14, leading=20, alignment=1) 
    style_cell_l = ParagraphStyle('L', fontName=font, fontSize=14, leading=20, alignment=0)
    style_note_hanging = ParagraphStyle('NH', fontName=font, fontSize=14, leading=20, alignment=0, leftIndent=28, firstLineIndent=-28)

    story.append(Paragraph(f"<b>{UNIT_TITLE}執行「防制危險駕車專案勤務」規劃表</b>", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def br(txt):
        if not txt: return ""
        s = str(txt).replace('\n', '<br/>').replace('\xa0', ' ')
        return re.sub(r'(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?[:：]?)', r'<b>\1</b>', s)

    data_cmd = [[Paragraph("<b>任 務 編 組</b>", style_th), '', '', ''], [Paragraph(f"<b>{h}</b>", style_th) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        if all(str(v).strip() == "" for v in r.values): continue
        data_cmd.append([Paragraph(br(r.get('職稱','')), style_cell), Paragraph(br(r.get('代號', '')), style_cell), Paragraph(br(r.get('姓名','')), style_cell), Paragraph(br(r.get('任務','')), style_cell_l)])
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.15, page_width*0.25, page_width*0.45])
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))]))
    story.append(t1); story.append(Spacer(1, 6*mm))

    data_ptl = [[Paragraph("<b>警 力 佈 署</b>", style_th), '', '', '', ''], [Paragraph(f"<b>交通快打指揮官：</b>{commander}", style_cell_l), '', '', '', ''], [Paragraph(f"<b>{h}</b>", style_th) for h in ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]]]
    for _, r in df_patrol.iterrows():
        if all(str(v).strip() == "" for v in r.values): continue
        data_ptl.append([Paragraph(br(r.get('勤務時段','')), style_cell), Paragraph(br(r.get('代號', '')), style_cell), Paragraph(br(r.get('編組','')), style_cell), Paragraph(br(r.get('服勤人員','')), style_cell), Paragraph(br(r.get('任務分工','')), style_cell_l)])

    t2 = Table(data_ptl, colWidths=[page_width*0.22, page_width*0.13, page_width*0.15, page_width*0.23, page_width*0.27])
    t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('SPAN',(0,1),(-1,1)), ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')), ('BACKGROUND',(0,2),(-1,2),colors.HexColor('#f2f2f2'))]))
    story.append(t2); story.append(Spacer(1, 6*mm))
    
    story.append(Paragraph("<b>📍 巡簽地點：</b>", style_cell_l))
    for l in CHECKIN_POINTS.split('\n'):
        if l.strip(): story.append(Paragraph(l.strip(), style_note_hanging))
    story.append(Spacer(1, 4*mm)); story.append(Paragraph("<b>📝 備註：</b>", style_cell_l))
    for l in NOTES.split('\n'):
        if l.strip(): story.append(Paragraph(l.strip(), style_note_hanging))

    doc.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)
    buf.seek(0); return buf

def get_preview_html(df_c, df_p, cmdr_n, time_s):
    cmd_rows = ""
    for _, r in df_c.iterrows():
        if all(str(v).strip() == "" for v in r.values): continue
        cmd_rows += f"<tr><td>{str(r.get('職稱','')).replace('\\n','<br>')}</td><td>{r.get('代號','')}</td><td>{str(r.get('姓名','')).replace('\\n','<br>').replace('\n','<br>')}</td><td>{r.get('任務','')}</td></tr>"
    ptl_rows = ""
    for _, r in df_p.iterrows():
        if all(str(v).strip() == "" for v in r.values): continue
        ptl_rows += f"<tr><td>{str(r.get('勤務時段','')).replace('\\n','<br>').replace('\n','<br>')}</td><td>{r.get('代號','')}</td><td>{str(r.get('編組','')).replace('\\n','<br>').replace('\n','<br>')}</td><td>{str(r.get('服勤人員','')).replace('\\n','<br>').replace('\n','<br>')}</td><td>{r.get('任務分工','')}</td></tr>"
    return f"""<style>table {{ width:100%; border-collapse:collapse; font-family:"標楷體"; }} th,td {{ border:1px solid black; padding:8px; text-align:center; }} th {{ background:#f2f2f2; font-size:16pt; }} td {{ font-size:14pt; }}</style>
    <h2 style='text-align:center;'>{UNIT_TITLE} 規劃表</h2><div style='text-align:right;'>勤務時間：{time_s}</div><br>
    <table><tr><th colspan="4">任 務 編 組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>{cmd_rows}</table><br>
    <table><tr><th colspan="5">警 力 佈 署</th></tr><tr><th colspan="5" style="text-align:left;">交通快打指揮官：{cmdr_n}</th></tr><tr><th>勤務時段</th><th>代號</th><th>編組</th><th>服勤人員</th><th>任務分工</th></tr>{ptl_rows}</table>"""

# ============================================================
# 5. 主介面與側邊欄設定
# ============================================================
st.sidebar.title("🛠️ 雲端設定")
if st.sidebar.button("初始化/檢查雲端分頁"): init_sheets()
if st.sidebar.button("⚠️ 強制重置為最新專案資料 (覆蓋雲端)"):
    with st.spinner("重置中..."):
        save_data(DEFAULT_TIME_VAL, DEFAULT_CMDR_VAL, DEFAULT_CMD, DEFAULT_SCHEDULE)
        st.cache_data.clear()
        st.rerun()

st.title("🚔 防制危險駕車專案勤務規劃表")

df_set, df_cmd_raw, df_ptl_raw, err = load_data()

if 'data_ptl' not in st.session_state:
    if err and err != "權限不足":
        st.warning(f"⚠️ 尚未初始化雲端資料，目前顯示預設底稿。請點擊左側「初始化/檢查雲端分頁」。")
        
    if df_set is not None and not df_set.empty:
        sd = dict(zip(df_set.iloc[:, 0].astype(str), df_set.iloc[:, 1].astype(str)))
        st.session_state.p_time = sd.get("plan_time", DEFAULT_TIME_VAL)
        st.session_state.cmdr = sd.get("commander", DEFAULT_CMDR_VAL)
        st.session_state.data_cmd = df_cmd_raw
        st.session_state.data_ptl = df_ptl_raw
    else:
        st.session_state.p_time = DEFAULT_TIME_VAL
        st.session_state.cmdr = DEFAULT_CMDR_VAL
        st.session_state.data_cmd = DEFAULT_CMD.copy()
        st.session_state.data_ptl = DEFAULT_SCHEDULE.copy()

if 'prev_p_time' not in st.session_state: st.session_state.prev_p_time = st.session_state.p_time
if 'prev_cmdr' not in st.session_state: st.session_state.prev_cmdr = st.session_state.cmdr
if 'last_ptl_len' not in st.session_state: st.session_state.last_ptl_len = len(st.session_state.data_ptl)

col1, col2 = st.columns(2)
p_time = col1.text_input("1. 勤務時間", st.session_state.p_time)
cmdr_input = col2.text_input("2. 交通快打指揮官", st.session_state.cmdr)

needs_sync_rerun = False
if p_time != st.session_state.prev_p_time:
    dedicated_time, normal_time = calc_time_strings(p_time)
    for i in range(len(st.session_state.data_ptl)):
        group = str(st.session_state.data_ptl.at[i, '編組']).strip()
        st.session_state.data_ptl.at[i, '勤務時段'] = dedicated_time if (i == 0 or "專責" in group) else normal_time
    st.session_state.prev_p_time = p_time; st.session_state.p_time = p_time; needs_sync_rerun = True

if cmdr_input != st.session_state.prev_cmdr:
    if len(st.session_state.data_ptl) > 0:
        daihao, bianzu = get_unit_details(cmdr_input)
        st.session_state.data_ptl.at[0, '代號'], st.session_state.data_ptl.at[0, '編組'] = daihao, bianzu
    st.session_state.prev_cmdr = cmdr_input; st.session_state.cmdr = cmdr_input; needs_sync_rerun = True

if needs_sync_rerun: st.rerun()

st.subheader("3. 任務編組")
res_cmd = st.data_editor(st.session_state.data_cmd, num_rows="dynamic", use_container_width=True).fillna("")
st.subheader("4. 警力佈署")
res_ptl_raw = st.data_editor(st.session_state.data_ptl, num_rows="dynamic", use_container_width=True).fillna("")

current_len = len(res_ptl_raw)
needs_editor_rerun = False
if current_len > st.session_state.last_ptl_len:
    _, normal_time = calc_time_strings(p_time)
    for i in range(st.session_state.last_ptl_len, current_len):
        res_ptl_raw.at[i, '勤務時段'], res_ptl_raw.at[i, '服勤人員'] = normal_time, ""
    st.session_state.data_ptl, st.session_state.last_ptl_len, needs_editor_rerun = res_ptl_raw, current_len, True
elif current_len < st.session_state.last_ptl_len:
    st.session_state.data_ptl, st.session_state.last_ptl_len = res_ptl_raw, current_len
else:
    st.session_state.data_ptl = res_ptl_raw

if len(st.session_state.data_ptl) > 0 and is_blank(st.session_state.data_ptl.at[0, '勤務時段']):
    dedicated_time, _ = calc_time_strings(p_time)
    st.session_state.data_ptl.at[0, '勤務時段'] = dedicated_time
    daihao, bianzu = get_unit_details(cmdr_input)
    st.session_state.data_ptl.at[0, '代號'], st.session_state.data_ptl.at[0, '編組'] = daihao, bianzu
    needs_editor_rerun = True

if needs_editor_rerun: st.rerun()

date_match = re.search(r'(?:(\d+)年)?(\d+)月(\d+)日', p_time)
date_fn = f"{date_match.group(1) if date_match.group(1) else str(datetime.now().year - 1911)}{str(date_match.group(2)).zfill(2)}{str(date_match.group(3)).zfill(2)}" if date_match else datetime.now().strftime('%m%d')
final_filename = f"防制危險駕車勤務規劃表_{date_fn}"

st.markdown("---")
with st.expander("📄 預覽勤務規劃表"):
    html_preview = get_preview_html(res_cmd, st.session_state.data_ptl, cmdr_input, p_time)
    st.components.v1.html(html_preview, height=600, scrolling=True)

col_sync, col_pdf = st.columns(2)

if col_sync.button("💾 同步雲端並發送備份郵件", type="primary", use_container_width=True):
    with st.spinner("處理中..."):
        if save_data(p_time, cmdr_input, res_cmd, st.session_state.data_ptl):
            mail_ok, mail_err = send_report_email(p_time, cmdr_input, res_cmd, st.session_state.data_ptl, final_filename)
            if mail_ok: st.success("✅ 同步與郵件發送成功！")
            else: st.warning(f"⚠️ 雲端已同步，但郵件失敗: {mail_err}")
        else: st.error("❌ 雲端同步失敗，請先點擊左側「初始化/檢查雲端分頁」。")

pdf_buf = generate_pdf(p_time, cmdr_input, res_cmd, st.session_state.data_ptl)
col_pdf.download_button(label="📝 下載規劃表 PDF", data=pdf_buf.getvalue(), file_name=f"{final_filename}.pdf", mime="application/pdf", use_container_width=True)
