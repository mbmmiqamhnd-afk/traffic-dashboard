import streamlit as st

# --- 1. 頁面設定 (必須放在全站最頂端第一個 Streamlit 指令) ---
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

from menu import show_sidebar
show_sidebar()

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import io
import os
import re
import smtplib
import traceback
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

CHECKIN_POINTS = """1. 中油高原交流道站（龍源路2-20號）
2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）
3. 7-11龍潭佳園門市（中正路三坑段776號）
4. 旭日路三坑自然生態公園停車場
5. 旭日路與大溪區交界處"""

NOTES = """一、各編組執行前由帶班人員在駐地實施勤前教育。
二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。
三、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。
四、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"""

# ★★★ 5月22日 專案專屬精準底稿資料庫 ★★★
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
        "編組": "專責警力\n（龍潭所輪值）", 
        "服勤人員": "00-02時段:\n警員廖怡惠\n警員劉柏延\n\n02-04時段:\n警員林軒宇\n警員廖怡惠", 
        "任務分工": "「加強防制」勤務，在文化路、中正路三坑段、龍源路及旭日路來回巡邏，隨機攔檢改裝(噪音)車輛（每2小時至責任區域內指定巡簽地點巡簽1次並守望10分鐘，將守望情形拍照上傳LINE「龍潭分局聯絡平臺」群組）"
    },
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安80", "編組": "石門所", "服勤人員": "線上巡邏組合警力兼任", "任務分工": "「區域聯防」勤務，於中正路、文化路、中豐路、龍源路及旭日路巡邏(每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次)"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安90", "編組": "高平所", "服勤人員": "線上巡邏組合警力兼任", "任務分工": "「區域聯防」勤務，於中豐路及龍源路巡邏(每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次)"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安990", "編組": "龍潭交通分隊", "服勤人員": "線上巡邏組合警力兼任", "任務分工": "「區域聯防\"勤務，於龍源路及及旭日路巡邏(每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次)"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安50", "編組": "聖亭所", "服勤人員": "線上巡邏組合警力兼任", "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安60", "編組": "龍潭所", "服勤人員": "線上巡邏組合警力兼任", "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安70", "編組": "中興所", "服勤人員": "線上巡邏組合警力兼任", "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"}
])

# --- 工具函式 ---
def normalize(s):
    return str(s).replace('\n', '').replace('\r', '').replace(' ', '').strip()

def is_blank(val):
    return normalize(val) in ["", "None", "nan"]

# ─────────────── ★ 關鍵連線修正：高相容 PEM 金鑰自動重組機制 ★ ───────────────

@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    try:
        # 1. 把 Secrets 轉為可修改的獨立字典，保護後台不被污染
        creds_dict = dict(st.secrets["gcp_service_account"])
        
        # 2. 自動修正 PEM 多行與雙斜線金鑰地雷：
        #    不論後台是含有真實鍵盤換行、帶有文字 \n、或是多餘的空格與隱形字符，
        #    此段邏輯會強制進行「乾淨清洗與標準重組」，確保封包完全符合 cryptography 模組之 PEM 規範。
        if "private_key" in creds_dict and isinstance(creds_dict["private_key"], str):
            pk = creds_dict["private_key"].strip()
            # 移除所有可能的雙斜線 \n 文字，並用真正的換行符號還原
            pk = pk.replace("\\n", "\n")
            
            # 如果發現頭尾標籤以外的內文換行被破壞，進行正規化安全補強
            if "-----BEGIN PRIVATE KEY-----" in pk and "-----END PRIVATE KEY-----" in pk:
                header = "-----BEGIN PRIVATE KEY-----"
                footer = "-----END PRIVATE KEY-----"
                # 抽出內文，把內文所有的換行、回車與空格全部濾乾淨
                body = pk.replace(header, "").replace(footer, "").replace("\n", "").replace("\r", "").replace(" ", "")
                # 每 64 個字元強制補上一個標準 PEM 斷行（完美契合 PEM 檔案標準規範）
                clean_body = "\n".join([body[i:i+64] for i in range(0, len(body), 64)])
                # 重新拼回最標準、絕不噴錯的 PEM 結構
                creds_dict["private_key"] = f"{header}\n{clean_body}\n{footer}\n"
        
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Google 授權失敗：{e}")
        return None

def clean_df_to_list(df):
    return df.astype(str).values.tolist()

@st.cache_data(ttl=10)
def load_from_cloud():
    try:
        client = get_client()
        if not client: return None, None, None
        sh = client.open_by_key(SHEET_ID)
        ws_list = sh.worksheets()
        
        ws_set = next((w for w in ws_list if w.title == "危駕_設定"), None)
        ws_cmd = next((w for w in ws_list if w.title == "危駕_指揮組"), None)
        ws_ptl = next((w for w in ws_list if w.title == "危駕_警力佈署"), None)
        
        s = pd.DataFrame(ws_set.get_all_records()).fillna("") if ws_set else None
        c = pd.DataFrame(ws_cmd.get_all_records()).fillna("") if ws_cmd else pd.DataFrame()
        p = pd.DataFrame(ws_ptl.get_all_records()).fillna("") if ws_ptl else pd.DataFrame()
        return s, c, p
    except: 
        return None, None, None

def save_to_cloud(p_time, cmdr, df_c, df_p):
    try:
        client = get_client()
        if not client: return False
        sh = client.open_by_key(SHEET_ID)
        
        # 1. 危駕_設定
        try:
            ws_set = sh.worksheet("危駕_設定")
        except Exception:
            ws_set = sh.add_worksheet(title="危駕_設定", rows="50", cols="5")
        ws_set.clear()
        ws_set.update(range_name='A1', values=[["Key", "Value"], ["plan_time", p_time], ["commander", cmdr]])
        
        # 2. 危駕_指揮組
        try:
            ws_cmd = sh.worksheet("危駕_指揮組")
        except Exception:
            ws_cmd = sh.add_worksheet(title="危駕_指揮組", rows="100", cols="20")
        ws_cmd.clear()
        clean_cmd = df_c.dropna(how="all").fillna("")
        if not clean_cmd.empty:
            ws_cmd.update(range_name='A1', values=[clean_cmd.columns.tolist()] + clean_df_to_list(clean_cmd))
            
        # 3. 危駕_警力佈署
        try:
            ws_ptl = sh.worksheet("危駕_警力佈署")
        except Exception:
            ws_ptl = sh.add_worksheet(title="危駕_警力佈署", rows="100", cols="20")
        ws_ptl.clear()
        clean_ptl = df_p.dropna(how="all").fillna("")
        if not clean_ptl.empty:
            ws_ptl.update(range_name='A1', values=[clean_ptl.columns.tolist()] + clean_df_to_list(clean_ptl))
            
        load_from_cloud.clear()
        return True
    except Exception as e:
        st.error(f"❌ 同步失敗原因：{e}")
        st.code(traceback.format_exc())
        return False

# --- 寄信功能 ---
def send_report_email(time_str, commander, df_cmd, df_patrol, custom_filename):
    try:
        if "email" not in st.secrets: return False, "未在 secrets 中設定 email 資訊"
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        
        pdf_buf = generate_pdf(time_str, commander, df_cmd, df_patrol)
        pdf_bytes = pdf_buf.getvalue()
        
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender
        msg["Subject"] = custom_filename
        msg.attach(MIMEText(f"附件為「{custom_filename}」PDF 規劃表。", "plain", "utf-8"))
        
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        filename_encoded = _ul.quote(f"{custom_filename}.pdf")
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{filename_encoded}")
        msg.attach(part)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# --- 邏輯運算函式 ---
def calc_time_strings(p_time):
    date_match = re.search(r'(?:(\d+)年)?(\d+)月(\d+)日(.*)', p_time)
    if not date_match: return "", ""
    y, m, d = date_match.group(1), int(date_match.group(2)), int(date_match.group(3))
    time_part = date_match.group(4).strip() or "22時至翌日6時"
    y_tw = int(y) if y else (datetime.now().year - 1911)
    base_dt = datetime(y_tw + 1911, m, d)
    next_dt = base_dt + timedelta(days=1)
    dedicated_time = f"{next_dt.month}月{next_dt.day}日\n零時至4時"
    normal_time = f"{m}月{d}日\n{time_part}"
    return dedicated_time, normal_time

def get_unit_details(cmdr_input):
    if "分隊" in cmdr_input: unit_base = "隆安99"
    elif "石門" in cmdr_input: unit_base = "隆安8"
    elif "高平" in cmdr_input: unit_base = "隆安9"
    elif "聖亭" in cmdr_input: unit_base = "隆安5"
    elif "龍潭" in cmdr_input: unit_base = "隆安6"
    elif "中興" in cmdr_input: unit_base = "隆安7"
    else: unit_base = "隆安"
    
    suffix = "1" if ("所長" in cmdr_input and "副" not in cmdr_input) else "2"
    unit_match = re.search(r'([\u4e00-\u9fa5]+?(?:派出所|所|分隊|警備隊))', cmdr_input)
    unit_name = re.sub(r'派出所$', '所', unit_match.group(1)) if unit_match else cmdr_input[:3]
    
    return unit_base + suffix, f"專責警力\n（{unit_name}輪值）"

# --- PDF 字型與產生 ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    for p in ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/custom/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
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
        page_num_text = f"- {doc.page} -"
        canvas.drawCentredString(A4[0] / 2, 10 * mm, page_num_text)
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
        s = re.sub(r'(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?[:：]?)', r'<b>\1</b>', s)
        return s

    data_cmd = [[Paragraph("<b>任　務　編　組</b>", style_th), '', '', ''], 
                [Paragraph(f"<b>{h}</b>", style_th) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        if all(str(v).strip() == "" for v in r.values): continue
        data_cmd.append([
            Paragraph(br(r.get('職稱','')), style_cell), 
            Paragraph(br(r.get('代號', '')), style_cell), 
            Paragraph(br(r.get('姓名','')), style_cell), 
            Paragraph(br(r.get('任務','')), style_cell_l)
        ])
    
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.15, page_width*0.25, page_width*0.45])
    t1.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), 
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), 
        ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))
    ]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    data_ptl = [[Paragraph("<b>警　力　佈　署</b>", style_th), '', '', '', ''], 
                [Paragraph(f"<b>交通快打指揮官：</b>{commander}", style_cell_l), '', '', '', ''], 
                [Paragraph(f"<b>{h}</b>", style_th) for h in ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]]]
    for _, r in df_patrol.iterrows():
        if all(str(v).strip() == "" for v in r.values): continue
        data_ptl.append([
            Paragraph(br(r.get('勤務時段','')), style_cell), 
            Paragraph(br(r.get('代號', '')), style_cell), 
            Paragraph(br(r.get('編組','')), style_cell), 
            Paragraph(br(r.get('服勤人員','')), style_cell), 
            Paragraph(br(r.get('任務分工','')), style_cell_l)
        ])

    t2 = Table(data_ptl, colWidths=[page_width*0.22, page_width*0.13, page_width*0.15, page_width*0.23, page_width*0.27])
    t2.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), 
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('SPAN',(0,1),(-1,1)), 
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')), ('BACKGROUND',(0,2),(-1,2),colors.HexColor('#f2f2f2'))
    ]))
    story.append(t2)
    
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph("<b>📍 巡簽地點：</b>", style_cell_l))
    for l in CHECKIN_POINTS.split('\n'):
        if l.strip(): story.append(Paragraph(l.strip(), style_note_hanging))
        
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph("<b>📝 備註：</b>", style_cell_l))
    for l in NOTES.split('\n'):
        if l.strip(): story.append(Paragraph(l.strip(), style_note_hanging))

    doc.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)
    buf.seek(0)
    return buf

# --- HTML 預覽 ---
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
# 主介面
# ============================================================
st.title("🚔 防制危險駕車專案勤務規劃表")

# 1. 狀態初始化 (對齊二合一規格：若雲端為空，自動導入預設 5/22 底稿)
if 'data_ptl' not in st.session_state:
    s, c, p = load_from_cloud()
    if s is not None and not s.empty:
        sd = dict(zip(s.iloc[:, 0].astype(str), s.iloc[:, 1].astype(str)))
        st.session_state.p_time = sd.get("plan_time", DEFAULT_TIME_VAL)
        st.session_state.cmdr = sd.get("commander", DEFAULT_CMDR_VAL)
        st.session_state.data_cmd, st.session_state.data_ptl = c, p
    else:
        st.session_state.p_time = DEFAULT_TIME_VAL
        st.session_state.cmdr = DEFAULT_CMDR_VAL
        st.session_state.data_cmd = DEFAULT_CMD.copy()
        st.session_state.data_ptl = DEFAULT_SCHEDULE.copy()

# 確保快取監聽變數存在
if 'prev_p_time' not in st.session_state: st.session_state.prev_p_time = st.session_state.p_time
if 'prev_cmdr' not in st.session_state: st.session_state.prev_cmdr = st.session_state.cmdr
if 'last_ptl_len' not in st.session_state: st.session_state.last_ptl_len = len(st.session_state.data_ptl)

# 2. 輸入區塊
col1, col2 = st.columns(2)
with col1: p_time = st.text_input("1. 勤務時間", st.session_state.p_time)
with col2: cmdr_input = st.text_input("2. 交通快打指揮官", st.session_state.cmdr)

# 3. 變更監聽器
needs_sync_rerun = False

if p_time != st.session_state.prev_p_time:
    dedicated_time, normal_time = calc_time_strings(p_time)
    for i in range(len(st.session_state.data_ptl)):
        group = str(st.session_state.data_ptl.at[i, '編組']).strip()
        if i == 0 or "專責" in group:
            st.session_state.data_ptl.at[i, '勤務時段'] = dedicated_time
        else:
            st.session_state.data_ptl.at[i, '勤務時段'] = normal_time
    st.session_state.prev_p_time = p_time
    st.session_state.p_time = p_time
    needs_sync_rerun = True

if cmdr_input != st.session_state.prev_cmdr:
    if len(st.session_state.data_ptl) > 0:
        daihao, bianzu = get_unit_details(cmdr_input)
        st.session_state.data_ptl.at[0, '代號'] = daihao
        st.session_state.data_ptl.at[0, '編組'] = bianzu
    st.session_state.prev_cmdr = cmdr_input
    st.session_state.cmdr = cmdr_input
    needs_sync_rerun = True

if needs_sync_rerun:
    st.rerun()

# 4. 編輯器顯示
dedicated_time, normal_time = calc_time_strings(p_time)
st.subheader("3. 任務編組")
res_cmd = st.data_editor(st.session_state.data_cmd, num_rows="dynamic", use_container_width=True).fillna("")
st.subheader("4. 警力佈署")
res_ptl_raw = st.data_editor(st.session_state.data_ptl, num_rows="dynamic", use_container_width=True).fillna("")

# 5. 表格行數防呆
current_len = len(res_ptl_raw)
needs_editor_rerun = False

if current_len > st.session_state.last_ptl_len:
    for i in range(st.session_state.last_ptl_len, current_len):
        res_ptl_raw.at[i, '勤務時段'], res_ptl_raw.at[i, '服勤人員'] = normal_time, ""
    st.session_state.data_ptl, st.session_state.last_ptl_len, needs_editor_rerun = res_ptl_raw, current_len, True
elif current_len < st.session_state.last_ptl_len:
    st.session_state.data_ptl, st.session_state.last_ptl_len = res_ptl_raw, current_len
else:
    st.session_state.data_ptl = res_ptl_raw

if len(st.session_state.data_ptl) > 0 and is_blank(st.session_state.data_ptl.at[0, '勤務時段']):
    st.session_state.data_ptl.at[0, '勤務時段'] = dedicated_time
    daihao, bianzu = get_unit_details(cmdr_input)
    st.session_state.data_ptl.at[0, '代號'] = daihao
    st.session_state.data_ptl.at[0, '編組'] = bianzu
    needs_editor_rerun = True

if needs_editor_rerun: 
    st.rerun()

# 6. 檔名生成與輸出區
date_match = re.search(r'(?:(\d+)年)?(\d+)月(\d+)日', p_time)
if date_match:
    y_fn = date_match.group(1) if date_match.group(1) else str(datetime.now().year - 1911)
    m_fn = str(date_match.group(2)).zfill(2)
    d_fn = str(date_match.group(3)).zfill(2)
    date_fn = f"{y_fn}{m_fn}{d_fn}"
else:
    date_fn = datetime.now().strftime('%m%d')
final_filename = f"防制危險駕車勤務規劃表_{date_fn}"

st.markdown("---")
with st.expander("📄 預覽勤務規劃表"):
    html_preview = get_preview_html(res_cmd, st.session_state.data_ptl, cmdr_input, p_time)
    st.components.v1.html(html_preview, height=600, scrolling=True)

col_pdf, col_sync = st.columns(2)

# --- 兩階段載入提示按鈕區 ---
if col_sync.button("💾 同步雲端並寄信", use_container_width=True):
    with st.spinner("同步中，請稍候…"):
        sync_ok = save_to_cloud(p_time, cmdr_input, res_cmd, st.session_state.data_ptl)
    if sync_ok:
        with st.spinner("同步成功，正在寄送郵件…"):
            mail_ok, mail_err = send_report_email(p_time, cmdr_input, res_cmd, st.session_state.data_ptl, final_filename)
        if mail_ok: 
            st.success(f"✅ 資料已同步至 Google Sheets，郵件發送成功！")
        else: 
            st.warning(f"⚠️ 同步成功，但郵件發送失敗：{mail_err}")
    else: 
        st.error("❌ 雲端同步失敗，請檢查網路或密鑰設定。")

# 下載按鈕單獨抽出
pdf_buf = generate_pdf(p_time, cmdr_input, res_cmd, st.session_state.data_ptl)
col_pdf.download_button(
    label="📝 下載規劃表 PDF", 
    data=pdf_buf.getvalue(), 
    file_name=f"{final_filename}.pdf",
    mime="application/pdf",
    use_container_width=True
)
