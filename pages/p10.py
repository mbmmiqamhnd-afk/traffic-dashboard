import streamlit as st
# 1. 頁面設定必須在最前面
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

# --- 工具函式 ---
def normalize(s):
    return str(s).replace('\n', '').replace('\r', '').replace(' ', '').strip()

def is_blank(val):
    return normalize(val) in ["", "None", "nan"]

def format_staff_only(val):
    if pd.isna(val) or str(val).strip() in ["None", "nan", ""]:
        return ""
    s = str(val).replace('\\', '\n').replace('、', '\n').replace('\xa0', ' ')
    s = re.sub(r'(\d{2}[:：]?\d{0,2}\s*-\s*\d{2}[:：]?\d{0,2}[時]?[:：])\s*([^\n\s])', r'\1\n\2', s)
    return '\n'.join([l.strip() for l in s.split('\n') if l.strip()])

@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)
    except:
        return None

def load_from_cloud():
    client = get_client()
    if not client: return None, None, None
    try:
        sh = client.open_by_key(SHEET_ID)
        s = pd.DataFrame(sh.worksheet("危駕_設定").get_all_records())
        c = pd.DataFrame(sh.worksheet("危駕_指揮組").get_all_records())
        p = pd.DataFrame(sh.worksheet("危駕_警力佈署").get_all_records())
        return s, c, p
    except: return None, None, None

def save_to_cloud(p_time, cmdr, df_c, df_p):
    client = get_client()
    if not client: return False
    try:
        sh = client.open_by_key(SHEET_ID)
        sh.worksheet("危駕_設定").clear()
        sh.worksheet("危駕_設定").update(range_name='A1', values=[["Key", "Value"], ["plan_time", p_time], ["commander", cmdr]])
        for name, df in [("危駕_指揮組", df_c), ("危駕_警力佈署", df_p)]:
            ws = sh.worksheet(name)
            ws.clear()
            data = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
            ws.update(range_name='A1', values=data)
        return True
    except: return False

# --- 寄信功能 ---
def send_report_email(time_str, commander, df_cmd, df_patrol, custom_filename):
    try:
        if "email" not in st.secrets: return False, "未在 secrets 中設定 email 資訊"
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        
        pdf_buf = generate_pdf(df_cmd, df_patrol, commander, time_str)
        pdf_bytes = pdf_buf.getvalue()
        
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender  # 寄給自己
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

# --- 日期解析與時段字串計算 ---
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

# --- PDF 字型與產生 ---
def register_font():
    font_paths = ["kaiu.ttf", os.path.join(os.path.dirname(__file__), "kaiu.ttf"), "C:/Windows/Fonts/kaiu.ttf"]
    for fp in font_paths:
        if os.path.exists(fp):
            pdfmetrics.registerFont(TTFont("BiauKai", fp))
            return True
    return False

FONT_AVAILABLE = register_font()
FONT_NAME = "BiauKai" if FONT_AVAILABLE else "Helvetica"

def generate_pdf(df_c, df_p, cmdr_n, time_s):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=12*mm, bottomMargin=12*mm)
    styles = {"title": ParagraphStyle("title", fontName=FONT_NAME, fontSize=14, leading=20, alignment=1),
              "cell": ParagraphStyle("cell", fontName=FONT_NAME, fontSize=9, leading=13)}
    W = A4[0] - 30*mm
    def make_table(data, col_widths):
        t = Table(data, colWidths=col_widths)
        t.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),FONT_NAME),("FONTSIZE",(0,0),(-1,-1),9),("GRID",(0,0),(-1,-1),0.5,colors.black),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(0,0),(-1,-1),"CENTER")]))
        return t
    story = [Paragraph(f"{UNIT_TITLE} 防制危險駕車專案勤務規劃表", styles["title"]), Spacer(1, 4*mm)]
    story.append(make_table([[Paragraph(f"勤務時間：{time_s}", styles["cell"]), Paragraph(f"指揮官：{cmdr_n}", styles["cell"])]], [W*0.6, W*0.4]))
    story.append(Spacer(1, 3*mm))
    story.append(make_table([["職稱","代號","姓名","任務"]]+[[Paragraph(str(r[c]), styles["cell"]) for c in ["職稱","代號","姓名","任務"]] for _,r in df_c.iterrows()], [W*0.15,W*0.15,W*0.2,W*0.5]))
    story.append(Spacer(1, 3*mm))
    story.append(make_table([["勤務時段","代號","編組","服勤人員","任務分工"]]+[[Paragraph(str(r[c]).replace('\n','<br/>'), styles["cell"]) for c in ["勤務時段","代號","編組","服勤人員","任務分工"]] for _,r in df_p.iterrows()], [W*0.15,W*0.1,W*0.15,W*0.3,W*0.3]))
    doc.build(story)
    buf.seek(0)
    return buf

# ============================================================
# 主介面
# ============================================================
st.title("🚔 防制危險駕車專案勤務規劃表")

if 'data_ptl' not in st.session_state:
    s, c, p = load_from_cloud()
    if s is not None:
        sd = dict(zip(s.iloc[:, 0].astype(str), s.iloc[:, 1].astype(str)))
        st.session_state.p_time = sd.get("plan_time", "115年4月30日22時至翌日6時")
        st.session_state.cmdr = sd.get("commander", "石門所副所長林榮裕")
        st.session_state.data_cmd, st.session_state.data_ptl = c, p
    else:
        st.session_state.p_time = "115年4月30日22時至翌日6時"
        st.session_state.cmdr = "石門所副所長林榮裕"
        st.session_state.data_cmd = pd.DataFrame(columns=["職稱", "代號", "姓名", "任務"])
        st.session_state.data_ptl = pd.DataFrame([["", "", "", "", ""]], columns=["勤務時段", "代號", "編組", "服勤人員", "任務分工"])
if 'last_ptl_len' not in st.session_state:
    st.session_state.last_ptl_len = len(st.session_state.data_ptl)

col1, col2 = st.columns(2)
with col1: p_time = st.text_input("1. 勤務時間", st.session_state.p_time)
with col2: cmdr_input = st.text_input("2. 交通快打指揮官", st.session_state.cmdr)

dedicated_time, normal_time = calc_time_strings(p_time)
st.subheader("3. 任務編組")
res_cmd = st.data_editor(st.session_state.data_cmd, num_rows="dynamic", use_container_width=True).fillna("")
st.subheader("4. 警力佈署")
res_ptl_raw = st.data_editor(st.session_state.data_ptl, num_rows="dynamic", use_container_width=True).fillna("")

# 核心邏輯：自動校正新增列日期
current_len = len(res_ptl_raw)
needs_rerun = False
if current_len > st.session_state.last_ptl_len:
    for i in range(st.session_state.last_ptl_len, current_len):
        res_ptl_raw.at[i, '勤務時段'], res_ptl_raw.at[i, '服勤人員'] = normal_time, ""
    st.session_state.data_ptl, st.session_state.last_ptl_len, needs_rerun = res_ptl_raw, current_len, True
elif current_len < st.session_state.last_ptl_len:
    st.session_state.data_ptl, st.session_state.last_ptl_len = res_ptl_raw, current_len
else:
    st.session_state.data_ptl = res_ptl_raw

if len(st.session_state.data_ptl) > 0 and is_blank(st.session_state.data_ptl.at[0, '勤務時段']):
    st.session_state.data_ptl.at[0, '勤務時段'] = dedicated_time
    unit_base = "隆安8" if "石門" in cmdr_input else "隆安6" if "龍潭" in cmdr_input else "隆安"
    st.session_state.data_ptl.at[0, '代號'] = unit_base + ("1" if "所長" in cmdr_input and "副" not in cmdr_input else "2")
    st.session_state.data_ptl.at[0, '編組'], needs_rerun = f"專責警力\n（{cmdr_input[:3]}輪值）", True

if needs_rerun: st.rerun()

# --- 按鈕區 ---
st.markdown("---")
col_pdf, col_sync = st.columns(2)
with col_pdf:
    if st.button("📥 下載 PDF"):
        pdf_buf = generate_pdf(res_cmd, st.session_state.data_ptl, cmdr_input, p_time)
        st.download_button("點此下載", pdf_buf, f"危駕規劃_{datetime.now().strftime('%m%d')}.pdf", "application/pdf")

with col_sync:
    if st.button("💾 同步雲端並寄信", type="primary"):
        with st.spinner("處理中..."):
            # 1. 同步
            sync_ok = save_to_cloud(p_time, cmdr_input, res_cmd, st.session_state.data_ptl)
            # 2. 寄信
            date_fn = "".join(re.findall(r'\d+', p_time))[:8]
            final_filename = f"防制危險駕車勤務規劃表_{date_fn}"
            mail_ok, mail_err = send_report_email(p_time, cmdr_input, res_cmd, st.session_state.data_ptl, final_filename)
            
            if sync_ok and mail_ok: st.success(f"✅ 同步與郵件寄送成功！")
            elif sync_ok: st.warning(f"⚠️ 雲端已同步，但郵件失敗：{mail_err}")
            else: st.error("❌ 雲端同步失敗")
