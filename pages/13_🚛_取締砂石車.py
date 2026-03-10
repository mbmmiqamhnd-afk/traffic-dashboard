import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib
import io
import os
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# --- 1. 頁面設定 ---
st.set_page_config(page_title="取締砂石車專案勤務", layout="wide", page_icon="🚛")
st.title("🚛 取締砂石（大型貨）車重點違規專案勤務規劃表")

SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# 💡 固定備註內容 (分列一、二)
CORRECT_NOTES = (
    "一、執行前由各單位帶班人員在駐地實施勤前教育。\n"
    "二、加強取締砂石（大型貨）車超載、車速、酒醉駕車、闖紅燈、無照駕車、爭道行駛、"
    "違反禁行路線、變更車斗、未使用專用車箱及未裝設行車紀錄器（行車視野輔助器）等違規，"
    "以共同消弭不法行為，保障用路人生命財產安全。"
)

# --- 2. Google Sheets 連線 ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("砂石_設定")
        ws_cmd = sh.worksheet("砂石_指揮組")
        ws_sch = sh.worksheet("砂石_勤務表")
        
        df_set = pd.DataFrame(ws_set.get_all_records())
        df_cmd = pd.DataFrame(ws_cmd.get_all_records())
        df_sch = pd.DataFrame(ws_sch.get_all_records())
        
        # 🛡️ 物理性攔截：在資料進入 UI 之前強制清理
        sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
        briefing = str(sd.get("briefing", ""))
        
        # 只要包含舊文字，強制清空為空字串
        if "時間：" in briefing or "地點：" in briefing or "各單位執行前" in briefing:
            briefing = ""
            
        return df_set, df_cmd, df_sch, briefing, None
    except Exception as e:
        return None, None, None, "", str(e)

def save_data(month, briefing, df_cmd, df_schedule):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("砂石_設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["month", month], ["briefing", briefing]])
        
        for ws_name, df in [("砂石_指揮組", df_cmd), ("砂石_勤務表", df_schedule)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df = df.fillna("")
            ws.update([df.columns.tolist()] + df.values.tolist())
        return True
    except:
        return False

# --- 3. 字型與 PDF 產生 ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    for p in ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf(month, briefing, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=12*mm, bottomMargin=12*mm)
    W = A4[0] - 24*mm
    story = []
    s_title  = ParagraphStyle("t", fontName=font, fontSize=16, alignment=1, spaceAfter=8, leading=22)
    s_th     = ParagraphStyle("th", fontName=font, fontSize=16, alignment=1, leading=22)
    s_cell   = ParagraphStyle("c", fontName=font, fontSize=14, leading=18, alignment=1)
    s_left   = ParagraphStyle("l", fontName=font, fontSize=14, leading=18, alignment=0)
    s_section= ParagraphStyle("sec",fontName=font, fontSize=14, leading=20, spaceAfter=4)
    s_note   = ParagraphStyle("n", fontName=font, fontSize=12, leading=16, spaceAfter=4)
    
    def c(txt, style=s_cell): return Paragraph(str(txt).replace("\n","<br/>"), style)
    story.append(Paragraph(f"<b>{UNIT}執行{month}「取締砂石（大型貨）車重點違規」專案勤務規劃表</b>", s_title))
    
    # 指揮編組
    cw1 = [W*0.15, W*0.12, W*0.28, W*0.45]
    data1 = [[Paragraph("<b>任　務　編　組</b>", s_th), '', '', ''], [Paragraph(f"<b>{h}</b>", s_th) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, row in df_cmd.iterrows():
        data1.append([c(f"<b>{row.get('職稱','')}</b>"), c(row.get('代號','')), c(row.get('姓名','').replace("、","<br/>")), c(row.get('任務',''), s_left)])
    t1 = Table(data1, colWidths=cw1, repeatRows=2); t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))]))
    story.append(t1); story.append(Spacer(1, 6*mm))
    story.append(Paragraph(f"<b>📢 勤前教育：</b><br/>{briefing.replace(chr(10), '<br/>')}", s_section)); story.append(Spacer(1, 4*mm))
    
    # 警力佈署
    col_date = '勤務日期'
    cw2 = [W*0.28, W*0.16, W*0.12, W*0.44]
    data2 = [[Paragraph("<b>警　力　佈　署</b>", s_th), '', '', ''], [Paragraph(f"<b>{h}</b>", s_th) for h in ["勤務日期", "執行單位", "執行人數", "執行路段"]]]
    for _, row in df_schedule.iterrows():
        data2.append([c(row.get(col_date, '')), c(row.get('執行單位','')), c(row.get('執行人數','')), c(row.get('執行路段', ''), s_left)])
    t2 = Table(data2, colWidths=cw2, repeatRows=2); t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))]))
    story.append(KeepTogether([t2])); story.append(Spacer(1, 6*mm))
    
    story.append(Paragraph(f"<b>備註：</b><br/>{CORRECT_NOTES.replace(chr(10),'<br/>')}", s_note))
    doc.build(story); return buf.getvalue()

# --- 4. 寄信 ---
def send_report_email(subject, month, briefing, df_cmd, df_schedule):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        pdf_bytes = generate_pdf(month, briefing, df_cmd, df_schedule)
        msg = MIMEMultipart(); msg["From"] = sender; msg["To"] = sender; msg["Subject"] = subject
        msg.attach(MIMEText("附件為最新的專案勤務表 PDF。", "plain", "utf-8"))
        part = MIMEBase("application", "pdf"); part.set_payload(pdf_bytes); encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(subject)}.pdf")
        msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd); server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# --- 5. 介面 ---
df_set, df_cmd_raw, df_sch_raw, filtered_brief, err = load_data()

# 初始化
cur_month = "115年3月份"
if not err and df_set is not None:
    sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    cur_month = sd.get("month", cur_month)
    df_c, df_s = df_cmd_raw, df_sch_raw
    if "日期" in df_s.columns: df_s.rename(columns={"日期": "勤務日期"}, inplace=True)
    df_s["執行人數"] = df_s["執行人數"].astype(str).replace("", "2至4人").replace("nan", "2至4人")
else:
    df_c, df_s = pd.DataFrame([{"職稱": "指揮官", "代號": "隆安1", "姓名": "施宇峰", "任務": "核定"}]), pd.DataFrame()

st.subheader("1. 基礎資訊")
c1, c2 = st.columns(2)
month_val = c1.text_input("月份", value=cur_month)
# 💡 這裡直接使用 filtered_brief，舊文字在這裡就被擋掉了
brief_info = c2.text_area("📢 勤前教育", value=filtered_brief, height=80)

st.subheader("2. 任務編組")
ed_cmd = st.data_editor(df_c, num_rows="dynamic", use_container_width=True)
st.subheader("3. 警力佈署")
ed_sch = st.data_editor(df_s, num_rows="dynamic", use_container_width=True)

# --- 6. HTML 預覽 ---
def get_html():
    parts = ["<style>body{font-family:'標楷體';padding:10px;} th{border:1px solid black;padding:5px;background-color:#f2f2f2;} td{border:1px solid black;padding:5px;} .note{font-size:12pt;margin-top:10px;} table{width:100%;border-collapse:collapse;}</style>"]
    parts.append(f"<h2 style='text-align:center;'>{UNIT}執行{month_val}「取締砂石（大型貨）車重點違規」專案勤務規劃表</h2>")
    parts.append("<table><tr><th colspan='4'>任 務 編 組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>")
    for _, r in ed_cmd.iterrows():
        parts.append(f"<tr><td><b>{r.get('職稱','')}</b></td><td>{r.get('代號','')}</td><td>{str(r.get('姓名','')).replace('、','<br>')}</td><td style='text-align:left'>{r.get('任務','')}</td></tr>")
    parts.append("</table>")
    parts.append(f"<p><b>📢 勤前教育：</b><br>{brief_info.replace(chr(10), '<br>')}</p>")
    parts.append("<table><tr><th colspan='4'>警 力 佈 署</th></tr><tr><th>勤務日期</th><th>執行單位</th><th>執行人數</th><th>執行路段</th></tr>")
    for _, row in ed_sch.iterrows():
        parts.append(f"<tr><td>{str(row.get('勤務日期','')).replace(chr(10),'<br>')}</td><td>{row.get('執行單位','')}</td><td>{row.get('執行人數','')}</td><td style='text-align:left'>{row.get('執行路段','')}</td></tr>")
    parts.append("</table>")
    parts.append(f"<div class='note'><b>備註：</b><br>{CORRECT_NOTES.replace(chr(10), '<br>')}</div>")
    return "".join(parts)

st.markdown("---")
st.components.v1.html(get_html(), height=600, scrolling=True)

if st.button("💾 同步雲端、寄信並下載 PDF", type="primary"):
    if save_data(month_val, brief_info, ed_cmd, ed_sch):
        subject = f"取締砂石車專案勤務表_{datetime.now().strftime('%m%d')}"
        ok, mail_err = send_report_email(subject, month_val, brief_info, ed_cmd, ed_sch)
        if ok: st.success("📧 雲端已更新，PDF 已寄至信箱！")
        else: st.error(f"❌ 雲端已更新，但寄信失敗：{mail_err}")
        st.download_button("📥 下載 PDF", data=generate_pdf(month_val, brief_info, ed_cmd, ed_sch), file_name=f"{subject}.pdf")
