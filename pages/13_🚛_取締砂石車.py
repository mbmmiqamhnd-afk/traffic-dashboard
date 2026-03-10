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

# 💡 強制設定備註內容 (分列一、二)
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

@st.cache_data(ttl=5)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("砂石_設定")
        ws_cmd = sh.worksheet("砂石_指揮組")
        ws_sch = sh.worksheet("砂石_勤務表")
        return pd.DataFrame(ws_set.get_all_records()), pd.DataFrame(ws_cmd.get_all_records()), pd.DataFrame(ws_sch.get_all_records()), None
    except Exception as e:
        return None, None, None, str(e)

def save_data(month, briefing, df_cmd, df_schedule):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("砂石_設定")
        ws_set.clear()
        # 存入雲端時強制清空或更新
        ws_set.update([["Key", "Value"], ["month", month], ["briefing", briefing]])
        
        for ws_name, df in [("砂石_指揮組", df_cmd), ("砂石_勤務表", df_schedule)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df = df.fillna("")
            ws.update([df.columns.tolist()] + df.values.tolist())
        st.cache_data.clear()
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

    s_title  = ParagraphStyle("t",  fontName=font, fontSize=16, alignment=1, spaceAfter=8, leading=22)
    s_th     = ParagraphStyle("th", fontName=font, fontSize=16, alignment=1, leading=22)
    s_cell   = ParagraphStyle("c",  fontName=font, fontSize=14, leading=18, alignment=1)
    s_left   = ParagraphStyle("l",  fontName=font, fontSize=14, leading=18, alignment=0)
    s_section= ParagraphStyle("sec",fontName=font, fontSize=14, leading=20, spaceAfter=4)
    s_note   = ParagraphStyle("n",  fontName=font, fontSize=12, leading=16, spaceAfter=4)
    
    def c(txt, style=s_cell):
        return Paragraph(str(txt).replace("\n","<br/>"), style)

    story.append(Paragraph(f"<b>{UNIT}執行{month}「取締砂石（大型貨）車重點違規」專案勤務規劃表</b>", s_title))

    # 任務編組
    cw1 = [W*0.15, W*0.12, W*0.28, W*0.45]
    data1 = [[Paragraph("<b>任　務　編　組</b>", s_th), '', '', '']]
    data1.append([Paragraph(f"<b>{h}</b>", s_th) for h in ["職稱", "代號", "姓名", "任務"]])
    for _, row in df_cmd.iterrows():
        data1.append([c(f"<b>{row.get('職稱','')}</b>"), c(row.get('代號','')), c(row.get('姓名','').replace("、","<br/>")), c(row.get('任務',''), s_left)])
    t1 = Table(data1, colWidths=cw1, repeatRows=2)
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')), ('TOPPADDING',(0,0),(-1,-1),6), ('BOTTOMPADDING',(0,0),(-1,-1),6)]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    story.append(Paragraph(f"<b>📢 勤前教育：</b><br/>{briefing.replace(chr(10), '<br/>')}", s_section))
    story.append(Spacer(1, 4*mm))

    # 警力佈署 (日期合併處理)
    col_date = '勤務日期'
    cw2 = [W*0.28, W*0.16, W*0.12, W*0.44]
    data2 = [[Paragraph("<b>警　力　佈　署</b>", s_th), '', '', '']]
    data2.append([Paragraph(f"<b>{h}</b>", s_th) for h in ["勤務日期", "執行單位", "執行人數", "執行路段"]])
    for _, row in df_schedule.iterrows():
        data2.append([c(row.get(col_date, '')), c(row.get('執行單位','')), c(row.get('執行人數','')), c(row.get('執行路段', ''), s_left)])

    table_styles = [('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')), ('TOPPADDING',(0,0),(-1,-1),6), ('BOTTOMPADDING',(0,0),(-1,-1),6)]
    
    # 日期欄位自動合併
    non_empty = [idx for idx, val in enumerate(df_schedule[col_date]) if str(val).strip() != ""]
    non_empty.append(len(df_schedule))
    for k in range(len(non_empty) - 1):
        start, end = non_empty[k], non_empty[k+1] - 1
        if end > start:
            table_styles.append(('SPAN', (0, start + 2), (0, end + 2)))
            table_styles.append(('VALIGN', (0, start + 2), (0, end + 2), 'MIDDLE'))

    t2 = Table(data2, colWidths=cw2, repeatRows=2)
    t2.setStyle(TableStyle(table_styles))
    story.append(KeepTogether([t2]))
    story.append(Spacer(1, 6*mm))

    # 備註 (強制使用新格式)
    story.append(Paragraph(f"<b>備註：</b><br/>{CORRECT_NOTES.replace(chr(10),'<br/>')}", s_note))
    doc.build(story)
    return buf.getvalue()

# --- 4. 寄信功能 ---
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

# --- 5. 主介面與資料處理 ---
df_set, df_cmd_raw, df_sch_raw, err = load_data()

# 💡 初始化資料：不論雲端內容，強制清空「時間地點」並確保備註正確
cur_month = "115年3月份"
cur_brief = "" # 這裡強制設為空

if not err and df_set is not None:
    sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    cur_month = sd.get("month", cur_month)
    # 如果雲端有舊資料且包含「時間：」，則強制改回空字串
    cloud_brief = str(sd.get("briefing", ""))
    if "時間：" in cloud_brief or "地點：" in cloud_brief:
        cur_brief = ""
    else:
        cur_brief = cloud_brief

    df_c, df_s = df_cmd_raw, df_sch_raw
    if "日期" in df_s.columns: df_s.rename(columns={"日期": "勤務日期"}, inplace=True)
    if "執行人數" in df_s.columns: df_s["執行人數"] = df_s["執行人數"].replace("", "2至4人")
else:
    # 離線或讀取失敗時的預設值
    df_c = pd.DataFrame([{"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"}]) # 僅範例
    df_s = pd.DataFrame([{"勤務日期": "115年3月13日", "執行單位": "聖亭派出所", "執行人數": "2至4人", "執行路段": "中豐路"}]) # 僅範例

st.subheader("1. 基礎資訊")
c1, c2 = st.columns(2)
cur_month = c1.text_input("月份", value=cur_month)
# 顯示在編輯器裡的「勤前教育」
brief_info = c2.text_area("📢 勤前教育", value=cur_brief, height=80, help="原本的時間地點文字已被強制移除，若需要請自行輸入新內容")

st.subheader("2. 任務編組")
ed_cmd = st.data_editor(df_c, num_rows="dynamic", use_container_width=True)

st.subheader("3. 警力佈署")
ed_sch = st.data_editor(df_s, num_rows="dynamic", use_container_width=True)

# --- 6. HTML 預覽 ---
def get_html():
    parts = []
    parts.append("<style>body{font-family:'標楷體';padding:20px;} th{border:1px solid black;padding:8px;font-size:16pt;text-align:center;background-color:#f2f2f2;} td{border:1px solid black;padding:8px;font-size:14pt;text-align:center;} .note{font-size:12pt;margin:15px 0;line-height:1.6;text-align:left;} .section{font-size:14pt;margin:15px 0;text-align:left;}</style>")
    parts.append(f"<html><body><h2 style='text-align:center;'><b>{UNIT}<br>執行{cur_month}「取締砂石（大型貨）車重點違規」專案勤務規劃表</b></h2><br>")
    
    # 任務編組
    parts.append("<table><tr><th colspan='4'>任 務 編 組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>")
    for _, r in ed_cmd.iterrows():
        parts.append(f"<tr><td><b>{r.get('職稱','')}</b></td><td>{r.get('代號','')}</td><td>{str(r.get('姓名','')).replace('、','<br>')}</td><td style='text-align:left'>{r.get('任務','')}</td></tr>")
    parts.append("</table>")
    
    parts.append(f"<div class='section'><b>📢 勤前教育：</b><br>{brief_info.replace(chr(10), '<br>')}</div>")
    
    # 警力佈署
    parts.append("<table><tr><th colspan='4'>警 力 佈 署</th></tr><tr><th>勤務日期</th><th>執行單位</th><th>執行人數</th><th>執行路段</th></tr>")
    for _, row in ed_sch.iterrows():
        parts.append(f"<tr><td>{str(row.get('勤務日期','')).replace(chr(10),'<br>')}</td><td>{row.get('執行單位','')}</td><td>{row.get('執行人數','')}</td><td style='text-align:left'>{row.get('執行路段','')}</td></tr>")
    parts.append("</table>")
    
    # 💡 強制 HTML 預覽顯示正確備註
    parts.append(f"<div class='note'><b>備註：</b><br>{CORRECT_NOTES.replace(chr(10), '<br>')}</div>")
    parts.append("</body></html>")
    return "".join(parts)

st.markdown("---")
st.subheader("📄 預覽與輸出")
st.components.v1.html(get_html(), height=700, scrolling=True)

if st.button("同步雲端、寄信並下載 PDF 💾", type="primary"):
    # 儲存時，雲端就會被刷成乾淨的版本
    save_data(cur_month, brief_info, ed_cmd, ed_sch)
    subject = f"取締砂石車專案勤務規劃表_{datetime.now().strftime('%Y%m%d')}"
    ok, mail_err = send_report_email(subject, cur_month, brief_info, ed_cmd, ed_sch)
    if ok: st.success("📧 雲端同步成功，報表已寄至信箱！")
    else: st.error(f"❌ 寄信失敗：{mail_err}")
    pdf_out = generate_pdf(cur_month, brief_info, ed_cmd, ed_sch)
    st.download_button("點此下載 PDF", data=pdf_out, file_name=f"{subject}.pdf")
