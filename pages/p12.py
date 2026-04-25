import streamlit as st

# --- 1. 頁面設定 (必須放在全站最頂端第一個 Streamlit 指令) ---
st.set_page_config(page_title="行人及護老交通安全", layout="wide", page_icon="🚶")

# 呼叫側邊欄
from menu import show_sidebar
show_sidebar()

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

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# --- 預設範本 ---
DEFAULT_MONTH = "115年3月份"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "駐區督察", "姓名": "孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、秘書室巡官 陳鵬翔（代理人：警員張庭溱）、人事室警員 陳明祥", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"}
])

DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "聖亭派出所", "路段": "中豐路、聖亭路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "龍潭派出所", "路段": "中豐路、中正路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "中興派出所", "路段": "中興路、福龍路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "石門派出所", "路段": "中正、文化路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "高平派出所", "路段": "中豐、中原路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "三和派出所", "路段": "龍新路、楊銅路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "警備隊", "路段": "校園周邊道路 or 轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "龍潭交通分隊", "路段": "校園周邊道路 or 轄區行人易肇事路口"}
])

DEFAULT_NOTES = """壹、警察局規劃3月份「行人及護老交通安全專案勤務」期程：
一、3月6日（星期五）6至10時、16至20時。
二、3月12日（星期四）6至10時、16至20時。
三、3月24日（星期二）6至10時、16至20時。
四、3月30日（星期一）6至10時、16至20時。
貳、執行本專案勤務視轄區狀況及執勤警力，擇定轄區易肇事路口（段）及校園周邊道路，依上揭日期妥適編排勤務協助維護行人、學童及高齡者通行安全，並加強取締「車不讓人」、「未依規定停讓」、「違規停車」等。"""

# --- 2. gspread 連線 ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except:
        return None

# --- 3. 讀取 ---
@st.cache_data(ttl=60)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        ws_list = sh.worksheets()
        ws_set = next((w for w in ws_list if w.title == "護老_設定"), None)
        ws_cmd = next((w for w in ws_list if w.title == "護老_指揮組"), None)
        ws_sch = next((w for w in ws_list if w.title == "護老_勤務表"), None)
        if not all([ws_set, ws_cmd, ws_sch]): return None, None, None, None, "缺工作表"
        
        df_settings = pd.DataFrame(ws_set.get_all_records())
        df_cmd      = pd.DataFrame(ws_cmd.get_all_records())
        df_schedule = pd.DataFrame(ws_sch.get_all_records())
        
        sd = dict(zip(df_settings.iloc[:,0], df_settings.iloc[:,1])) if not df_settings.empty else {}
        notes = sd.get("notes", DEFAULT_NOTES)
        return df_settings, df_cmd, df_schedule, notes, None
    except Exception as e: return None, None, None, None, str(e)

# --- 4. 寫入 ---
def save_data(month, df_cmd, df_schedule, notes):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        
        ws_set = sh.worksheet("護老_設定")
        ws_set.clear()
        ws_set.update(range_name='A1', values=[["Key", "Value"], ["month", month], ["notes", notes]])
        
        ws_cmd = sh.worksheet("護老_指揮組")
        ws_cmd.clear()
        ws_cmd.update(range_name='A1', values=[df_cmd.columns.tolist()] + df_cmd.fillna("").astype(str).values.tolist())
        
        ws_sch = sh.worksheet("護老_勤務表")
        ws_sch.clear()
        ws_sch.update(range_name='A1', values=[df_schedule.columns.tolist()] + df_schedule.fillna("").astype(str).values.tolist())
        
        load_data.clear()
        return True
    except: return False

# --- 5. 字型與 PDF 產生 ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    font_paths = ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]
    for p in font_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf(month, df_cmd, df_schedule, notes_content):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=12*mm, bottomMargin=18*mm)
    W = A4[0] - 24*mm
    story = []

    def draw_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 10)
        page_num = f"第 {doc.page} 頁"
        canvas.drawCentredString(A4[0]/2.0, 10*mm, page_num)
        canvas.restoreState()

    s_title  = ParagraphStyle("t",  fontName=font, fontSize=16, alignment=1, spaceAfter=8, leading=22, wordWrap='CJK')
    s_th     = ParagraphStyle("th", fontName=font, fontSize=16, alignment=1, leading=22, wordWrap='CJK')
    s_cell   = ParagraphStyle("c",  fontName=font, fontSize=14, leading=18, alignment=1, wordWrap='CJK')
    s_left   = ParagraphStyle("l",  fontName=font, fontSize=14, leading=18, alignment=0, wordWrap='CJK')
    s_note   = ParagraphStyle("n",  fontName=font, fontSize=12, leading=18, leftIndent=8.5*mm, firstLineIndent=-8.5*mm, spaceAfter=4, wordWrap='CJK')
    
    def c(txt, style=s_cell):
        return Paragraph(str(txt).replace("\n","<br/>"), style)

    header_text = f"<b>{UNIT}{month}執行「行人及護老交通安全」專案勤務規劃表</b>"
    story.append(Paragraph(header_text, s_title))

    cw1 = [W*0.15, W*0.12, W*0.28, W*0.45]
    data1 = [[Paragraph("<b>任　務　編　組</b>", s_th), '', '', ''],
             [Paragraph(f"<b>{h}</b>", s_th) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, row in df_cmd.iterrows():
        data1.append([c(f"<b>{row.get('職稱','')}</b>"), c(row.get('代號','')), c(str(row.get('姓名','')).replace("、","<br/>")), c(row.get('任務',''), s_left)])
    t1 = Table(data1, colWidths=cw1, repeatRows=2)
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    col_date = '日期（6時至10時、16時至20時）'
    cw2 = [W*0.28, W*0.16, W*0.56]
    data2 = [[Paragraph("<b>警　力　佈　署</b>", s_th), '', ''],
             [Paragraph(f"<b>執行勤務日期</b>", s_th), Paragraph("<b>單位</b>", s_th), Paragraph("<b>路段</b>", s_th)]]
    for _, row in df_schedule.iterrows():
        data2.append([c(row.get(col_date, '')), c(row.get('單位','')), c(row.get('路段', ''), s_left)])

    t2_styles = [('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))]
    
    if not df_schedule.empty:
        non_empty = [i for i, v in enumerate(df_schedule[col_date]) if str(v).strip() != ""]
        non_empty.append(len(df_schedule))
        for k in range(len(non_empty) - 1):
            s, e = non_empty[k], non_empty[k+1] - 1
            if e > s: t2_styles.append(('SPAN', (0, s+2), (0, e+2)))
    
    t2 = Table(data2, colWidths=cw2, repeatRows=2)
    t2.setStyle(TableStyle(t2_styles))
    story.append(t2)
    story.append(Spacer(1, 6*mm))

    story.append(Paragraph("<b>備註：</b>", s_note))
    for line in notes_content.split('\n'):
        if line.strip(): story.append(Paragraph(line, s_note))
            
    doc.build(story, onFirstPage=draw_page_number, onLaterPages=draw_page_number)
    return buf.getvalue()

# --- 6. 寄信功能 ---
def send_report_email(subject, month, df_cmd, df_schedule, notes):
    try:
        sender = st.secrets["email"]["user"]; pwd = st.secrets["email"]["password"]
        pdf_bytes = generate_pdf(month, df_cmd, df_schedule, notes)
        msg = MIMEMultipart(); msg["From"] = sender; msg["To"] = sender; msg["Subject"] = subject
        msg.attach(MIMEText("附件為最新的護老交通安全勤務規劃表 PDF。", "plain", "utf-8"))
        part = MIMEBase("application", "pdf"); part.set_payload(pdf_bytes); encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(subject)}.pdf")
        msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd); server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# --- 7. 主介面邏輯 ---
df_set, df_cmd_raw, df_sch_raw, db_notes, err = load_data()
if err or df_set is None:
    c_month = DEFAULT_MONTH; ed_cmd, ed_sch = DEFAULT_CMD.copy(), DEFAULT_SCHEDULE.copy(); current_notes = DEFAULT_NOTES
else:
    sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    c_month = sd.get("month", DEFAULT_MONTH); ed_cmd, ed_sch = df_cmd_raw, df_sch_raw; current_notes = db_notes

st.title("🚶 行人及護老交通安全專案勤務規劃表")
c_month = st.text_input("1. 月份", value=c_month)
st.subheader("2. 任務編組")
ed_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True)
st.subheader("3. 警力佈署")
ed_sch = st.data_editor(ed_sch, num_rows="dynamic", use_container_width=True)
st.subheader("4. 備註編輯")
ed_notes = st.text_area("編輯備註內容", value=current_notes, height=250)

full_header_name = f"{UNIT}{c_month}執行「行人及護老交通安全」專案勤務規劃表"

# --- 8. HTML 預覽 ---
def get_html(notes_content):
    parts = ["<style>body{font-family:'標楷體';padding:10px;} table{width:100%;border-collapse:collapse;} th,td{border:1px solid black;padding:8px;text-align:center;} th{background:#f2f2f2;font-size:16pt;} .note{font-size:12pt;margin-top:10px;line-height:1.6;}</style>"]
    parts.append(f"<html><body><h2 style='text-align:center;'>{full_header_name}</h2>")
    parts.append("<table><tr><th colspan='4'>任 務 編 組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>")
    for _, r in ed_cmd.iterrows():
        parts.append(f"<tr><td><b>{r.get('職稱','')}</b></td><td>{r.get('代號','')}</td><td>{str(r.get('姓名','')).replace('、','<br>')}</td><td style='text-align:left'>{r.get('任務','')}</td></tr>")
    parts.append("</table><br><table><tr><th colspan='3'>警 力 佈 署</th></tr><tr><th>勤務日期</th><th>單位</th><th>路段</th></tr>")
    for _, r in ed_sch.iterrows():
        parts.append(f"<tr><td>{str(r.get('日期（6時至10時、16時至20時）','')).replace('\n','<br>')}</td><td>{r.get('單位','')}</td><td style='text-align:left'>{str(r.get('路段','')).replace('\n','<br>')}</td></tr>")
    parts.append("</table><div class='note'><b>備註：</b><br>")
    for line in notes_content.split('\n'):
        if line.strip(): parts.append(f"{line}<br>")
    parts.append("</div></body></html>")
    return "".join(parts)

st.markdown("---")
st.components.v1.html(get_html(ed_notes), height=600, scrolling=True)

if st.button("同步雲端、寄信並下載 PDF 💾", type="primary"):
    if save_data(c_month, ed_cmd, ed_sch, ed_notes):
        ok, mail_err = send_report_email(full_header_name, c_month, ed_cmd, ed_sch, ed_notes)
        if ok: st.success("📧 同步成功，報表已寄至信箱！")
        else: st.warning(f"⚠️ 雲端已同步，但寄信失敗：{mail_err}")
        pdf_out = generate_pdf(c_month, ed_cmd, ed_sch, ed_notes)
        st.download_button("點此下載 PDF", data=pdf_out, file_name=f"{full_header_name}.pdf")
    else:
        st.error("❌ 雲端同步失敗，請檢查權限設定。")
