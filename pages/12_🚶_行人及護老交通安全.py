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
st.set_page_config(page_title="行人及護老交通安全", layout="wide", page_icon="🚶")

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
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、秘書室巡官 陳鵬翔（代理人：警員張庭溱）、人事室警員 陳明祥、行政組警務佐 曾威仁", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"}
])

DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "聖亭派出所", "路段": "中豐路、聖亭路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "龍潭派出所", "路段": "中豐路、中正路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "中興派出所", "路段": "中興路、福龍路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "石門派出所", "路段": "中正、文化路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "高平派出所", "路段": "中豐、中原路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "三和派出所", "路段": "龍新路、楊銅路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "警備隊", "路段": "校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "龍潭交通分隊", "路段": "校園周邊道路或轄區行人易肇事路口"}
])

NOTES = """壹、警察局規劃3月份「行人及護老交通安全專案勤務」期程：
一、3月6日（星期五）6至10時、16至20時。
二、3月12日（星期四）6至10時、16至20時。
三、3月24日（星期二）6至10時、16至20時。
四、3月30日（星期一）6至10時、16至20時。
貳、執行本專案勤務視轄區狀況及執勤警力，擇定轄區易肇事路口（段）及校園周邊道路，依上揭日期妥適編排勤務（必要時得另行規劃專案）協助維護行人、學童及高齡者通行安全，並加強取締「車不讓人」、「未依規定停讓」、「違規（臨時）停車」、「行人違反路權」及「道路障礙」等違規，必要時得合併相關勤務實施，以達「一種勤務多種功能」之效益。
叁、執行「行人及護老交通安全實施計畫」合強化違規取締項目：
一、車不讓人（第44條第1項第2款、第2項、第3項、第45條第1項第6款）
二、違規（臨時）停車（第55條、第56條）
三、行人（含代步器、電動輪椅）違反路權（第78條、第80條）
四、道路障礙（第82條）"""

# --- 2. gspread 連線 ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

# --- 3. 讀取 ---
@st.cache_data(ttl=60)
def load_data():
    try:
        client = get_client()
        if client is None:
            return None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        ws_list = sh.worksheets()
        ws_set = next((w for w in ws_list if w.title == "護老_設定"), None)
        ws_cmd = next((w for w in ws_list if w.title == "護老_指揮組"), None)
        ws_sch = next((w for w in ws_list if w.title == "護老_勤務表"), None)
        if not all([ws_set, ws_cmd, ws_sch]):
            return None, None, None, "缺工作表 (需有: 護老_設定, 護老_指揮組, 護老_勤務表)"
        df_settings = pd.DataFrame(ws_set.get_all_records())
        df_cmd      = pd.DataFrame(ws_cmd.get_all_records())
        df_schedule = pd.DataFrame(ws_sch.get_all_records())
        return df_settings, df_cmd, df_schedule, None
    except Exception as e:
        return None, None, None, str(e)

# --- 4. 寫入 ---
def save_data(month, df_cmd, df_schedule):
    try:
        client = get_client()
        if client is None:
            return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("護老_設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["month", month]])
        
        ws_cmd = sh.worksheet("護老_指揮組")
        ws_cmd.clear()
        df_cmd = df_cmd.fillna("")
        ws_cmd.update([df_cmd.columns.tolist()] + df_cmd.values.tolist())
        
        ws_sch = sh.worksheet("護老_勤務表")
        ws_sch.clear()
        df_schedule = df_schedule.fillna("")
        ws_sch.update([df_schedule.columns.tolist()] + df_schedule.values.tolist())
        load_data.clear()
        return True
    except:
        return False

# --- 5. 字型與 PDF 產生 ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames():
        return fname
    for p in ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf(month, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=12*mm, bottomMargin=12*mm)
    W = A4[0] - 24*mm
    story = []

    # 字體大小設定 (標題、表頭16，內文14，備註12)
    s_title  = ParagraphStyle("t",  fontName=font, fontSize=16, alignment=1, spaceAfter=8, leading=22)
    s_th     = ParagraphStyle("th", fontName=font, fontSize=16, alignment=1, leading=22)
    s_cell   = ParagraphStyle("c",  fontName=font, fontSize=14, leading=18, alignment=1)
    s_left   = ParagraphStyle("l",  fontName=font, fontSize=14, leading=18, alignment=0)
    # 將備註字體設定為 12，行高適度縮小為 16
    s_note   = ParagraphStyle("n",  fontName=font, fontSize=12, leading=16, spaceAfter=4)
    
    def c(txt, style=s_cell):
        return Paragraph(str(txt).replace("\n","<br/>"), style)

    story.append(Paragraph(f"<b>{UNIT}{month}執行「行人及護老交通安全」專案勤務規劃表</b>", s_title))

    # ==================== 任務編組 ====================
    cw1 = [W*0.15, W*0.12, W*0.28, W*0.45]
    data1 = [[Paragraph("<b>任　務　編　組</b>", s_th), '', '', '']]
    data1.append([Paragraph(f"<b>{h}</b>", s_th) for h in ["職稱", "代號", "姓名", "任務"]])
    
    for _, row in df_cmd.iterrows():
        data1.append([
            c(f"<b>{row.get('職稱','')}</b>"),
            c(row.get('代號','')),
            c(row.get('姓名','').replace("、","<br/>").replace(",","<br/>")), 
            c(row.get('任務',''), s_left)
        ])
        
    t1 = Table(data1, colWidths=cw1, repeatRows=2)
    t1.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)),
        ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')),
        ('TOPPADDING',(0,0),(-1,-1),6), ('BOTTOMPADDING',(0,0),(-1,-1),6),
    ]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    # ==================== 警力佈署 ====================
    col_date = '日期（6時至10時、16時至20時）'
    cw2 = [W*0.28, W*0.16, W*0.56]
    data2 = [[Paragraph("<b>警　力　佈　署</b>", s_th), '', '']]
    data2.append([Paragraph(f"<b>{h}</b>", s_th) for h in ["執行勤務日期<br/>（6時至10時、16時至20時）", "單位", "路段"]])
    
    for _, row in df_schedule.iterrows():
        data2.append([
            c(row.get(col_date, '')),
            c(row.get('單位','')),
            c(row.get('路段', ''), s_left)
        ])

    table_styles = [
        ('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)),
        ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')),
        ('TOPPADDING',(0,0),(-1,-1),6), ('BOTTOMPADDING',(0,0),(-1,-1),6),
    ]
    
    # 自動合併日期欄
    non_empty = [i for i, v in enumerate(df_schedule[col_date]) if str(v).strip() != ""]
    non_empty.append(len(df_schedule))
    OFFSET = 2
    for k in range(len(non_empty) - 1):
        s, e = non_empty[k], non_empty[k+1] - 1
        if e > s:
            table_styles.append(('SPAN',   (0, s+OFFSET), (0, e+OFFSET)))
            table_styles.append(('VALIGN', (0, s+OFFSET), (0, e+OFFSET), 'MIDDLE'))

    t2 = Table(data2, colWidths=cw2, repeatRows=2)
    t2.setStyle(TableStyle(table_styles))
    story.append(KeepTogether([t2]))
    story.append(Spacer(1, 6*mm))

    # 備註 (字體12pt)
    story.append(Paragraph(f"<b>備註：</b><br/>{NOTES.replace(chr(10),'<br/>')}", s_note))

    doc.build(story)
    return buf.getvalue()

# --- 6. 寄信功能 ---
def send_report_email(subject, month, df_cmd, df_schedule):
    try:
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        pdf_bytes = generate_pdf(month, df_cmd, df_schedule)
        
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender
        msg["Subject"] = subject
        msg.attach(MIMEText("附件為最新的護老交通安全勤務規劃表 PDF。", "plain", "utf-8"))
        
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        filename = _ul.quote(f"{subject}.pdf")
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{filename}")
        msg.attach(part)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# --- 7. 主介面邏輯 ---
df_set, df_cmd, df_sch, err = load_data()
if err or df_set is None:
    c_month = DEFAULT_MONTH
    ed_cmd, ed_sch = DEFAULT_CMD.copy(), DEFAULT_SCHEDULE.copy()
else:
    sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    c_month = sd.get("month", DEFAULT_MONTH)
    ed_cmd, ed_sch = df_cmd, df_sch

st.title("🚶 行人及護老交通安全專案勤務規劃表")

st.subheader("1. 基礎資訊")
c_month = st.text_input("月份", value=c_month)

st.subheader("2. 任務編組")
with st.expander("編輯名單", expanded=True):
    ed_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True)

st.subheader("3. 執行勤務日期、單位及路段")
ed_sch = st.data_editor(ed_sch, num_rows="dynamic", use_container_width=True)

st.subheader("4. 備註（固定）")
st.text(NOTES)

# --- 8. HTML 安全產生器 (表頭16pt、內文14pt、備註12pt) ---
def get_html():
    parts = []
    # CSS 中將 .note 的 font-size 設定為 12pt
    parts.append("<style>body{font-family:'標楷體';padding:20px;} th{border:1px solid black;padding:8px;font-size:16pt;text-align:center;line-height:1.5;background-color:#f2f2f2;} td{border:1px solid black;padding:8px;font-size:14pt;text-align:center;line-height:1.5;} .note{font-size:12pt;margin:15px 0;line-height:1.6;}</style>")
    parts.append(f"<html><body><h2 style='text-align:center;font-size:16pt;'><b>{UNIT}{c_month}執行「行人及護老交通安全」專案勤務規劃表</b></h2><br>")
    
    # 任務編組
    parts.append("<table><tr><th colspan='4'>任 務 編 組</th></tr>")
    parts.append("<tr><th width='15%'>職稱</th><th width='12%'>代號</th><th width='28%'>姓名</th><th width='45%'>任務</th></tr>")
    for _, r in ed_cmd.iterrows():
        name = str(r.get('姓名', '')).replace('、', '<br>')
        parts.append(f"<tr><td><b>{r.get('職稱','')}</b></td><td>{r.get('代號','')}</td>")
        parts.append(f"<td>{name}</td><td style='text-align:left'>{r.get('任務','')}</td></tr>")
    parts.append("</table>")
    
    # 警力佈署
    parts.append("<table><tr><th colspan='3'>警 力 佈 署</th></tr>")
    parts.append("<tr><th width='28%'>執行勤務日期（6時至10時、16時至20時）</th><th width='16%'>單位</th><th width='56%'>路段</th></tr>")
    
    col_date = '日期（6時至10時、16時至20時）'
    total_rows = len(ed_sch)
    row_spans = {}
    skip_rows = set()
    
    i = 0
    while i < total_rows:
        d_val = str(ed_sch.iloc[i][col_date]).strip()
        if d_val != "":
            span = 1
            for j in range(i + 1, total_rows):
                if str(ed_sch.iloc[j][col_date]).strip() == "": span += 1
                else: break
            row_spans[i] = span
            for k in range(1, span): skip_rows.add(i + k)
            i += span
        else: i += 1

    for idx, row in ed_sch.iterrows():
        parts.append("<tr>")
        date_str = str(row.get(col_date,'')).replace('\n', '<br>')
        if idx in row_spans:
            parts.append(f"<td rowspan='{row_spans[idx]}'>{date_str}</td>")
        elif idx in skip_rows: pass
        else:
            parts.append(f"<td>{date_str}</td>")
            
        parts.append(f"<td>{str(row.get('單位','')).replace(chr(10), '<br>')}</td><td style='text-align:left'>{str(row.get('路段','')).replace(chr(10), '<br>')}</td></tr>")
        
    parts.append("</table>")
    # 備註將套用上面的 .note 樣式 (12pt)
    parts.append(f"<div class='note'><b>備註：</b><br>{NOTES.replace(chr(10), '<br>')}</div>")
    parts.append("</body></html>")
    return "".join(parts)

st.markdown("---")
st.subheader("📄 預覽與輸出")
st.components.v1.html(get_html(), height=700, scrolling=True)

if st.button("同步雲端、寄信並下載 PDF 💾", type="primary"):
    save_data(c_month, ed_cmd, ed_sch)
    subject = f"護老交通安全勤務規劃表_{datetime.now().strftime('%Y%m%d')}"
    ok, mail_err = send_report_email(subject, c_month, ed_cmd, ed_sch)
    if ok: 
        st.success("📧 雲端同步成功，報表已寄至信箱！")
    else: 
        st.error(f"❌ 雲端已同步，但寄信失敗：{mail_err}")
    
    pdf_out = generate_pdf(c_month, ed_cmd, ed_sch)
    st.download_button("點此下載 PDF", data=pdf_out, file_name=f"護老勤務表_{datetime.now().strftime('%Y%m%d')}.pdf")
