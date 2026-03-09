import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib, io, os
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

# --- 1. 頁面設定 ---
st.set_page_config(page_title="行人及護老交通安全", layout="wide", page_icon="🚶")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# --- 預設範本資料 ---
DEFAULT_MONTH = "115年3月份"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官",       "代號": "隆安1",    "姓名": "分局長 施宇峰",                                           "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官",     "代號": "隆安2",    "姓名": "副分局長 何憶雯",                                         "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官",     "代號": "隆安3",    "姓名": "副分局長 蔡志明",                                         "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官",   "代號": "駐區督察", "姓名": "孫三陽",                                                      "任務": "重點機動督導。"},
    {"職稱": "督導組",       "代號": "隆安6",    "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組",       "代號": "隆安684",  "姓名": "督察組教官 郭文義",                                         "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13",   "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、秘書室巡官 陳鵬翔（代理人：警員張庭溱）、人事室警員 陳明祥、行政組警務佐 曾威仁", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組",       "代號": "隆安",     "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興",            "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "聖亭派出所",   "路段": "中豐路、聖亭路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "龍潭派出所",   "路段": "中豐路、中正路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "中興派出所",   "路段": "中興路、福龍路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "石門派出所",   "路段": "中正、文化路段\n校園周邊道路 or 轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "高平派出所",   "路段": "中豐、中原路段\n校園周邊道路 or 轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "三和派出所",   "路段": "龍新路、楊銅路段\n校園周邊道路 or 轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "警備隊",       "路段": "校園周邊道路 or 轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "龍潭交通分隊", "路段": "校園周邊道路 or 轄區行人易肇事路口"},
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

# --- 2. 建立 gspread 連線 ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

# --- 3. 讀取與寫入函數 ---
@st.cache_data(ttl=60)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, "未設定 Secrets"
        sh = client.open_by_key(SHEET_ID)
        ws_list = sh.worksheets()
        ws_set = next((w for w in ws_list if w.title == "護老_設定"), None)
        ws_cmd = next((w for w in ws_list if w.title == "護老_指揮組"), None)
        ws_sch = next((w for w in ws_list if w.title == "護老_勤務表"), None)
        if not all([ws_set, ws_cmd, ws_sch]): return None, None, None, "缺工作表"
        return pd.DataFrame(ws_set.get_all_records()), pd.DataFrame(ws_cmd.get_all_records()), pd.DataFrame(ws_sch.get_all_records()), None
    except Exception as e: return None, None, None, str(e)

def save_data(month, df_cmd, df_schedule):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        sh.worksheet("護老_設定").clear(); sh.worksheet("護老_設定").update([["Key", "Value"], ["month", month]])
        sh.worksheet("護老_指揮組").clear(); sh.worksheet("護老_指揮組").update([df_cmd.columns.tolist()] + df_cmd.fillna("").values.tolist())
        sh.worksheet("護老_勤務表").clear(); sh.worksheet("護老_勤務表").update([df_schedule.columns.tolist()] + df_schedule.fillna("").values.tolist())
        load_data.clear(); return True
    except Exception as e: return False

# --- 4. PDF 生成 (強化表格間距與合併標題) ---
def _get_font():
    fname = "kaiu"
    font_paths = ["kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/kaiu.ttf"]
    for p in font_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf_from_data(month, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 30*mm
    story = []
    
    # 樣式設定
    style_t = ParagraphStyle('T', fontName=font, fontSize=16, alignment=1, spaceAfter=10)
    style_c = ParagraphStyle('C', fontName=font, fontSize=10, alignment=1, leading=14)
    style_l = ParagraphStyle('L', fontName=font, fontSize=10, alignment=0, leading=14)
    style_h = ParagraphStyle('H', fontName=font, fontSize=13, alignment=1, leading=18)

    # 1. 主標題
    story.append(Paragraph(f"<b>{UNIT}{month}執行「行人及護老交通安全」專案勤務規劃表</b>", style_t))

    # 2. 任務編組
    data1 = [[Paragraph("<b>任　務　編　組</b>", style_h), "", "", ""],
             [Paragraph("<b>職稱</b>", style_c), Paragraph("<b>代號</b>", style_c), Paragraph("<b>姓名</b>", style_c), Paragraph("<b>任務</b>", style_c)]]
    for _, r in df_cmd.iterrows():
        data1.append([Paragraph(f"<b>{r['職稱']}</b>", style_c), Paragraph(str(r['代號']), style_c), Paragraph(str(r['姓名']).replace("、","<br/>"), style_c), Paragraph(str(r['任務']), style_l)])
    
    t1 = Table(data1, colWidths=[page_width*0.15, page_width*0.1, page_width*0.25, page_width*0.5])
    t1.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.5,colors.black),('SPAN',(0,0),(3,0)),('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('BACKGROUND',(0,0),(-1,1),colors.whitesmoke)]))
    story.append(t1)

    # --- 關鍵修正：增加間距到 10mm (明確空一行) ---
    story.append(Spacer(1, 10*mm))

    # 3. 警力佈署 (標題同步)
    data2 = [[Paragraph("<b>警　力　佈　署</b>", style_h), "", ""],
             [Paragraph("<b>日期</b>", style_c), Paragraph("<b>單位</b>", style_c), Paragraph("<b>路段</b>", style_c)]]
    for _, r in df_schedule.iterrows():
        data2.append([Paragraph(str(r.iloc[0]), style_c), Paragraph(str(r.iloc[1]), style_c), Paragraph(str(r.iloc[2]).replace("\n","<br/>"), style_l)])
    
    t2 = Table(data2, colWidths=[page_width*0.25, page_width*0.2, page_width*0.55])
    # 自動合併日期欄位邏輯
    styles2 = [('GRID',(0,0),(-1,-1),0.5,colors.black),('SPAN',(0,0),(2,0)),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('BACKGROUND',(0,0),(-1,1),colors.whitesmoke)]
    non_empty = [i for i, v in enumerate(df_schedule.iloc[:,0]) if str(v).strip() != ""] + [len(df_schedule)]
    for k in range(len(non_empty)-1):
        if non_empty[k+1] - non_empty[k] > 1:
            styles2.append(('SPAN', (0, non_empty[k]+2), (0, non_empty[k+1]+1)))
    t2.setStyle(TableStyle(styles2))
    story.append(t2)

    # 4. 備註
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph("<b>備註：</b>", style_l))
    story.append(Paragraph(NOTES.replace("\n", "<br/>"), style_l))

    doc.build(story); return buf.getvalue()

# --- 5. UI 與 寄信功能 ---
def send_email(month, df_cmd, df_sch):
    try:
        pdf = generate_pdf_from_data(month, df_cmd, df_sch)
        msg = MIMEMultipart()
        msg["From"] = st.secrets["email"]["user"]; msg["To"] = msg["From"]; msg["Subject"] = f"{UNIT}{month}勤務規劃表"
        msg.attach(MIMEText("請查收附件勤務報表。", "plain"))
        att = MIMEBase("application", "pdf"); att.set_payload(pdf); encoders.encode_base64(att)
        att.add_header("Content-Disposition", f"attachment; filename=Report_{datetime.now().strftime('%m%d')}.pdf")
        msg.attach(att)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(msg["From"], st.secrets["email"]["password"]); s.sendmail(msg["From"], msg["To"], msg.as_string())
        return True
    except: return False

# --- 6. 主程式執行 ---
df_set, df_cmd_raw, df_sch_raw, err = load_data()
if err: 
    month, df_cmd, df_sch = DEFAULT_MONTH, DEFAULT_CMD, DEFAULT_SCHEDULE
else:
    month = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1])).get("month", DEFAULT_MONTH)
    df_cmd, df_sch = df_cmd_raw, df_sch_raw

st.title("🚶 行人及護老交通安全專案")
month = st.text_input("月份", month)
st.subheader("任務編組")
e_cmd = st.data_editor(df_cmd, num_rows="dynamic", use_container_width=True)
st.subheader("警力佈署")
e_sch = st.data_editor(df_sch, num_rows="dynamic", use_container_width=True)

if st.button("💾 儲存並寄送 PDF 報表", type="primary"):
    if save_data(month, e_cmd, e_sch):
        if send_email(month, e_cmd, e_sch): st.success("✅ 雲端存檔成功，PDF 已寄出！")
        else: st.warning("⚠️ 存檔成功，但 PDF 寄送失敗。")
    else: st.error("❌ 存檔失敗")
