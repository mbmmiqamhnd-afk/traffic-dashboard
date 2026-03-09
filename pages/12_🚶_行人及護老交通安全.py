import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib
import os
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- ReportLab 相關引用 (PDF 生成核心) ---
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- 1. 基礎設定 ---
st.set_page_config(page_title="行人及護老交通安全", layout="wide")
st.title("🚶 行人及護老交通安全專案勤務規劃表")

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
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "聖亭派出所", "路段": "中豐路、聖亭路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "龍潭派出所", "路段": "中豐路、中正路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "中興派出所", "路段": "中興路、福龍路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "石門派出所", "路段": "中正、文化路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "高平派出所", "路段": "中豐、中原路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "三和派出所", "路段": "龍新路、楊銅路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "警備隊", "路段": "校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "龍潭交通分隊", "路段": "校園周邊道路或轄區行人易肇事路口"},
])

NOTES = """壹、警察局規劃3月份「行人及護老交通安全專案勤務」期程：
一、3月6日（星期五）6至10時、16至20時。
二、3月12日（星期四）6至10時、16至20時。
三、3月24日（星期二）6至10時、16至20時。
四、3月30日（星期一）6至10時、16至20時。
貳、執行本專案勤務視轄區狀況及執勤警力，擇定轄區易肇事路口（段）及校園周邊道路...（略）"""

# --- 2. 工具函式：PDF 與 Email ---
def _get_font():
    fname = "kaiu"
    paths = [os.path.join(os.getcwd(), 'kaiu.ttf'), 'kaiu.ttf', '/mount/src/traffic-dashboard/kaiu.ttf']
    for p in paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf(month, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    W = A4[0] - 30*mm
    story = []
    s_title = ParagraphStyle("t", fontName=font, fontSize=14, alignment=1, leading=20)
    s_cell = ParagraphStyle("c", fontName=font, fontSize=10, alignment=1, leading=14)
    s_left = ParagraphStyle("l", fontName=font, fontSize=10, alignment=0, leading=14)
    
    story.append(Paragraph(f"{UNIT} {month} 執行「行人及護老交通安全」專案勤務規劃表", s_title))
    story.append(Spacer(1, 5*mm))

    # 任務編組表
    data1 = [[Paragraph("<b>任 務 編 組</b>", s_title), '', '', ''], [Paragraph("職稱", s_cell), Paragraph("代號", s_cell), Paragraph("姓名", s_cell), Paragraph("任務", s_cell)]]
    for _, r in df_cmd.iterrows():
        data1.append([Paragraph(f"<b>{r['職稱']}</b>", s_cell), Paragraph(r['代號'], s_cell), Paragraph(r['姓名'].replace('、', '<br/>'), s_cell), Paragraph(r['任務'], s_left)])
    
    t1 = Table(data1, colWidths=[W*0.15, W*0.1, W*0.25, W*0.5], repeatRows=2)
    t1.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.5,colors.black), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.lightgrey), ('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t1)
    
    doc.build(story)
    return buf.getvalue()

def send_report_email(month, df_cmd, df_schedule):
    try:
        sender, pw = st.secrets["email"]["user"], st.secrets["email"]["password"]
        pdf_data = generate_pdf(month, df_cmd, df_schedule)
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender
        msg["Subject"] = f"{UNIT} {month} 勤務規劃表"
        msg.attach(MIMEText("附件為標楷體格式報表 PDF。", "plain", "utf-8"))
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_data); encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename=report.pdf")
        msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(sender, pw); s.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# --- 3. Google Sheets 核心 ---
def get_client():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    return gspread.authorize(creds)

def load_data():
    try:
        client = get_client(); sh = client.open_by_key(SHEET_ID)
        df_set = pd.DataFrame(sh.worksheet("護老_設定").get_all_records())
        df_cmd = pd.DataFrame(sh.worksheet("護老_指揮組").get_all_records())
        df_sch = pd.DataFrame(sh.worksheet("護老_勤務表").get_all_records())
        return df_set, df_cmd, df_sch, None
    except Exception as e: return None, None, None, str(e)

def save_data(month, df_cmd, df_schedule):
    try:
        client = get_client(); sh = client.open_by_key(SHEET_ID)
        sh.worksheet("護老_設定").update([["Key", "Value"], ["month", month]])
        for ws_name, df in [("護老_指揮組", df_cmd), ("護老_勤務表", df_schedule)]:
            ws = sh.worksheet(ws_name); ws.clear()
            ws.update([df.columns.tolist()] + df.fillna("").values.tolist())
        st.toast("✅ 雲端存檔成功！")
    except Exception as e: st.error(f"❌ 存檔失敗：{e}")

# --- 4. 主程式 UI ---
df_set, df_cmd_raw, df_sch_raw, err = load_data()
if err:
    st.warning("使用預設範本"); cur_month = DEFAULT_MONTH; df_c = DEFAULT_CMD; df_s = DEFAULT_SCHEDULE
else:
    cur_month = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1])).get("month", DEFAULT_MONTH)
    df_c, df_s = df_cmd_raw, df_sch_raw

cur_month = st.text_input("月份標題", cur_month)
ed_cmd = st.data_editor(df_c, num_rows="dynamic", use_container_width=True)
ed_sch = st.data_editor(df_s, num_rows="dynamic", use_container_width=True)

# --- 5. HTML 標楷體預覽 ---
def generate_html(month, df_c, df_s):
    # 強制指定標楷體字型家族
    font_family = "'標楷體', 'DFKai-SB', 'BiauKai', 'KaiTi', serif"
    style = f"""
    <style>
        body {{ font-family: {font_family}; color: #000; }}
        .container {{ max-width: 800px; margin: auto; border: 1px solid #ccc; padding: 20px; }}
        h2 {{ text-align: center; font-weight: bold; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ border: 1px solid black; padding: 8px; text-align: center; }}
        th {{ background-color: #f2f2f2; }}
        .left {{ text-align: left; }}
    </style>
    """
    html = f"<html><head>{style}</head><body><div class='container'><h2>{UNIT} {month} 勤務規劃表</h2>"
    html += "<table><tr><th colspan='4'>任 務 編 組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>"
    for _, r in df_c.iterrows():
        html += f"<tr><td><b>{r['職稱']}</b></td><td>{r['代號']}</td><td>{str(r['姓名']).replace('、','<br>')}</td><td class='left'>{r['任務']}</td></tr>"
    html += "</table><br>"
    html += "<table><tr><th colspan='3'>警 力 佈 署</th></tr><tr><th>日期</th><th>單位</th><th>路段</th></tr>"
    for _, r in df_s.iterrows():
        html += f"<tr><td>{r.iloc[0]}</td><td>{r['單位']}</td><td class='left'>{str(r['路段']).replace('\\n','<br>')}</td></tr>"
    html += f"</table><p><b>備註：</b><br>{NOTES.replace('\\n','<br>')}</p></div></body></html>"
    return html

st.subheader("📄 標楷體即時預覽")
st.components.v1.html(generate_html(cur_month, ed_cmd, ed_sch), height=600, scrolling=True)

if st.button("同步存檔並寄送 PDF 📧"):
    save_data(cur_month, ed_cmd, ed_sch)
    ok, mail_err = send_report_email(cur_month, ed_cmd, ed_sch)
    if ok: st.success("報表已寄出！")
    else: st.error(f"寄信失敗: {mail_err}")
