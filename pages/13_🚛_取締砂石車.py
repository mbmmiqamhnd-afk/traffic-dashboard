import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# --- 1. 頁面設定 ---
st.set_page_config(page_title="取締砂石車專案勤務", layout="wide")
st.title("🚛 取締砂石（大型貨）車重點違規專案勤務規劃表")
st.caption("資料與 Google Sheets 即時連線，手機、電腦皆可編輯")

SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# --- 預設範本 ---
DEFAULT_MONTH  = "115年3月份"
DEFAULT_BRIEF  = "時間：各單位執行前\n地點：現地勤教"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官",       "代號": "隆安1",    "姓名": "分局長 施宇峰",                                       "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官",     "代號": "隆安2",    "姓名": "副分局長 何憶雯",                                     "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官",     "代號": "隆安3",    "姓名": "副分局長 蔡志明",                                     "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官",   "代號": "建興",     "姓名": "駐區督察 孫三陽",                                     "任務": "重點機動督導。"},
    {"職稱": "督導組",       "代號": "隆安6",    "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組",       "代號": "隆安684",  "姓名": "督察組教官 郭文義",                                   "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13",   "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、保安民防組巡官 陳鵬翔（代理人：警員張庭溱）、人事室警員 陳明祥、行政組警務佐 曾威仁", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組",       "代號": "隆安",     "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興",            "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期": "115年3月13日（星期五）00時至24時", "執行單位": "聖亭派出所",   "執行人數": "2至4人", "執行路段": "中豐路、聖亭路段等砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "龍潭派出所",   "執行人數": "",       "執行路段": "大昌路、中豐路段砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "中興派出所",   "執行人數": "",       "執行路段": "中興路、福龍路及龍平路段砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "石門派出所",   "執行人數": "",       "執行路段": "中正路、龍源路及民族路段砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "高平派出所",   "執行人數": "",       "執行路段": "中豐路、龍源路段砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "三和派出所",   "執行人數": "",       "執行路段": "楊銅路、龍新路段砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "警備隊",       "執行人數": "",       "執行路段": "中豐路、龍源路、聖亭路段砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "龍潭交通分隊", "執行人數": "",       "執行路段": "中豐路、龍源路、聖亭路段砂石（大型貨）車行經路段"},
    {"日期": "115年3月27日（星期五）00時至24時", "執行單位": "聖亭派出所",   "執行人數": "2至4人", "執行路段": "中豐路、聖亭路段等砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "龍潭派出所",   "執行人數": "",       "執行路段": "大昌路、中豐路段砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "中興派出所",   "執行人數": "",       "執行路段": "中興路、福龍路及龍平路段砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "石門派出所",   "執行人數": "",       "執行路段": "中正路、龍源路及民族路段砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "高平派出所",   "執行人數": "",       "執行路段": "中豐路、龍源路段砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "三和派出所",   "執行人數": "",       "執行路段": "楊銅路、龍新路段砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "警備隊",       "執行人數": "",       "執行路段": "中豐路、龍源路、聖亭路段砂石（大型貨）車行經路段"},
    {"日期": "",                                  "執行單位": "龍潭交通分隊", "執行人數": "",       "執行路段": "中豐路、龍源路、聖亭路段砂石（大型貨）車行經路段"},
])

NOTES = "※ 加強取締砂石（大型貨）車超載、車速、酒醉駕車、闖紅燈、無照駕車、爭道行駛、違反禁行路線、變更車斗、未使用專用車箱及未裝設行車紀錄器（行車視野輔助器）等違規，以共同消弭不法行為，保障用路人生命財產安全。"

# --- 寄信函數 ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames():
        return fname
    for p in ['/mount/src/traffic-dashboard/kaiu.ttf', 'kaiu.ttf', './kaiu.ttf']:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(fname, p))
                return fname
            except Exception:
                pass
    return "Helvetica"

def generate_pdf(month, briefing, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    W = A4[0] - 30*mm
    story = []
    s_title = ParagraphStyle("t", fontName=font, fontSize=13, alignment=1, spaceAfter=2, leading=18)
    s_cell  = ParagraphStyle("c", fontName=font, fontSize=9,  leading=13, alignment=1)
    s_left  = ParagraphStyle("l", fontName=font, fontSize=9,  leading=13, alignment=0)
    s_note  = ParagraphStyle("n", fontName=font, fontSize=9,  leading=14)

    def c(txt, style=None):
        txt = str(txt).replace("\n","<br/>").replace("、","<br/>").replace(",","<br/>")
        return Paragraph(txt, style or s_cell)

    story.append(Paragraph(f"{UNIT}執行{month}「取締砂石（大型貨）車重點違規」專案勤務規劃表", s_title))
    story.append(Spacer(1, 2*mm))

    cw1 = [W*0.15, W*0.10, W*0.25, W*0.50]
    data1 = [[Paragraph("<b>任　務　編　組</b>", s_title),'','','']]
    data1.append([c("<b>職稱</b>"),c("<b>代號</b>"),c("<b>姓名</b>"),c("<b>任務</b>")])
    for _, row in df_cmd.iterrows():
        data1.append([c(f"<b>{row.get('職稱','')}</b>"),c(row.get('代號','')),
                      c(row.get('姓名','')),c(row.get('任務',''),s_left)])
    t1 = Table(data1, colWidths=cw1, repeatRows=2)
    t1.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('SPAN',(0,0),(-1,0)),
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),
        ('BACKGROUND',(0,1),(-1,1),colors.HexColor('#f2f2f2')),
        ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
    ]))
    story.append(t1)
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(f"<b>📢 勤前教育：</b>{briefing}", s_note))
    story.append(Spacer(1, 2*mm))

    cw2 = [W*0.25, W*0.20, W*0.10, W*0.45]
    data2 = [[Paragraph("<b>警　力　佈　署</b>", s_title),'','','']]
    data2.append([c("<b>日期</b>"),c("<b>執行單位</b>"),c("<b>執行人數</b>"),c("<b>執行路段</b>")])
    for _, row in df_schedule.iterrows():
        data2.append([c(row.get('日期','')),c(row.get('執行單位','')),
                      c(row.get('執行人數','')),c(row.get('執行路段',''),s_left)])
    t2 = Table(data2, colWidths=cw2, repeatRows=2)
    t2.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('SPAN',(0,0),(-1,0)),
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),
        ('BACKGROUND',(0,1),(-1,1),colors.HexColor('#f2f2f2')),
        ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
    ]))
    story.append(t2)
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(f"<b>備註：</b>{NOTES}", s_note))
    doc.build(story)
    return buf.getvalue()

def send_report_email(html_content, subject, month, briefing, df_cmd, df_schedule):
    import urllib.parse as _ul
    try:
        sender   = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        receiver = sender
        pdf_bytes = generate_pdf(month, briefing, df_cmd, df_schedule)
        msg = MIMEMultipart()
        msg["From"]=sender; msg["To"]=receiver; msg["Subject"]=subject
        msg.attach(MIMEText("請見附件 PDF 報表。", "plain", "utf-8"))
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        encoded_name = _ul.quote(f"{subject}.pdf", safe='')
        part.add_header("Content-Disposition",
            f"attachment; filename=\"report.pdf\"; filename*=UTF-8\'\'{encoded_name}")
        msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, receiver, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)






# --- 2. gspread 連線 ---
def get_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

# --- 3. 讀取 ---
def load_data():
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        df_settings = pd.DataFrame(sh.worksheet("砂石_設定").get_all_records())
        df_cmd      = pd.DataFrame(sh.worksheet("砂石_指揮組").get_all_records())
        df_schedule = pd.DataFrame(sh.worksheet("砂石_勤務表").get_all_records())
        return df_settings, df_cmd, df_schedule, None
    except Exception as e:
        return None, None, None, str(e)

# --- 4. 寫入 ---
def save_data(month, briefing, df_cmd, df_schedule):
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)

        ws_set = sh.worksheet("砂石_設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["month", month], ["briefing", briefing]])

        ws_cmd = sh.worksheet("砂石_指揮組")
        ws_cmd.clear()
        df_cmd = df_cmd.fillna("")
        ws_cmd.update([df_cmd.columns.tolist()] + df_cmd.values.tolist())

        ws_sch = sh.worksheet("砂石_勤務表")
        ws_sch.clear()
        df_schedule = df_schedule.fillna("")
        ws_sch.update([df_schedule.columns.tolist()] + df_schedule.values.tolist())

        st.toast("✅ 雲端存檔成功！", icon="☁️")
        return True
    except Exception as e:
        st.error(f"❌ 存檔失敗：{e}")
        return False

# --- 5. 初始化 ---
df_set, df_cmd, df_sch, error_msg = load_data()

if error_msg or df_set is None or df_set.empty:
    if error_msg:
        st.error(f"❌ 無法讀取 Google Sheets：\n{error_msg}")
    st.info("💡 已載入預設範本，請修改後按「下載報表」自動儲存。")
    current_month    = DEFAULT_MONTH
    current_brief    = DEFAULT_BRIEF
    df_cmd_edit      = DEFAULT_CMD.copy()
    df_schedule_edit = DEFAULT_SCHEDULE.copy()
else:
    try:
        sd = dict(zip(df_set.iloc[:, 0], df_set.iloc[:, 1]))
        current_month    = sd.get("month",    DEFAULT_MONTH)
        current_brief    = sd.get("briefing", DEFAULT_BRIEF)
        df_cmd_edit      = df_cmd if not df_cmd.empty else DEFAULT_CMD.copy()
        df_schedule_edit = df_sch if not df_sch.empty else DEFAULT_SCHEDULE.copy()
    except Exception as e:
        st.error(f"資料格式解析失敗：{e}")
        st.stop()

# --- 6. 介面 ---
st.subheader("1. 基礎資訊")
c1, c2 = st.columns(2)
current_month = c1.text_input("月份", value=current_month)
brief_info    = c2.text_area("📢 勤前教育", value=current_brief, height=80)

st.subheader("2. 任務編組")
st.caption("💡 姓名若有多人，請用「、」分隔。")
with st.expander("編輯名單", expanded=True):
    edited_cmd = st.data_editor(
        df_cmd_edit,
        num_rows="dynamic",
        use_container_width=True,
        column_config={"任務": None}
    )
    if "任務" not in edited_cmd.columns:
        edited_cmd["任務"] = df_cmd_edit["任務"]

st.subheader("3. 執行任務單位、時間及路段")
edited_schedule = st.data_editor(df_schedule_edit, num_rows="dynamic", use_container_width=True)

st.subheader("4. 備註（固定）")
st.text(NOTES)

# --- 7. 產生 HTML ---
def generate_html(month, briefing, df_cmd, df_schedule):
    import os as _os
    _kaiu = '/mount/src/traffic-dashboard/kaiu.ttf'
    if _os.path.exists(_kaiu):
        _font_face = "@font-face { font-family: 'BiauKai'; src: url('file://" + _kaiu + "') format('truetype'); }"
        _font_css = "body { font-family: 'BiauKai', serif;"
    else:
        _font_face = ""
        _font_css = "body { font-family: serif;"
    style = """
    <style>
        """ + _font_face + """
        """ + _font_css + """ color: #000; font-size: 14px; }
        .container { width: 100%; max-width: 800px; margin: 0 auto; padding: 20px; }
        h2 { text-align: left; margin-bottom: 5px; letter-spacing: 2px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
        th, td { border: 1px solid black; padding: 5px; text-align: center; font-size: 14px; vertical-align: middle; }
        th { background-color: #f2f2f2; }
        .left-align { text-align: left; }
        .section { margin-bottom: 10px; line-height: 1.8; }
        .notes { white-space: pre-wrap; font-size: 13px; line-height: 1.8; }
        @media print { .no-print { display: none; } body { -webkit-print-color-adjust: exact; } }
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2>{UNIT}執行{month}「取締砂石（大型貨）車重點違規」專案勤務規劃表</h2>"

    # 任務編組
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr>"
    html += "<tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, row in df_cmd.iterrows():
        name = str(row.get('姓名', '')).replace("、", "<br>").replace(",", "<br>")
        html += f"<tr><td><b>{row.get('職稱','')}</b></td><td>{row.get('代號','')}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務','')}</td></tr>"
    html += "</table>"

    # 勤前教育
    html += f"<div class='section'><b>📢 勤前教育：</b><span style='white-space:pre-wrap'>{briefing}</span></div>"

    # 執行任務表
    html += "<div class='section'><b>執行任務單位、時間及路段</b></div>"
    html += "<table><tr><th width='25%'>日期</th><th width='20%'>執行單位</th><th width='10%'>執行人數</th><th width='45%'>執行路段</th></tr>"
    for _, row in df_schedule.iterrows():
        html += f"<tr><td>{row.get('日期','')}</td><td>{row.get('執行單位','')}</td><td>{row.get('執行人數','')}</td><td class='left-align'>{row.get('執行路段','')}</td></tr>"
    html += "</table>"

    # 備註
    html += f"<div class='section'><b>備註：</b><span class='notes'>{NOTES}</span></div>"

    html += "</div></body></html>"
    return html

html_out = generate_html(current_month, brief_info, edited_cmd, edited_schedule)

# --- 8. 輸出 ---
st.markdown("---")
col_view, col_dl = st.columns([3, 1])
with col_view:
    st.subheader("📄 即時預覽")
    st.components.v1.html(html_out, height=800, scrolling=True)
with col_dl:
    st.subheader("📥 輸出")
    if st.download_button(
        label="下載報表並同步雲端 💾",
        data=html_out.encode("utf-8"),
        file_name=f"砂石車勤務表_{datetime.now().strftime('%Y%m%d')}.html",
        mime="text/html; charset=utf-8",
        type="primary"
    ):
        save_data(current_month, brief_info, edited_cmd, edited_schedule)
        subject = f"取締砂石車勤務規劃表_{datetime.now().strftime('%Y%m%d')}"
        ok, err = send_report_email(html_out, subject, current_month, brief_info, edited_cmd, edited_schedule)
        if ok:
            st.toast("📧 報表已寄出至信箱！", icon="✉️")
        else:
            st.error(f"❌ 寄信失敗：{err}")
    st.info("💡 下載後打開檔案，按 Ctrl+P 列印。")
