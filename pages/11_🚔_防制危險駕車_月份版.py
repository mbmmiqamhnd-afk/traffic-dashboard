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
st.set_page_config(page_title="防制危險駕車勤務", layout="wide")
st.title("🚔 防制危險駕車專案勤務規劃表")
st.caption("資料與 Google Sheets 即時連線，手機、電腦皆可編輯")

SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

UNIT = "桃園市政府警察局龍潭分局"

# --- 預設範本 ---
DEFAULT_TIME        = "115年3月6日22時至翌日6時"
DEFAULT_BRIEF       = "時間：各編組執行前\n地點：現地勤教"
DEFAULT_COMMANDER   = "石門所副所長林榮裕"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官",   "代號": "隆安1",   "姓名": "分局長 施宇峰",       "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2",   "姓名": "副分局長 何憶雯",     "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3",   "姓名": "副分局長 蔡志明",     "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "業務組",   "代號": "隆安13",  "姓名": "交通組警務員 葉佳媛", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報群聚飆車狀況。"},
    {"職稱": "督導組",   "代號": "隆安681", "姓名": "督察組督察員 黃中彥", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "通訊組",   "代號": "隆安",    "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "監看群聚告警訊息、指揮、調度及通報本勤務事宜。"},
])

DEFAULT_PATROL = pd.DataFrame([
    {"勤務時段": "3月7日零時至4時",      "無線電": "隆安82",  "編組": "專責警力（石門所輪值）", "服勤人員": "00-02時：副所長林榮裕、警員王耀民\n02-04時：副所長林榮裕、警員陳欣妤", "任務分工": "「加強防制」勤務，在文化路、中正路三坑段、龍源路及旭日路來回巡邏，隨機攔檢改裝（噪音）車輛（每2小時至責任區域內指定巡簽地點巡簽1次並守望10分鐘，將守望情形拍照上傳LINE「龍潭分局聯絡平臺」群組）"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安80",  "編組": "石門所線上巡邏組合警力兼任",     "服勤人員": "",  "任務分工": "「區域聯防」勤務，於中正路、文化路、中豐路、龍源路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次），並加強查緝毒駕"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安90",  "編組": "高平所線上巡邏組合警力兼任",     "服勤人員": "",  "任務分工": "「區域聯防」勤務，於中豐路及龍源路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安990", "編組": "龍潭交通分隊線上巡邏組合警力兼任", "服勤人員": "",  "任務分工": "「區域聯防」勤務，於龍源路及溪州橋旁新建道路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安50",  "編組": "聖亭所線上巡邏組合警力兼任",     "服勤人員": "",  "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安60",  "編組": "龍潭所線上巡邏組合警力兼任",     "服勤人員": "",  "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安70",  "編組": "中興所線上巡邏組合警力兼任",     "服勤人員": "",  "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
])

CHECKIN_POINTS = """1. 中油高原交流道站（龍源路2-20號）
2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）
3. 7-11龍潭佳園門市（中正路三坑段776號）
4. 旭日路三坑自然生態公園停車場
5. 旭日路與大溪區交界處"""

NOTES = """一、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。
二、駕駛巡邏車應開啟警示燈，如發現有危險駕車行為「勿追車」，並立即向勤指中心報告，請求以優勢警力執行攔截圍捕。
三、針對下列違法、違規事項加強攔查：
（一）道路交通管理處罰條例第16條（改裝排管）、第18條（改裝車體設備）、第21條（無照駕駛）及43條各款項（蛇行、嚴重超速、逼車、任意減速、拆除消音器、以其他方式造成噪音、兩車以上競速等）及第35條1項2款（毒駕）。
（二）違反刑法185條公共危險罪（以他法致生往來危險者）。
（三）違反社會秩序維護法第72條妨害安寧者，同法第64條聚眾不解散。"""

# --- 字型 & PDF & 寄信函數 ---
def _get_font():
    fname = "kaiu"
    if fname not in pdfmetrics.getRegisteredFontNames():
        for p in ["kaiu.ttf", "./kaiu.ttf"]:
            if os.path.exists(p):
                pdfmetrics.registerFont(TTFont(fname, p))
                return fname
        return "Helvetica"
    return fname

def _html_table_to_pdf(html_content, page_title):
    """把 HTML 報表轉成 PDF bytes（解析 <tr><td> 重排）"""
    import re as _re
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=12*mm, rightMargin=12*mm,
        topMargin=12*mm, bottomMargin=12*mm)
    W = A4[0] - 24*mm

    title_s  = ParagraphStyle("t",  fontName=font, fontSize=13, alignment=1, spaceAfter=4)
    cell_s   = ParagraphStyle("c",  fontName=font, fontSize=9,  leading=13)
    small_s  = ParagraphStyle("sm", fontName=font, fontSize=8,  leading=11)

    def p(txt, style=None):
        txt = _re.sub(r'<br\s*/?>', '\n', str(txt))
        txt = _re.sub(r'<[^>]+>', '', txt).strip()
        return Paragraph(txt.replace('\n','<br/>'), style or cell_s)

    story = [Paragraph(page_title, title_s), Spacer(1, 3*mm)]

    # 解析所有 <table>
    tables_raw = _re.findall(r'<table[^>]*>(.*?)</table>', html_content, _re.DOTALL|_re.IGNORECASE)
    for tbl_html in tables_raw:
        rows_raw = _re.findall(r'<tr[^>]*>(.*?)</tr>', tbl_html, _re.DOTALL|_re.IGNORECASE)
        data = []
        for row_html in rows_raw:
            cells = _re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', row_html, _re.DOTALL|_re.IGNORECASE)
            if cells:
                data.append([p(c) for c in cells])
        if not data:
            continue
        col_n = max(len(r) for r in data)
        col_w = [W / col_n] * col_n
        t = Table(data, colWidths=col_w, repeatRows=1)
        t.setStyle(TableStyle([
            ('FONTNAME',    (0,0),(-1,-1), font),
            ('FONTSIZE',    (0,0),(-1,-1), 9),
            ('GRID',        (0,0),(-1,-1), 0.5, colors.black),
            ('VALIGN',      (0,0),(-1,-1), 'MIDDLE'),
            ('BACKGROUND',  (0,0),(-1, 0), colors.HexColor('#f2f2f2')),
            ('TOPPADDING',  (0,0),(-1,-1), 3),
            ('BOTTOMPADDING',(0,0),(-1,-1), 3),
        ]))
        story.append(t)
        story.append(Spacer(1, 3*mm))

    # 備註（<table> 外的文字）
    plain = _re.sub(r'<table[^>]*>.*?</table>', '', html_content, flags=_re.DOTALL|_re.IGNORECASE)
    plain = _re.sub(r'<br\s*/?>', '\n', plain)
    plain = _re.sub(r'<[^>]+>', '', plain).strip()
    if plain:
        story.append(Paragraph(plain.replace('\n','<br/>'), small_s))

    doc.build(story)
    return buf.getvalue()

def send_report_email(html_content, subject):
    try:
        sender   = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        receiver = sender
        pdf_bytes = _html_table_to_pdf(html_content, subject)
        msg = MIMEMultipart()
        msg["From"]    = sender
        msg["To"]      = receiver
        msg["Subject"] = subject
        msg.attach(MIMEText("請見附件 PDF 報表。", "plain", "utf-8"))
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{subject}.pdf"')
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
        df_settings = pd.DataFrame(sh.worksheet("危駕_設定").get_all_records())
        df_cmd      = pd.DataFrame(sh.worksheet("危駕_指揮組").get_all_records())
        df_patrol   = pd.DataFrame(sh.worksheet("危駕_警力佈署").get_all_records())
        return df_settings, df_cmd, df_patrol, None
    except Exception as e:
        return None, None, None, str(e)

# --- 4. 寫入 ---
def save_data(time_str, briefing, commander, df_cmd, df_patrol):
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)

        ws_set = sh.worksheet("危駕_設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"],
                       ["plan_time",   time_str],
                       ["briefing",    briefing],
                       ["commander",   commander]])

        ws_cmd = sh.worksheet("危駕_指揮組")
        ws_cmd.clear()
        df_cmd = df_cmd.fillna("")
        ws_cmd.update([df_cmd.columns.tolist()] + df_cmd.values.tolist())

        ws_ptl = sh.worksheet("危駕_警力佈署")
        ws_ptl.clear()
        df_patrol = df_patrol.fillna("")
        ws_ptl.update([df_patrol.columns.tolist()] + df_patrol.values.tolist())

        st.toast("✅ 雲端存檔成功！", icon="☁️")
        return True
    except Exception as e:
        st.error(f"❌ 存檔失敗：{e}")
        return False

# --- 5. 初始化 ---
df_set, df_cmd, df_ptl, error_msg = load_data()

if error_msg or df_set is None or df_set.empty:
    if error_msg:
        st.error(f"❌ 無法讀取 Google Sheets：\n{error_msg}")
    st.info("💡 已載入預設範本，請修改後按「下載報表」自動儲存。")
    current_time      = DEFAULT_TIME
    current_brief     = DEFAULT_BRIEF
    current_commander = DEFAULT_COMMANDER
    df_cmd_edit       = DEFAULT_CMD.copy()
    df_patrol_edit    = DEFAULT_PATROL.copy()
else:
    try:
        sd = dict(zip(df_set.iloc[:, 0], df_set.iloc[:, 1]))
        current_time      = sd.get("plan_time",  DEFAULT_TIME)
        current_brief     = sd.get("briefing",   DEFAULT_BRIEF)
        current_commander = sd.get("commander",  DEFAULT_COMMANDER)
        df_cmd_edit       = df_cmd    if not df_cmd.empty    else DEFAULT_CMD.copy()
        df_patrol_edit    = df_ptl    if not df_ptl.empty    else DEFAULT_PATROL.copy()
    except Exception as e:
        st.error(f"資料格式解析失敗：{e}")
        st.stop()

# --- 6. 介面 ---
st.subheader("1. 基礎資訊")
c1, c2 = st.columns(2)
plan_time  = c1.text_input("勤務時間", value=current_time)
brief_info = c2.text_area("📢 勤前教育", value=current_brief, height=80)

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

st.subheader("3. 警力佈署")
commander = st.text_input("交通快打指揮官", value=current_commander)
edited_patrol = st.data_editor(df_patrol_edit, num_rows="dynamic", use_container_width=True)

st.subheader("4. 巡簽地點（固定）")
st.text(CHECKIN_POINTS)

st.subheader("5. 備註（固定）")
st.text(NOTES)

# --- 7. 產生 HTML ---
def generate_html(time_str, briefing, commander, df_cmd, df_patrol):
    style = """
    <style>
        body { font-family: 'DFKai-SB', 'BiauKai', '標楷體', serif; color: #000; font-size: 14px; }
        .container { width: 100%; max-width: 800px; margin: 0 auto; padding: 20px; }
        h2 { text-align: left; margin-bottom: 5px; letter-spacing: 2px; }
        .info { text-align: right; font-weight: bold; margin-bottom: 15px; font-size: 14px; }
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
    html += f"<h2>{UNIT}執行「防制危險駕車專案勤務」規劃表</h2>"
    html += f"<div class='info'>勤務時間：{time_str}</div>"

    # 任務編組
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr>"
    html += "<tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, row in df_cmd.iterrows():
        name = str(row.get('姓名', '')).replace("、", "<br>").replace(",", "<br>")
        html += f"<tr><td><b>{row.get('職稱','')}</b></td><td>{row.get('代號','')}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務','')}</td></tr>"
    html += "</table>"

    # 勤前教育
    html += f"<div class='section'><b>📢 勤前教育：</b><span style='white-space:pre-wrap'>{briefing}</span></div>"

    # 警力佈署
    html += f"<div class='section'><b>警力佈署</b>　交通快打指揮官：{commander}</div>"
    html += "<table><tr><th width='15%'>勤務時段</th><th width='10%'>代號</th><th width='18%'>編組</th><th width='20%'>服勤人員</th><th width='37%'>任務分工</th></tr>"
    for _, row in df_patrol.iterrows():
        personnel = str(row.get('服勤人員', '')).replace("、", "<br>").replace("\n", "<br>")
        html += f"<tr><td>{row.get('勤務時段','')}</td><td>{row.get('無線電','')}</td><td>{row.get('編組','')}</td><td style='line-height:1.4'>{personnel}</td><td class='left-align'>{row.get('任務分工','')}</td></tr>"
    html += "</table>"

    # 巡簽地點
    html += f"<div class='section'><b>巡簽地點</b><br><span style='white-space:pre-wrap'>{CHECKIN_POINTS}</span></div>"

    # 備註
    html += f"<div class='section'><b>備註</b><br><span class='notes'>{NOTES}</span></div>"

    html += "</div></body></html>"
    return html

html_out = generate_html(plan_time, brief_info, commander, edited_cmd, edited_patrol)

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
        file_name=f"危駕勤務表_{datetime.now().strftime('%Y%m%d')}.html",
        mime="text/html; charset=utf-8",
        type="primary"
    ):
        save_data(plan_time, brief_info, commander, edited_cmd, edited_patrol)
        subject = f"防制危險駕車勤務規劃表_{datetime.now().strftime('%Y%m%d')}"
        ok, err = send_report_email(html_out, subject)
        if ok:
            st.toast("📧 報表已寄出至信箱！", icon="✉️")
        else:
            st.error(f"❌ 寄信失敗：{err}")
    st.info("💡 下載後打開檔案，按 Ctrl+P 列印。")
