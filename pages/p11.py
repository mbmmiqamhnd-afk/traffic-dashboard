import streamlit as st
import pafrom menu import show_sidebar
show_sidebar()ndas as pd
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
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# --- 1. 頁面設定 ---
st.set_page_config(page_title="防制危險駕車月份版", layout="wide", page_icon="🗓️")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# --- 預設範本資料 ---
DEFAULT_MONTH = "115年5月份"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "建興", "姓名": "駐區督察 孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、秘書室巡官 陳鵬翔（代理人：警員張庭溱）、人事室警員 陳明祥、行政組警務佐 曾威仁", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"}
])

DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（22時至翌日6時）": "115年4月30日～\n5月2日", "單位": "石門派出所", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "高平派出所", "分工": "於中豐路及龍源路巡邏（每1小時巡邏人員至下列轄區巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "龍潭交通分隊", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "聖亭派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "", "單位": "龍潭派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "", "單位": "中興派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "115年5月8日～\n5月9日", "單位": "石門派出所", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "高平派出所", "分工": "於中豐路及龍源路巡邏（每1小時巡邏人員至下列轄區巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "龍潭交通分隊", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "聖亭派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "", "單位": "龍潭派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "", "單位": "中興派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "115年5月15日～\n5月16日", "單位": "石門派出所", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "高平派出所", "分工": "於中豐路及龍源路巡邏（每1小時巡邏人員至下列轄區巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "龍潭交通分隊", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "聖亭派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "", "單位": "龍潭派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "", "單位": "中興派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "115年5月22日～\n5月23日", "單位": "石門派出所", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "高平派出所", "分工": "於中豐路及龍源路巡邏（每1小時巡邏人員至下列轄區巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "龍潭交通分隊", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "聖亭派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "", "單位": "龍潭派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "", "單位": "中興派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "115年5月29日～\n5月30日", "單位": "石門派出所", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "高平派出所", "分工": "於中豐路及龍源路巡邏（每1小時巡邏人員至下列轄區巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "龍潭交通分隊", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "", "單位": "聖亭派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "", "單位": "龍潭派出所", "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "", "單位": "中興派出所", "分工": "於轄內易發生危險駕車路段巡邏"}
])

CHECKIN_POINTS = """1. 中油高原交流道站（龍源路2-20號）
2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）
3. 7-11龍潭佳園門市（中正路三坑段776號）
4. 旭日路三坑自然生態公園停車場
5. 旭日路與大溪區交界處"""

NOTES = """一、各編組執行前由帶班人員在駐地實施勤前教育。
二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。
三、駕駛巡邏車應開啟警示燈，如發現有危險駕車行為「勿追車」，並立即向勤指中心報告，請求以優勢警力執行攔截圍捕。
四、針對下列違法、違規事項加強攔查：
（一）道路交通管理處罰條例第16條（改裝排氣管）、第18條（改裝車體設備）、第21條（無照駕駛）及43條各款項（蛇行、嚴重超速、逼車、任意減速、拆除消音器、以其他方式造成噪音、兩車以上競速等）。
（二）違反刑法185條妨害公眾往來安全罪。
（三）違反社會秩序維護法第72條妨害安寧者，同法第64條聚眾不解散。"""

# --- 2. 建立連線與讀取 ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_data(ttl=60)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("危駕月_設定")
        ws_cmd = sh.worksheet("危駕月_指揮組")
        ws_sch = sh.worksheet("危駕月_勤務表")
        return pd.DataFrame(ws_set.get_all_records()), pd.DataFrame(ws_cmd.get_all_records()), pd.DataFrame(ws_sch.get_all_records()), None
    except Exception as e: return None, None, None, str(e)

def save_data(month, df_cmd, df_schedule):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("危駕月_設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["month", month]])
        for ws_name, df in [("危駕月_指揮組", df_cmd), ("危駕月_勤務表", df_schedule)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df = df.fillna("")
            ws.update([df.columns.tolist()] + df.values.tolist())
        load_data.clear()
        return True
    except: return False

# --- 3. PDF 生成 (含頁碼功能) ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    for p in ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf_from_data(full_title, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []
    
    # 頁碼回標函數
    def add_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 10)
        page_num = f"第 {canvas.getPageNumber()} 頁"
        # 繪製在底部中央 (距離底部 10mm)
        canvas.drawCentredString(A4[0]/2.0, 10*mm, page_num)
        canvas.restoreState()

    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, leading=22, alignment=1, spaceAfter=10)
    style_th = ParagraphStyle('THeader', fontName=font, fontSize=16, alignment=1, leading=22)
    style_col_header = ParagraphStyle('ColHeader', fontName=font, fontSize=16, leading=20, alignment=1)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=18, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=18, alignment=0)
    style_hanging = ParagraphStyle('Hanging', fontName=font, fontSize=14, leading=20, leftIndent=8.5*mm, firstLineIndent=-8.5*mm, spaceAfter=5)
    style_section = ParagraphStyle('Section', fontName=font, fontSize=14, leading=20, spaceAfter=4)

    story.append(Paragraph(f"<b>{full_title}</b>", style_title))
    def clean(txt): return str(txt).replace("\n", "<br/>").replace("、", "<br/>")

    # 任務編組
    data_cmd = [[Paragraph("<b>任　務　編　組</b>", style_th), '', '', ''],
                [Paragraph(f"<b>{h}</b>", style_col_header) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([Paragraph(f"<b>{r.get('職稱','')}</b>", style_cell), Paragraph(str(r.get('代號','')), style_cell), Paragraph(clean(r.get('姓名','')), style_cell), Paragraph(str(r.get('任務','')), style_cell_left)])
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.12, page_width*0.28, page_width*0.45], repeatRows=2)
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')), ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6)]))
    story.append(t1); story.append(Spacer(1, 6*mm))

    # 警力佈署
    data_sch = [[Paragraph("<b>警　力　佈　署</b>", style_th), '', ''],
                [Paragraph(f"<b>{h}</b>", style_col_header) for h in ["日期（22時至翌日6時）", "單位", "分工"]]]
    for _, r in df_schedule.iterrows():
        data_sch.append([Paragraph(clean(r.get('日期（22時至翌日6時）','')), style_cell), Paragraph(clean(r.get('單位','')), style_cell), Paragraph(str(r.get('分工','')), style_cell_left)])
    t2 = Table(data_sch, colWidths=[page_width*0.22, page_width*0.22, page_width*0.56], repeatRows=2)
    t2_styles = [('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')), ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6)]
    date_col = '日期（22時至翌日6時）'
    non_empty = [i for i, val in enumerate(df_schedule[date_col]) if str(val).strip() != ""]
    non_empty.append(len(df_schedule))
    for k in range(len(non_empty) - 1):
        s, e = non_empty[k], non_empty[k+1] - 1
        if e > s: t2_styles.append(('SPAN', (0, s + 2), (0, e + 2))); t2_styles.append(('VALIGN', (0, s + 2), (0, e + 2), 'MIDDLE'))
    t2.setStyle(TableStyle(t2_styles))
    story.append(t2); story.append(Spacer(1, 6*mm))

    # 巡簽與備註
    story.append(Paragraph("<b>📍 巡簽地點：</b>", style_section))
    for line in CHECKIN_POINTS.split('\n'):
        if line.strip(): story.append(Paragraph(line, style_hanging))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph("<b>📝 備註：</b>", style_section))
    for line in NOTES.split('\n'):
        if line.strip(): story.append(Paragraph(line, style_hanging))

    # 執行 Build 並套用頁碼
    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    return buf.getvalue()

# --- 4. 寄信功能 ---
def send_report_email(full_title, df_cmd, df_schedule):
    try:
        sender = st.secrets["email"]["user"]; pwd = st.secrets["email"]["password"]
        pdf_bytes = generate_pdf_from_data(full_title, df_cmd, df_schedule)
        msg = MIMEMultipart(); msg["From"] = sender; msg["To"] = sender; msg["Subject"] = full_title
        msg.attach(MIMEText(f"附件為最新的「{full_title}」報表 PDF。", "plain", "utf-8"))
        part = MIMEBase("application", "pdf"); part.set_payload(pdf_bytes); encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(full_title)}.pdf")
        msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd); server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# --- 5. 主介面邏輯 ---
df_set, df_cmd_raw, df_sch_raw, err = load_data()
if not (err or df_set is None):
    sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    c_month = sd.get("month", DEFAULT_MONTH)
    ed_cmd, ed_sch = df_cmd_raw, df_sch_raw
else:
    c_month = DEFAULT_MONTH
    ed_cmd, ed_sch = DEFAULT_CMD.copy(), DEFAULT_SCHEDULE.copy()

st.title("🚔 防制危險駕車專案勤務規劃表（月份版）")
c_month = st.text_input("1. 月份資訊", c_month)
full_title = f"{UNIT}{c_month}執行「防制危險駕車」專案勤務規劃表"

st.subheader("2. 任務編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True)
st.subheader("3. 警力佈署")
res_sch = st.data_editor(ed_sch, num_rows="dynamic", use_container_width=True)

# HTML 預覽 (含 rowspan 自動合併邏輯)
def get_html():
    parts = ["<style>body{font-family:'標楷體';padding:20px;} table{border-collapse:collapse;width:100%;} th,td{border:1px solid black;padding:8px;font-size:14pt;text-align:center;line-height:1.5;} th{font-size:16pt;background-color:#f2f2f2;} .note-section{font-size:14pt;font-weight:bold;margin-top:15px;} .hanging-note{font-size:14pt;padding-left:2.2em;text-indent:-2.2em;margin-bottom:5px;line-height:1.6;}</style>"]
    parts.append(f"<html><body><h2 style='text-align:center;font-size:16pt;'><b>{full_title}</b></h2><br>")
    parts.append("<table><tr><th colspan='4'>任 務 編 組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>")
    for _, r in res_cmd.iterrows():
        parts.append(f"<tr><td><b>{r.get('職稱','')}</b></td><td>{r.get('代號','')}</td><td>{str(r.get('姓名','')).replace('、','<br>')}</td><td style='text-align:left'>{r.get('任務','')}</td></tr>")
    parts.append("</table><br><table><tr><th colspan='3'>警 力 佈 署</th></tr><tr><th>日期</th><th>單位</th><th>分工</th></tr>")
    col_date = '日期（22時至翌日6時）'
    dates = res_sch[col_date].tolist()
    non_empty_indices = [i for i, val in enumerate(dates) if str(val).strip() != ""]
    non_empty_indices.append(len(dates))
    rowspans = {}
    for k in range(len(non_empty_indices) - 1):
        s, e = non_empty_indices[k], non_empty_indices[k+1]
        rowspans[s] = e - s
    for i, row in res_sch.reset_index(drop=True).iterrows():
        parts.append("<tr>")
        if i in rowspans:
            rs = rowspans[i]
            date_val = str(row.get(col_date, '')).replace('\n', '<br>')
            parts.append(f"<td rowspan='{rs}'>{date_val}</td>")
        parts.append(f"<td>{str(row.get('單位','')).replace(chr(10),'<br>')}</td><td style='text-align:left'>{str(row.get('分工','')).replace(chr(10),'<br>')}</td></tr>")
    parts.append("</table><div class='note-section'>📍 巡簽地點：</div>")
    for line in CHECKIN_POINTS.split('\n'):
        if line.strip(): parts.append(f"<div class='hanging-note'>{line}</div>")
    parts.append("<div class='note-section'>📝 備註：</div>")
    for line in NOTES.split('\n'):
        if line.strip(): parts.append(f"<div class='hanging-note'>{line}</div>")
    parts.append("</body></html>")
    return "".join(parts)

st.markdown("---")
st.components.v1.html(get_html(), height=600, scrolling=True)

if st.button("同步雲端、寄信並下載 PDF 💾", type="primary"):
    save_data(c_month, res_cmd, res_sch)
    ok, mail_err = send_report_email(full_title, res_cmd, res_sch)
    if ok: st.success(f"📧 雲端同步成功，報表「{full_title}」已寄至信箱！")
    else: st.error(f"❌ 同步成功但寄信失敗：{mail_err}")
    pdf_out = generate_pdf_from_data(full_title, res_cmd, res_sch)
    st.download_button(label="下載 PDF", data=pdf_out, file_name=f"{full_title}.pdf", mime="application/pdf")
