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
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# --- 預設範本資料 ---
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

# 修改重點：將原本空白的服勤人員欄位，預設填入「線上巡邏警力兼任」
DEFAULT_PATROL = pd.DataFrame([
    {"勤務時段": "3月7日零時至4時",       "無線電": "隆安82",  "編組": "專責警力（石門所輪值）", "服勤人員": "00-02時：副所長林榮裕、警員王耀民\n02-04時：副所長林榮裕、警員陳欣妤", "任務分工": "「加強防制」勤務，在文化路、中正路三坑段、龍源路及旭日路來回巡邏，隨機攔檢改裝（噪音）車輛（每2小時至責任區域內指定巡簽地點巡簽1次並守望10分鐘，將守望情形拍照上傳LINE「龍潭分局聯絡平臺」群組）"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安80",  "編組": "石門所線上巡邏組合警力兼任",      "服勤人員": "線上巡邏警力兼任",  "任務分工": "「區域聯防」勤務，於中正路、文化路、中豐路、龍源路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次），並加強查緝毒駕"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安90",  "編組": "高平所線上巡邏組合警力兼任",      "服勤人員": "線上巡邏警力兼任",  "任務分工": "「區域聯防」勤務，於中豐路及龍源路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安990", "編組": "龍潭交通分隊線上巡邏組合警力兼任", "服勤人員": "線上巡邏警力兼任",  "任務分工": "「區域聯防」勤務，於龍源路及溪州橋旁新建道路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安50",  "編組": "聖亭所線上巡邏組合警力兼任",      "服勤人員": "線上巡邏警力兼任",  "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安60",  "編組": "龍潭所線上巡邏組合警力兼任",      "服勤人員": "線上巡邏警力兼任",  "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "3月6日22時至翌日6時", "無線電": "隆安70",  "編組": "中興所線上巡邏組合警力兼任",      "服勤人員": "線上巡邏警力兼任",  "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
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

# --- 2. 建立 gspread 連線 (Cache Resource) ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

# --- 3. 讀取函數 (Cache Data) ---
@st.cache_data(ttl=60)
def load_data():
    try:
        client = get_client()
        if client is None:
            return None, None, None, "未設定 Secrets (離線模式)"
        
        sh = client.open_by_key(SHEET_ID)
        ws_list = sh.worksheets()
        
        ws_set = next((w for w in ws_list if w.title == "危駕_設定"), None)
        ws_cmd = next((w for w in ws_list if w.title == "危駕_指揮組"), None)
        ws_ptl = next((w for w in ws_list if w.title == "危駕_警力佈署"), None)

        if not all([ws_set, ws_cmd, ws_ptl]):
            return None, None, None, "缺工作表 (需有: 危駕_設定, 危駕_指揮組, 危駕_警力佈署)"

        df_settings = pd.DataFrame(ws_set.get_all_records())
        df_cmd      = pd.DataFrame(ws_cmd.get_all_records())
        df_patrol   = pd.DataFrame(ws_ptl.get_all_records())
        return df_settings, df_cmd, df_patrol, None
    except Exception as e:
        return None, None, None, str(e)

# --- 4. 寫入函數 ---
def save_data(time_str, briefing, commander, df_cmd, df_patrol):
    try:
        client = get_client()
        if client is None:
            st.warning("⚠️ 離線模式無法存檔至雲端，僅能下載 PDF。")
            return False

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
        
        load_data.clear()
        st.toast("✅ 雲端存檔成功！", icon="☁️")
        return True
    except Exception as e:
        st.error(f"❌ 存檔失敗：{e}")
        return False

# --- 5. PDF 生成函數 (含表格修正) ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames():
        return fname
    font_paths = ["kaiu.ttf", "./kaiu.ttf", "font/kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]
    font_path = None
    for p in font_paths:
        if os.path.exists(p):
            font_path = p
            break   
    if font_path:
        try:
            pdfmetrics.registerFont(TTFont(fname, font_path))
            return fname
        except Exception:
            pass
    return "Helvetica"

def generate_pdf_from_data(time_str, briefing, commander, df_cmd, df_patrol):
    font = _get_font()
    buf = io.BytesIO()
    
    # 邊距設定
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=15*mm, bottomMargin=15*mm,
        title=f"{UNIT}執行「防制危險駕車專案勤務」規劃表")
        
    page_width = A4[0] - 30*mm
    story = []
    
    # --- 樣式定義 ---
    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, leading=22, spaceAfter=6)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=12)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=10, leading=13, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=10, leading=13, alignment=0)
    style_section = ParagraphStyle('Section', fontName=font, fontSize=11, leading=16, spaceAfter=4)
    style_note = ParagraphStyle('Note', fontName=font, fontSize=10, leading=14, spaceAfter=2)
    style_table_header = ParagraphStyle('TableHeader', fontName=font, fontSize=14, alignment=1, leading=18)

    # 1. 標題與時間
    story.append(Paragraph(f"{UNIT}執行「防制危險駕車專案勤務」規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def clean_text(txt):
        return str(txt).replace("\n", "<br/>").replace("、", "<br/>")

    # ====================
    # 2. 任務編組表格
    # ====================
    col_widths_cmd = [page_width * 0.15, page_width * 0.10, page_width * 0.25, page_width * 0.50]
    headers_cmd = ["職稱", "代號", "姓名", "任務"]
    
    data_cmd = []
    # Row 0: 標題列
    data_cmd.append([Paragraph("<b>任　務　編　組</b>", style_table_header), '', '', ''])
    # Row 1: 欄位名
    data_cmd.append([Paragraph(f"<b>{h}</b>", style_cell) for h in headers_cmd])
    
    for _, row in df_cmd.iterrows():
        job = Paragraph(f"<b>{row.get('職稱','')}</b>", style_cell)
        code = Paragraph(str(row.get('代號','')), style_cell)
        name = Paragraph(clean_text(row.get('姓名','')), style_cell)
        task = Paragraph(str(row.get('任務','')), style_cell_left)
        data_cmd.append([job, code, name, task])

    t1 = Table(data_cmd, colWidths=col_widths_cmd, repeatRows=2)
    t1.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        # Row 0
        ('SPAN', (0,0), (-1,0)),
        ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#f2f2f2')),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        # Row 1
        ('BACKGROUND', (0,1), (-1, 1), colors.HexColor('#f2f2f2')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    # ====================
    # 3. 勤前教育
    # ====================
    brief_clean = briefing.replace("\n", "<br/>")
    story.append(Paragraph(f"<b>📢 勤前教育：</b><br/>{brief_clean}", style_section))
    story.append(Spacer(1, 4*mm))

    # ====================
    # 4. 警力佈署表格 (指揮官整併入表格)
    # ====================
    col_widths_ptl = [page_width * 0.15, page_width * 0.10, page_width * 0.18, page_width * 0.20, page_width * 0.37]
    headers_ptl = ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]
    
    data_ptl = []
    
    # [Row 0] 警力佈署 (灰底, 跨欄)
    data_ptl.append([Paragraph("<b>警　力　佈　署</b>", style_table_header), '', '', '', ''])
    
    # [Row 1] 指揮官 (白底, 跨欄)
    cmd_text = f"<b>交通快打指揮官：</b>{commander}"
    data_ptl.append([Paragraph(cmd_text, style_cell_left), '', '', '', ''])
    
    # [Row 2] 欄位名稱 (灰底)
    data_ptl.append([Paragraph(f"<b>{h}</b>", style_cell) for h in headers_ptl])
    
    for _, row in df_patrol.iterrows():
        time_p = Paragraph(clean_text(row.get('勤務時段','')), style_cell)
        code = Paragraph(str(row.get('無線電','')), style_cell)
        group = Paragraph(clean_text(row.get('編組','')), style_cell)
        ppl = Paragraph(clean_text(row.get('服勤人員','')), style_cell)
        task = Paragraph(str(row.get('任務分工','')), style_cell_left)
        data_ptl.append([time_p, code, group, ppl, task])

    # repeatRows=3: 標題、指揮官、欄位名 重複
    t2 = Table(data_ptl, colWidths=col_widths_ptl, repeatRows=3)
    
    t2.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        
        # Row 0: 警力佈署 (灰底)
        ('SPAN', (0,0), (-1,0)),
        ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#f2f2f2')),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        
        # Row 1: 指揮官 (白底)
        ('SPAN', (0,1), (-1,1)),
        ('BACKGROUND', (0,1), (-1, 1), colors.white),
        ('ALIGN', (0,1), (-1,1), 'LEFT'),
        ('LEFTPADDING', (0,1), (-1,1), 6),
        
        # Row 2: 欄位名 (灰底)
        ('BACKGROUND', (0,2), (-1, 2), colors.HexColor('#f2f2f2')),
        
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t2)
    story.append(Spacer(1, 6*mm))

    # ====================
    # 5. 巡簽地點 & 備註
    # ====================
    story.append(Paragraph("<b>巡簽地點：</b>", style_section))
    check_pts = CHECKIN_POINTS.replace("\n", "<br/>")
    story.append(Paragraph(check_pts, style_note))
    story.append(Spacer(1, 4*mm))
    
    story.append(Paragraph("<b>備註：</b>", style_section))
    notes_clean = NOTES.replace("\n", "<br/>")
    story.append(Paragraph(notes_clean, style_note))

    try:
        doc.build(story)
        return buf.getvalue()
    except Exception as e:
        print(f"PDF Build Error: {e}")
        return None

def send_report_email(html_content, subject, time_str, briefing, commander, df_cmd, df_patrol):
    import urllib.parse as _ul
    try:
        sender   = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        receiver = sender
        
        pdf_bytes = generate_pdf_from_data(time_str, briefing, commander, df_cmd, df_patrol)
        if pdf_bytes is None:
            return False, "PDF 生成失敗 (請檢查 kaiu.ttf 字型)"

        msg = MIMEMultipart()
        msg["From"]    = sender
        msg["To"]      = receiver
        msg["Subject"] = subject
        msg.attach(MIMEText("請見附件 PDF 報表。\n\n本郵件由雲端勤務系統自動發送。", "plain", "utf-8"))
        
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        encoded_name = _ul.quote(f"{subject}.pdf", safe='')
        part.add_header(
            "Content-Disposition",
            f"attachment; filename=\"report.pdf\"; filename*=UTF-8''{encoded_name}"
        )
        msg.attach(part)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, receiver, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# --- 6. 主程式邏輯 ---

df_set, df_cmd, df_ptl, error_msg = load_data()

if error_msg:
    st.error(f"❌ Google Sheets 讀取失敗：{error_msg}")
    st.warning("⚠️ 啟用離線範本模式。")
    current_time      = DEFAULT_TIME
    current_brief     = DEFAULT_BRIEF
    current_commander = DEFAULT_COMMANDER
    df_cmd_edit       = DEFAULT_CMD.copy()
    df_patrol_edit    = DEFAULT_PATROL.copy()
elif df_set is None:
    st.info("💡 資料庫無資料，載入預設範本。")
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

# 介面顯示
st.title("🚔 防制危險駕車專案勤務規劃表")
st.caption("資料與 Google Sheets 即時連線，手機、電腦皆可編輯")

st.subheader("1. 基礎資訊")
c1, c2 = st.columns(2)
plan_time  = c1.text_input("勤務時間", value=current_time)
brief_info = c2.text_area("📢 勤前教育", value=current_brief, height=80)

st.subheader("2. 任務編組")
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

# HTML 預覽產生器 (同步 PDF 樣式)
def generate_html_preview():
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
        .cmd-row { background-color: white; text-align: left; padding-left: 10px; }
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2>{UNIT}執行「防制危險駕車專案勤務」規劃表</h2>"
    html += f"<div class='info'>勤務時間：{plan_time}</div>"

    # 任務編組表格
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr>"
    html += "<tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, row in edited_cmd.iterrows():
        name = str(row.get('姓名', '')).replace("、", "<br>").replace(",", "<br>")
        html += f"<tr><td><b>{row.get('職稱','')}</b></td><td>{row.get('代號','')}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務','')}</td></tr>"
    html += "</table>"

    html += f"<div class='section'><b>📢 勤前教育：</b><span style='white-space:pre-wrap'>{brief_info}</span></div>"
    
    # 警力佈署表格 (含指揮官列)
    html += "<table>"
    html += "<tr><th colspan='5'>警　力　佈　署</th></tr>"
    html += f"<tr><td colspan='5' class='cmd-row'><b>交通快打指揮官：</b>{commander}</td></tr>"
    html += "<tr><th width='15%'>勤務時段</th><th width='10%'>代號</th><th width='18%'>編組</th><th width='20%'>服勤人員</th><th width='37%'>任務分工</th></tr>"
    for _, row in edited_patrol.iterrows():
        personnel = str(row.get('服勤人員', '')).replace("、", "<br>").replace("\n", "<br>")
        html += f"<tr><td>{row.get('勤務時段','')}</td><td>{row.get('無線電','')}</td><td>{row.get('編組','')}</td><td style='line-height:1.4'>{personnel}</td><td class='left-align'>{row.get('任務分工','')}</td></tr>"
    html += "</table>"

    html += f"<div class='section'><b>巡簽地點</b><br><span style='white-space:pre-wrap'>{CHECKIN_POINTS}</span></div>"
    html += f"<div class='section'><b>備註</b><br><span class='notes'>{NOTES}</span></div>"
    html += "</div></body></html>"
    return html

html_out = generate_html_preview()

# 輸出區域
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
        save_success = save_data(plan_time, brief_info, commander, edited_cmd, edited_patrol)
        if save_success:
            subject = f"防制危險駕車勤務規劃表_{datetime.now().strftime('%Y%m%d')}"
            ok, err = send_report_email(html_out, subject, plan_time, brief_info, commander, edited_cmd, edited_patrol)
            if ok:
                st.toast("📧 報表已寄出至信箱！", icon="✉️")
            else:
                st.error(f"❌ 寄信失敗：{err}")
    st.info("💡 提示：請確保專案目錄下有 `kaiu.ttf`，否則 PDF 中文會顯示異常。")
