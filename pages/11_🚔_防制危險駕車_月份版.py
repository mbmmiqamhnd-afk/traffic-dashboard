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
st.set_page_config(page_title="防制危險駕車月份版", layout="wide", page_icon="🗓️")

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
    {"日期（22時至翌日6時）": "115年3月6日～3月7日",   "單位": "石門派出所",   "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "",                        "單位": "高平派出所",   "分工": "於中豐路及龍源路巡邏（每1小時巡邏人員至下列轄區巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "",                        "單位": "龍潭交通分隊", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "",                        "單位": "聖亭派出所",   "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "",                        "單位": "龍潭派出所",   "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "",                        "單位": "中興派出所",   "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "115年3月13日～3月14日",  "單位": "石門派出所",   "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "",                        "單位": "高平派出所",   "分工": "於中豐路及龍源路巡邏（每1小時巡邏人員至下列轄區巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "",                        "單位": "龍潭交通分隊", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "",                        "單位": "聖亭派出所",   "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "",                        "單位": "龍潭派出所",   "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "",                        "單位": "中興派出所",   "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "115年3月20日～3月21日",  "單位": "石門派出所",   "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "",                        "單位": "高平派出所",   "分工": "於中豐路及龍源路巡邏（每1小時巡邏人員至下列轄區巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "",                        "單位": "龍潭交通分隊", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "",                        "單位": "聖亭派出所",   "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "",                        "單位": "龍潭派出所",   "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "",                        "單位": "中興派出所",   "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "115年3月27日～3月28日",  "單位": "石門派出所",   "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "",                        "單位": "高平派出所",   "分工": "於中豐路及龍源路巡邏（每1小時巡邏人員至下列轄區巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "",                        "單位": "龍潭交通分隊", "分工": "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"},
    {"日期（22時至翌日6時）": "",                        "單位": "聖亭派出所",   "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "",                        "單位": "龍潭派出所",   "分工": "於轄內易發生危險駕車路段巡邏"},
    {"日期（22時至翌日6時）": "",                        "單位": "中興派出所",   "分工": "於轄內易發生危險駕車路段巡邏"},
])

CHECKIN_POINTS = """1. 中油高原交流道站（龍源路2-20號）
2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）
3. 7-11龍潭佳園門市（中正路三坑段776號）
4. 旭日路三坑自然生態公園停車場
5. 旭日路與大溪區交界處"""

NOTES = """一、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。
二、駕駛巡邏車應開啟警示燈，如發現有危險駕車行為「勿追車」，並立即向勤指中心報告，請求以優勢警力執行攔截圍捕。
三、針對下列違法、違規事項加強攔查：
（一）道路交通管理處罰條例第16條（改裝排氣管）、第18條（改裝車體設備）、第21條（無照駕駛）及43條各款項（蛇行、嚴重超速、逼車、任意減速、拆除消音器、以其他方式造成噪音、兩車以上競速等）。
（二）違反刑法185條妨害公眾往來安全罪。
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
        
        ws_set = next((w for w in ws_list if w.title == "危駕月_設定"), None)
        ws_cmd = next((w for w in ws_list if w.title == "危駕月_指揮組"), None)
        ws_sch = next((w for w in ws_list if w.title == "危駕月_勤務表"), None)

        if not all([ws_set, ws_cmd, ws_sch]):
            return None, None, None, "缺工作表 (需有: 危駕月_設定, 危駕月_指揮組, 危駕月_勤務表)"

        df_settings = pd.DataFrame(ws_set.get_all_records())
        df_cmd      = pd.DataFrame(ws_cmd.get_all_records())
        df_schedule = pd.DataFrame(ws_sch.get_all_records())
        return df_settings, df_cmd, df_schedule, None
    except Exception as e:
        return None, None, None, str(e)

# --- 4. 寫入函數 ---
def save_data(month, df_cmd, df_schedule):
    try:
        client = get_client()
        if client is None:
            st.warning("⚠️ 離線模式無法存檔至雲端，僅能下載 PDF。")
            return False

        sh = client.open_by_key(SHEET_ID)

        ws_set = sh.worksheet("危駕月_設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["month", month]])

        ws_cmd = sh.worksheet("危駕月_指揮組")
        ws_cmd.clear()
        df_cmd = df_cmd.fillna("")
        ws_cmd.update([df_cmd.columns.tolist()] + df_cmd.values.tolist())

        ws_sch = sh.worksheet("危駕月_勤務表")
        ws_sch.clear()
        df_schedule = df_schedule.fillna("")
        ws_sch.update([df_schedule.columns.tolist()] + df_schedule.values.tolist())

        load_data.clear()
        st.toast("✅ 雲端存檔成功！", icon="☁️")
        return True
    except Exception as e:
        st.error(f"❌ 存檔失敗：{e}")
        return False

# --- 5. PDF 生成函數 (含自動合併日期欄位) ---
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

def generate_pdf_from_data(month, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=15*mm, bottomMargin=15*mm,
        title=f"{UNIT}{month}防制危險駕車專案勤務規劃表")
        
    page_width = A4[0] - 30*mm
    story = []
    
    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, leading=22, spaceAfter=10)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=10, leading=13, alignment=1) # 置中
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=10, leading=13, alignment=0) # 靠左
    style_section = ParagraphStyle('Section', fontName=font, fontSize=11, leading=16, spaceAfter=4)
    style_note = ParagraphStyle('Note', fontName=font, fontSize=10, leading=14, spaceAfter=2)
    style_table_header = ParagraphStyle('TableHeader', fontName=font, fontSize=14, alignment=1, leading=18)

    # 1. 標題
    story.append(Paragraph(f"{UNIT}{month}執行「防制危險駕車」專案勤務規劃表", style_title))
    
    def clean_text(txt):
        return str(txt).replace("\n", "<br/>").replace("、", "<br/>")

    # ====================
    # 2. 任務編組表格
    # ====================
    col_widths_cmd = [page_width * 0.15, page_width * 0.10, page_width * 0.25, page_width * 0.50]
    headers_cmd = ["職稱", "代號", "姓名", "任務"]
    
    data_cmd = []
    data_cmd.append([Paragraph("<b>任　務　編　組</b>", style_table_header), '', '', ''])
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
        ('SPAN', (0,0), (-1,0)),
        ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#f2f2f2')),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('BACKGROUND', (0,1), (-1, 1), colors.HexColor('#f2f2f2')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    # ====================
    # 3. 勤務表 (日期、單位、分工) - 含自動合併功能
    # ====================
    col_widths_sch = [page_width * 0.20, page_width * 0.20, page_width * 0.60]
    headers_sch = ["日期（22時至翌日6時）", "單位", "分工"]
    
    data_sch = []
    # Row 0: 標題列
    data_sch.append([Paragraph("<b>警　力　佈　署</b>", style_table_header), '', ''])
    # Row 1: 欄位名
    data_sch.append([Paragraph(f"<b>{h}</b>", style_cell) for h in headers_sch])
    
    for _, row in df_schedule.iterrows():
        date_p = Paragraph(clean_text(row.get('日期（22時至翌日6時）','')), style_cell)
        unit_p = Paragraph(clean_text(row.get('單位','')), style_cell)
        task_p = Paragraph(str(row.get('分工','')), style_cell_left)
        data_sch.append([date_p, unit_p, task_p])

    t2 = Table(data_sch, colWidths=col_widths_sch, repeatRows=2)
    
    # 基礎樣式
    table_styles = [
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('SPAN', (0,0), (-1,0)),
        ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#f2f2f2')),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('BACKGROUND', (0,1), (-1, 1), colors.HexColor('#f2f2f2')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]

    # --- 自動計算日期欄位合併 (PDF SPAN) ---
    # 邏輯：找出所有非空白日期的索引位置，然後將其與後面的空白列合併
    date_col = '日期（22時至翌日6時）'
    
    # 找出所有有資料的列索引 (index)
    non_empty_indices = [i for i, val in enumerate(df_schedule[date_col]) if str(val).strip() != ""]
    # 補上最後一個邊界
    non_empty_indices.append(len(df_schedule))
    
    header_offset = 2 # 因為表格前兩列是標題和欄位名稱
    
    for k in range(len(non_empty_indices) - 1):
        start_row = non_empty_indices[k]
        end_row = non_empty_indices[k+1] - 1
        
        # 如果這個區間大於 0，表示需要合併
        if end_row > start_row:
            # ReportLab 的座標是 (col, row)，這裡 col=0 是日期欄
            # 加上 header_offset 因為 data_sch 前兩列是 Headers
            span_cmd = ('SPAN', (0, start_row + header_offset), (0, end_row + header_offset))
            table_styles.append(span_cmd)
            # 確保合併後垂直置中
            valign_cmd = ('VALIGN', (0, start_row + header_offset), (0, end_row + header_offset), 'MIDDLE')
            table_styles.append(valign_cmd)

    t2.setStyle(TableStyle(table_styles))
    story.append(t2)
    story.append(Spacer(1, 6*mm))

    # ====================
    # 4. 巡簽地點 & 備註
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

def send_report_email(html_content, subject, month, df_cmd, df_schedule):
    import urllib.parse as _ul
    try:
        sender   = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        receiver = sender
        
        pdf_bytes = generate_pdf_from_data(month, df_cmd, df_schedule)
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

df_set, df_cmd, df_sch, error_msg = load_data()

if error_msg:
    st.error(f"❌ Google Sheets 讀取失敗：{error_msg}")
    st.warning("⚠️ 啟用離線範本模式。")
    current_month    = DEFAULT_MONTH
    df_cmd_edit      = DEFAULT_CMD.copy()
    df_schedule_edit = DEFAULT_SCHEDULE.copy()
elif df_set is None:
    st.info("💡 資料庫無資料，載入預設範本。")
    current_month    = DEFAULT_MONTH
    df_cmd_edit      = DEFAULT_CMD.copy()
    df_schedule_edit = DEFAULT_SCHEDULE.copy()
else:
    try:
        sd = dict(zip(df_set.iloc[:, 0], df_set.iloc[:, 1]))
        current_month    = sd.get("month", DEFAULT_MONTH)
        df_cmd_edit      = df_cmd if not df_cmd.empty else DEFAULT_CMD.copy()
        df_schedule_edit = df_sch if not df_sch.empty else DEFAULT_SCHEDULE.copy()
    except Exception as e:
        st.error(f"資料格式解析失敗：{e}")
        st.stop()

# 介面顯示
st.title("🚔 防制危險駕車專案勤務規劃表（月份版）")
st.caption("資料與 Google Sheets 即時連線，手機、電腦皆可編輯")

st.subheader("1. 基礎資訊")
current_month = st.text_input("月份", value=current_month)

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
edited_schedule = st.data_editor(df_schedule_edit, num_rows="dynamic", use_container_width=True)

st.subheader("4. 巡簽地點（固定）")
st.text(CHECKIN_POINTS)

st.subheader("5. 備註（固定）")
st.text(NOTES)

# HTML 預覽產生器 (同步 PDF 樣式 + HTML rowspan)
def generate_html_preview():
    style = """
    <style>
        body { font-family: 'DFKai-SB', 'BiauKai', '標楷體', serif; color: #000; font-size: 14px; }
        .container { width: 100%; max-width: 800px; margin: 0 auto; padding: 20px; }
        h2 { text-align: left; margin-bottom: 5px; letter-spacing: 2px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
        th, td { border: 1px solid black; padding: 5px; text-align: center; font-size: 14px; vertical-align: middle; }
        th { background-color: #f2f2f2; }
        .left-align { text-align: left; }
        .section { margin-bottom: 10px; line-height: 1.8; }
        .notes { white-space: pre-wrap; font-size: 13px; line-height: 1.8; }
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2>{UNIT}{current_month}執行「防制危險駕車」專案勤務規劃表</h2>"

    # 任務編組
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr>"
    html += "<tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, row in edited_cmd.iterrows():
        name = str(row.get('姓名', '')).replace("、", "<br>").replace(",", "<br>")
        html += f"<tr><td><b>{row.get('職稱','')}</b></td><td>{row.get('代號','')}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務','')}</td></tr>"
    html += "</table>"

    # 警力佈署 (含日期合併)
    html += "<table>"
    html += "<tr><th colspan='3'>警　力　佈　署</th></tr>"
    html += "<tr><th width='20%'>日期（22時至翌日6時）</th><th width='20%'>單位</th><th width='60%'>分工</th></tr>"
    
    col_date = '日期（22時至翌日6時）'
    total_rows = len(edited_schedule)
    
    # 建立合併資訊 (HTML rowspan)
    row_spans = {} # {row_index: span_count}
    skip_rows = set()
    
    i = 0
    while i < total_rows:
        date_val = str(edited_schedule.iloc[i][col_date]).strip()
        if date_val != "":
            # 計算接下來有多少空白
            span = 1
            for j in range(i + 1, total_rows):
                if str(edited_schedule.iloc[j][col_date]).strip() == "":
                    span += 1
                else:
                    break
            row_spans[i] = span
            for k in range(1, span):
                skip_rows.add(i + k)
            i += span
        else:
            i += 1

    for idx, row in edited_schedule.iterrows():
        html += "<tr>"
        
        # 處理日期欄位 (合併邏輯)
        if idx in row_spans:
            rowspan = row_spans[idx]
            html += f"<td rowspan='{rowspan}'>{row.get(col_date,'')}</td>"
        elif idx in skip_rows:
            pass # 被合併掉了，不輸出 td
        else:
            # 理論上不會跑到這，除非資料邏輯有誤，保險起見印出單格
            html += f"<td>{row.get(col_date,'')}</td>"
            
        html += f"<td>{row.get('單位','')}</td><td class='left-align'>{row.get('分工','')}</td></tr>"
    
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
        file_name=f"危駕月份勤務表_{datetime.now().strftime('%Y%m%d')}.html",
        mime="text/html; charset=utf-8",
        type="primary"
    ):
        save_success = save_data(current_month, edited_cmd, edited_schedule)
        if save_success:
            subject = f"防制危險駕車月份勤務規劃表_{datetime.now().strftime('%Y%m%d')}"
            ok, err = send_report_email(html_out, subject, current_month, edited_cmd, edited_schedule)
            if ok:
                st.toast("📧 報表已寄出至信箱！", icon="✉️")
            else:
                st.error(f"❌ 寄信失敗：{err}")
    st.info("💡 提示：請確保專案目錄下有 `kaiu.ttf`，否則 PDF 中文會顯示異常。")
