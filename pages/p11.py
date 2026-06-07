import streamlit as st

st.set_page_config(page_title="防制危險駕車月份版", layout="wide", page_icon="🗓️")

try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    pass

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import smtplib
import io
import os
import urllib.parse as _ul
import re
from datetime import datetime, timedelta
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

# =========================
# 常數與設定
# =========================
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
WS_MAP = {"set": "危駕月_設定", "cmd": "危駕月_指揮組", "sch": "危駕月_勤務表"}
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"
CMD_COLS = ["職稱", "代號", "姓名", "任務"]
SCH_COLS = ["日期（22時至翌日6時）", "單位", "巡邏路段"]

# 固定6單位與對應巡邏路段
FIXED_UNITS = [
    ("石門派出所", "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"),
    ("高平派出所", "於中豐路及龍源路巡邏（每1小時巡邏人員至下列轄區巡簽地點巡簽1次）"),
    ("龍潭交通分隊", "於中正路、文化路、中豐路、龍源路及旭日巡邏（每1小時巡邏人員至下列巡簽地點巡簽1次）"),
    ("聖亭派出所", "於轄內易發生危險駕車路段巡邏"),
    ("龍潭派出所", "於轄內易發生危險駕車路段巡邏"),
    ("中興派出所", "於轄內易發生危險駕車路段巡邏"),
]

# =========================
# 預設資料
# =========================
DEFAULT_MONTH = "115年5月份"
DEFAULT_HOLIDAYS = "4/30, 5/1, 5/2, 5/8, 5/9"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "建興", "姓名": "駐區督察 孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝裝備及勤務紀律。"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、秘書室巡官 陳鵬翔（代理人：警員張庭溱）、人事室警員 陳明祥、行政組警務佐 曾威仁", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"}
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

# =========================
# 假日日期解析與自動分組
# =========================
def parse_holidays(holiday_str, roc_year_str):
    """
    解析輸入的假日字串（如 '4/30, 5/1, 5/2, 5/8, 5/9'）
    自動將連續日期合併為同一區間，回傳 DataFrame（SCH_COLS 格式）
    """
    # 從月份資訊擷取民國年，轉為西元年
    year_match = re.search(r"(\d+)年", roc_year_str)
    if year_match:
        ce_year = int(year_match.group(1)) + 1911
    else:
        ce_year = datetime.now().year

    # 解析日期字串
    dates = []
    for part in holiday_str.replace("，", ",").split(","):
        part = part.strip()
        m = re.match(r"(\d+)[/月](\d+)", part)
        if m:
            try:
                d = datetime(ce_year, int(m.group(1)), int(m.group(2)))
                dates.append(d)
            except ValueError:
                pass

    if not dates:
        return pd.DataFrame(columns=SCH_COLS)

    dates = sorted(set(dates))

    # 將連續日期分組
    groups = []
    group = [dates[0]]
    for i in range(1, len(dates)):
        if (dates[i] - dates[i-1]).days == 1:
            group.append(dates[i])
        else:
            groups.append(group)
            group = [dates[i]]
    groups.append(group)

    # 組建 DataFrame
    rows = []
    for group in groups:
        if len(group) == 1:
            label = f"{group[0].month}月{group[0].day}日"
        else:
            start, end = group[0], group[-1]
            if start.month == end.month:
                label = f"{start.month}月{start.day}日～{end.day}日"
            else:
                label = f"{start.month}月{start.day}日～\n{end.month}月{end.day}日"

        for i, (unit, patrol) in enumerate(FIXED_UNITS):
            rows.append({
                "日期（22時至翌日6時）": label if i == 0 else "",
                "單位": unit,
                "巡邏路段": patrol,
            })

    return pd.DataFrame(rows, columns=SCH_COLS)

# =========================
# Google Sheets
# =========================
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        st.error("❌ 找不到 gcp_service_account，請確認 Secrets 設定。")
        return None
    try:
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Google 授權失敗：{e}")
        return None

def init_sheets():
    client = get_client()
    if client is None: return
    sh = client.open_by_key(SHEET_ID)
    headers = {WS_MAP["set"]: [["Key", "Value"]], WS_MAP["cmd"]: [CMD_COLS], WS_MAP["sch"]: [SCH_COLS]}
    for name, header in headers.items():
        try:
            sh.worksheet(name)
        except:
            sh.add_worksheet(title=name, rows="200", cols="20").update(header)
    st.success("✅ 初始化完成")
    st.cache_data.clear()
    st.rerun()

@st.cache_data(ttl=600)
def load_data():
    try:
        client = get_client()
        if client is None:
            return None, None, None, {}, "授權失敗"
        sh = client.open_by_key(SHEET_ID)
        set_df = pd.DataFrame(sh.worksheet(WS_MAP["set"]).get_all_records()).fillna("")
        cmd_df = pd.DataFrame(sh.worksheet(WS_MAP["cmd"]).get_all_records()).fillna("")
        sch_df = pd.DataFrame(sh.worksheet(WS_MAP["sch"]).get_all_records()).fillna("")
        if not sch_df.empty:
            if "分工" in sch_df.columns and "巡邏路段" not in sch_df.columns:
                sch_df = sch_df.rename(columns={"分工": "巡邏路段"})
            sch_df = sch_df.reindex(columns=SCH_COLS, fill_value="")
        settings = {}
        if not set_df.empty and set_df.shape[1] >= 2:
            settings = dict(zip(set_df.iloc[:,0].astype(str), set_df.iloc[:,1].astype(str)))
        return set_df, cmd_df, sch_df, settings, None
    except Exception as e:
        return None, None, None, {}, str(e)

def save_data(settings_dict, cmd, sch):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet(WS_MAP["set"])
        ws_set.clear()
        ws_set.update([["Key", "Value"]] + [[k, v] for k, v in settings_dict.items()])
        for ws_name, df, cols in [(WS_MAP["cmd"], cmd, CMD_COLS), (WS_MAP["sch"], sch, SCH_COLS)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df_clean = df[cols].fillna("")
            if not df_clean.empty:
                ws.update([df_clean.columns.tolist()] + df_clean.values.tolist())
        load_data.clear()
        return True
    except Exception as e:
        st.error(f"❌ 儲存失敗：{e}")
        return False

# =========================
# PDF 生成
# =========================
@st.cache_resource
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    for p in ["kaiu.ttf", "./kaiu.ttf", "/usr/share/fonts/truetype/kaiu.ttf", "/app/kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf(full_title, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []

    def add_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 10)
        canvas.drawCentredString(A4[0]/2.0, 10*mm, f"第 {canvas.getPageNumber()} 頁")
        canvas.restoreState()

    s_title    = ParagraphStyle('Title',     fontName=font, fontSize=16, leading=22, alignment=1, spaceAfter=10)
    s_th       = ParagraphStyle('THeader',   fontName=font, fontSize=16, alignment=1, leading=22)
    s_col      = ParagraphStyle('ColHeader', fontName=font, fontSize=14, leading=20, alignment=1)
    s_cell     = ParagraphStyle('Cell',      fontName=font, fontSize=12, leading=18, alignment=1)
    s_left     = ParagraphStyle('CellLeft',  fontName=font, fontSize=12, leading=18, alignment=0)
    s_hanging  = ParagraphStyle('Hanging',   fontName=font, fontSize=12, leading=20, leftIndent=8.5*mm, firstLineIndent=-8.5*mm, spaceAfter=5)
    s_section  = ParagraphStyle('Section',   fontName=font, fontSize=13, leading=20, spaceAfter=4)

    def clean(txt): return str(txt).replace("\n", "<br/>").replace("、", "<br/>")

    story.append(Paragraph(f"<b>{full_title}</b>", s_title))

    # 任務編組
    data_cmd = [[Paragraph("<b>任 務 編 組</b>", s_th), '', '', ''],
                [Paragraph(f"<b>{h}</b>", s_col) for h in CMD_COLS]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([
            Paragraph(f"<b>{r.get('職稱','')}</b>", s_cell),
            Paragraph(str(r.get('代號','')), s_cell),
            Paragraph(clean(r.get('姓名','')), s_cell),
            Paragraph(str(r.get('任務','')), s_left)
        ])
    t1 = Table(data_cmd, colWidths=[page_width*0.14, page_width*0.11, page_width*0.28, page_width*0.45], repeatRows=2)
    t1.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)),
        ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')),
        ('TOPPADDING',(0,0),(-1,-1),5), ('BOTTOMPADDING',(0,0),(-1,-1),5)
    ]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    # 警力佈署
    data_sch = [[Paragraph("<b>警 力 佈 署</b>", s_th), '', ''],
                [Paragraph(f"<b>{h}</b>", s_col) for h in SCH_COLS]]
    for _, r in df_schedule.iterrows():
        data_sch.append([
            Paragraph(clean(r.get('日期（22時至翌日6時）','')), s_cell),
            Paragraph(clean(r.get('單位','')), s_cell),
            Paragraph(str(r.get('巡邏路段','')), s_left)
        ])

    t2_styles = [
        ('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)),
        ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')),
        ('TOPPADDING',(0,0),(-1,-1),5), ('BOTTOMPADDING',(0,0),(-1,-1),5)
    ]

    # 自動合併日期欄（連續空白格合併至有值的格）
    date_col = '日期（22時至翌日6時）'
    if not df_schedule.empty:
        non_empty = [i for i, val in enumerate(df_schedule[date_col]) if str(val).strip() != ""]
        non_empty.append(len(df_schedule))
        for k in range(len(non_empty) - 1):
            s, e = non_empty[k], non_empty[k+1] - 1
            if e > s:
                t2_styles.append(('SPAN',   (0, s+2), (0, e+2)))
                t2_styles.append(('VALIGN', (0, s+2), (0, e+2), 'MIDDLE'))

    t2 = Table(data_sch, colWidths=[page_width*0.22, page_width*0.22, page_width*0.55], repeatRows=2)
    t2.setStyle(TableStyle(t2_styles))
    story.append(t2)
    story.append(Spacer(1, 6*mm))

    # 巡簽地點
    story.append(Paragraph("<b>巡簽地點：</b>", s_section))
    for line in CHECKIN_POINTS.split('\n'):
        if line.strip(): story.append(Paragraph(line, s_hanging))
    story.append(Spacer(1, 4*mm))

    # 備註
    story.append(Paragraph("<b>備註：</b>", s_section))
    for line in NOTES.split('\n'):
        if line.strip(): story.append(Paragraph(line, s_hanging))

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    buf.seek(0)
    return buf

# =========================
# Email
# =========================
def send_email(subject, pdf_buf, filename):
    try:
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender
        msg["Subject"] = subject
        msg.attach(MIMEText("附件為最新勤務規劃表（月份版）。", "plain", "utf-8"))
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_buf.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(filename)}.pdf")
        msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(sender, pwd)
            s.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

# =========================
# 主畫面
# =========================
st.title("🚔 防制危險駕車專案勤務規劃表（月份版）")

if st.sidebar.button("🔧 初始化工作表"):
    init_sheets()
if st.sidebar.button("🔄 強制重新載入"):
    st.cache_data.clear()
    st.rerun()

df_set, df_cmd_raw, df_sch_raw, settings, err = load_data()
if err:
    st.warning(f"⚠️ 無法連線 Google Sheets（{err}），顯示預設底稿。")

# --- 基本設定 ---
st.subheader("1. 基本設定")
col_a, col_b = st.columns([1, 2])
c_month   = col_a.text_input("月份資訊", value=settings.get("month", DEFAULT_MONTH))
c_holidays = col_b.text_input(
    "假日日期（用逗號分隔，例：4/30, 5/1, 5/2, 5/8, 5/9）",
    value=settings.get("holidays", DEFAULT_HOLIDAYS),
    help="輸入當月所有執行日期，連續日期會自動合併為同一區間"
)

full_title = f"{UNIT}{c_month}執行「防制危險駕車」專案勤務規劃表"

# --- 自動產生警力佈署 ---
auto_sch = parse_holidays(c_holidays, c_month)

# 若雲端有資料且假日輸入與雲端一致則優先用雲端（允許手動微調後保留）
# 若假日欄位變動或雲端無資料，則用自動產生的結果
saved_holidays = settings.get("holidays", "")
if (df_sch_raw is not None and not df_sch_raw.empty) and (saved_holidays == c_holidays):
    use_sch = df_sch_raw
else:
    use_sch = auto_sch

# --- 任務編組 ---
st.subheader("2. 任務編組")
ed_cmd = df_cmd_raw if (df_cmd_raw is not None and not df_cmd_raw.empty) else DEFAULT_CMD.copy()
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True)

# --- 警力佈署 ---
st.subheader("3. 警力佈署")
st.caption("💡 修改上方假日日期後，表格會自動重新產生")
res_sch = st.data_editor(
    use_sch,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "日期（22時至翌日6時）": st.column_config.TextColumn(disabled=True),
    }
)
st.markdown("---")

if st.button("💾 儲存至雲端並發送 Email", use_container_width=True):
    with st.spinner("正在執行儲存與發信作業，請稍候..."):
        s = {"month": c_month, "holidays": c_holidays}
        save_ok = save_data(s, res_cmd, res_sch)
        if save_ok:
            st.success("✅ 雲端試算表資料儲存成功！")
            pdf_buf = generate_pdf(full_title, res_cmd, res_sch)
            mail_ok, mail_err = send_email(full_title, pdf_buf, full_title)
            if mail_ok:
                st.success("📧 最新勤務表 PDF 已成功寄出！")
            else:
                st.error(f"❌ Email 發送失敗：{mail_err}")
        else:
            st.error("❌ 雲端儲存失敗，已中止發送 Email 作業。")
