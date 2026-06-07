import streamlit as st

st.set_page_config(page_title="取締砂石車專案勤務", layout="wide", page_icon="🚛")

try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    pass

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib, io, os
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import re

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

WEEKDAY_ZH = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]

# 固定8單位（依 PDF 順序）
FIXED_UNITS = [
    ("聖亭派出所",   "2至4人", "中豐路、聖亭路段等砂石（大型貨）車行經路段"),
    ("龍潭派出所",   "2至4人", "大昌路、中豐路段砂石（大型貨）車行經路段"),
    ("中興派出所",   "2至4人", "中興路、福龍路及龍平路段砂石（大型貨）車行經路段"),
    ("石門派出所",   "2至4人", "中正路、龍源路及民族路段砂石（大型貨）車行經路段"),
    ("高平派出所",   "2至4人", "中豐路、龍源路段砂石（大型貨）車行經路段"),
    ("三和派出所",   "2至4人", "楊銅路、龍新路段砂石（大型貨）車行經路段"),
    ("警備隊",       "2至4人", "中豐路、龍源路、聖亭路段砂石（大型貨）車行經路段"),
    ("龍潭交通分隊", "2至4人", "中豐路、龍源路、聖亭路段砂石（大型貨）車行經路段"),
]

CORRECT_NOTES = (
    "一、執行前由各單位帶班人員在駐地實施勤前教育。<br/>"
    "二、加強取締砂石（大型貨）車超載、車速、酒醉駕車、闖紅燈、無照駕車、爭道行駛、"
    "違反禁行路線、變更車斗、未使用專用車箱及未裝設行車紀錄器（行車視野輔助器）等違規，"
    "以共同消弭不法行為，保障用路人生命財產安全。"
)

DEFAULT_MONTH    = "115年6月份"
DEFAULT_HOLIDAYS = "6/18, 6/26"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官",     "代號": "隆安1",   "姓名": "分局長 施宇峰",   "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官",   "代號": "隆安2",   "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官",   "代號": "隆安3",   "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "建興",    "姓名": "駐區督察 孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組",     "代號": "隆安6",   "姓名": "督察組組長 黃長旗\n督察組督察員 黃中彥\n督察組警務員 陳冠彰\n督察組巡官 古家杰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組",     "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟\n交通組警務員 盧冠仁\n交通組警務員 李峯甫\n交通組警務員 葉佳媛\n交通組巡官 郭勝隆\n交通組警員 吳享運\n保安民防組巡官 陳鵬翔（代理人：警員張庭溱）\n人事室警員 陳明祥\n行政組警務佐 曾威仁", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組",     "代號": "隆安",    "姓名": "主任 蔡奇青\n執勤官 李文章\n執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"},
])

# =========================
# 假日日期解析與自動產生勤務表
# =========================
def parse_holidays(holiday_str, roc_year_str):
    """
    解析輸入的假日字串（如 '6/18, 6/26'）
    每個日期獨立一組（不合併），格式：115年6月18日（星期四）00時至24時
    回傳 DataFrame（勤務日期、執行單位、執行人數、執行路段）
    """
    year_match = re.search(r"(\d+)年", roc_year_str)
    ce_year = int(year_match.group(1)) + 1911 if year_match else datetime.now().year

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
        return pd.DataFrame(columns=["勤務日期", "執行單位", "執行人數", "執行路段"])

    dates = sorted(set(dates))

    rows = []
    for i, date in enumerate(dates):
        roc_year = date.year - 1911
        weekday  = WEEKDAY_ZH[date.weekday()]
        label    = f"{roc_year}年{date.month}月{date.day}日（{weekday}）00時至24時"
        for j, (unit, count, patrol) in enumerate(FIXED_UNITS):
            rows.append({
                "勤務日期": label if j == 0 else "",
                "執行單位": unit,
                "執行人數": count,
                "執行路段": patrol,
            })

    return pd.DataFrame(rows, columns=["勤務日期", "執行單位", "執行人數", "執行路段"])

# --- Google Sheets ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        st.error("❌ 找不到 gcp_service_account，請確認 Secrets 設定。")
        return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Google 授權失敗：{e}")
        return None

@st.cache_data(ttl=600)
def load_data():
    try:
        client = get_client()
        if client is None:
            return None, None, None, "授權失敗"
        sh  = client.open_by_key(SHEET_ID)
        df_set = pd.DataFrame(sh.worksheet("砂石_設定").get_all_records()).fillna("")
        df_cmd = pd.DataFrame(sh.worksheet("砂石_指揮組").get_all_records()).fillna("")
        df_sch = pd.DataFrame(sh.worksheet("砂石_勤務表").get_all_records()).fillna("")
        if not df_sch.empty and "日期" in df_sch.columns:
            df_sch.rename(columns={"日期": "勤務日期"}, inplace=True)
        sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1])) if not df_set.empty else {}
        return df_cmd, df_sch, sd, None
    except Exception as e:
        return None, None, {}, str(e)

def save_data(month, holidays, df_cmd, df_schedule):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("砂石_設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["month", month], ["holidays", holidays]])
        for ws_name, df in [("砂石_指揮組", df_cmd), ("砂石_勤務表", df_schedule)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df_cleaned = df.dropna(how='all').fillna("")
            if not df_cleaned.empty:
                ws.update([df_cleaned.columns.tolist()] + df_cleaned.values.tolist())
        load_data.clear()
        return True
    except Exception as e:
        st.error(f"❌ 同步失敗：{e}")
        return False

# --- PDF 產生 ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    for p in ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/custom/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf(month, df_cmd, df_schedule, title_full):
    font = _get_font()
    buf  = io.BytesIO()
    doc  = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=12*mm, bottomMargin=12*mm)
    W    = A4[0] - 24*mm

    s_title = ParagraphStyle("t",   fontName=font, fontSize=16, alignment=1, spaceAfter=8, leading=22, wordWrap='CJK')
    s_th    = ParagraphStyle("th",  fontName=font, fontSize=16, alignment=1, leading=22,   wordWrap='CJK')
    s_cell  = ParagraphStyle("c",   fontName=font, fontSize=14, leading=18,  alignment=1,  wordWrap='CJK')
    s_left  = ParagraphStyle("l",   fontName=font, fontSize=14, leading=18,  alignment=0,  wordWrap='CJK')
    s_note  = ParagraphStyle("n",   fontName=font, fontSize=12, leading=16,  spaceAfter=4, wordWrap='CJK')

    def c(txt, style=s_cell): return Paragraph(str(txt).replace("\n", "<br/>"), style)

    story = []
    story.append(Paragraph(f"<b>{title_full}</b>", s_title))

    # 任務編組
    data1 = [
        [Paragraph("<b>任 務 編 組</b>", s_th), '', '', ''],
        [Paragraph(f"<b>{h}</b>", s_th) for h in ["職稱", "代號", "姓名", "任務"]]
    ]
    for _, row in df_cmd.iterrows():
        data1.append([
            c(f"<b>{row.get('職稱','')}</b>"),
            c(row.get('代號', '')),
            c(str(row.get('姓名', '')).replace('、', '<br/>')),
            c(row.get('任務', ''), s_left)
        ])
    t1 = Table(data1, colWidths=[W*0.15, W*0.12, W*0.28, W*0.45], repeatRows=2)
    t1.setStyle(TableStyle([
        ('FONTNAME',   (0,0), (-1,-1), font),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('SPAN',       (0,0), (-1,0)),
        ('BACKGROUND', (0,0), (-1,1),  colors.HexColor('#f2f2f2')),
    ]))
    story.append(t1)
    story.append(Spacer(1, 4*mm))

    # 警力佈署
    col_date = "勤務日期"
    data2 = [
        [Paragraph("<b>警 力 佈 署</b>", s_th), '', '', ''],
        [Paragraph(f"<b>{h}</b>", s_th) for h in ["勤務日期", "執行單位", "執行人數", "執行路段"]]
    ]
    for _, row in df_schedule.iterrows():
        data2.append([
            c(row.get(col_date, '')),
            c(row.get('執行單位', '')),
            c(row.get('執行人數', '')),
            c(row.get('執行路段', ''), s_left)
        ])

    t_style = [
        ('FONTNAME',   (0,0), (-1,-1), font),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('SPAN',       (0,0), (-1,0)),
        ('BACKGROUND', (0,0), (-1,1),  colors.HexColor('#f2f2f2')),
    ]
    if not df_schedule.empty:
        non_empty = [i for i, v in enumerate(df_schedule[col_date]) if str(v).strip() != ""]
        non_empty.append(len(df_schedule))
        for k in range(len(non_empty) - 1):
            s_idx, e_idx = non_empty[k], non_empty[k+1] - 1
            if e_idx > s_idx:
                t_style.append(('SPAN', (0, s_idx+2), (0, e_idx+2)))

    t2 = Table(data2, colWidths=[W*0.30, W*0.14, W*0.12, W*0.44], repeatRows=2)
    t2.setStyle(TableStyle(t_style))
    story.append(KeepTogether([t2]))
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph(f"<b>備註：</b><br/>{CORRECT_NOTES}", s_note))

    doc.build(story)
    return buf.getvalue()

# --- 寄信 ---
def send_report_email(subject, pdf_bytes, pdf_filename):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"] = sender; msg["To"] = sender; msg["Subject"] = subject
        msg.attach(MIMEText("附件為最新勤務規劃表 PDF。", "plain", "utf-8"))
        part = MIMEBase("application", "pdf"); part.set_payload(pdf_bytes); encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(pdf_filename)}.pdf")
        msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd); server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# --- 主程式 ---
st.title("🚛 取締砂石（大型貨）車重點違規專案勤務規劃表")

if st.sidebar.button("🔧 初始化工作表"):
    pass  # 工作表已存在，略過
if st.sidebar.button("🔄 強制重新載入"):
    st.cache_data.clear()
    st.rerun()

df_cmd_raw, df_sch_raw, sd, err = load_data()
if err:
    st.warning(f"⚠️ 無法連線 Google Sheets（{err}），顯示預設資料。")

# --- 基本設定 ---
st.subheader("1. 基本設定")
col_a, col_b = st.columns([1, 2])
month_val    = col_a.text_input("月份", value=sd.get("month", DEFAULT_MONTH))
holiday_val  = col_b.text_input(
    "執行日期（用逗號分隔，例：6/18, 6/26）",
    value=sd.get("holidays", DEFAULT_HOLIDAYS),
    help="每個日期各自獨立一組，自動帶入星期與時間格式"
)

full_table_title = f"{UNIT}執行{month_val}「取締砂石（大型貨）車重點違規」專案勤務規劃表"

# 自動產生警力佈署
auto_sch = parse_holidays(holiday_val, month_val)

# 假日沒變動時優先用雲端資料（保留手動調整）
saved_holidays = sd.get("holidays", "")
if (df_sch_raw is not None and not df_sch_raw.empty) and (saved_holidays == holiday_val):
    use_sch = df_sch_raw
else:
    use_sch = auto_sch

# --- 任務編組 ---
st.subheader("2. 任務編組")
ed_cmd = df_cmd_raw if (df_cmd_raw is not None and not df_cmd_raw.empty) else DEFAULT_CMD.copy()
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True)

# --- 警力佈署 ---
st.subheader("3. 警力佈署")
st.caption("💡 修改上方執行日期後，表格會自動重新產生")
res_sch = st.data_editor(
    use_sch,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "勤務日期": st.column_config.TextColumn(disabled=True),
    }
)

st.markdown("---")

if st.button("💾 同步雲端並發送備份郵件", type="primary", use_container_width=True):
    with st.spinner("處理中..."):
        res_cmd_clean = res_cmd.dropna(how="all").fillna("")
        res_sch_clean = res_sch.dropna(how="all").fillna("")
        if save_data(month_val, holiday_val, res_cmd_clean, res_sch_clean):
            pdf_bytes = generate_pdf(month_val, res_cmd_clean, res_sch_clean, full_table_title)
            ok, mail_err = send_report_email(full_table_title, pdf_bytes, full_table_title)
            if ok:
                st.success("✅ 同步與郵件發送成功！")
            else:
                st.warning(f"⚠️ 雲端已同步，但郵件失敗：{mail_err}")
        else:
            st.error("❌ 雲端同步失敗，請檢查權限設定。")
