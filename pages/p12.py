import streamlit as st

st.set_page_config(page_title="行人及護老交通安全", layout="wide", page_icon="🚶")

try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    st.sidebar.warning("找不到 menu.py，跳過側邊欄載入。")

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import calendar
import smtplib
import io
import os
import traceback
import urllib.parse as _ul
import re
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

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"
WEEKDAY_ZH = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]

# 固定8單位
FIXED_UNITS = [
    ("聖亭派出所",   "中豐路、聖亭路段\n校園周邊道路或轄區行人易肇事路口"),
    ("龍潭派出所",   "中豐路、中正路段\n校園周邊道路或轄區行人易肇事路口"),
    ("中興派出所",   "中興路、福龍路段\n校園周邊道路或轄區行人易肇事路口"),
    ("石門派出所",   "中正、文化路段\n校園周邊道路或轄區行人易肇事路口"),
    ("高平派出所",   "中豐、中原路段\n校園周邊道路或轄區行人易肇事路口"),
    ("三和派出所",   "龍新路、楊銅路段\n校園周邊道路或轄區行人易肇事路口"),
    ("警備隊",       "校園周邊道路或轄區行人易肇事路口"),
    ("龍潭交通分隊", "校園周邊道路或轄區行人易肇事路口"),
]

DEFAULT_MONTH    = "115年3月份"

# ✨ 各月份預設假日對照表（輸入月份時會自動帶出，並保留手動修改權限）
MONTH_HOLIDAYS_MAP = {
    "115年1月份": "1/1, 1/19, 1/20, 1/21, 1/22, 1/23, 1/26",  # 元旦與春節連假
    "115年2月份": "2/27",                                     # 228連假
    "115年3月份": "3/3, 3/10",                                 # 預設範例
    "115年4月份": "4/2, 4/3",                                 # 清明連假
    "115年5月份": "5/1",                                      # 勞動節
    "115年6月份": "6/19",                                     # 端午節
    "115年7月份": "",
    "115年8月份": "",
    "115年9月份": "9/25",                                     # 中秋節
    "115年10月份": "10/9",                                    # 國慶連假
    "115年11月份": "",
    "115年12月份": ""
}
DEFAULT_HOLIDAYS = "3/3, 3/10"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官",       "代號": "隆安1",   "姓名": "分局長 施宇峰",   "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官",     "代號": "隆安2",   "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官",     "代號": "隆安3",   "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官",   "代號": "駐區督察", "姓名": "孫三陽",         "任務": "重點機動督導。"},
    {"職稱": "督導組",       "代號": "隆安6",   "姓名": "督察組組長 黃長旗\n督察組督察員 黃中彥\n督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組",       "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13",  "姓名": "交通組組長 楊孟竟\n交通組警務員 盧冠仁\n交通組警務員 李峯甫\n交通組巡官 郭勝隆\n交通組巡官 羅千金\n交通組警員 吳享運\n秘書室巡官 陳鵬翔（代理人：警員張庭溱）\n人事室警員 陳明祥", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組",       "代號": "隆安",    "姓名": "主任 蔡奇青\n執勤官 李文章\n執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_NOTES = """壹、警察局規劃本月份「行人及護老交通安全專案勤務」，各編組依規劃日期執行。
貳、執行本專案勤務視轄區狀況及執勤警力，擇定轄區易肇事路口（段）及校園周邊道路，依上揭日期妥適編排勤務協助維護行人、學童及高齡者通行安全，並加強取締「車不讓人」、「未依規定停讓」、「違規停車」等。"""

# ====================================================
# 自動產生上班日文字
# ====================================================
def generate_workday_label(month_str, holiday_str):
    """
    依月份與放假日，產生上班日清單文字
    格式：3月3日（星期一）、4日（星期二）、…
    支援：+月/日 代表週六日強行補班（例如 +3/21）
    """
    year_match = re.search(r"(\d+)年", month_str)
    mon_match  = re.search(r"(\d+)月", month_str)
    if not year_match or not mon_match:
        return ""

    ce_year = int(year_match.group(1)) + 1911
    month   = int(mon_match.group(1))

    # 解析放假日與補班日
    holidays = set()
    work_weekends = set()
    
    for part in holiday_str.replace("，", ",").split(","):
        part = part.strip()
        
        if part.startswith("+"):
            m = re.match(r"\+(\d+)[/月](\d+)", part)
            if m:
                try:
                    d = datetime(ce_year, int(m.group(1)), int(m.group(2)))
                    work_weekends.add(d.date())
                except ValueError:
                    pass
        else:
            m = re.match(r"(\d+)[/月](\d+)", part)
            if m:
                try:
                    d = datetime(ce_year, int(m.group(1)), int(m.group(2)))
                    holidays.add(d.date())
                except ValueError:
                    pass

    # 取得當月所有日期
    _, days_in_month = calendar.monthrange(ce_year, month)
    workdays = []
    for day in range(1, days_in_month + 1):
        d = datetime(ce_year, month, day)
        current_date = d.date()
        
        if current_date in work_weekends:
            workdays.append(d)
        elif d.weekday() < 5 and current_date not in holidays:
            workdays.append(d)

    if not workdays:
        return "（本月無上班日）"

    # 排序日期
    workdays.sort()

    # 組成文字：第一天含月份，後續只顯示日
    parts = []
    for i, d in enumerate(workdays):
        wd = WEEKDAY_ZH[d.weekday()]
        if i == 0:
            parts.append(f"{month}月{d.day}日（{wd}）")
        else:
            parts.append(f"{d.day}日（{wd}）")

    # 僅回傳純粹的日期與星期組合文字
    return "、".join(parts)


def build_schedule_df(month_str, holiday_str):
    """產生警力佈署 DataFrame，日期欄第一列填入，其餘留空"""
    label = generate_workday_label(month_str, holiday_str)
    # 欄位標題精準指定時間區段
    col   = "執行勤務日期（6時至10時，16時至20時）"
    rows  = []
    for i, (unit, patrol) in enumerate(FIXED_UNITS):
        rows.append({
            col:    label if i == 0 else "",
            "單位": unit,
            "路段": patrol,
        })
    return pd.DataFrame(rows, columns=[col, "單位", "路段"])


# --- Google Sheets ---
@st.cache_resource
def get_client():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Google 授權失敗：{e}")
        return None

def clean_df_to_list(df):
    return df.astype(str).values.tolist()

@st.cache_data(ttl=600)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, None, {}, "權限不足或未設定 Secrets"
        sh       = client.open_by_key(SHEET_ID)
        ws_list  = sh.worksheets()
        ws_set   = next((w for w in ws_list if w.title == "護老_設定"), None)
        ws_cmd   = next((w for w in ws_list if w.title == "護老_指揮組"), None)
        ws_sch   = next((w for w in ws_list if w.title == "護老_勤務表"), None)
        df_set   = pd.DataFrame(ws_set.get_all_records()).fillna("") if ws_set else None
        df_cmd   = pd.DataFrame(ws_cmd.get_all_records()).fillna("") if ws_cmd else pd.DataFrame()
        df_sch   = pd.DataFrame(ws_sch.get_all_records()).fillna("") if ws_sch else pd.DataFrame()
        sd       = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1])) if (df_set is not None and not df_set.empty) else {}
        notes    = sd.get("notes", DEFAULT_NOTES)
        return df_set, df_cmd, df_sch, notes, sd, None
    except Exception as e:
        return None, None, None, DEFAULT_NOTES, {}, str(e)

def save_data(month, holidays, df_cmd, df_schedule, notes):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)

        try:
            ws_set = sh.worksheet("護老_設定")
        except Exception:
            ws_set = sh.add_worksheet(title="護老_設定", rows="50", cols="5")
        ws_set.clear()
        ws_set.update(range_name='A1', values=[
            ["Key", "Value"],
            ["month", month],
            ["holidays", holidays],
            ["notes", notes],
        ])

        for ws_title, df in [("護老_指揮組", df_cmd), ("護老_勤務表", df_schedule)]:
            try:
                ws = sh.worksheet(ws_title)
            except Exception:
                ws = sh.add_worksheet(title=ws_title, rows="100", cols="20")
            ws.clear()
            clean = df.dropna(how="all").fillna("")
            if not clean.empty:
                ws.update(range_name='A1', values=[clean.columns.tolist()] + clean_df_to_list(clean))

        load_data.clear()
        return True
    except Exception as e:
        st.error(f"❌ 雲端同步失敗：{e}")
        st.code(traceback.format_exc())
        return False

# --- PDF ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    for p in ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/custom/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf(month, df_cmd, df_schedule, notes_content):
    font = _get_font()
    buf  = io.BytesIO()
    doc  = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=12*mm, bottomMargin=18*mm)
    W    = A4[0] - 24*mm

    def draw_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 10)
        canvas.drawCentredString(A4[0]/2.0, 10*mm, f"第 {doc.page} 頁")
        canvas.restoreState()

    s_title = ParagraphStyle("t",  fontName=font, fontSize=16, alignment=1, spaceAfter=8, leading=22, wordWrap='CJK')
    s_th    = ParagraphStyle("th", fontName=font, fontSize=16, alignment=1, leading=22,   wordWrap='CJK')
    s_cell  = ParagraphStyle("c",  fontName=font, fontSize=14, leading=18,  alignment=1,  wordWrap='CJK')
    s_left  = ParagraphStyle("l",  fontName=font, fontSize=14, leading=18,  alignment=0,  wordWrap='CJK')
    s_note  = ParagraphStyle("n",  fontName=font, fontSize=12, leading=18,
                              leftIndent=8.5*mm, firstLineIndent=-8.5*mm, spaceAfter=4, wordWrap='CJK')

    def c(txt, style=s_cell):
        return Paragraph(str(txt).replace("\n", "<br/>"), style)

    story = []
    story.append(Paragraph(f"<b>{UNIT}{month}執行「行人及護老交通安全」專案勤務規劃表</b>", s_title))

    # 任務編組
    data1 = [
        [Paragraph("<b>任 務 編 組</b>", s_th), '', '', ''],
        [Paragraph(f"<b>{h}</b>", s_th) for h in ["職稱", "代號", "姓名", "任務"]]
    ]
    for _, row in df_cmd.iterrows():
        data1.append([
            c(f"<b>{row.get('職稱','')}</b>"),
            c(row.get('代號', '')),
            c(str(row.get('姓名', '')).replace("、", "<br/>")),
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
    story.append(Spacer(1, 6*mm))

    # 警力佈署
    # ✨ 修正：這裡的欄位總標題已帶入完整含有時間區段的變數名稱
    col_date = "執行勤務日期（6時至10時，16時至20時）"
    data2 = [
        [Paragraph("<b>警 力 佈 署</b>", s_th), '', ''],
        [Paragraph(f"<b>{col_date}</b>", s_th), Paragraph("<b>單位</b>", s_th), Paragraph("<b>路段</b>", s_th)]
    ]
    for _, row in df_schedule.iterrows():
        data2.append([
            c(row.get(col_date, '')),
            c(row.get('單位', '')),
            c(row.get('路段', ''), s_left)
        ])

    t2_styles = [
        ('FONTNAME',   (0,0), (-1,-1), font),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('SPAN',       (0,0), (-1,0)),
        ('BACKGROUND', (0,0), (-1,1),  colors.HexColor('#f2f2f2')),
    ]
    if not df_schedule.empty and col_date in df_schedule.columns:
        non_empty = [i for i, v in enumerate(df_schedule[col_date]) if str(v).strip() != ""]
        non_empty.append(len(df_schedule))
        for k in range(len(non_empty) - 1):
            s, e = non_empty[k], non_empty[k+1] - 1
            if e > s:
                t2_styles.append(('SPAN', (0, s+2), (0, e+2)))

    t2 = Table(data2, colWidths=[W*0.28, W*0.16, W*0.56], repeatRows=2)
    t2.setStyle(TableStyle(t2_styles))
    story.append(t2)
    story.append(Spacer(1, 6*mm))

    story.append(Paragraph("<b>備註：</b>", s_note))
    for line in notes_content.split('\n'):
        if line.strip():
            story.append(Paragraph(line, s_note))

    doc.build(story, onFirstPage=draw_page_number, onLaterPages=draw_page_number)
    return buf.getvalue()

# --- 寄信 ---
def send_report_email(subject, month, df_cmd, df_schedule, notes):
    try:
        sender = st.secrets["email"]["user"]
        pwd    = st.secrets["email"]["password"]
        pdf_bytes = generate_pdf(month, df_cmd, df_schedule, notes)
        msg = MIMEMultipart()
        msg["From"] = sender; msg["To"] = sender; msg["Subject"] = subject
        msg.attach(MIMEText("附件為最新的護老交通安全勤務規劃表 PDF。", "plain", "utf-8"))
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(subject)}.pdf")
        msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)


# ====================================================
# --- 主介面 ---
# ====================================================
st.title("🚶 行人及護老交通安全專案勤務規劃表")

if st.sidebar.button("🔄 強制重新載入"):
    st.cache_data.clear()
    if "current_holidays" in st.session_state:
        del st.session_state["current_holidays"]
    st.rerun()

df_set, df_cmd_raw, df_sch_raw, db_notes, sd, err = load_data()
if err and err != "權限不足或未設定 Secrets":
    st.warning(f"⚠️ 無法連線 Google Sheets（{err}），顯示預設資料。")


# ----------------------------------------------------
# 核心控制：Session State 月份與放假日連端緩存機制
# ----------------------------------------------------
init_month = sd.get("month", DEFAULT_MONTH)

# 初始化假日快取
if "current_holidays" not in st.session_state:
    if "holidays" in sd:
        st.session_state["current_holidays"] = sd.get("holidays")
    else:
         st.session_state["current_holidays"] = MONTH_HOLIDAYS_MAP.get(init_month, DEFAULT_HOLIDAYS)

# 當手動改變月份輸入框時觸發
def on_month_change():
    new_month = st.session_state["selected_month"]
    st.session_state["current_holidays"] = MONTH_HOLIDAYS_MAP.get(new_month, "")


# --- 基本設定 ---
st.subheader("1. 基本設定")
col_a, col_b = st.columns([1, 2])

# 月份欄位
c_month   = col_a.text_input("月份", value=init_month, key="selected_month", on_change=on_month_change)

# 假日欄位
c_holidays = col_b.text_input(
    "本月放假日（用逗號分隔，例：3/3, 3/10，補班日請用加號如 +3/21）",
    key="current_holidays",
    help="輸入國定假日、補假等放假日期。週六週日會自動排除不需填寫。若週六日需補班請在前面加個 + 號（如 +3/21）"
)

# 即時預覽上班日
if c_month:
    preview = generate_workday_label(c_month, c_holidays)
    if preview:
        st.caption(f"📅 本月上班日：{preview}")

full_header_name = f"{UNIT}{c_month}執行「行人及護老交通安全」專案勤務規劃表"

# 自動產生警力佈署
auto_sch = build_schedule_df(c_month, c_holidays)

# 假日沒變動且雲端有資料時，優先用雲端
saved_holidays = sd.get("holidays", "")
if (df_sch_raw is not None and not df_sch_raw.empty) and (saved_holidays == c_holidays) and (sd.get("month","") == c_month):
    use_sch = df_sch_raw
else:
    use_sch = auto_sch

# --- 任務編組 ---
st.subheader("2. 任務編組")
ed_cmd = df_cmd_raw if (df_cmd_raw is not None and not df_cmd_raw.empty) else DEFAULT_CMD.copy()
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True)

# --- 警力佈署 ---
st.subheader("3. 警力佈署")
st.caption("💡 修改上方月份或放假日後，日期欄會自動重新產生")
col_date = "執行勤務日期（6時至10時，16時至20時）"
res_sch = st.data_editor(
    use_sch,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        col_date: st.column_config.TextColumn(disabled=True),
    }
)

# --- 備註 ---
st.subheader("4. 備註編輯")
current_notes = db_notes if db_notes else DEFAULT_NOTES
ed_notes = st.text_area("編輯備註內容", value=current_notes, height=250)

st.markdown("---")

if st.button("💾 同步雲端並發送電子郵件備份", type="primary", use_container_width=True):
    with st.spinner("同步中，請稍候…"):
        res_cmd_clean = res_cmd.dropna(how="all").fillna("")
        res_sch_clean = res_sch.dropna(how="all").fillna("")
        if save_data(c_month, c_holidays, res_cmd_clean, res_sch_clean, ed_notes):
            with st.spinner("同步成功，正在寄送郵件…"):
                ok, mail_err = send_report_email(full_header_name, c_month, res_cmd_clean, res_sch_clean, ed_notes)
            if ok:
                st.success("📧 資料已同步至 Google Sheets，規劃表 PDF 已成功寄出！")
            else:
                st.error(f"❌ 雲端已更新，但寄信失敗：{mail_err}")
        else:
            st.error("❌ 雲端同步失敗，請檢查網路、Secrets 金鑰或試算表權限。")
