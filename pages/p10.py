import streamlit as st
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    pass

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import io, os, re, smtplib
import urllib.parse as _ul

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# =========================
# 基本設定
# =========================
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
WS_MAP = {
    "set": "危駕_設定",
    "cmd": "危駕_指揮組",
    "ptl": "危駕_警力佈署"
}
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
UNIT_TITLE = "桃園市政府警察局龍潭分局"

CMD_COLS = ["職稱", "代號", "姓名", "任務"]
PTL_COLS = ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]

# =========================
# 字體
# =========================
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames():
        return fname
    font_paths = [
        "./kaiu.ttf", "kaiu.ttf",
        "/usr/share/fonts/truetype/custom/kaiu.ttf",
        "C:/Windows/Fonts/kaiu.ttf",
    ]
    for p in font_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

# =========================
# Google Sheets 連線
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

# =========================
# Sheets 初始化
# =========================
def init_sheets():
    client = get_client()
    if client is None:
        return
    sh = client.open_by_key(SHEET_ID)
    headers = {
        WS_MAP["set"]: [["Key", "Value"]],
        WS_MAP["cmd"]: [CMD_COLS],
        WS_MAP["ptl"]: [PTL_COLS],
    }
    for name, header in headers.items():
        try:
            sh.worksheet(name)
        except Exception:
            sh.add_worksheet(title=name, rows="100", cols="20").update(header)
    st.success("初始化完成")
    st.cache_data.clear()
    st.rerun()

# =========================
# 讀取資料
# =========================
@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        if client is None:
            return pd.DataFrame(), pd.DataFrame(columns=CMD_COLS), pd.DataFrame(columns=PTL_COLS), "", "授權失敗"
        sh = client.open_by_key(SHEET_ID)

        try:
            set_df = pd.DataFrame(sh.worksheet(WS_MAP["set"]).get_all_records()).fillna("")
        except Exception:
            set_df = pd.DataFrame()

        try:
            cmd_df = pd.DataFrame(sh.worksheet(WS_MAP["cmd"]).get_all_records()).fillna("")
        except Exception:
            cmd_df = pd.DataFrame(columns=CMD_COLS)

        try:
            ptl_df = pd.DataFrame(sh.worksheet(WS_MAP["ptl"]).get_all_records()).fillna("")
        except Exception:
            ptl_df = pd.DataFrame(columns=PTL_COLS)

        time_val = ""
        if not set_df.empty and set_df.shape[1] >= 2:
            d = dict(zip(set_df.iloc[:, 0].astype(str), set_df.iloc[:, 1].astype(str)))
            time_val = d.get("time", "")

        return set_df, cmd_df, ptl_df, time_val, None
    except Exception as e:
        return pd.DataFrame(), pd.DataFrame(columns=CMD_COLS), pd.DataFrame(columns=PTL_COLS), "", str(e)

# =========================
# 儲存資料
# =========================
def save_data(time_str, cmd, ptl):
    try:
        client = get_client()
        if client is None:
            return False
        sh = client.open_by_key(SHEET_ID)

        ws_set = sh.worksheet(WS_MAP["set"])
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["time", time_str]])

        for ws_name, df in [(WS_MAP["cmd"], cmd), (WS_MAP["ptl"], ptl)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df_clean = df.dropna(how="all").fillna("")
            if not df_clean.empty:
                ws.update([df_clean.columns.tolist()] + df_clean.values.tolist())

        load_data.clear()
        return True
    except Exception as e:
        st.error(f"❌ 儲存失敗：{e}")
        return False

# =========================
# PDF 產生
# =========================
def generate_pdf(time_str, cmd_df, ptl_df):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=12*mm, rightMargin=12*mm,
        topMargin=12*mm, bottomMargin=12*mm,
    )
    W = A4[0] - 24*mm
    story = []

    s_title = ParagraphStyle("title", fontName=font, fontSize=16, alignment=1, leading=24, spaceAfter=6, wordWrap="CJK")
    s_th    = ParagraphStyle("th",    fontName=font, fontSize=14, alignment=1, leading=20, wordWrap="CJK")
    s_cell  = ParagraphStyle("cell",  fontName=font, fontSize=13, alignment=1, leading=18, wordWrap="CJK")
    s_left  = ParagraphStyle("left",  fontName=font, fontSize=13, alignment=0, leading=18, wordWrap="CJK")

    def c(txt, style=s_cell):
        return Paragraph(str(txt).replace("\n", "<br/>"), style)

    story.append(Paragraph(f"<b>{UNIT_TITLE}防制危險駕車勤務規劃表</b>", s_title))
    story.append(Paragraph(f"勤務時間：{time_str}", s_th))
    story.append(Spacer(1, 4*mm))

    # ── 指揮組表格（防空資料列不足造成 ValueError）──
    cmd_clean = cmd_df.dropna(how="all").fillna("")
    if cmd_clean.empty:
        # 補一列空白，避免 ReportLab Table 報錯
        cmd_clean = pd.DataFrame([{col: "" for col in CMD_COLS}])

    data_cmd = [[Paragraph("<b>任 務 編 組</b>", s_th), "", "", ""],
                [Paragraph(f"<b>{h}</b>", s_th) for h in CMD_COLS]]
    for _, row in cmd_clean.iterrows():
        data_cmd.append([
            c(f"<b>{row.get('職稱', '')}</b>"),
            c(row.get("代號", "")),
            c(str(row.get("姓名", "")).replace("、", "<br/>")),
            c(row.get("任務", ""), s_left),
        ])
    t1 = Table(data_cmd, colWidths=[W*0.15, W*0.12, W*0.28, W*0.45], repeatRows=2)
    t1.setStyle(TableStyle([
        ("FONTNAME",   (0,0), (-1,-1), font),
        ("GRID",       (0,0), (-1,-1), 0.5, colors.black),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("SPAN",       (0,0), (-1,0)),
        ("BACKGROUND", (0,0), (-1,1), colors.HexColor("#f2f2f2")),
    ]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    # ── 警力佈署表格 ──
    ptl_clean = ptl_df.dropna(how="all").fillna("")
    if ptl_clean.empty:
        ptl_clean = pd.DataFrame([{col: "" for col in PTL_COLS}])

    data_ptl = [[Paragraph("<b>警 力 佈 署</b>", s_th), "", "", "", ""],
                [Paragraph(f"<b>{h}</b>", s_th) for h in PTL_COLS]]
    for _, row in ptl_clean.iterrows():
        data_ptl.append([
            c(row.get("勤務時段", "")),
            c(row.get("代號", "")),
            c(row.get("編組", "")),
            c(str(row.get("服勤人員", "")).replace("、", "<br/>")),
            c(row.get("任務分工", ""), s_left),
        ])
    t2 = Table(data_ptl, colWidths=[W*0.20, W*0.12, W*0.13, W*0.20, W*0.35], repeatRows=2)
    t2.setStyle(TableStyle([
        ("FONTNAME",   (0,0), (-1,-1), font),
        ("GRID",       (0,0), (-1,-1), 0.5, colors.black),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("SPAN",       (0,0), (-1,0)),
        ("BACKGROUND", (0,0), (-1,1), colors.HexColor("#e6e6e6")),
    ]))
    story.append(t2)

    doc.build(story)
    buf.seek(0)
    return buf

# =========================
# Email
# =========================
def send_email(time_str, cmd_df, ptl_df, filename):
    try:
        sender = st.secrets["email"]["user"]
        pwd    = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"]    = sender
        msg["To"]      = sender
        msg["Subject"] = filename
        msg.attach(MIMEText("附件為最新勤務規劃表。", "plain", "utf-8"))

        pdf = generate_pdf(time_str, cmd_df, ptl_df)
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf.read())
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
# UI
# =========================
st.title("🚔 防制危險駕車勤務規劃")

if st.sidebar.button("🔧 初始化工作表"):
    init_sheets()

if st.sidebar.button("🔄 強制更新資料"):
    st.cache_data.clear()
    st.rerun()

set_df, cmd_df, ptl_df, saved_time, err = load_data()

if err:
    st.warning(f"⚠️ 無法連線 Google Sheets（{err}），顯示預設資料。")

# 確保欄位結構正確
if cmd_df.empty or not all(c in cmd_df.columns for c in CMD_COLS):
    cmd_df = pd.DataFrame(columns=CMD_COLS)
if ptl_df.empty or not all(c in ptl_df.columns for c in PTL_COLS):
    ptl_df = pd.DataFrame(columns=PTL_COLS)

time_val = st.text_input("勤務時間", value=saved_time or "22時至翌日6時")

st.subheader("1. 任務編組")
res_cmd = st.data_editor(cmd_df, num_rows="dynamic", use_container_width=True).dropna(how="all").fillna("")

st.subheader("2. 警力佈署")
res_ptl = st.data_editor(ptl_df, num_rows="dynamic", use_container_width=True).dropna(how="all").fillna("")

st.markdown("---")
col1, col2, col3 = st.columns(3)

with col1:
    if st.button("💾 儲存至雲端", use_container_width=True):
        if save_data(time_val, res_cmd, res_ptl):
            st.success("✅ 已儲存")
        else:
            st.error("❌ 儲存失敗")

with col2:
    pdf_buf = generate_pdf(time_val, res_cmd, res_ptl)
    st.download_button(
        "📄 下載 PDF",
        data=pdf_buf,
        file_name=f"{UNIT_TITLE}防制危險駕車勤務規劃表.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

with col3:
    if st.button("📧 發送 Email", use_container_width=True):
        ok, mail_err = send_email(time_val, res_cmd, res_ptl, f"{UNIT_TITLE}防制危險駕車勤務規劃表")
        if ok:
            st.success("✅ 已寄出")
        else:
            st.error(f"❌ 發送失敗：{mail_err}")
