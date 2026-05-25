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

DEFAULT_SIGN_POINTS = (
    "巡簽地點：\n"
    "1. 中油高原交流道站（龍源路2-20號）\n"
    "2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）\n"
    "3. 7-11龍潭佳園門市（中正路三坑段776號）\n"
    "4. 旭日路三坑自然生態公園停車場\n"
    "5. 旭日路與大溪區交界處"
)

DEFAULT_NOTES = (
    "一、各編組執行前由帶班人員在駐地實施勤前教育。\n"
    "二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。\n"
    "三、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。\n"
    "四、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"
)

# =========================
# 字體
# =========================
@st.cache_resource
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
            return pd.DataFrame(), pd.DataFrame(columns=CMD_COLS), pd.DataFrame(columns=PTL_COLS), {}, "授權失敗"
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

        settings = {}
        if not set_df.empty and set_df.shape[1] >= 2:
            settings = dict(zip(set_df.iloc[:, 0].astype(str), set_df.iloc[:, 1].astype(str)))

        return set_df, cmd_df, ptl_df, settings, None
    except Exception as e:
        return pd.DataFrame(), pd.DataFrame(columns=CMD_COLS), pd.DataFrame(columns=PTL_COLS), {}, str(e)

# =========================
# 儲存資料
# =========================
def save_data(settings_dict, cmd, ptl):
    try:
        client = get_client()
        if client is None:
            return False
        sh = client.open_by_key(SHEET_ID)

        ws_set = sh.worksheet(WS_MAP["set"])
        ws_set.clear()
        rows = [["Key", "Value"]] + [[k, v] for k, v in settings_dict.items()]
        ws_set.update(rows)

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
def generate_pdf(time_str, project_name, fast_cmd, cmd_df, ptl_df, sign_points, notes):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=12*mm, rightMargin=12*mm,
        topMargin=12*mm, bottomMargin=15*mm,
    )
    W = A4[0] - 24*mm
    story = []

    s_title  = ParagraphStyle("title", fontName=font, fontSize=15, alignment=1,  leading=22, spaceAfter=4,    wordWrap="CJK")
    s_sub    = ParagraphStyle("sub",   fontName=font, fontSize=13, alignment=1,  leading=20, spaceAfter=4,    wordWrap="CJK")
    s_th     = ParagraphStyle("th",    fontName=font, fontSize=13, alignment=1,  leading=18,                  wordWrap="CJK")
    s_cell   = ParagraphStyle("cell",  fontName=font, fontSize=12, alignment=1,  leading=17,                  wordWrap="CJK")
    s_left   = ParagraphStyle("left",  fontName=font, fontSize=12, alignment=0,  leading=17,                  wordWrap="CJK")
    s_note   = ParagraphStyle("note",  fontName=font, fontSize=11, alignment=0,  leading=16, spaceBefore=2,   wordWrap="CJK",
                               leftIndent=10, firstLineIndent=-10)

    def c(txt, style=s_cell):
        return Paragraph(str(txt).replace("\n", "<br/>"), style)

    # ── 標題
    story.append(Paragraph(
        f"<b>{UNIT_TITLE}執行「{project_name}」規劃表</b>", s_title))
    story.append(Paragraph(f"勤務時間：{time_str}", s_sub))
    story.append(Spacer(1, 3*mm))

    # ── 任務編組表
    cmd_clean = cmd_df.dropna(how="all").fillna("")
    if cmd_clean.empty:
        cmd_clean = pd.DataFrame([{col: "" for col in CMD_COLS}])

    data_cmd = [
        [Paragraph("<b>任 務 編 組</b>", s_th), "", "", ""],
        [Paragraph(f"<b>{h}</b>", s_th) for h in ["職稱", "代號", "姓名", "任務"]],
    ]
    for _, row in cmd_clean.iterrows():
        data_cmd.append([
            c(f"<b>{row.get('職稱', '')}</b>"),
            c(row.get("代號", "")),
            c(str(row.get("姓名", "")).replace("、", "<br/>")),
            c(row.get("任務", ""), s_left),
        ])
    t_cmd = Table(data_cmd, colWidths=[W*0.13, W*0.11, W*0.25, W*0.51], repeatRows=2)
    t_cmd.setStyle(TableStyle([
        ("FONTNAME",   (0,0), (-1,-1), font),
        ("GRID",       (0,0), (-1,-1), 0.5, colors.black),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("SPAN",       (0,0), (-1,0)),
        ("BACKGROUND", (0,0), (-1,1), colors.HexColor("#f2f2f2")),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    story.append(t_cmd)
    story.append(Spacer(1, 4*mm))

    # ── 快打指揮官（獨立一行）
    if fast_cmd.strip():
        story.append(Paragraph(f"交通快打指揮官：{fast_cmd}", s_sub))
        story.append(Spacer(1, 2*mm))

    # ── 警力佈署表
    ptl_clean = ptl_df.dropna(how="all").fillna("")
    if ptl_clean.empty:
        ptl_clean = pd.DataFrame([{col: "" for col in PTL_COLS}])

    data_ptl = [
        [Paragraph("<b>警 力 佈 署</b>", s_th), "", "", "", ""],
        [Paragraph(f"<b>{h}</b>", s_th) for h in ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]],
    ]

    # 同一勤務時段合併第一欄
    ptl_rows = ptl_clean.reset_index(drop=True)
    merge_groups = []
    prev_val, grp_start = None, 1
    for i, row in ptl_rows.iterrows():
        val = str(row.get("勤務時段", "")).strip()
        tbl_row = i + 2  # header 佔 2 行
        if val != prev_val:
            if prev_val is not None:
                merge_groups.append((grp_start, tbl_row - 1))
            prev_val = val
            grp_start = tbl_row
        data_ptl.append([
            c(row.get("勤務時段", "")),
            c(row.get("代號", "")),
            c(row.get("編組", "")),
            c(str(row.get("服勤人員", "")).replace("、", "<br/>")),
            c(row.get("任務分工", ""), s_left),
        ])
    if prev_val is not None:
        merge_groups.append((grp_start, len(ptl_rows) + 1))

    ts_ptl = [
        ("FONTNAME",   (0,0), (-1,-1), font),
        ("GRID",       (0,0), (-1,-1), 0.5, colors.black),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("SPAN",       (0,0), (-1,0)),
        ("BACKGROUND", (0,0), (-1,1), colors.HexColor("#e6e6e6")),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]
    for (rs, re) in merge_groups:
        if re > rs:
            ts_ptl.append(("SPAN", (0, rs), (0, re)))

    t_ptl = Table(data_ptl, colWidths=[W*0.18, W*0.10, W*0.12, W*0.18, W*0.42], repeatRows=2)
    t_ptl.setStyle(TableStyle(ts_ptl))
    story.append(t_ptl)
    story.append(Spacer(1, 4*mm))

    # ── 巡簽地點
    if sign_points.strip():
        for line in sign_points.strip().split("\n"):
            if line.strip():
                story.append(Paragraph(line.strip(), s_note))
        story.append(Spacer(1, 3*mm))

    # ── 備註
    story.append(Paragraph("<b>備註：</b>", s_sub))
    for line in notes.strip().split("\n"):
        if line.strip():
            story.append(Paragraph(line.strip(), s_note))

    # ── 頁碼
    def add_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 10)
        canvas.drawCentredString(A4[0] / 2, 8*mm, f"- {canvas.getPageNumber()} -")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    buf.seek(0)
    return buf

# =========================
# Email
# =========================
def send_email(subject, pdf_buf, filename):
    try:
        sender = st.secrets["email"]["user"]
        pwd    = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"]    = sender
        msg["To"]      = sender
        msg["Subject"] = subject
        msg.attach(MIMEText("附件為最新勤務規劃表。", "plain", "utf-8"))
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
# UI
# =========================
st.title("🚔 防制危險駕車勤務規劃")

if st.sidebar.button("🔧 初始化工作表"):
    init_sheets()
if st.sidebar.button("🔄 強制更新資料"):
    st.cache_data.clear()
    st.rerun()

set_df, cmd_df, ptl_df, settings, err = load_data()
if err:
    st.warning(f"⚠️ 無法連線 Google Sheets（{err}），顯示預設資料。")

# 確保欄位結構
if cmd_df.empty or not all(c in cmd_df.columns for c in CMD_COLS):
    cmd_df = pd.DataFrame(columns=CMD_COLS)
if ptl_df.empty or not all(c in ptl_df.columns for c in PTL_COLS):
    ptl_df = pd.DataFrame(columns=PTL_COLS)

# ── 基本設定區
st.subheader("基本設定")
col_a, col_b = st.columns(2)
project_name = col_a.text_input("專案名稱", value=settings.get("project_name", "防制危險駕車專案勤務"))
time_val     = col_b.text_input("勤務時間", value=settings.get("time", "115年5月22日22時至翌日6時"))
fast_cmd     = st.text_input("交通快打指揮官", value=settings.get("fast_cmd", "龍潭所副所長全楚文"))

# ── 編組資料
st.subheader("1. 任務編組")
res_cmd = st.data_editor(cmd_df, num_rows="dynamic", use_container_width=True).dropna(how="all").fillna("")

st.subheader("2. 警力佈署")
st.caption("💡 相同勤務時段請填相同文字，PDF 輸出時第一欄會自動合併。")
res_ptl = st.data_editor(ptl_df, num_rows="dynamic", use_container_width=True).dropna(how="all").fillna("")

# ── 巡簽地點與備註
st.subheader("3. 巡簽地點與備註")
col_c, col_d = st.columns(2)
sign_points = col_c.text_area("巡簽地點", value=settings.get("sign_points", DEFAULT_SIGN_POINTS), height=160)
notes       = col_d.text_area("備註", value=settings.get("notes", DEFAULT_NOTES), height=160)

st.markdown("---")
col1, col2, col3 = st.columns(3)

# ── 儲存
with col1:
    if st.button("💾 儲存至雲端", use_container_width=True):
        s = {
            "project_name": project_name,
            "time":         time_val,
            "fast_cmd":     fast_cmd,
            "sign_points":  sign_points,
            "notes":        notes,
        }
        if save_data(s, res_cmd, res_ptl):
            st.success("✅ 已儲存")

# ── 下載 PDF
with col2:
    pdf_buf = generate_pdf(time_val, project_name, fast_cmd, res_cmd, res_ptl, sign_points, notes)
    filename = f"{UNIT_TITLE}執行「{project_name}」規劃表"
    st.download_button(
        "📄 下載 PDF",
        data=pdf_buf,
        file_name=f"{filename}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

# ── 發送 Email
with col3:
    if st.button("📧 發送 Email", use_container_width=True):
        pdf_buf2 = generate_pdf(time_val, project_name, fast_cmd, res_cmd, res_ptl, sign_points, notes)
        ok, mail_err = send_email(filename, pdf_buf2, filename)
        if ok:
            st.success("✅ 已寄出")
        else:
            st.error(f"❌ 發送失敗：{mail_err}")
