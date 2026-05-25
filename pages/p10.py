import streamlit as st
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import io, os, re, smtplib, urllib.parse as _ul

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

# =========================
# 安全 PEM 修復（重點）
# =========================
def clean_private_key(pk: str) -> str:
    pk = pk.replace("\\n", "\n")

    if "-----BEGIN PRIVATE KEY-----" not in pk:
        return pk

    body = pk.split("-----BEGIN PRIVATE KEY-----")[-1].split("-----END PRIVATE KEY-----")[0]
    body = re.sub(r'[\s\r\n"\']', '', body)
    body = "\n".join([body[i:i+64] for i in range(0, len(body), 64)])

    return f"-----BEGIN PRIVATE KEY-----\n{body}\n-----END PRIVATE KEY-----\n"


# =========================
# Google Sheets Client
# =========================
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        return None

    info = dict(st.secrets["gcp_service_account"])

    if "private_key" in info:
        info["private_key"] = clean_private_key(info["private_key"])

    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


# =========================
# Sheets 初始化
# =========================
def init_sheets():
    client = get_client()
    sh = client.open_by_key(SHEET_ID)

    headers = {
        WS_MAP["set"]: [["Key", "Value"]],
        WS_MAP["cmd"]: [["職稱", "代號", "姓名", "任務"]],
        WS_MAP["ptl"]: [["勤務時段", "代號", "編組", "服勤人員", "任務分工"]]
    }

    for name, header in headers.items():
        try:
            sh.worksheet(name)
        except:
            sh.add_worksheet(title=name, rows="100", cols="20").update(header)

    st.success("初始化完成")
    st.cache_data.clear()
    st.rerun()


# =========================
# 讀寫資料
# =========================
def load_data():
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)

        set_ws = pd.DataFrame(sh.worksheet(WS_MAP["set"]).get_all_records())
        cmd_ws = pd.DataFrame(sh.worksheet(WS_MAP["cmd"]).get_all_records())
        ptl_ws = pd.DataFrame(sh.worksheet(WS_MAP["ptl"]).get_all_records())

        return set_ws.fillna(""), cmd_ws.fillna(""), ptl_ws.fillna("")
    except:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


def save_data(time_str, cmd, ptl):
    client = get_client()
    sh = client.open_by_key(SHEET_ID)

    set_ws = sh.worksheet(WS_MAP["set"])
    set_ws.clear()
    set_ws.update([
        ["Key", "Value"],
        ["time", time_str]
    ])

    for name, df in [(WS_MAP["cmd"], cmd), (WS_MAP["ptl"], ptl)]:
        ws = sh.worksheet(name)
        ws.clear()
        ws.update([df.columns.tolist()] + df.values.tolist())


# =========================
# PDF
# =========================
def _font():
    name = "kaiu"
    if name in pdfmetrics.getRegisteredFontNames():
        return name
    return "Helvetica"


def generate_pdf(time_str, cmd, ptl):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4)

    font = _font()

    story = []

    style = ParagraphStyle(name="s", fontName=font, fontSize=12)

    story.append(Paragraph(f"{UNIT_TITLE} 勤務規劃表", style))
    story.append(Paragraph(f"時間：{time_str}", style))
    story.append(Spacer(1, 10))

    cmd_data = [cmd.columns.tolist()] + cmd.values.tolist()
    ptl_data = [ptl.columns.tolist()] + ptl.values.tolist()

    t1 = Table(cmd_data)
    t1.setStyle(TableStyle([("GRID", (0,0), (-1,-1), 0.5, colors.black)]))

    t2 = Table(ptl_data)
    t2.setStyle(TableStyle([("GRID", (0,0), (-1,-1), 0.5, colors.black)]))

    story.append(t1)
    story.append(Spacer(1, 10))
    story.append(t2)

    doc.build(story)
    buf.seek(0)
    return buf


# =========================
# Email
# =========================
def send_email(time_str, cmd, ptl, filename):
    sender = st.secrets["email"]["user"]
    pwd = st.secrets["email"]["password"]

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = sender
    msg["Subject"] = filename

    msg.attach(MIMEText("附件勤務表", "plain"))

    pdf = generate_pdf(time_str, cmd, ptl)

    part = MIMEBase("application", "pdf")
    part.set_payload(pdf.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={filename}.pdf")

    msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(sender, pwd)
        s.send_message(msg)


# =========================
# UI
# =========================
st.title("🚔 防制危險駕車勤務")

if st.sidebar.button("初始化"):
    init_sheets()

set_df, cmd_df, ptl_df = load_data()

time_val = st.text_input("勤務時間", "22時至翌日6時")

cmd_df = st.data_editor(cmd_df, num_rows="dynamic")
ptl_df = st.data_editor(ptl_df, num_rows="dynamic")

col1, col2 = st.columns(2)

with col1:
    if st.button("💾 儲存"):
        save_data(time_val, cmd_df, ptl_df)
        st.success("已儲存")

with col2:
    pdf = generate_pdf(time_val, cmd_df, ptl_df)
    st.download_button("📄 下載PDF", pdf, file_name="勤務表.pdf")

if st.button("📧 發送Email"):
    send_email(time_val, cmd_df, ptl_df, "勤務表")
    st.success("已寄出")
