import streamlit as st
# 1. 頁面設定必須在最前面
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

from menu import show_sidebar
show_sidebar()

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import io
import os
import re
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
UNIT_TITLE = "桃園市政府警察局龍潭分局"

CHECKIN_POINTS = """1. 中油高原交流道站（龍源路2-20號）
2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）
3. 7-11龍潭佳園門市（中正路三坑段776號）
4. 旭日路三坑自然生態公園停車場
5. 旭日路與大溪區交界處"""

NOTES = """一、各編組執行前由帶班人員在駐地實施勤前教育。
二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。
三、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。
四、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"""

# --- 工具函式 ---
def normalize(s):
    return str(s).replace('\n', '').replace('\r', '').replace(' ', '').strip()

def is_blank(val):
    return normalize(val) in ["", "None", "nan"]

def format_staff_only(val):
    if pd.isna(val) or str(val).strip() in ["None", "nan", ""]:
        return ""
    s = str(val).replace('\\', '\n').replace('、', '\n').replace('\xa0', ' ')
    s = re.sub(r'(\d{2}[:：]?\d{0,2}\s*-\s*\d{2}[:：]?\d{0,2}[時]?[:：])\s*([^\n\s])', r'\1\n\2', s)
    return '\n'.join([l.strip() for l in s.split('\n') if l.strip()])

@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)
    except:
        return None

def load_from_cloud():
    client = get_client()
    if not client:
        return None, None, None
    try:
        sh = client.open_by_key(SHEET_ID)
        s = pd.DataFrame(sh.worksheet("危駕_設定").get_all_records())
        c = pd.DataFrame(sh.worksheet("危駕_指揮組").get_all_records())
        p = pd.DataFrame(sh.worksheet("危駕_警力佈署").get_all_records())
        return s, c, p
    except:
        return None, None, None

def save_to_cloud(p_time, cmdr, df_c, df_p):
    client = get_client()
    if not client:
        return False
    try:
        sh = client.open_by_key(SHEET_ID)
        sh.worksheet("危駕_設定").clear()
        sh.worksheet("危駕_設定").update(
            range_name='A1',
            values=[["Key", "Value"], ["plan_time", p_time], ["commander", cmdr]]
        )
        for name, df in [("危駕_指揮組", df_c), ("危駕_警力佈署", df_p)]:
            ws = sh.worksheet(name)
            ws.clear()
            data = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
            ws.update(range_name='A1', values=data)
        return True
    except:
        return False

# --- 日期解析與時段字串計算 ---
def calc_time_strings(p_time):
    """
    回傳 (第一列專用的隔日, 第二列以後專用的當日)
    """
    date_match = re.search(r'(?:(\d+)年)?(\d+)月(\d+)日(.*)', p_time)
    if not date_match:
        return "", ""

    y = date_match.group(1)
    m = int(date_match.group(2))
    d = int(date_match.group(3))
    
    # 抓取日期後面的時段字串 (例如 "22時至翌日6時")
    time_part = date_match.group(4).strip()
    if not time_part:
        time_part = "22時至翌日6時"

    y_tw = int(y) if y else (datetime.now().year - 1911)
    base_dt = datetime(y_tw + 1911, m, d)
    next_dt = base_dt + timedelta(days=1)

    # 第一列：勤務時間的隔日 零時至4時
    dedicated_time = f"{next_dt.month}月{next_dt.day}日\n零時至4時"
    # 第二列以後：勤務時間的當日 + 使用者輸入的時段
    normal_time = f"{m}月{d}日\n{time_part}"

    return dedicated_time, normal_time

# --- PDF 字型與產生 ---
def register_font():
    font_paths = [
        "kaiu.ttf",
        os.path.join(os.path.dirname(__file__), "kaiu.ttf"),
        "C:/Windows/Fonts/kaiu.ttf",
        "/usr/share/fonts/truetype/kaiu.ttf",
    ]
    for fp in font_paths:
        if os.path.exists(fp):
            pdfmetrics.registerFont(TTFont("BiauKai", fp))
            return True
    return False

FONT_AVAILABLE = register_font()
FONT_NAME = "BiauKai" if FONT_AVAILABLE else "Helvetica"

def generate_pdf(df_c, df_p, cmdr_n, time_s):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
    )
    styles = {
        "title": ParagraphStyle("title", fontName=FONT_NAME, fontSize=14, leading=20, alignment=1),
        "sub":   ParagraphStyle("sub",   fontName=FONT_NAME, fontSize=11, leading=16, alignment=1),
        "cell":  ParagraphStyle("cell",  fontName=FONT_NAME, fontSize=9,  leading=13),
        "note":  ParagraphStyle("note",  fontName=FONT_NAME, fontSize=8,  leading=12),
    }

    W = A4[0] - 30 * mm

    def make_table(data, col_widths, row_heights=None):
        t = Table(data, colWidths=col_widths, rowHeights=row_heights)
        t.setStyle(TableStyle([
            ("FONTNAME",      (0, 0), (-1, -1), FONT_NAME),
            ("FONTSIZE",      (0, 0), (-1, -1), 9),
            ("GRID",          (0, 0), (-1, -1), 0.5, colors.black),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING",    (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("BACKGROUND",    (0, 0), (-1, 0),  colors.lightgrey),
        ]))
        return t

    story = []
    story.append(Paragraph(UNIT_TITLE, styles["title"]))
    story.append(Paragraph("防制危險駕車專案勤務規劃表", styles["title"]))
    story.append(Spacer(1, 4 * mm))

    info_data = [[
        Paragraph(f"勤務時間：{time_s}", styles["cell"]),
        Paragraph(f"交通快打指揮官：{cmdr_n}", styles["cell"]),
    ]]
    story.append(make_table(info_data, [W * 0.55, W * 0.45], [16]))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph("【任務編組】", styles["sub"]))
    if len(df_c) > 0:
        cmd_header = [["職稱", "代號", "姓名", "任務"]]
        cmd_rows = [
            [Paragraph(str(r.get("職稱", "")), styles["cell"]),
             Paragraph(str(r.get("代號", "")), styles["cell"]),
             Paragraph(str(r.get("姓名", "")), styles["cell"]),
             Paragraph(str(r.get("任務", "")), styles["cell"])]
            for _, r in df_c.iterrows()
        ]
        story.append(make_table(cmd_header + cmd_rows, [W * 0.2, W * 0.15, W * 0.2, W * 0.45]))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph("【警力佈署】", styles["sub"]))
    if len(df_p) > 0:
        ptl_header = [["勤務時段", "代號", "編組", "服勤人員", "任務分工"]]
        ptl_rows = [
            [Paragraph(str(r.get("勤務時段", "")).replace('\n', '<br/>'), styles["cell"]),
             Paragraph(str(r.get("代號", "")),    styles["cell"]),
             Paragraph(str(r.get("編組", "")).replace('\n', '<br/>'),    styles["cell"]),
             Paragraph(format_staff_only(r.get("服勤人員", "")).replace('\n', '<br/>'), styles["cell"]),
             Paragraph(str(r.get("任務分工", "")).replace('\n', '<br/>'), styles["cell"])]
            for _, r in df_p.iterrows()
        ]
        story.append(make_table(ptl_header + ptl_rows, [W * 0.14, W * 0.1, W * 0.16, W * 0.3, W * 0.3]))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph("【打卡地點】", styles["sub"]))
    checkin_data = [[Paragraph(CHECKIN_POINTS.replace('\n', '<br/>'), styles["cell"])]]
    story.append(make_table(checkin_data, [W]))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph("【注意事項】", styles["sub"]))
    note_data = [[Paragraph(NOTES.replace('\n', '<br/>'), styles["note"])]]
    story.append(make_table(note_data, [W]))

    doc.build(story)
    buf.seek(0)
    return buf

# --- HTML 預覽 ---
def get_preview_html(df_c, df_p, cmdr_n, time_s):
    def df_to_html(df, cols):
        rows_html = ""
        for _, r in df.iterrows():
            cells = "".join(
                f"<td style='padding:4px 8px;border:1px solid #ccc;white-space:pre-wrap'>{str(r.get(c, ''))}</td>"
                for c in cols
            )
            rows_html += f"<tr>{cells}</tr>"
        header = "".join(f"<th style='padding:4px 8px;border:1px solid #ccc;background:#eee'>{c}</th>" for c in cols)
        return f"<table style='border-collapse:collapse;width:100%;font-size:12px'><thead><tr>{header}</tr></thead><tbody>{rows_html}</tbody></table>"

    return f"""
    <div style="font-family:serif;font-size:13px;padding:12px">
      <h3 style="text-align:center;margin:4px 0">{UNIT_TITLE}</h3>
      <h4 style="text-align:center;margin:4px 0">防制危險駕車專案勤務規劃表</h4>
      <p><b>勤務時間：</b>{time_s}　<b>指揮官：</b>{cmdr_n}</p>
      <p><b>【任務編組】</b></p>
      {df_to_html(df_c, ["職稱","代號","姓名","任務"])}
      <p><b>【警力佈署】</b></p>
      {df_to_html(df_p, ["勤務時段","代號","編組","服勤人員","任務分工"])}
    </div>
    """

# ============================================================
# 主介面
# ============================================================
st.title("🚔 防制危險駕車專案勤務規劃表")

# --- 初始化 Session State ---
if 'data_ptl' not in st.session_state:
    s, c, p = load_from_cloud()
    if s is not None:
        sd = dict(zip(s.iloc[:, 0].astype(str), s.iloc[:, 1].astype(str)))
        st.session_state.p_time   = sd.get("plan_time", "115年4月30日22時至翌日6時")
        st.session_state.cmdr     = sd.get("commander", "石門所副所長林榮裕")
        st.session_state.data_cmd = c
        st.session_state.data_ptl = p
    else:
        st.session_state.p_time   = "115年4月30日22時至翌日6時"
        st.session_state.cmdr     = "石門所副所長林榮裕"
        st.session_state.data_cmd = pd.DataFrame(columns=["職稱", "代號", "姓名", "任務"])
        st.session_state.data_ptl = pd.DataFrame([["", "", "", "", ""]], columns=["勤務時段", "代號", "編組", "服勤人員", "任務分工"])
    
    # 紀錄表格的初始行數，用來比對有沒有被點擊「新增一列」
    st.session_state.last_ptl_len = len(st.session_state.data_ptl)

col1, col2 = st.columns(2)
with col1:
    p_time = st.text_input("1. 勤務時間", st.session_state.p_time)
with col2:
    cmdr_input = st.text_input("2. 交通快打指揮官", st.session_state.cmdr)

# 計算兩種時段字串
dedicated_time, normal_time = calc_time_strings(p_time)

# --- 任務編組 ---
st.subheader("3. 任務編組")
res_cmd = st.data_editor(st.session_state.data_cmd, num_rows="dynamic", use_container_width=True).fillna("")

# --- 警力佈署 ---
st.subheader("4. 警力佈署")

# 表格渲染
res_ptl_raw = st.data_editor(st.session_state.data_ptl, num_rows="dynamic", use_container_width=True).fillna("")

# ============================================================
# 核心修正：行數偵測攔截 (打擊 Streamlit 自動數字 +1 的 Bug)
# ============================================================
current_len = len(res_ptl_raw)
needs_rerun = False

if current_len > st.session_state.last_ptl_len:
    # 代表使用者剛剛點擊了「新增一列」
    # 強制將剛新增的所有列覆蓋為「當日 22時至翌日6時」
    for i in range(st.session_state.last_ptl_len, current_len):
        res_ptl_raw.at[i, '勤務時段'] = normal_time
        res_ptl_raw.at[i, '服勤人員'] = ""  # 順便清空姓名，避免複製到上一列的人員
    
    st.session_state.data_ptl = res_ptl_raw
    st.session_state.last_ptl_len = current_len
    needs_rerun = True

elif current_len < st.session_state.last_ptl_len:
    # 代表使用者刪除了列
    st.session_state.data_ptl = res_ptl_raw
    st.session_state.last_ptl_len = current_len

else:
    # 沒有新增也沒有刪除，正常更新資料
    st.session_state.data_ptl = res_ptl_raw

# 針對第一列：如果是空白的，給予專屬的隔日零時預設值
if len(st.session_state.data_ptl) > 0 and is_blank(st.session_state.data_ptl.at[0, '勤務時段']):
    st.session_state.data_ptl.at[0, '勤務時段'] = dedicated_time
    # 自動給予預設代號與編組
    unit_base = "隆安8" if "石門" in cmdr_input else "隆安6" if "龍潭" in cmdr_input else "隆安"
    st.session_state.data_ptl.at[0, '代號'] = unit_base + ("1" if "所長" in cmdr_input and "副" not in cmdr_input else "2")
    st.session_state.data_ptl.at[0, '編組'] = f"專責警力\n（{cmdr_input[:3]}輪值）"
    needs_rerun = True

# 只要有修正過錯誤的預設值，立刻刷新畫面
if needs_rerun:
    st.rerun()

res_ptl = st.session_state.data_ptl
st.session_state.p_time = p_time
st.session_state.cmdr = cmdr_input
st.session_state.data_cmd = res_cmd

# --- 預覽 ---
st.markdown("---")
with st.expander("📄 預覽勤務規劃表"):
    html_preview = get_preview_html(res_cmd, res_ptl, cmdr_input, p_time)
    st.components.v1.html(html_preview, height=600, scrolling=True)

# --- 下載 PDF ＆ 同步雲端 ---
col_pdf, col_sync = st.columns(2)

with col_pdf:
    if st.button("📥 產生並下載 PDF"):
        with st.spinner("產生 PDF 中..."):
            pdf_buf = generate_pdf(res_cmd, res_ptl, cmdr_input, p_time)
            st.download_button(
                label="點此下載 PDF",
                data=pdf_buf,
                file_name=f"危駕勤務規劃_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
            )

with col_sync:
    if st.button("💾 同步至雲端", type="primary"):
        with st.spinner("同步中..."):
            if save_to_cloud(p_time, cmdr_input, res_cmd, res_ptl):
                st.success("✅ 雲端同步成功！")
            else:
                st.error("❌ 同步失敗。請檢查 Google Sheets 權限或 Secrets 設定。")
