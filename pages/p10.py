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
import io, os, smtplib
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

# =========================
# 基本設定
# =========================
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
WS_MAP = {"set": "危駕_設定", "cmd": "危駕_指揮組", "ptl": "危駕_警力佈署"}
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
UNIT_TITLE = "桃園市政府警察局龍潭分局"
CMD_COLS = ["職稱", "代號", "姓名", "任務"]
PTL_COLS = ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]

# =========================
# 預設底稿
# =========================
DEFAULT_PROJECT = "防制危險駕車專案勤務"
DEFAULT_TIME = "115年5月22日22時至翌日6時"
DEFAULT_FAST_CMD = "龍潭所副所長全楚文"

DEFAULT_CMD = pd.DataFrame([ ... ])  # 保持不變，省略以節省篇幅

DEFAULT_PTL = pd.DataFrame([
    {"勤務時段": "5月23日\n零時至4時", "代號": "隆安62", "編組": "專責警力\n（龍潭所輪值）", 
     "服勤人員": "00-02時段：\n警員廖怡惠\n警員劉柏延\n02-04時段：\n警員林軒宇\n警員廖怡惠", 
     "任務分工": "「加強防制」勤務，在文化路、中正路三坑段、龍源路及旭日路來回巡邏，隨機攔檢改裝（噪音）車輛（每2小時至責任區域內指定巡簽地點巡簽1次並守望10分鐘，將守望情形拍照上傳LINE「龍潭分局聯絡平臺」群組）"},
    
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安80", "編組": "石門所線上巡邏組合警力兼任", "服勤人員": "", 
     "任務分工": "「區域聯防」勤務，於中正路、文化路、中豐路、龍源路及旭日路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安90", "編組": "高平所線上巡邏組合警力兼任", "服勤人員": "", 
     "任務分工": "「區域聯防」勤務，於中豐路及龍源路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安990", "編組": "龍潭交通分隊線上巡邏組合警力兼任", "服勤人員": "", 
     "任務分工": "「區域聯防」勤務，於龍源路及旭日路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安50", "編組": "聖亭所線上巡邏組合警力兼任", "服勤人員": "", 
     "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安60", "編組": "龍潭所線上巡邏組合警力兼任", "服勤人員": "", 
     "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安70", "編組": "中興所線上巡邏組合警力兼任", "服勤人員": "", 
     "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
])

DEFAULT_SIGN_POINTS = "巡簽地點：\n1. 中油高原交流道站（龍源路2-20號）\n2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）\n3. 7-11龍潭佳園門市（中正路三坑段776號）\n4. 旭日路三坑自然生態公園停車場\n5. 旭日路與大溪區交界處"
DEFAULT_NOTES = "一、各編組執行前由帶班人員在駐地實施勤前教育。\n二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。\n三、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。\n四、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"

# =========================
# 字體與 Google Sheets 函數（略，保持之前版本）
# =========================
# ... (字體、get_client、init_sheets、load_data、save_data 保持不變) ...

# =========================
# PDF 生成 - 最終穩定版
# =========================
def generate_pdf(time_str, project_name, fast_cmd, cmd_df, ptl_df, sign_points, notes):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=12*mm, bottomMargin=15*mm)
    W = A4[0] - 24*mm
    story = []

    s_title = ParagraphStyle("title", fontName=font, fontSize=15, alignment=1, leading=22, spaceAfter=4, wordWrap="CJK")
    s_sub = ParagraphStyle("sub", fontName=font, fontSize=13, alignment=1, leading=20, spaceAfter=4, wordWrap="CJK")
    s_th = ParagraphStyle("th", fontName=font, fontSize=13, alignment=1, leading=18, wordWrap="CJK")
    s_cell = ParagraphStyle("cell", fontName=font, fontSize=12, alignment=1, leading=17, wordWrap="CJK")
    s_left = ParagraphStyle("left", fontName=font, fontSize=12, alignment=0, leading=17, wordWrap="CJK")
    s_note = ParagraphStyle("note", fontName=font, fontSize=11, alignment=0, leading=16, spaceBefore=2, wordWrap="CJK", leftIndent=10, firstLineIndent=-10)

    def c(txt, style=s_cell):
        return Paragraph(str(txt).replace("\n", "<br/>"), style)

    # 標題與任務編組（略）

    # ==================== 警力佈署 ====================
    ptl_clean = ptl_df.copy()
    # 強制清理：保留所有有勤務時段的行
    ptl_clean = ptl_clean[ptl_clean["勤務時段"].astype(str).str.strip() != ""].fillna("")
    
    if len(ptl_clean) == 0:
        ptl_clean = DEFAULT_PTL.copy()

    data_ptl = [
        [Paragraph("<b>警 力 佈 署</b>", s_th), "", "", "", ""],
        [Paragraph(f"<b>{h}</b>", s_th) for h in PTL_COLS]
    ]

    for _, row in ptl_clean.iterrows():
        data_ptl.append([
            c(row.get("勤務時段", "")),
            c(row.get("代號", "")),
            c(row.get("編組", "")),
            c(str(row.get("服勤人員", "")).replace("、", "<br/>")),
            c(row.get("任務分工", ""), s_left)
        ])

    # 安全合併邏輯
    merge_groups = []
    if len(ptl_clean) > 0:
        i = 0
        while i < len(ptl_clean):
            j = i + 1
            curr_time = str(ptl_clean.iloc[i]["勤務時段"]).strip()
            while j < len(ptl_clean) and str(ptl_clean.iloc[j]["勤務時段"]).strip() == curr_time:
                j += 1
            if j - i > 1:
                merge_groups.append((i + 2, i + j + 1))
            i = j

    ts_ptl = [
        ("FONTNAME", (0,0), (-1,-1), font),
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("SPAN", (0,0), (-1,0)),
        ("BACKGROUND", (0,0), (-1,1), colors.HexColor("#e6e6e6")),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]

    for start, end in merge_groups:
        ts_ptl.append(("SPAN", (0, start), (0, end-1)))

    t_ptl = Table(data_ptl, colWidths=[W*0.18, W*0.10, W*0.14, W*0.18, W*0.40], repeatRows=2)
    t_ptl.setStyle(TableStyle(ts_ptl))
    story.append(t_ptl)
    story.append(Spacer(1, 4*mm))

    # 巡簽與備註（略）
    # ... 保持不變 ...

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    buf.seek(0)
    return buf
