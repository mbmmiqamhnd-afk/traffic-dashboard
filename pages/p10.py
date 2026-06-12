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
import numpy as np
from datetime import datetime, timedelta
import re

# =========================
# 基本設定
# =========================
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
WS_MAP = {"set": "危駕_設定", "cmd": "危駕_指揮組", "ptl": "危駕_警力佈署"}
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
UNIT_TITLE = "桃園市政府警察局龍潭分局"
CMD_COLS = ["職稱", "代號", "姓名", "任務"]
# 這裡將「任務分工」改為「巡邏路段」
PTL_COLS = ["勤務時段", "代號", "編組", "服勤人員", "巡邏路段"]

# =========================
# 預設底稿
# =========================
DEFAULT_PROJECT = "防制危險駕車專案勤務"
DEFAULT_TIME = "115年5月22日22時至翌日6時"
DEFAULT_FAST_CMD = "龍潭所副所長全楚文"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "業務組", "代號": "隆安13", "姓名": "交通組巡官 郭勝隆", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報群聚飆車狀況。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青\n執勤官 李文章\n執勤員 黃文興", "任務": "監看群聚告警訊息、指揮、調度及通報本勤務事宜。"},
])

DEFAULT_PTL = pd.DataFrame([
    {"勤務時段": "5月23日\n零時至4時", "代號": "隆安62", "編組": "專責警力\n（龍潭所）", 
     "服勤人員": "00-02時段：\n警員廖怡惠\n警員劉柏延\n02-04時段：\n警員林軒宇\n警員廖怡惠", 
     "巡邏路段": "「加強防制」勤務，在文化路、中正路三坑段、龍源路及旭日路來回巡邏，隨機攔檢改裝（噪音）車輛（每2小時至責任區域內指定巡簽地點巡簽1次並守望10分鐘，將守望情形拍照上傳LINE「龍潭分局聯絡平臺」群組）"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安80", "編組": "石門所", 
     "服勤人員": "線上巡邏組合警力兼任", 
     "巡邏路段": "「區域聯防」勤務，於中正路、文化路、中豐路、龍源路及旭日路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安90", "編組": "高平所", 
     "服勤人員": "線上巡邏組合警力兼任", 
     "巡邏路段": "「區域聯防」勤務，於中豐路及龍源路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安990", "編組": "龍潭交通分隊", 
     "服勤人員": "線上巡邏組合警力兼任", 
     "巡邏路段": "「區域聯防」勤務，於龍源路及旭日路巡邏（每1小時巡邏人員至責任區域內指定巡簽地點巡簽1次）"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安50", "編組": "聖亭所", 
     "服勤人員": "線上巡邏組合警力兼任", 
     "巡邏路段": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安60", "編組": "龍潭所", 
     "服勤人員": "線上巡邏組合警力兼任", 
     "巡邏路段": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
    {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安70", "編組": "中興所", 
     "服勤人員": "線上巡邏組合警力兼任", 
     "巡邏路段": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"},
])

DEFAULT_SIGN_POINTS = "巡簽地點：\n1. 中油高原交流道站（龍源路2-20號）\n2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）\n3. 7-11龍潭佳園門市（中正路三坑段776號）\n4. 旭日路三坑自然生態公園停車場\n5. 旭日路與大溪區交界處"
DEFAULT_NOTES = "一、各編組執行前由帶班人員在駐地實施勤前教育。\n二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。\n三、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。\n四、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"

# =========================
# 字體與 Google Sheets
# =========================
@st.cache_resource
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames():
        return fname
    for p in ["./kaiu.ttf", "kaiu.ttf", "/usr/share/fonts/truetype/kaiu.ttf", "/app/kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(fname, p))
                return fname
            except:
                pass
    return "Helvetica"

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
    headers = {WS_MAP["set"]: [["Key", "Value"]], WS_MAP["cmd"]: [CMD_COLS], WS_MAP["ptl"]: [PTL_COLS]}
    for name, header in headers.items():
        try:
            sh.worksheet(name)
        except:
            sh.add_worksheet(title=name, rows="200", cols="20").update(header)
    st.success("初始化完成")
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
        ptl_df = pd.DataFrame(sh.worksheet(WS_MAP["ptl"]).get_all_records()).fillna("")
        
        if not ptl_df.empty:
            # ==========================================
            # 自動偵測並把舊的「任務分工」替換成「巡邏路段」
            # ==========================================
            if "任務分工" in ptl_df.columns:
                ptl_df = ptl_df.rename(columns={"任務分工": "巡邏路段"})
            
            ptl_df = ptl_df.reindex(columns=PTL_COLS, fill_value="")
            ptl_df = ptl_df[ptl_df["勤務時段"].astype(str).str.strip() != ""].reset_index(drop=True)
        
        settings = {}
        if not set_df.empty and set_df.shape[1] >= 2:
            settings = dict(zip(set_df.iloc[:,0].astype(str), set_df.iloc[:,1].astype(str)))
        return set_df, cmd_df, ptl_df, settings, None
    except Exception as e:
        return None, None, None, {}, str(e)

def save_data(settings_dict, cmd, ptl):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet(WS_MAP["set"])
        ws_set.clear()
        ws_set.update([["Key", "Value"]] + [[k, v] for k, v in settings_dict.items()])
        
        for ws_name, df, cols in [(WS_MAP["cmd"], cmd, CMD_COLS), (WS_MAP["ptl"], ptl, PTL_COLS)]:
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
# PDF 生成 - 交通快打指揮官在第2列（跨全部欄位）
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

    story.append(Paragraph(f"<b>{UNIT_TITLE}執行「{project_name}」規劃表</b>", s_title))
    story.append(Paragraph(f"勤務時間：{time_str}", s_sub))
    story.append(Spacer(1, 3*mm))

    # 任務編組表
    cmd_clean = cmd_df.dropna(how="all").fillna("")
    data_cmd = [[Paragraph("<b>任 務 編 組</b>", s_th), "", "", ""]]
    data_cmd.append([Paragraph(f"<b>{h}</b>", s_th) for h in CMD_COLS])
    for _, row in cmd_clean.iterrows():
        data_cmd.append([c(f"<b>{row.get('職稱','')}</b>"), c(row.get("代號","")), 
                         c(str(row.get("姓名","")).replace("、","<br/>")), c(row.get("任務",""), s_left)])
    t_cmd = Table(data_cmd, colWidths=[W*0.13, W*0.11, W*0.25, W*0.51], repeatRows=2)
    t_cmd.setStyle(TableStyle([("FONTNAME",(0,0),(-1,-1),font), ("GRID",(0,0),(-1,-1),0.5,colors.black),
                               ("VALIGN",(0,0),(-1,-1),"MIDDLE"), ("SPAN",(0,0),(-1,0)), 
                               ("BACKGROUND",(0,0),(-1,1),colors.HexColor("#f2f2f2"))]))
    story.append(t_cmd)
    story.append(Spacer(1, 4*mm))

    # 警力佈署表格
    ptl_clean = ptl_df.copy()
    if len(ptl_clean) < 3:
        ptl_clean = DEFAULT_PTL.copy()
    else:
        ptl_clean = ptl_clean[ptl_clean["勤務時段"].astype(str).str.strip() != ""].reset_index(drop=True)

    data_ptl = [[Paragraph("<b>警 力 佈 署</b>", s_th), "", "", "", ""]]

    # 第2列：交通快打指揮官（跨全部欄位）
    data_ptl.append([Paragraph(f"<b>交通快打指揮官：{fast_cmd}</b>", s_th), "", "", "", ""])

    # 第3列：欄位標題
    data_ptl.append([Paragraph(f"<b>{h}</b>", s_th) for h in PTL_COLS])

    # 警力資料
    for _, row in ptl_clean.iterrows():
        data_ptl.append([
            c(row.get("勤務時段", "")),
            c(row.get("代號", "")),
            c(row.get("編組", "")),
            c(str(row.get("服勤人員", "")).replace("、", "<br/>")),
            c(row.get("巡邏路段", ""), s_left)  # 這裡也對應修改為巡邏路段
        ])

    # 合併設定
    merge_groups = [(1, 2)]  # 交通快打指揮官跨全部欄位
    i = 3
    while i < len(data_ptl):
        j = i + 1
        curr = str(ptl_clean.iloc[i-3]["勤務時段"]).strip() if i-3 < len(ptl_clean) else ""
        while j < len(data_ptl) and (j-3 < len(ptl_clean)) and str(ptl_clean.iloc[j-3]["勤務時段"]).strip() == curr:
            j += 1
        if j - i > 1:
            merge_groups.append((i, j))
        i = j

    ts_ptl = [
        ("FONTNAME", (0,0), (-1,-1), font),
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("SPAN", (0,0), (-1,0)),           # 總標題
        ("SPAN", (0,1), (-1,1)),           # 交通快打指揮官跨全部欄位
        ("BACKGROUND", (0,0), (-1,2), colors.HexColor("#e6e6e6")),
    ]
    for start, end in merge_groups:
        if start > 2:  # 只對警力資料做合併
            ts_ptl.append(("SPAN", (0, start), (0, end-1)))

    t_ptl = Table(data_ptl, colWidths=[W*0.18, W*0.10, W*0.14, W*0.18, W*0.40], repeatRows=3)
    t_ptl.setStyle(TableStyle(ts_ptl))
    story.append(t_ptl)
    story.append(Spacer(1, 4*mm))

    # 巡簽與備註
    if sign_points.strip():
        for line in sign_points.strip().split("\n"):
            if line.strip():
                story.append(Paragraph(line.strip(), s_note))
        story.append(Spacer(1, 3*mm))

    story.append(Paragraph("<b>備註：</b>", s_note))
    for line in notes.strip().split("\n"):
        if line.strip():
            story.append(Paragraph(line.strip(), s_note))

    def add_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 10)
        canvas.drawCentredString(A4[0] / 2, 8*mm, f"- {canvas.getPageNumber()} -")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    buf.seek(0)
    return buf

# Email
def send_email(subject, pdf_buf, filename):
    try:
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender
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
# 主畫面
# =========================
st.title("🚔 防制危險駕車勤務規劃")

if st.sidebar.button("🔧 初始化工作表"):
    init_sheets()
if st.sidebar.button("🔄 強制重新載入"):
    st.cache_data.clear()
    st.rerun()

set_df, cmd_df, ptl_df, settings, err = load_data()
if err:
    st.warning(f"⚠️ 無法連線 Google Sheets（{err}），顯示預設底稿。")

use_cmd = cmd_df if (cmd_df is not None and not cmd_df.empty) else DEFAULT_CMD.copy()
use_ptl = ptl_df if (ptl_df is not None and not ptl_df.empty) else DEFAULT_PTL.copy()

use_ptl = use_ptl[use_ptl["勤務時段"].astype(str).str.strip() != ""].reset_index(drop=True)

st.subheader("基本設定")
col_a, col_b = st.columns(2)
project_name = col_a.text_input("專案名稱", value=settings.get("project_name", DEFAULT_PROJECT))
time_val = col_b.text_input("勤務時間", value=settings.get("time", DEFAULT_TIME))
fast_cmd = st.text_input("交通快打指揮官", value=settings.get("fast_cmd", DEFAULT_FAST_CMD))

# ========================================================
# 根據「勤務時間」與「指揮官」自動連動更新「警力部署」
# ========================================================
date_updated = False

# 1. 自動從「交通快打指揮官」擷取單位名稱 (自動濾除帶班幹部、所長等職稱，僅保留單位名稱)
unit_match = re.search(r"(.+?(所|分隊))", fast_cmd)
cmd_unit = unit_match.group(1) if unit_match else "該單位"

# 2. 建立各單位「專責警力」的代號「字首」對照表
unit_prefix_map = {
    "聖亭所": "隆安5",
    "龍潭所": "隆安6",
    "中興所": "隆安7",
    "石門所": "隆安8",
    "高平所": "隆安9",
    "龍潭交通分隊": "隆安99"
}

# 取得單位代號字首
base_code = unit_prefix_map.get(cmd_unit, "請確認代號")

# 動態判斷職稱以決定尾數
if base_code != "請確認代號":
    # 優先判斷明確為副手或幹部的職稱
    if "副所長" in fast_cmd or "副分隊長" in fast_cmd or "小隊長" in fast_cmd:
        cmd_code = f"{base_code}2"
    # 再判斷正主管
    elif "所長" in fast_cmd or "分隊長" in fast_cmd:
        cmd_code = f"{base_code}1"
    else:
        cmd_code = f"{base_code}2" # 預設巡邏網尾數
else:
    cmd_code = base_code

# 3. 嘗試解析完整的民國年月日格式 (例如: 115年5月25日)
date_match = re.search(r"(\d+)年(\d+)月(\d+)日", time_val)
if date_match:
    try:
        roc_year = int(date_match.group(1))
        m = int(date_match.group(2))
        d = int(date_match.group(3))
        
        # 換算為西元年以準確處理跨月與閏年
        ce_year = roc_year + 1911
        current_date = datetime(ce_year, m, d)
        next_date = current_date + timedelta(days=1)
        
        m_str = str(m)
        d_str = str(d)
        next_m_str = str(next_date.month)
        next_d_str = str(next_date.day)
        
        for idx in range(len(use_ptl)):
            group_name = str(use_ptl.loc[idx, "編組"])
            if "專責警力" in group_name:
                use_ptl.loc[idx, "勤務時段"] = f"{next_m_str}月{next_d_str}日\n零時至4時"
                # 連動編組名稱，僅呈現單位名稱，移除帶班職稱
                use_ptl.loc[idx, "編組"] = f"專責警力\n（{cmd_unit}）"
                use_ptl.loc[idx, "代號"] = cmd_code
            else:
                use_ptl.loc[idx, "勤務時段"] = f"{m_str}月{d_str}日\n22時至翌日6時"
        date_updated = True
    except ValueError:
        pass

# 4. 備用方案：若使用者未輸入年份，僅輸入「X月X日」
if not date_updated:
    date_match_simple = re.search(r"(\d+)月(\d+)日", time_val)
    if date_match_simple:
        m_str = date_match_simple.group(1)
        d = int(date_match_simple.group(2))
        next_d = d + 1
        
        for idx in range(len(use_ptl)):
            group_name = str(use_ptl.loc[idx, "編組"])
            if "專責警力" in group_name:
                use_ptl.loc[idx, "勤務時段"] = f"{m_str}月{next_d}日\n零時至4時"
                use_ptl.loc[idx, "編組"] = f"專責警力\n（{cmd_unit}）"
                use_ptl.loc[idx, "代號"] = cmd_code
            else:
                use_ptl.loc[idx, "勤務時段"] = f"{m_str}月{d}日\n22時至翌日6時"
# ========================================================

st.subheader("1. 任務編組")
res_cmd = st.data_editor(use_cmd, num_rows="dynamic", use_container_width=True)

st.subheader("2. 警力佈署")
st.caption("💡 相同勤務時段會自動合併")
res_ptl = st.data_editor(use_ptl, num_rows="fixed", use_container_width=True, height=420)

st.subheader("3. 巡簽地點與備註")
col_c, col_d = st.columns(2)
sign_points = col_c.text_area("巡簽地點", value=settings.get("sign_points", DEFAULT_SIGN_POINTS), height=160)
notes = col_d.text_area("備註", value=settings.get("notes", DEFAULT_NOTES), height=160)

st.markdown("---")

# ========================================================
# 提取檔名用的日期（自動抓取「年月日」或「月日」）
# ========================================================
file_date_str = ""
date_match_file = re.search(r"(\d+年)?\d+月\d+日", time_val)
if date_match_file:
    file_date_str = f"({date_match_file.group(0)})"

# 動態生成包含日期的檔案名稱
dynamic_filename = f"{UNIT_TITLE}執行「{project_name}」規劃表{file_date_str}"
# ========================================================

if st.button("💾 同步雲端並發送郵件", use_container_width=True):
    # 1. 儲存至雲端
    s = {"project_name": project_name, "time": time_val, "fast_cmd": fast_cmd, "sign_points": sign_points, "notes": notes}
    save_ok = save_data(s, res_cmd, res_ptl)
    
    # 2. 生成 PDF 並發送郵件
    pdf_buf = generate_pdf(time_val, project_name, fast_cmd, res_cmd, res_ptl, sign_points, notes)
    mail_ok, mail_err = send_email(dynamic_filename, pdf_buf, dynamic_filename)
    
    # 3. 根據結果顯示提示訊息
    if save_ok and mail_ok:
        st.success("✅ 已成功同步至雲端並寄出郵件！")
    elif save_ok and not mail_ok:
        st.warning(f"⚠️ 已儲存至雲端，但發送郵件失敗：{mail_err}")
    elif not save_ok and mail_ok:
        st.warning("⚠️ 郵件已寄出，但未能儲存至雲端！")
    else:
        st.error(f"❌ 儲存與發送均失敗 (發送錯誤：{mail_err})")
