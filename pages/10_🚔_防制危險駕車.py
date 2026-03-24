import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import smtplib
import io
import os
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

# --- 1. 頁面設定 ---
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# --- 預設範本資料 ---
DEFAULT_TIME = "115年3月6日22時至翌日6時"
DEFAULT_COMMANDER = "石門所副所長林榮裕"

# --- 2. 建立連線與讀取 ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_data(ttl=60)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        return pd.DataFrame(sh.worksheet("危駕_設定").get_all_records()), \
               pd.DataFrame(sh.worksheet("危駕_指揮組").get_all_records()), \
               pd.DataFrame(sh.worksheet("危駕_警力佈署").get_all_records()), None
    except Exception as e: return None, None, None, str(e)

def save_data(time_str, commander, df_cmd, df_patrol):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        sh.worksheet("危駕_設定").update([["Key", "Value"], ["plan_time", time_str], ["commander", commander]])
        for ws_name, df in [("危駕_指揮組", df_cmd), ("危駕_警力佈署", df_patrol)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            ws.update([df.columns.tolist()] + df.fillna("").values.tolist())
        load_data.clear()
        return True
    except: return False

def auto_format_personnel(val):
    if pd.isna(val) or str(val).strip() in ["None", "nan", ""]: return ""
    s = str(val).replace('\\n', '\n').replace('、', '\n')
    s = re.sub(r'([^\n])\s*(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?)', r'\1\n\2', s)
    s = re.sub(r'(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?)[：:\s]*', r'\1：\n', s)
    return '\n'.join([line.strip() for line in s.split('\n') if line.strip()])

# --- 3. PDF 生成相關 ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    for p in ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf_from_data(time_str, commander, df_cmd, df_patrol):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []
    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, alignment=1, spaceAfter=8)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=10)
    style_th = ParagraphStyle('THeader', fontName=font, fontSize=16, alignment=1)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=16, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=16, alignment=0)

    story.append(Paragraph(f"<b>{UNIT}執行「防制危險駕車專案勤務」規劃表</b>", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def clean(txt):
        s = str(txt) if not pd.isna(txt) else ""
        s = re.sub(r'(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?：?)', r'<b>\1</b>', s)
        return s.replace('\n', '<br/>')

    # 指揮組表格
    data_cmd = [[Paragraph("<b>任　務　編　組</b>", style_th), '', '', ''], ["職稱", "代號", "姓名", "任務"]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([Paragraph(f"<b>{clean(r['職稱'])}</b>", style_cell), clean(r['代號']), clean(r['姓名']).replace("、", "<br/>"), Paragraph(clean(r['任務']), style_cell_left)])
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.15, page_width*0.25, page_width*0.45])
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    # 警力佈署表格
    data_ptl = [[Paragraph("<b>警　力　佈　署</b>", style_th), '', '', '', ''], [Paragraph(f"<b>交通快打指揮官：</b>{commander}", style_cell_left), '', '', '', ''], ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]]
    for _, r in df_patrol.iterrows():
        data_ptl.append([Paragraph(clean(r['勤務時段']), style_cell), clean(r['無線電']), Paragraph(clean(r['編組']), style_cell), Paragraph(clean(r['服勤人員']), style_cell), Paragraph(clean(r['任務分工']), style_cell_left)])
    t2 = Table(data_ptl, colWidths=[page_width*0.20, page_width*0.10, page_width*0.15, page_width*0.25, page_width*0.30])
    t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('SPAN',(0,1),(-1,1)), ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')), ('BACKGROUND',(0,2),(-1,2),colors.HexColor('#f2f2f2'))]))
    story.append(t2)
    doc.build(story)
    return buf.getvalue()

# --- 4. 主介面邏輯 ---
df_set, df_cmd, df_ptl, err = load_data()
if err or df_set is None:
    t, cmdr = DEFAULT_TIME, DEFAULT_COMMANDER
    ed_cmd, ed_ptl = pd.DataFrame(), pd.DataFrame()
else:
    sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    t, cmdr = sd.get("plan_time", DEFAULT_TIME), sd.get("commander", DEFAULT_COMMANDER)
    ed_cmd, ed_ptl = df_cmd, df_ptl

st.title("🚔 防制危險駕車專案勤務規劃表")

st.subheader("1. 基礎資訊")
p_time = st.text_input("勤務時間", t)
cmdr_input = st.text_input("交通快打指揮官", cmdr)

# 🎯【核心修正：強制更正單位名稱】🎯
# 只要偵測到「XX所/分隊/分局」，就強行將其與「輪值」組合，無視雲端舊資料
unit_match = re.search(r'([\u4e00-\u9fa5]+(?:所|分隊|分局))', cmdr_input)
if unit_match and len(ed_ptl) > 0:
    pure_unit = unit_match.group(1)
    
    # 使用 .iat 或 .at 進行絕對定址修改，確保 Streamlit 刷新
    ed_ptl.at[0, '編組'] = f"專責警力\n（{pure_unit}輪值）"
    
    # 無線電代號連動
    unit_map = {"石門": "隆安8", "高平": "隆安9", "聖亭": "隆安5", "龍潭": "隆安6", "中興": "隆安7", "分隊": "隆安99"}
    base = next((v for k, v in unit_map.items() if k in pure_unit), "隆安")
    suffix = "2" if re.search(r'副|小隊長', cmdr_input) else "1"
    ed_ptl.at[0, '無線電'] = base + suffix

    # 服勤人員連動
    name_only = cmdr_input.replace(pure_unit, "").strip()
    if name_only:
        staff_text = str(ed_ptl.at[0, '服勤人員'])
        ed_ptl.at[0, '服勤人員'] = re.sub(r'(\d{2}-\d{2}時：?)\n?.*', f'\\1\n{name_only}', staff_text)

# 日期自動推算
date_m = re.search(r'(\d+)月(\d+)日', p_time)
if date_m and len(ed_ptl) > 0:
    m, d = int(date_m.group(1)), int(date_m.group(2))
    dt_next = datetime(datetime.now().year, m, d) + timedelta(days=1)
    ed_ptl.at[0, '勤務時段'] = f"{dt_next.month}月{dt_next.day}日\n零時至4時"

st.subheader("2. 任務編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True).fillna("")

st.subheader("3. 警力佈署")
# 顯示前最後一次確保格式
if '服勤人員' in ed_ptl.columns:
    ed_ptl['服勤人員'] = ed_ptl['服勤人員'].apply(auto_format_personnel)

res_ptl = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True).fillna("")

# --- 5. 預覽與下載 ---
st.markdown("---")
if st.button("同步雲端、寄信並下載 PDF 💾", type="primary"):
    save_data(p_time, cmdr_input, res_cmd, res_ptl)
    st.success("✅ 已同步至雲端試算表！")
    pdf_out = generate_pdf_from_data(p_time, cmdr_input, res_cmd, res_ptl)
    st.download_button("點此下載 PDF 報表", data=pdf_out, file_name=f"危駕勤務表_{datetime.now().strftime('%m%d')}.pdf")

# HTML 預覽 (簡化版)
st.subheader("📄 報表預覽")
preview_html = f"<b>{UNIT}規劃表</b><br>指揮官：{cmdr_input}<br><table border='1' style='border-collapse:collapse;width:100%'><tr><th>編組</th><th>人員</th></tr><tr><td>{res_ptl.iloc[0,2] if len(res_ptl)>0 else ''}</td><td>{res_ptl.iloc[0,3] if len(res_ptl)>0 else ''}</td></tr></table>"
st.components.v1.html(preview_html, height=200)
