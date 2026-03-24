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

CHECKIN_POINTS = """1. 中油高原交流道站（龍源路2-20號）
2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）
3. 7-11龍潭佳園門市（中正路三坑段776號）
4. 旭日路三坑自然生態公園停車場
5. 旭日路與大溪區交界處"""

NOTES = """一、各編組執行前由帶班人員在駐地實施勤前教育。
二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。
三、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。
四、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"""

# --- 2. 格式化工具 ---
def auto_format_personnel(val):
    if pd.isna(val) or str(val).strip() in ["None", "nan", ""]: return ""
    s = str(val).replace('\\n', '\n').replace('、', '\n')
    s = re.sub(r'([^\n])\s*(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?)', r'\1\n\2', s)
    s = re.sub(r'(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?)[：:\s]*', r'\1：\n', s)
    return '\n'.join([l.strip() for l in s.split('\n') if l.strip()])

# --- 3. 雲端與 PDF 引擎 ---
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
    
    style_title = ParagraphStyle('T', fontName=font, fontSize=16, alignment=1, spaceAfter=8)
    style_info = ParagraphStyle('I', fontName=font, fontSize=12, alignment=2, spaceAfter=10)
    style_th = ParagraphStyle('H', fontName=font, fontSize=16, alignment=1, leading=22)
    style_cell = ParagraphStyle('C', fontName=font, fontSize=14, leading=18, alignment=1) # 內文 14, 置中
    style_cell_l = ParagraphStyle('L', fontName=font, fontSize=14, leading=18, alignment=0) # 內文 14, 置左

    story.append(Paragraph(f"<b>{UNIT}執行「防制危險駕車專案勤務」規劃表</b>", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def br(txt):
        s = str(txt).replace('\n', '<br/>').replace('、', '<br/>')
        s = re.sub(r'(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?：?)', r'<b>\1</b>', s)
        return s

    # 1. 任務編組 (處理姓名垂直並列)
    data_cmd = [[Paragraph("<b>任　務　編　組</b>", style_th), '', '', ''], 
                [Paragraph(f"<b>{h}</b>", style_th) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        # 關鍵：將姓名中的頓號換成 <br/> 實現垂直並列
        name_vertical = br(r['姓名'])
        data_cmd.append([
            Paragraph(f"<b>{br(r['職稱'])}</b>", style_cell), 
            br(r['代號']), 
            Paragraph(name_vertical, style_cell), # 姓名垂直並列且置中
            Paragraph(br(r['任務']), style_cell_l)
        ])
    
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.15, page_width*0.25, page_width*0.45])
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    # 2. 警力佈署
    data_ptl = [[Paragraph("<b>警　力　佈　署</b>", style_th), '', '', '', ''], 
                [Paragraph(f"<b>交通快打指揮官：</b>{commander}", style_cell_l), '', '', '', ''], 
                [Paragraph(f"<b>{h}</b>", style_th) for h in ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]]]
    for _, r in df_patrol.iterrows():
        data_ptl.append([Paragraph(br(r['勤務時段']), style_cell), br(r['無線電']), Paragraph(br(r['編組']), style_cell), Paragraph(br(r['服勤人員']), style_cell), Paragraph(br(r['任務分工']), style_cell_l)])

    t2 = Table(data_ptl, colWidths=[page_width*0.20, page_width*0.10, page_width*0.15, page_width*0.25, page_width*0.30])
    t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('SPAN',(0,1),(-1,1)), ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')), ('BACKGROUND',(0,2),(-1,2),colors.HexColor('#f2f2f2'))]))
    story.append(t2)
    
    # 備註
    story.append(Spacer(1, 6*mm)); story.append(Paragraph("<b>📍 巡簽地點：</b>", style_cell_l))
    story.append(Paragraph(br(CHECKIN_POINTS), style_cell_l)); story.append(Spacer(1, 4*mm))
    story.append(Paragraph("<b>📝 備註：</b>", style_cell_l))
    for l in NOTES.split('\n'): story.append(Paragraph(l.strip(), style_cell_l))

    doc.build(story)
    return buf.getvalue()

# --- 4. 介面與邏輯 ---
df_set, df_cmd, df_ptl, err = load_data()
if err or df_set is None:
    t, cmdr = "115年3月6日22時至翌日6時", "石門所副所長林榮裕"
    ed_cmd, ed_ptl = pd.DataFrame(), pd.DataFrame()
else:
    sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1])); t, cmdr = sd.get("plan_time", ""), sd.get("commander", "")
    ed_cmd, ed_ptl = df_cmd, df_ptl

st.title("🚔 防制危險駕車專案勤務規劃表")
p_time = st.text_input("1. 勤務時間", t)
cmdr_input = st.text_input("2. 交通快打指揮官", cmdr)

# 自動去職稱邏輯
u_m = re.search(r'([\u4e00-\u9fa5]+(?:所|分隊|分局))', cmdr_input)
if u_m and len(ed_ptl) > 0:
    pu = u_m.group(1); ed_ptl.at[0, '編組'] = f"專責警力\n（{pu}輪值）"
    if '編組' in ed_ptl.columns: ed_ptl['編組'] = ed_ptl['編組'].apply(lambda x: re.sub(r'([\u4e00-\u9fa5]+(?:所|分隊|分局))(?:副所長|所長|分隊長|小隊長|警員)', r'\1輪值', str(x)))

st.subheader("3. 任務編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True).fillna("")
st.subheader("4. 警力佈署")
if '服勤人員' in ed_ptl.columns: ed_ptl['服勤人員'] = ed_ptl['服勤人員'].apply(auto_format_personnel)
res_ptl = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True).fillna("")

# --- 5. 預覽與下載 ---
def get_preview(df_c, df_p, cmdr_n, time_s):
    cmd_h = "".join([f"<tr><td>{r['職稱']}</td><td>{r['代號']}</td><td>{str(r['姓名']).replace('、','<br>')}</td><td>{r['任務']}</td></tr>" for _, r in df_c.iterrows()])
    ptl_h = "".join([f"<tr><td>{str(r['勤務時段']).replace('\n','<br>')}</td><td>{r['無線電']}</td><td>{str(r['編組']).replace('\n','<br>')}</td><td>{str(r['服勤人員']).replace('\n','<br>')}</td><td>{r['任務分工']}</td></tr>" for _, r in df_p.iterrows()])
    return f"""<style>table {{ width:100%; border-collapse:collapse; font-family:"標楷體"; }} th,td {{ border:1px solid black; padding:8px; text-align:center; }} th {{ background:#f2f2f2; font-size:16pt; }} td {{ font-size:14pt; }}</style>
    <h2 style='text-align:center;'>{UNIT} 規劃表</h2><div style='text-align:right;'>指揮官：{cmdr_n}</div><br>
    <table><tr><th colspan="4">任 務 編 組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>{cmd_h}</table><br>
    <table><tr><th colspan="5">警 力 佈 署</th></tr><tr><th>勤務時段</th><th>代號</th><th>編組</th><th>服勤人員</th><th>任務分工</th></tr>{ptl_h}</table>"""

st.components.v1.html(get_preview(res_cmd, res_ptl, cmdr_input, p_time), height=500, scrolling=True)

if st.button("💾 同步、寄信並下載", type="primary"):
    pdf = generate_pdf_from_data(p_time, cmdr_input, res_cmd, res_ptl)
    st.success("✅ 已同步！"); st.download_button("📥 下載 PDF", data=pdf, file_name="危駕勤務.pdf")
