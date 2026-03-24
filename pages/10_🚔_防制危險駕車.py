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
        df_set = pd.DataFrame(sh.worksheet("危駕_設定").get_all_records())
        df_cmd = pd.DataFrame(sh.worksheet("危駕_指揮組").get_all_records())
        df_ptl = pd.DataFrame(sh.worksheet("危駕_警力佈署").get_all_records())
        return df_set, df_cmd, df_ptl, None
    except Exception as e: return None, None, None, str(e)

def save_data(time_str, commander, df_cmd, df_patrol):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        sh.worksheet("危駕_設定").clear()
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

# --- 3. PDF 生成 ---
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
    style_th = ParagraphStyle('H', fontName=font, fontSize=16, alignment=1)
    style_c = ParagraphStyle('C', fontName=font, fontSize=14, leading=16, alignment=1)
    style_l = ParagraphStyle('L', fontName=font, fontSize=14, leading=16, alignment=0)

    story.append(Paragraph(f"<b>{UNIT}執行「防制危險駕車專案勤務」規劃表</b>", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def clean(txt):
        s = str(txt) if not pd.isna(txt) else ""
        s = re.sub(r'(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?：?)', r'<b>\1</b>', s)
        return s.replace('\n', '<br/>')

    data_cmd = [[Paragraph("<b>任　務　編　組</b>", style_th), '', '', ''], ["職稱", "代號", "姓名", "任務"]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([Paragraph(f"<b>{clean(r['職稱'])}</b>", style_c), clean(r['代號']), clean(r['姓名']).replace("、", "<br/>"), Paragraph(clean(r['任務']), style_l)])
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.15, page_width*0.25, page_width*0.45])
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    data_ptl = [[Paragraph("<b>警　力　佈　署</b>", style_th), '', '', '', ''], [Paragraph(f"<b>交通快打指揮官：</b>{commander}", style_l), '', '', '', ''], ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]]
    for _, r in df_patrol.iterrows():
        data_ptl.append([Paragraph(clean(r['勤務時段']), style_c), clean(r['無線電']), Paragraph(clean(r['編組']), style_c), Paragraph(clean(r['服勤人員']), style_c), Paragraph(clean(r['任務分工']), style_l)])
    t2 = Table(data_ptl, colWidths=[page_width*0.20, page_width*0.10, page_width*0.15, page_width*0.25, page_width*0.30])
    t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('SPAN',(0,1),(-1,1)), ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')), ('BACKGROUND',(0,2),(-1,2),colors.HexColor('#f2f2f2'))]))
    story.append(t2)
    doc.build(story)
    return buf.getvalue()

# --- 4. 主流程 ---
df_set, df_cmd, df_ptl, err = load_data()
if err or df_set is None:
    t_val, cmdr_val = "115年3月6日22時至翌日6時", "石門所副所長林榮裕"
    ed_cmd, ed_ptl = pd.DataFrame(), pd.DataFrame()
else:
    sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    t_val, cmdr_val = sd.get("plan_time", ""), sd.get("commander", "")
    ed_cmd, ed_ptl = df_cmd, df_ptl

st.title("🚔 防制危險駕車專案勤務規劃表")

p_time = st.text_input("1. 勤務時間", t_val)
cmdr_input = st.text_input("2. 交通快打指揮官", cmdr_val)

# ====== 🎯 暴力修正引擎：解決「副所長」文字殘留問題 ======
if len(ed_ptl) > 0:
    # A. 提取單位名稱 (如: 石門所)
    m = re.search(r'([\u4e00-\u9fa5]+(?:所|分隊|分局))', cmdr_input)
    if m:
        unit = m.group(1)
        # B. 無論雲端存什麼，強制複寫第一列
        # 這裡用 loc 確保精準寫入
        ed_ptl.loc[0, '編組'] = f"專責警力\n（{unit}輪值）"
        
        # C. 無線電連動
        umap = {"石門": "隆安8", "高平": "隆安9", "聖亭": "隆安5", "龍潭": "隆安6", "中興": "隆安7", "分隊": "隆安99"}
        base = next((v for k, v in umap.items() if k in unit), "隆安")
        suffix = "2" if any(x in cmdr_input for x in ["副", "小隊長"]) else "1"
        ed_ptl.loc[0, '無線電'] = base + suffix

        # D. 姓名連動
        name = cmdr_input.replace(unit, "").strip()
        if name:
            old_staff = str(ed_ptl.loc[0, '服勤人員'])
            ed_ptl.loc[0, '服勤人員'] = re.sub(r'(\d{2}-\d{2}時：?)\n?.*', f'\\1\n{name}', old_staff)

# E. 最後一道防線：遍歷全表，只要「編組」欄位有職稱，通通切掉
if '編組' in ed_ptl.columns:
    def force_clean(x):
        s = str(x)
        # 如果裡面有「所」又有「副所長」，把副所長等職稱刪掉
        return re.sub(r'(副所長|所長|分隊長|小隊長)', '', s)
    ed_ptl['編組'] = ed_ptl['編組'].apply(force_clean)

# 日期自動跳隔天
dm = re.search(r'(\d+)月(\d+)日', p_time)
if dm and len(ed_ptl) > 0:
    m, d = int(dm.group(1)), int(dm.group(2))
    dt_n = datetime(datetime.now().year, m, d) + timedelta(days=1)
    ed_ptl.loc[0, '勤務時段'] = f"{dt_n.month}月{dt_n.day}日\n零時至4時"

# --- 顯示編輯器 ---
st.subheader("3. 任務編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True).fillna("")

st.subheader("4. 警力佈署")
if '服勤人員' in ed_ptl.columns:
    ed_ptl['服勤人員'] = ed_ptl['服勤人員'].apply(auto_format_personnel)
res_ptl = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True).fillna("")

# --- 存檔與下載 ---
st.markdown("---")
if st.button("💾 同步雲端、下載 PDF", type="primary"):
    if save_data(p_time, cmdr_input, res_cmd, res_ptl):
        st.success("✅ 雲端同步成功！")
        pdf = generate_pdf_from_data(p_time, cmdr_input, res_cmd, res_ptl)
        st.download_button("📥 下載 PDF 報表", data=pdf, file_name=f"危駕勤務_{datetime.now().strftime('%m%d')}.pdf")
