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

# --- 2. 核心排版引擎 (時段自動換行) ---
def auto_format_personnel(val):
    if pd.isna(val) or str(val).strip() in ["None", "nan", ""]: 
        return ""
    s = str(val).replace('\\n', '\n').replace('、', '\n')
    # 遇到時段強制斷行並補上冒號
    s = re.sub(r'([^\n])\s*(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?)', r'\1\n\2', s)
    s = re.sub(r'(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?)[：:\s]*', r'\1：\n', s)
    lines = [line.strip() for line in s.split('\n') if line.strip()]
    return '\n'.join(lines)

# --- 3. 雲端連動 ---
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
        sh.worksheet("危駕_設定").clear()
        sh.worksheet("危駕_設定").update([["Key", "Value"], ["plan_time", time_str], ["commander", commander]])
        for ws_name, df in [("危駕_指揮組", df_cmd), ("危駕_警力佈署", df_patrol)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            ws.update([df.columns.tolist()] + df.fillna("").values.tolist())
        load_data.clear()
        return True
    except: return False

# --- 4. PDF 生成 (精準字體設定：標題 16 / 內文 14) ---
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
    
    # 🎯 字體大小邏輯設定
    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, alignment=1, spaceAfter=8)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=10)
    style_th = ParagraphStyle('THeader', fontName=font, fontSize=16, alignment=1, leading=22) # 標題 16
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=18, alignment=1) # 內文 14
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=18, alignment=0)
    style_section = ParagraphStyle('Section', fontName=font, fontSize=14, leading=20, spaceAfter=4)

    story.append(Paragraph(f"<b>{UNIT}執行「防制危險駕車專案勤務」規劃表</b>", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def clean(txt):
        s = str(txt) if not pd.isna(txt) else ""
        s = re.sub(r'(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?：?)', r'<b>\1</b>', s)
        return s.replace('\n', '<br/>')

    # 任務編組
    data_cmd = [[Paragraph("<b>任　務　編　組</b>", style_th), '', '', ''], 
                [Paragraph(f"<b>{h}</b>", style_th) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([Paragraph(f"<b>{clean(r['職稱'])}</b>", style_cell), clean(r['代號']), clean(r['姓名']).replace("、", "<br/>"), Paragraph(clean(r['任務']), style_cell_left)])
    
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.15, page_width*0.25, page_width*0.45])
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    # 警力佈署
    data_ptl = [[Paragraph("<b>警　力　佈　署</b>", style_th), '', '', '', ''], 
                [Paragraph(f"<b>交通快打指揮官：</b>{commander}", style_cell_left), '', '', '', ''], 
                [Paragraph(f"<b>{h}</b>", style_th) for h in ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]]]
    for _, r in df_patrol.iterrows():
        data_ptl.append([Paragraph(clean(r['勤務時段']), style_cell), clean(r['無線電']), Paragraph(clean(r['編組']), style_cell), Paragraph(clean(r['服勤人員']), style_cell), Paragraph(clean(r['任務分工']), style_cell_left)])

    t2 = Table(data_ptl, colWidths=[page_width*0.20, page_width*0.10, page_width*0.15, page_width*0.25, page_width*0.30])
    t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('SPAN',(0,1),(-1,1)), ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')), ('BACKGROUND',(0,2),(-1,2),colors.HexColor('#f2f2f2'))]))
    story.append(t2)
    
    # 備註區塊
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph("<b>📍 巡簽地點：</b>", style_section))
    story.append(Paragraph(CHECKIN_POINTS.replace("\n", "<br/>"), style_cell_left))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph("<b>📝 備註：</b>", style_section))
    for line in (re.sub(r'^\s+', '', l) for l in NOTES.split('\n') if l.strip()):
        story.append(Paragraph(line, style_cell_left))

    doc.build(story)
    return buf.getvalue()

# --- 5. 寄信功能 ---
def send_report_email(time_str, commander, df_cmd, df_patrol, file_date_str):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        pdf_bytes = generate_pdf_from_data(time_str, commander, df_cmd, df_patrol)
        msg = MIMEMultipart()
        msg["From"], msg["To"], msg["Subject"] = sender, sender, f"防制危險駕車勤務規劃表_{file_date_str}"
        msg.attach(MIMEText("附件為最新的防制危險駕車勤務規劃表 PDF。", "plain", "utf-8"))
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        filename = _ul.quote(f"{msg['Subject']}.pdf")
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{filename}")
        msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# --- 6. 主介面邏輯 ---
df_set, df_cmd, df_ptl, err = load_data()
if err or df_set is None:
    t, cmdr = "115年3月6日22時至翌日6時", "石門所副所長林榮裕"
    ed_cmd, ed_ptl = pd.DataFrame(), pd.DataFrame()
else:
    sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    t, cmdr = sd.get("plan_time", ""), sd.get("commander", "")
    ed_cmd, ed_ptl = df_cmd, df_ptl

st.title("🚔 防制危險駕車專案勤務規劃表")

p_time = st.text_input("1. 勤務時間", t)
cmdr_input = st.text_input("2. 交通快打指揮官", cmdr)

# 🎯 核心連動：暴力去職稱 (自動修正 石門所輪值)
unit_match = re.search(r'([\u4e00-\u9fa5]+(?:所|分隊|分局))', cmdr_input)
if unit_match and len(ed_ptl) > 0:
    pure_unit = unit_match.group(1)
    ed_ptl.at[0, '編組'] = f"專責警力\n（{pure_unit}輪值）"
    umap = {"石門": "隆安8", "高平": "隆安9", "聖亭": "隆安5", "龍潭": "隆安6", "中興": "隆安7", "分隊": "隆安99"}
    base = next((v for k, v in umap.items() if k in pure_unit), "隆安")
    suffix = "2" if "副" in cmdr_input or "小隊長" in cmdr_input else "1"
    ed_ptl.at[0, '無線電'] = base + suffix
    name_only = cmdr_input.replace(pure_unit, "").strip()
    if name_only:
        ed_ptl.at[0, '服勤人員'] = re.sub(r'(\d{2}-\d{2}時：?)\n?.*', f'\\1\n{name_only}', str(ed_ptl.at[0, '服勤人員']))

# 全表去職稱掃描
if '編組' in ed_ptl.columns:
    ed_ptl['編組'] = ed_ptl['編組'].apply(lambda x: re.sub(r'([\u4e00-\u9fa5]+(?:所|分隊|分局))(?:副所長|所長|分隊長|小隊長|警員)', r'\1輪值', str(x)))

st.subheader("3. 任務編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True).fillna("")

st.subheader("4. 警力佈署")
if '服勤人員' in ed_ptl.columns:
    ed_ptl['服勤人員'] = ed_ptl['服勤人員'].apply(auto_format_personnel)
res_ptl = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True).fillna("")

# --- 7. HTML 完整預覽 (標題 16pt / 內文 14pt) ---
st.markdown("---")
st.subheader("📄 報表預覽 (標題 16pt / 內文 14pt)")

def get_html_preview(df_c, df_p, cmdr_name, time_str):
    cmd_h = "".join([f"<tr><td>{r['職稱']}</td><td>{r['代號']}</td><td>{r['姓名']}</td><td>{r['任務']}</td></tr>" for _, r in df_c.iterrows()])
    ptl_h = "".join([f"<tr><td>{str(r['勤務時段']).replace('\n','<br>')}</td><td>{r['無線電']}</td><td>{str(r['編組']).replace('\n','<br>')}</td><td>{str(r['服勤人員']).replace('\n','<br>')}</td><td>{r['任務分工']}</td></tr>" for _, r in df_p.iterrows()])
    return f"""<style>table {{ width: 100%; border-collapse: collapse; font-family: "標楷體"; }} th, td {{ border: 1px solid black; padding: 8px; text-align: center; }} th {{ background-color: #f2f2f2; font-size: 16pt; }} td {{ font-size: 14pt; }}</style>
    <h2 style='text-align:center;'>{UNIT} 規劃表</h2><div style='text-align:right;'>時間：{time_str} | 指揮官：{cmdr_name}</div><br>
    <table><tr><th colspan="4">任 務 編 組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>{cmd_h}</table><br>
    <table><tr><th colspan="5">警 力 佈 署</th></tr><tr><th>勤務時段</th><th>代號</th><th>編組</th><th>服勤人員</th><th>任務分工</th></tr>{ptl_h}</table>"""

st.components.v1.html(get_html_preview(res_cmd, res_ptl, cmdr_input, p_time), height=600, scrolling=True)

# --- 8. 儲存與寄信 ---
if st.button("💾 同步雲端、寄信並下載 PDF", type="primary"):
    dt_str = datetime.now().strftime('%Y%m%d')
    if save_data(p_time, cmdr_input, res_cmd, res_ptl):
        ok, err_m = send_report_email(p_time, cmdr_input, res_cmd, res_ptl, dt_str)
        if ok: st.success("✅ 同步成功並已寄送郵件！")
        else: st.error(f"⚠️ 同步成功但寄信失敗: {err_m}")
        st.download_button("📥 下載 PDF", data=generate_pdf_from_data(p_time, cmdr_input, res_cmd, res_ptl), file_name=f"危駕勤務_{dt_str}.pdf")
