import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib
import io
import os
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# --- 1. 頁面設定 ---
st.set_page_config(page_title="取締砂石車專案勤務", layout="wide", page_icon="🚛")
st.title("🚛 取締砂石（大型貨）車重點違規專案勤務規劃表")

SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# 💡 固定備註內容
CORRECT_NOTES = (
    "一、執行前由各單位帶班人員在駐地實施勤前教育。<br/>"
    "二、加強取締砂石（大型貨）車超載、車速、酒醉駕車、闖紅燈、無照駕車、爭道行駛、"
    "違反禁行路線、變更車斗、未使用專用車箱及未裝設行車紀錄器（行車視野輔助器）等違規，"
    "以共同消弭不法行為，保障用路人生命財產安全。"
)

# --- 2. Google Sheets 連線 ---
@st.cache_resource
def get_client():
    try:
        if "gcp_service_account" not in st.secrets: 
            return None
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Google Sheets 認證失敗: {e}")
        return None

def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, "", "認證資訊缺失"
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("砂石_設定")
        ws_cmd = sh.worksheet("砂石_指揮組")
        ws_sch = sh.worksheet("砂石_勤務表")
        
        df_set = pd.DataFrame(ws_set.get_all_records())
        df_cmd = pd.DataFrame(ws_cmd.get_all_records())
        df_sch = pd.DataFrame(ws_sch.get_all_records())
        
        sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
        briefing = str(sd.get("briefing", "")).strip()
        
        # 簡單清理
        if any(x in briefing for x in ["時間：", "地點：", "各單位執行前"]):
            briefing = ""
            
        return df_set, df_cmd, df_sch, briefing, None
    except Exception as e:
        return None, None, None, "", str(e)

def save_data(month, briefing, df_cmd, df_schedule):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        
        # 儲存設定
        ws_set = sh.worksheet("砂石_設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["month", month], ["briefing", briefing]])
        
        # 儲存表格
        for ws_name, df in [("砂石_指揮組", df_cmd), ("砂石_勤務表", df_schedule)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df_fill = df.fillna("")
            ws.update([df_fill.columns.tolist()] + df_fill.values.tolist())
        return True
    except Exception as e:
        st.error(f"儲存失敗: {e}")
        return False

# --- 3. PDF 產製 ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    # 搜尋常見標楷體路徑，若在 Linux 部署建議將字體放於專案目錄
    font_paths = ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/custom/kaiu.ttf"]
    for p in font_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica" # 墊底方案

def generate_pdf(month, briefing, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=12*mm, bottomMargin=12*mm)
    W = A4[0] - 24*mm
    story = []
    
    # 樣式設定
    s_title = ParagraphStyle("t", fontName=font, fontSize=16, alignment=1, spaceAfter=8, leading=22)
    s_th = ParagraphStyle("th", fontName=font, fontSize=15, alignment=1, leading=20)
    s_cell = ParagraphStyle("c", fontName=font, fontSize=14, leading=18, alignment=1)
    s_left = ParagraphStyle("l", fontName=font, fontSize=14, leading=18, alignment=0)
    s_section = ParagraphStyle("sec", fontName=font, fontSize=14, leading=20, spaceAfter=4)
    s_note = ParagraphStyle("n", fontName=font, fontSize=12, leading=16, spaceAfter=4)
    
    def c(txt, style=s_cell): return Paragraph(str(txt).replace("\n","<br/>"), style)

    # 標題
    title_text = f"<b>{UNIT}執行{month}「取締砂石（大型貨）車重點違規」專案勤務規劃表</b>"
    story.append(Paragraph(title_text, s_title))
    
    # 指揮組表格
    cw1 = [W*0.15, W*0.12, W*0.28, W*0.45]
    data1 = [[Paragraph("<b>任　務　編　組</b>", s_th), '', '', ''], 
             [Paragraph(f"<b>{h}</b>", s_th) for h in ["職稱", "代號", "姓名", "任務"]]]
    
    for _, row in df_cmd.iterrows():
        data1.append([
            c(f"<b>{row.get('職稱','')}</b>"), 
            c(row.get('代號','')), 
            c(str(row.get('姓名','')).replace('、','<br/>')), 
            c(row.get('任務',''), s_left)
        ])
    
    t1 = Table(data1, colWidths=cw1, repeatRows=2)
    t1.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),font), 
        ('GRID',(0,0),(-1,-1),0.5,colors.black), 
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'), 
        ('SPAN',(0,0),(-1,0)), 
        ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))
    ]))
    story.append(t1)
    story.append(Spacer(1, 4*mm))
    
    # 勤前教育
    if briefing.strip():
        story.append(Paragraph(f"<b>📢 勤前教育：</b><br/>{briefing.replace(chr(10), '<br/>')}", s_section))
        story.append(Spacer(1, 4*mm))
    
    # 勤務表
    col_date = '勤務日期'
    cw2 = [W*0.28, W*0.16, W*0.12, W*0.44]
    data2 = [[Paragraph("<b>警　力　佈　署</b>", s_th), '', '', ''], 
             [Paragraph(f"<b>{h}</b>", s_th) for h in ["勤務日期", "執行單位", "執行人數", "執行路段"]]]
    
    for _, row in df_schedule.iterrows():
        data2.append([c(row.get(col_date, '')), c(row.get('執行單位','')), c(row.get('執行人數','')), c(row.get('執行路段', ''), s_left)])
    
    t_style = [
        ('FONTNAME',(0,0),(-1,-1),font), 
        ('GRID',(0,0),(-1,-1),0.5,colors.black), 
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'), 
        ('SPAN',(0,0),(-1,0)), 
        ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))
    ]
    
    # 日期欄位自動合併邏輯
    non_empty_idxs = [i for i, v in enumerate(df_schedule[col_date]) if str(v).strip() != ""]
    non_empty_idxs.append(len(df_schedule))
    for k in range(len(non_empty_idxs)-1):
        s, e = non_empty_idxs[k], non_empty_idxs[k+1]-1
        if e > s: t_style.append(('SPAN', (0, s+2), (0, e+2)))
            
    t2 = Table(data2, colWidths=cw2, repeatRows=2)
    t2.setStyle(TableStyle(t_style))
    story.append(KeepTogether([t2]))
    story.append(Spacer(1, 6*mm))
    
    # 備註
    story.append(Paragraph(f"<b>備註：</b><br/>{CORRECT_NOTES}", s_note))
    
    doc.build(story)
    return buf.getvalue()

# --- 4. 寄信功能 ---
def send_report_email(subject, pdf_bytes, filename):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender # 寄給自己或指定對象
        msg["Subject"] = subject
        
        msg.attach(MIMEText(f"您好，附件為「{subject}」之 PDF 規劃表。", "plain", "utf-8"))
        
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        # 處理中文檔名
        safe_filename = _ul.quote(filename)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{safe_filename}.pdf")
        msg.attach(part)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# --- 5. 主程式流程 ---
df_set, df_cmd_raw, df_sch_raw, filtered_brief, err = load_data()

if err:
    st.warning(f"目前處於編輯模式（無法連線雲端：{err}）")
    cur_month = "115年3月份"
    df_c = pd.DataFrame(columns=["職稱", "代號", "姓名", "任務"])
    df_s = pd.DataFrame(columns=["勤務日期", "執行單位", "執行人數", "執行路段"])
else:
    sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    cur_month = sd.get("month", "115年3月份")
    df_c, df_s = df_cmd_raw, df_sch_raw
    if "日期" in df_s.columns: df_s.rename(columns={"日期": "勤務日期"}, inplace=True)

# 介面配置
st.subheader("1. 基礎資訊")
c1, c2 = st.columns([1, 2])
month_val = c1.text_input("月份標題", value=cur_month)
brief_info = c2.text_area("📢 勤前教育內容", value=filtered_brief, height=100)

st.subheader("2. 任務編組")
ed_cmd = st.data_editor(df_c, num_rows="dynamic", use_container_width=True, key="editor_cmd")

st.subheader("3. 警力佈署")
st.info("💡 撇步：第一欄『勤務日期』若連續相同，請僅保留第一列，後續留白即可自動合併單元格。")
ed_sch = st.data_editor(df_s, num_rows="dynamic", use_container_width=True, key="editor_sch")

# 動態產生預覽 HTML
def get_html():
    parts = ["<style>body{font-family:'Microsoft JhengHei', '標楷體';} th{border:1px solid black;background-color:#f2f2f2;padding:5px;} td{border:1px solid black;text-align:center;padding:5px;} table{width:100%;border-collapse:collapse;}</style>"]
    parts.append(f"<h3 style='text-align:center;'>{UNIT}執行{month_val}專案勤務規劃表</h3>")
    parts.append("<table><tr><th colspan='4'>任 務 編 組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>")
    for _, r in ed_cmd.iterrows():
        parts.append(f"<tr><td><b>{r.get('職稱','')}</b></td><td>{r.get('代號','')}</td><td>{str(r.get('姓名','')).replace('、','<br>')}</td><td style='text-align:left'>{r.get('任務','')}</td></tr>")
    parts.append("</table>")
    if brief_info.strip(): 
        parts.append(f"<div style='margin-top:10px;'><b>📢 勤前教育：</b><br>{brief_info.replace(chr(10), '<br>')}</div>")
    parts.append("<table style='margin-top:10px;'><tr><th colspan='4'>警 力 佈 署</th></tr><tr><th>勤務日期</th><th>執行單位</th><th>執行人數</th><th>執行路段</th></tr>")
    
    row_idx = 0
    while row_idx < len(ed_sch):
        date_val = str(ed_sch.iloc[row_idx].get('勤務日期','')).strip()
        span = 1
        if date_val:
            for nxt in range(row_idx+1, len(ed_sch)):
                if not str(ed_sch.iloc[nxt].get('勤務日期','')).strip(): span += 1
                else: break
        for i in range(span):
            curr = ed_sch.iloc[row_idx+i]
            parts.append("<tr>")
            if i == 0: parts.append(f"<td rowspan='{span}'>{date_val}</td>")
            parts.append(f"<td>{curr.get('執行單位','')}</td><td>{curr.get('執行人數','')}</td><td style='text-align:left'>{curr.get('執行路段','')}</td></tr>")
        row_idx += span
    parts.append(f"</table><div style='text-align:left; margin-top:10px; font-size:0.9em;'><b>備註：</b><br>{CORRECT_NOTES}</div>")
    return "".join(parts)

st.markdown("---")
st.markdown("#### 📄 即時預覽")
st.components.v1.html(get_html(), height=500, scrolling=True)

# 操作按鈕
colA, colB = st.columns(2)

# 定義統一的檔名與標題
full_title = f"{UNIT}執行{month_val}「取締砂石（大型貨）車重點違規」專案勤務規劃表"

if colA.button("💾 同步雲端並發送電子郵件", type="primary", use_container_width=True):
    with st.spinner("處理中..."):
        if save_data(month_val, brief_info, ed_cmd, ed_sch):
            pdf_bytes = generate_pdf(month_val, brief_info, ed_cmd, ed_sch)
            ok, mail_err = send_report_email(f"勤務規劃表_{month_val}", pdf_bytes, full_title)
            if ok: st.success("✅ 雲端已更新，PDF 已寄至信箱！")
            else: st.error(f"❌ 雲端已更新，但寄信失敗：{mail_err}")
        else:
            st.error("❌ 雲端同步失敗，請檢查權限設定。")

# PDF 下載按鈕
pdf_data = generate_pdf(month_val, brief_info, ed_cmd, ed_sch)
colB.download_button(
    label="📥 下載 PDF 報表檔案",
    data=pdf_data,
    file_name=f"{full_title}.pdf",
    mime="application/pdf",
    use_container_width=True
)
