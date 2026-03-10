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
DEFAULT_BRIEF = "時間：各編組執行前\n地點：現地勤教"
DEFAULT_COMMANDER = "石門所副所長林榮裕"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "業務組", "代號": "隆安13", "姓名": "交通組警務員 葉佳媛", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報群聚飆車狀況。"},
    {"職稱": "督導組", "代號": "隆安681", "姓名": "督察組督察員 黃中彥", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "監看群聚告警訊息、指揮、調度及通報本勤務事宜。"}
])

DEFAULT_PATROL = pd.DataFrame([
    {
        "勤務時段": "3月7日零時至4時", "無線電": "隆安82", "編組": "專責警力（石門所輪值）", 
        "服勤人員": "00-02時：副所長林榮裕、警員王耀民\n02-04時：副所長林榮裕、警員陳欣妤", 
        "任務分工": "「加強防制」勤務，在文化路、中正路三坑段、龍源路及旭日路來回巡邏，隨機攔檢改裝（噪音）車輛"
    },
    {
        "勤務時段": "3月6日22時至翌日6時", "無線電": "隆安80", "編組": "石門所", 
        "服勤人員": "線上巡邏警力兼任", 
        "任務分工": "「區域聯防」勤務，於中正路、文化路、中豐路、龍源路巡邏（每1小時巡簽1次），並加強查緝毒駕"
    },
    {
        "勤務時段": "3月6日22時至翌日6時", "無線電": "隆安90", "編組": "高平所", 
        "服勤人員": "線上巡邏警力兼任", 
        "任務分工": "「區域聯防」勤務，於中豐路及龍源路巡邏（每1小時巡簽1次）"
    },
    {
        "勤務時段": "3月6日22時至翌日6時", "無線電": "隆安990", "編組": "龍潭交通分隊", 
        "服勤人員": "線上巡邏警力兼任", 
        "任務分工": "「區域聯防」勤務，於龍源路及溪州橋旁新建道路巡邏（每1小時巡簽1次）"
    },
    {
        "勤務時段": "3月6日22時至翌日6時", "無線電": "隆安50", "編組": "聖亭所", 
        "服勤人員": "線上巡邏警力兼任", "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"
    },
    {
        "勤務時段": "3月6日22時至翌日6時", "無線電": "隆安60", "編組": "龍潭所", 
        "服勤人員": "線上巡邏警力兼任", "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"
    },
    {
        "勤務時段": "3月6日22時至翌日6時", "無線電": "隆安70", "編組": "中興所", 
        "服勤人員": "線上巡邏警力兼任", "任務分工": "「區域聯防」勤務，於轄內易發生危險駕車路段巡邏"
    }
])

CHECKIN_POINTS = """1. 中油高原交流道站（龍源路2-20號）
2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）
3. 7-11龍潭佳園門市（中正路三坑段776號）
4. 旭日路三坑自然生態公園停車場
5. 旭日路與大溪區交界處"""

NOTES = """一、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。
二、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。
三、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"""

# --- 2. 建立連線與讀取 ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_data(ttl=60)
def load_data():
    try:
        client = get_client()
        if client is None:
            return None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("危駕_設定")
        ws_cmd = sh.worksheet("危駕_指揮組")
        ws_ptl = sh.worksheet("危駕_警力佈署")
        return pd.DataFrame(ws_set.get_all_records()), pd.DataFrame(ws_cmd.get_all_records()), pd.DataFrame(ws_ptl.get_all_records()), None
    except Exception as e:
        return None, None, None, str(e)

def save_data(time_str, briefing, commander, df_cmd, df_patrol):
    try:
        client = get_client()
        if client is None:
            return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("危駕_設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["plan_time", time_str], ["briefing", briefing], ["commander", commander]])
        
        for ws_name, df in [("危駕_指揮組", df_cmd), ("危駕_警力佈署", df_patrol)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df = df.fillna("")
            ws.update([df.columns.tolist()] + df.values.tolist())
        load_data.clear()
        return True
    except:
        return False

# --- 3. PDF 生成 ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames():
        return fname
    for p in ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf_from_data(time_str, briefing, commander, df_cmd, df_patrol):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []
    
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=24, alignment=1, spaceAfter=8)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=10)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=18, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=18, alignment=0)
    style_section = ParagraphStyle('Section', fontName=font, fontSize=14, leading=20, spaceAfter=4)
    style_note = ParagraphStyle('Note', fontName=font, fontSize=14, leading=20, spaceAfter=5)
    style_th = ParagraphStyle('THeader', fontName=font, fontSize=16, alignment=1, leading=22)

    story.append(Paragraph(f"{UNIT}執行「防制危險駕車專案勤務」規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def clean(txt):
        return str(txt).replace("\n", "<br/>").replace("、", "<br/>")

    data_cmd = [[Paragraph("<b>任　務　編　組</b>", style_th), '', '', ''],
                [Paragraph(f"<b>{h}</b>", style_cell) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([
            Paragraph(f"<b>{r.get('職稱','')}</b>", style_cell), 
            Paragraph(str(r.get('代號','')), style_cell),
            Paragraph(clean(r.get('姓名','')), style_cell), 
            Paragraph(str(r.get('任務','')), style_cell_left)
        ])
                         
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.12, page_width*0.28, page_width*0.45], repeatRows=2)
    t1.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), 
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), 
        ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')), 
        ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6)
    ]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    story.append(Paragraph(f"<b>📢 勤前教育：</b><br/>{briefing.replace(chr(10), '<br/>')}", style_section))
    story.append(Spacer(1, 4*mm))

    data_ptl = [
        [Paragraph("<b>警　力　佈　署</b>", style_th), '', '', '', ''],
        [Paragraph(f"<b>交通快打指揮官：</b>{commander}", style_cell_left), '', '', '', ''],
        [Paragraph(f"<b>{h}</b>", style_cell) for h in ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]]
    ]
    
    for _, r in df_patrol.iterrows():
        data_ptl.append([
            str(r.get('勤務時段','')), 
            str(r.get('無線電','')), 
            Paragraph(clean(r.get('編組','')), style_cell), 
            Paragraph(clean(r.get('服勤人員','')), style_cell), 
            Paragraph(str(r.get('任務分工','')), style_cell_left)
        ])

    t2 = Table(data_ptl, colWidths=[page_width*0.28, page_width*0.10, page_width*0.14, page_width*0.18, page_width*0.30], repeatRows=3)
    t2.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),font), ('FONTSIZE',(0,0),(-1,-1),14),
        ('ALIGN',(0,3),(1,-1),'CENTER'), ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), 
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),
        ('SPAN',(0,1),(-1,1)), ('BACKGROUND',(0,1),(-1,1),colors.white), 
        ('LEFTPADDING',(0,1),(-1,1),6), ('BACKGROUND',(0,2),(-1,2),colors.HexColor('#f2f2f2')),
        ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6)
    ]))
    story.append(t2)
    story.append(Spacer(1, 6*mm))

    story.append(Paragraph("<b>📍 巡簽地點：</b>", style_section))
    story.append(Paragraph(CHECKIN_POINTS.replace("\n", "<br/>"), style_note))
    story.append(Spacer(1, 4*mm))
    story.append(Paragraph("<b>📝 備註：</b>", style_section))
    story.append(Paragraph(NOTES.replace("\n", "<br/>"), style_note))

    doc.build(story)
    return buf.getvalue()

# --- 4. 寄信功能 ---
def send_report_email(time_str, briefing, commander, df_cmd, df_patrol):
    try:
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        pdf_bytes = generate_pdf_from_data(time_str, briefing, commander, df_cmd, df_patrol)
        
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender
        msg["Subject"] = f"防制危險駕車勤務表_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為最新的防制危險駕車勤務表 PDF。", "plain", "utf-8"))
        
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
    except Exception as e:
        return False, str(e)

# --- 5. 主介面邏輯 ---
df_set, df_cmd, df_ptl, err = load_data()
if err or df_set is None:
    t, b, cmdr = DEFAULT_TIME, DEFAULT_BRIEF, DEFAULT_COMMANDER
    ed_cmd, ed_ptl = DEFAULT_CMD.copy(), DEFAULT_PATROL.copy()
else:
    sd = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    t = sd.get("plan_time", DEFAULT_TIME)
    b = sd.get("briefing", DEFAULT_BRIEF)
    cmdr = sd.get("commander", DEFAULT_COMMANDER)
    ed_cmd, ed_ptl = df_cmd, df_ptl

st.title("🚔 防制危險駕車專案勤務規劃表")
c1, c2 = st.columns(2)
p_time = c1.text_input("勤務時間", t)
b_info = c2.text_area("📢 勤前教育", b, height=80)

st.subheader("1. 任務編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True)

st.subheader("2. 警力佈署")
cmdr_input = st.text_input("交通快打指揮官", cmdr)
res_ptl = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True)

st.subheader("3. 巡簽地點與備註 (固定)")
st.info("此區塊將直接附加於報表末端")

# --- 重構：安全產生 HTML，避免單行過長 ---
def get_html():
    b_html = b_info.replace('\n', '<br>')
    chk_html = CHECKIN_POINTS.replace('\n', '<br>')
    note_html = NOTES.replace('\n', '<br>')
    
    parts = []
    parts.append("<style>body{font-family:'標楷體';padding:20px;} th,td{border:1px solid black;padding:8px;font-size:14pt;text-align:center;line-height:1.5;} .note{font-size:14pt;margin:15px 0;line-height:1.6;} .cmd-row{text-align:left;background-color:white;}</style>")
    parts.append(f"<html><body><h2 style='text-align:center'>{UNIT}<br>執行「防制危險駕車專案勤務」規劃表</h2>")
    parts.append(f"<div style='text-align:right'><b>時間：{p_time}</b></div><br>")
    
    # 任務編組表格
    parts.append("<table><tr><th colspan='4'>任 務 編 組</th></tr>")
    parts.append("<tr><th width='15%'>職稱</th><th width='12%'>代號</th><th width='28%'>姓名</th><th width='45%'>任務</th></tr>")
    
    for _, r in res_cmd.iterrows():
        name = str(r.get('姓名', '')).replace('、', '<br>')
        parts.append(f"<tr><td><b>{r.get('職稱','')}</b></td><td>{r.get('代號','')}</td>")
        parts.append(f"<td>{name}</td><td style='text-align:left'>{r.get('任務','')}</td></tr>")
    parts.append("</table>")
    
    parts.append(f"<div class='note'><b>📢 勤前教育：</b><br>{b_html}</div>")
    
    # 警力佈署表格
    parts.append("<table><tr><th colspan='5'>警 力 佈 署</th></tr>")
    parts.append(f"<tr><td colspan='5' class='cmd-row'><b>交通快打指揮官：</b>{cmdr_input}</td></tr>")
    parts.append("<tr><th width='28%'>勤務時段</th><th width='10%'>代號</th><th width='14%'>編組</th><th width='18%'>服勤人員</th><th width='30%'>任務分工</th></tr>")
    
    for _, r in res_ptl.iterrows():
        grp = str(r.get('編組', '')).replace('、', '<br>')
        ppl = str(r.get('服勤人員', '')).replace('\n', '<br>')
        # 把這行拆成多行，防止被截斷
        parts.append("<tr>")
        parts.append(f"<td style='white-space:nowrap;'>{r.get('勤務時段','')}</td>")
        parts.append(f"<td style='white-space:nowrap;'>{r.get('無線電','')}</td>")
        parts.append(f"<td>{grp}</td><td>{ppl}</td>")
        parts.append(f"<td style='text-align:left'>{r.get('任務分工','')}</td>")
        parts.append("</tr>")
        
    parts.append("</table>")
    parts.append(f"<div class='note'><b>📍 巡簽地點：</b><br>{chk_html}</div>")
    parts.append(f"<div class='note'><b>📝 備註：</b><br>{note_html}</div>")
    parts.append("</body></html>")
    
    return "".join(parts)

st.markdown("---")
st.subheader("📄 預覽與輸出")
st.components.v1.html(get_html(), height=600, scrolling=True)

if st.button("同步雲端、寄信並下載 PDF 💾", type="primary"):
    save_data(p_time, b_info, cmdr_input, res_cmd, res_ptl)
    ok, mail_err = send_report_email(p_time, b_info, cmdr_input, res_cmd, res_ptl)
    if ok:
        st.success("📧 雲端同步成功，報表已寄至信箱！")
    else:
        st.error(f"❌ 雲端已同步，但寄信失敗：{mail_err}")
    
    pdf_out = generate_pdf_from_data(p_time, b_info, cmdr_input, res_cmd, res_ptl)
    st.download_button("點此下載 PDF", data=pdf_out, file_name=f"防制危險駕車勤務_{datetime.now().strftime('%Y%m%d')}.pdf")
