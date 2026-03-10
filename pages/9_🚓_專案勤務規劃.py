import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib, io, os
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
st.set_page_config(page_title="雲端勤務規劃", layout="wide", page_icon="🚓")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年3月20日19至23時"
DEFAULT_PROJ    = "0320「取締改裝(噪音)車輛專案監、警、環聯合稽查勤務」"
DEFAULT_BRIEF   = "19時30分於分局二樓會議室召開"
DEFAULT_STATION = "環保局臨時檢驗站開設時間：20時至23時\n地點：桃園市龍潭區大昌路一段277號（龍潭區警政聯合辦公大樓）廣場"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "駐區督察", "姓名": "孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、勤指中心警員 張庭溱（代理人：巡官陳鵬翔）、行政組警務佐 曾威仁、人事室警員 陳明祥", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_PTL = pd.DataFrame([
    {"編組": "第一巡邏組", "無線電": "隆安54", "單位": "聖亭所", "服勤人員": "巡佐傅錫城、警員曾建凱", "任務分工": "於大昌路一段周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第二巡邏組", "無線電": "隆安62", "單位": "龍潭所", "服勤人員": "副所長全楚文、警員龔品璇", "任務分工": "於大昌路二段周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第三巡邏組", "無線電": "隆安72", "單位": "中興所", "服勤人員": "副所長薛德祥、警員冷柔萱", "任務分工": "於中興路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第四巡邏組", "無線電": "隆安83", "單位": "石門所", "服勤人員": "巡佐林偉政、警員盧瑾瑤", "任務分工": "於北龍路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第五巡邏組", "無線電": "隆安33", "單位": "三和所、高平所", "服勤人員": "警員唐銘聰、警員張湃柏", "任務分工": "於大昌路一、二段、北龍路及中興路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第六巡邏組", "無線電": "隆安994", "單位": "龍潭交通分隊", "服勤人員": "小隊長林振生、警員吳沛軒", "任務分工": "於大昌路一、二段、北龍路及中興路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
])

# --- 2. 建立連線與讀取 ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("設定")
        ws_cmd = sh.worksheet("指揮組")
        ws_ptl = sh.worksheet("巡邏組")
        return pd.DataFrame(ws_set.get_all_records()), pd.DataFrame(ws_cmd.get_all_records()), pd.DataFrame(ws_ptl.get_all_records()), None
    except Exception as e: return None, None, None, str(e)

def save_data(unit, time_str, project, briefing, station, df_cmd, df_ptl):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["unit_name", unit], ["plan_full_time", time_str], ["project_name", project], ["briefing_info", briefing], ["check_station", station]])
        
        for ws_name, df in [("指揮組", df_cmd), ("巡邏組", df_ptl)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            df = df.fillna("")
            ws.update([df.columns.tolist()] + df.values.tolist())
        load_data.clear()
        return True
    except: return False

# --- 3. PDF 生成 (字體統一為 14) ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    font_paths = ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]
    for p in font_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []
    
    # 字體設定
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=24, alignment=1, spaceAfter=8)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=10)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=18, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=18, alignment=0)
    style_note = ParagraphStyle('Note', fontName=font, fontSize=14, leading=20, spaceAfter=5)
    style_table_title = ParagraphStyle('TTitle', fontName=font, fontSize=16, alignment=1, leading=22)

    story.append(Paragraph(f"{unit}執行{project}規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def clean(t): return str(t).replace("\n", "<br/>").replace("、", "<br/>")

    # 表格 1：指揮組
    data_cmd = [[Paragraph("<b>任 務 編 組</b>", style_table_title), '', '', ''],
                [Paragraph(f"<b>{h}</b>", style_cell) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([Paragraph(f"<b>{r.get('職稱','')}</b>", style_cell), Paragraph(str(r.get('代號','')), style_cell),
                         Paragraph(clean(r.get('姓名','')), style_cell), Paragraph(str(r.get('任務','')), style_cell_left)])
    
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.12, page_width*0.28, page_width*0.45])
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('SPAN',(0,0),(-1,0)),
                            ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    # 中間文字 (字體 14)
    story.append(Paragraph(f"<b>📢 勤前教育：</b>{briefing}", style_note))
    story.append(Paragraph(f"<b>🚧 檢驗站資訊：</b><br/>{station.replace(chr(10), '<br/>')}", style_note))
    story.append(Spacer(1, 6*mm))

    # 表格 2：巡邏組 (修改前兩欄為純文字以避免換行)
    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "代號", "單位", "服勤人員", "任務分工"]]]
    for _, r in df_ptl.iterrows():
        task = f"{r.get('任務分工','')}<br/><font color='blue' size='11'>*雨備方案：轄區治安要點巡邏。</font>"
        # 第1, 2欄直接放入字串 (String) 強制不換行，第3, 4, 5欄放入 Paragraph 允許換行
        data_ptl.append([
            str(r.get('編組','')), 
            str(r.get('無線電','')),
            Paragraph(clean(r.get('單位','')), style_cell), 
            Paragraph(clean(r.get('服勤人員','')), style_cell), 
            Paragraph(task, style_cell_left)
        ])
    
    # 微調欄寬比例，給前兩欄更多空間
    t2 = Table(data_ptl, colWidths=[page_width*0.15, page_width*0.12, page_width*0.13, page_width*0.20, page_width*0.40])
    
    # 增加 FONTSIZE 與 ALIGN 設定來對應前兩欄的純文字
    t2.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,-1),font),
        ('FONTSIZE',(0,0),(-1,-1),14),           # 給純文字加上 14 號字
        ('ALIGN',(0,1),(1,-1),'CENTER'),         # 讓前兩欄的純文字置中對齊
        ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE')
    ]))
    story.append(t2)
    
    doc.build(story)
    return buf.getvalue()

# --- 4. 寄信功能 ---
def send_report_email(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    try:
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        pdf_bytes = generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl)
        
        msg = MIMEMultipart()
        msg["From"], msg["To"], msg["Subject"] = sender, sender, f"勤務規劃表_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為最新的勤務規劃表 PDF。", "plain", "utf-8"))
        
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

# --- 5. 主介面 ---
df_set, df_cmd, df_ptl, err = load_data()
if err or df_set is None:
    u, t, p, b, s = DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_PROJ, DEFAULT_BRIEF, DEFAULT_STATION
    ed_cmd, ed_ptl = DEFAULT_CMD.copy(), DEFAULT_PTL.copy()
else:
    d = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    u, t, p, b, s = d.get("unit_name", DEFAULT_UNIT), d.get("plan_full_time", DEFAULT_TIME), d.get("project_name", DEFAULT_PROJ), d.get("briefing_info", DEFAULT_BRIEF), d.get("check_station", DEFAULT_STATION)
    ed_cmd, ed_ptl = df_cmd, df_ptl

st.title("🚓 專案勤務規劃表")
c1, c2 = st.columns(2)
p_name = c1.text_input("專案名稱", p)
p_time = c2.text_input("勤務時間", t)

st.subheader("1. 指揮編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True)

c3, c4 = st.columns(2)
b_info = c3.text_area("📢 勤前教育", b, height=100)
s_info = c4.text_area("🚧 檢驗站資訊", s, height=100)

st.subheader("2. 巡邏編組")
res_ptl = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True)

# --- HTML 預覽 (加入 white-space:nowrap 防止換行) ---
def get_html():
    style = "<style>body{font-family:'標楷體';padding:20px;} th,td{border:1px solid black;padding:8px;font-size:14pt;text-align:center;} .note{font-size:14pt;margin:15px 0;line-height:1.6;}</style>"
    html = f"<html>{style}<body><h2 style='text-align:center'>{u}<br>{p_name}</h2><div style='text-align:right'><b>時間：{p_time}</b></div><br><table><tr><th colspan='4'>任 務 編 組</th></tr>"
    for _, r in res_cmd.iterrows():
        html += f"<tr><td><b>{r.get('職稱','')}</b></td><td>{r.get('代號','')}</td><td>{str(r.get('姓名','')).replace('、','<br>')}</td><td style='text-align:left'>{r.get('任務','')}</td></tr>"
    html += f"</table><div class='note'><b>📢 勤前教育：</b>{b_info}<br><b>🚧 檢驗站資訊：</b>{s_info.replace(chr(10),'<br>')}</div>"
    html += "<table><tr><th width='15%'>編組</th><th width='12%'>代號</th><th width='13%'>單位</th><th width='20%'>人員</th><th width='40%'>任務</th></tr>"
    
    for _, r in res_ptl.iterrows():
        # 在這裡的 <td> 加入 style='white-space: nowrap;' 強制不換行
        html += f"<tr><td style='white-space: nowrap;'>{r.get('編組','')}</td><td style='white-space: nowrap;'>{r.get('無線電','')}</td><td>{str(r.get('單位','')).replace('、','<br>')}</td><td>{str(r.get('服勤人員','')).replace('、','<br>')}</td><td style='text-align:left'>{r.get('任務分工','')}</td></tr>"
    return html + "</table></body></html>"

st.markdown("---")
st.subheader("📄 預覽與輸出")
st.components.v1.html(get_html(), height=500, scrolling=True)

if st.button("同步雲端、寄信並下載 PDF 💾", type="primary"):
    save_data(u, p_time, p_name, b_info, s_info, res_cmd, res_ptl)
    ok, mail_err = send_report_email(u, p_name, p_time, b_info, s_info, res_cmd, res_ptl)
    if ok: st.success("📧 雲端同步成功，報表已寄至信箱！")
    else: st.error(f"❌ 雲端已同步，但寄信失敗：{mail_err}")
    
    pdf_out = generate_pdf_from_data(u, p_name, p_time, b_info, s_info, res_cmd, res_ptl)
    st.download_button("點此下載 PDF", data=pdf_out, file_name=f"勤務規劃表_{datetime.now().strftime('%Y%m%d')}.pdf")
