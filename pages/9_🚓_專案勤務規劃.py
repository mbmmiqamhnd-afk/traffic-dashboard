import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib, io, os
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

# --- 常數與預設資料 ---
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

# --- 2. 建立 gspread 連線 ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets:
        return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

# --- 3. 讀取與寫入函數 ---
@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, "未設定 Secrets (離線模式)"
        sh = client.open_by_key(SHEET_ID)
        ws_list = sh.worksheets()
        ws_set = next((w for w in ws_list if w.title == "設定"), None)
        ws_cmd = next((w for w in ws_list if w.title == "指揮組"), None)
        ws_ptl = next((w for w in ws_list if w.title == "巡邏組"), None)
        if not all([ws_set, ws_cmd, ws_ptl]): return None, None, None, "缺工作表"
        return pd.DataFrame(ws_set.get_all_records()), pd.DataFrame(ws_cmd.get_all_records()), pd.DataFrame(ws_ptl.get_all_records()), None
    except Exception as e:
        return None, None, None, str(e)

def save_data(unit, time_str, project, briefing, station, df_cmd, df_ptl):
    try:
        client = get_client()
        if client is None:
            st.warning("⚠️ 離線模式無法存檔")
            return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["unit_name", unit], ["plan_full_time", time_str], ["project_name", project], ["briefing_info", briefing], ["check_station", station]])
        
        ws_cmd = sh.worksheet("指揮組")
        ws_cmd.clear()
        df_cmd = df_cmd.fillna("")
        ws_cmd.update([df_cmd.columns.tolist()] + df_cmd.values.tolist())

        ws_ptl = sh.worksheet("巡邏組")
        ws_ptl.clear()
        df_ptl = df_ptl.fillna("")
        ws_ptl.update([df_ptl.columns.tolist()] + df_ptl.values.tolist())
        
        load_data.clear()
        st.toast("✅ 雲端存檔成功！", icon="☁️")
        return True
    except Exception as e:
        st.error(f"❌ 存檔失敗：{e}")
        return False

# --- 4. PDF 生成 (字體大小改為 14) ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    font_paths = ["kaiu.ttf", "./kaiu.ttf", "font/kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]
    for p in font_paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(fname, p))
                return fname
            except: pass
    return "Helvetica"

def generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []
    
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=24, alignment=1, spaceAfter=10)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=12)
    # --- 關鍵修改：表格字體改為 14，Leading 增加為 18 ---
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=18, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=18, alignment=0)
    style_note = ParagraphStyle('Note', fontName=font, fontSize=12, leading=18, spaceAfter=4)
    style_table_title = ParagraphStyle('TableTitle', fontName=font, fontSize=16, alignment=1, leading=20) 

    story.append(Paragraph(f"{unit}執行{project}規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def clean_text(txt): return str(txt).replace("\n", "<br/>").replace("、", "<br/>")

    # 指揮組
    col_widths_cmd = [page_width * 0.15, page_width * 0.12, page_width * 0.28, page_width * 0.45]
    data_cmd = [[Paragraph("<b>任　務　編　組</b>", style_table_title), '', '', '']]
    data_cmd.append([Paragraph(f"<b>{h}</b>", style_cell) for h in ["職稱", "代號", "姓名", "任務"]])
    for _, row in df_cmd.iterrows():
        data_cmd.append([Paragraph(f"<b>{row.get('職稱','')}</b>", style_cell), Paragraph(str(row.get('代號','')), style_cell), 
                         Paragraph(clean_text(row.get('姓名','')), style_cell), Paragraph(str(row.get('任務','')), style_cell_left)])

    t1 = Table(data_cmd, colWidths=col_widths_cmd, repeatRows=2)
    t1.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), font), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                            ('SPAN', (0,0), (-1,0)), ('BACKGROUND', (0,0), (-1, 1), colors.HexColor('#f2f2f2')), ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6)]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph(f"<b>📢 勤前教育：</b>{briefing}", style_note))
    story.append(Paragraph(f"<b>🚧 檢驗站資訊：</b><br/>{station.replace(chr(10), '<br/>')}", style_note))
    story.append(Spacer(1, 6*mm))

    # 巡邏組
    col_widths_ptl = [page_width * 0.12, page_width * 0.10, page_width * 0.14, page_width * 0.20, page_width * 0.44]
    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "代號", "單位", "服勤人員", "任務分工"]]]
    for _, row in df_ptl.iterrows():
        task_text = f"{row.get('任務分工','')}<br/><font color='blue' size='11'>*雨備方案：轄區治安要點巡邏。</font>"
        data_ptl.append([Paragraph(str(row.get('編組','')), style_cell), Paragraph(str(row.get('無線電','')), style_cell), 
                         Paragraph(clean_text(row.get('單位','')), style_cell), Paragraph(clean_text(row.get('服勤人員','')), style_cell), Paragraph(task_text, style_cell_left)])

    t2 = Table(data_ptl, colWidths=col_widths_ptl, repeatRows=1)
    t2.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), font), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#f2f2f2')), 
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6)]))
    story.append(t2)
    doc.build(story)
    return buf.getvalue()

# --- 5. 主程式 ---
df_set, df_cmd, df_ptl, error_msg = load_data()
if error_msg or df_set is None:
    current_unit, current_time, current_proj, current_brief, current_station = DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_PROJ, DEFAULT_BRIEF, DEFAULT_STATION
    df_command_edit, df_patrol_edit = DEFAULT_CMD.copy(), DEFAULT_PTL.copy()
else:
    settings_dict = dict(zip(df_set.iloc[:, 0], df_set.iloc[:, 1]))
    current_unit = settings_dict.get("unit_name", DEFAULT_UNIT)
    current_time = settings_dict.get("plan_full_time", DEFAULT_TIME)
    current_proj = settings_dict.get("project_name", DEFAULT_PROJ)
    current_brief = settings_dict.get("briefing_info", DEFAULT_BRIEF)
    current_station = settings_dict.get("check_station", DEFAULT_STATION)
    df_command_edit, df_patrol_edit = (df_cmd if not df_cmd.empty else DEFAULT_CMD.copy()), (df_ptl if not df_ptl.empty else DEFAULT_PTL.copy())

st.title("🚓 專案勤務規劃表 (雲端同步版)")
c1, c2 = st.columns(2)
project_name = c1.text_input("專案名稱", value=current_proj)
plan_time = c2.text_input("勤務時間", value=current_time)

st.subheader("1. 指揮與幕僚編組")
edited_cmd = st.data_editor(df_command_edit, num_rows="dynamic", use_container_width=True, key="editor_cmd")

c3, c4 = st.columns(2)
brief_info = c3.text_area("📢 勤前教育", value=current_brief, height=100)
check_st = c4.text_area("🚧 檢驗站資訊", value=current_station, height=100)

st.subheader("2. 執行勤務編組 (巡邏組)")
edited_ptl = st.data_editor(df_patrol_edit, num_rows="dynamic", use_container_width=True, key="editor_ptl")

# HTML 產生器 (字體大小改為 14pt)
def generate_html_content():
    style = """
    <style>
        body { font-family: '標楷體', serif; color: #000; }
        .container { width: 100%; max-width: 900px; margin: 0 auto; padding: 10px; }
        th, td { border: 1px solid black; padding: 8px; text-align: center; font-size: 14pt; vertical-align: middle; line-height: 1.5; }
        th { background-color: #f2f2f2; }
        .left-align { text-align: left; }
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2 style='text-align:center'>{current_unit}執行<br>{project_name}規劃表</h2>"
    html += f"<div style='text-align:right; font-weight:bold;'>勤務時間：{plan_time}</div><br>"
    html += "<table><tr><th colspan='4' style='font-size:16pt'>任　務　編　組</th></tr><tr><th width='15%'>職稱</th><th width='12%'>代號</th><th width='28%'>姓名</th><th width='45%'>任務</th></tr>"
    for _, row in edited_cmd.iterrows():
        name = str(row.get('姓名', '')).replace("、", "<br>").replace(",", "<br>")
        html += f"<tr><td><b>{row.get('職稱','')}</b></td><td>{row.get('代號','')}</td><td>{name}</td><td class='left-align'>{row.get('任務','')}</td></tr>"
    html += "</table>"
    html += f"<div class='left-align' style='font-size:13pt; line-height:1.6'><b>📢 勤前教育：</b>{brief_info}<br><b>🚧 檢驗站資訊：</b>{check_st.replace(chr(10), '<br>')}</div><br>"
    html += "<table><tr><th width='12%'>編組</th><th width='10%'>代號</th><th width='14%'>單位</th><th width='20%'>服勤人員</th><th width='44%'>任務分工</th></tr>"
    for _, row in edited_ptl.iterrows():
        unit_c = str(row.get("單位","")).replace("、","<br>")
        ppl_c = str(row.get("服勤人員","")).replace("、","<br>")
        html += f"<tr><td>{row.get('編組','')}</td><td>{row.get('無線電','')}</td><td>{unit_c}</td><td>{ppl_c}</td><td class='left-align'>{row.get('任務分工','')}<br><span style='color:blue;font-size:0.9em'>*雨備方案：轄區治安要點巡邏。</span></td></tr>"
    html += "</table></div></body></html>"
    return html

html_out = generate_html_content()
st.markdown("---")
st.subheader("📄 即時預覽")
st.components.v1.html(html_out, height=600, scrolling=True)

if st.button("下載 PDF 並同步雲端 💾", type="primary"):
    if save_data(current_unit, plan_time, project_name, brief_info, check_st, edited_cmd, edited_ptl):
        pdf_data = generate_pdf_from_data(current_unit, project_name, plan_time, brief_info, check_st, edited_cmd, edited_ptl)
        st.download_button("點此儲存 PDF 檔案", data=pdf_data, file_name=f"勤務表_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf")
