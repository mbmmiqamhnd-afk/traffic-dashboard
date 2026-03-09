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

# --- 1. 頁面設定 (必須放在程式碼最上方) ---
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
    {"職稱": "指揮官",       "代號": "隆安1",    "姓名": "分局長 施宇峰",                                           "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官",     "代號": "隆安2",    "姓名": "副分局長 何憶雯",                                         "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官",     "代號": "隆安3",    "姓名": "副分局長 蔡志明",                                         "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官",   "代號": "駐區督察", "姓名": "孫三陽",                                                      "任務": "重點機動督導。"},
    {"職稱": "督導組",       "代號": "隆安6",    "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組",       "代號": "隆安684",  "姓名": "督察組教官 郭文義",                                         "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13",   "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、勤指中心警員 張庭溱（代理人：巡官陳鵬翔）、行政組警務佐 曾威仁、人事室警員 陳明祥", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組",       "代號": "隆安",     "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興",            "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_PTL = pd.DataFrame([
    {"編組": "第一巡邏組", "無線電": "隆安54",  "單位": "聖亭所",       "服勤人員": "巡佐傅錫城、警員曾建凱",       "任務分工": "於大昌路一段周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第二巡邏組", "無線電": "隆安62",  "單位": "龍潭所",       "服勤人員": "副所長全楚文、警員龔品璇",     "任務分工": "於大昌路二段周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第三巡邏組", "無線電": "隆安72",  "單位": "中興所",       "服勤人員": "副所長薛德祥、警員冷柔萱",     "任務分工": "於中興路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第四巡邏組", "無線電": "隆安83",  "單位": "石門所",       "服勤人員": "巡佐林偉政、警員盧瑾瑤",       "任務分工": "於北龍路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第五巡邏組", "無線電": "隆安33",  "單位": "三和所、高平所","服勤人員": "警員唐銘聰、警員張湃柏",      "任務分工": "於大昌路一、二段、北龍路及中興路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第六巡邏組", "無線電": "隆安994", "單位": "龍潭交通分隊", "服勤人員": "小隊長林振生、警員吳沛軒",    "任務分工": "於大昌路一、二段、北龍路及中興路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
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

# --- 3. 讀取函數 ---
@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        if client is None:
            return None, None, None, "未設定 Secrets (離線模式)"
        sh = client.open_by_key(SHEET_ID)
        ws_list = sh.worksheets()
        
        ws_set = next((w for w in ws_list if w.title == "設定"), None)
        ws_cmd = next((w for w in ws_list if w.title == "指揮組"), None)
        ws_ptl = next((w for w in ws_list if w.title == "巡邏組"), None)

        if not all([ws_set, ws_cmd, ws_ptl]):
            return None, None, None, "缺工作表"

        df_settings = pd.DataFrame(ws_set.get_all_records())
        df_command  = pd.DataFrame(ws_cmd.get_all_records())
        df_patrol   = pd.DataFrame(ws_ptl.get_all_records())
        return df_settings, df_command, df_patrol, None
    except Exception as e:
        return None, None, None, str(e)

# --- 4. 寫入函數 ---
def save_data(unit, time_str, project, briefing, station, df_cmd, df_ptl):
    try:
        client = get_client()
        if client is None:
            st.warning("⚠️ 離線模式無法存檔")
            return False

        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"],
                       ["unit_name", unit],
                       ["plan_full_time", time_str],
                       ["project_name",   project],
                       ["briefing_info",  briefing],
                       ["check_station",  station]])

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

# --- 5. PDF 生成 (包含任務編組第一列美化) ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames():
        return fname
    font_paths = ["kaiu.ttf", "./kaiu.ttf", "font/kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]
    font_path = None
    for p in font_paths:
        if os.path.exists(p):
            font_path = p
            break   
    if font_path:
        try:
            pdfmetrics.registerFont(TTFont(fname, font_path))
            return fname
        except Exception:
            pass
    return "Helvetica"

def generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=15*mm, bottomMargin=15*mm,
        title=f"{unit}執行{project}規劃表")
        
    page_width = A4[0] - 30*mm
    story = []
    
    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, leading=22, spaceAfter=6)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=12)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=10, leading=13, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=10, leading=13, alignment=0)
    style_note = ParagraphStyle('Note', fontName=font, fontSize=11, leading=16, spaceAfter=4)
    style_table_title = ParagraphStyle('TableTitle', fontName=font, fontSize=14, alignment=1, leading=18) 

    # 1. 標題
    story.append(Paragraph(f"{unit}執行{project}規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def clean_text(txt):
        return str(txt).replace("\n", "<br/>").replace("、", "<br/>")

    # 2. 指揮組表格 (含任務編組大標題)
    col_widths_cmd = [page_width * 0.15, page_width * 0.10, page_width * 0.25, page_width * 0.50]
    headers_cmd = ["職稱", "代號", "姓名", "任務"]
    data_cmd = []
    
    # [Row 0] 任務編組標題
    title_cell = Paragraph("<b>任　務　編　組</b>", style_table_title)
    data_cmd.append([title_cell, '', '', '']) 
    
    # [Row 1] 欄位名稱
    header_row = [Paragraph(f"<b>{h}</b>", style_cell) for h in headers_cmd]
    data_cmd.append(header_row)
    
    # [Row 2+] 資料
    for _, row in df_cmd.iterrows():
        job = Paragraph(f"<b>{row.get('職稱','')}</b>", style_cell)
        code = Paragraph(str(row.get('代號','')), style_cell)
        name = Paragraph(clean_text(row.get('姓名','')), style_cell)
        task = Paragraph(str(row.get('任務','')), style_cell_left)
        data_cmd.append([job, code, name, task])

    t1 = Table(data_cmd, colWidths=col_widths_cmd, repeatRows=2)
    t1.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        # Row 0 (任務編組)
        ('SPAN', (0,0), (-1,0)),
        ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#f2f2f2')),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        # Row 1 (Header)
        ('BACKGROUND', (0,1), (-1, 1), colors.HexColor('#f2f2f2')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    # 3. 勤教與檢驗站
    story.append(Paragraph(f"<b>📢 勤前教育：</b>{briefing}", style_note))
    st_text = str(station).replace("\n", "<br/>")
    story.append(Paragraph(f"<b>🚧 檢驗站資訊：</b><br/>{st_text}", style_note))
    story.append(Spacer(1, 6*mm))

    # 4. 巡邏組表格
    col_widths_ptl = [page_width * 0.10, page_width * 0.08, page_width * 0.12, page_width * 0.18, page_width * 0.52]
    headers_ptl = ["編組", "代號", "單位", "服勤人員", "任務分工"]
    data_ptl = []
    data_ptl.append([Paragraph(f"<b>{h}</b>", style_cell) for h in headers_ptl])
    
    for _, row in df_ptl.iterrows():
        group = Paragraph(str(row.get('編組','')), style_cell)
        code = Paragraph(str(row.get('無線電','')), style_cell)
        unit_p = Paragraph(clean_text(row.get('單位','')), style_cell)
        ppl = Paragraph(clean_text(row.get('服勤人員','')), style_cell)
        
        task_text = str(row.get('任務分工',''))
        full_task = f"{task_text}<br/><font color='blue' size='9'>*雨備方案：轄區治安要點巡邏。</font>"
        task_cell = Paragraph(full_task, style_cell_left)
        data_ptl.append([group, code, unit_p, ppl, task_cell])

    t2 = Table(data_ptl, colWidths=col_widths_ptl, repeatRows=1)
    t2.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#f2f2f2')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t2)

    try:
        doc.build(story)
        return buf.getvalue()
    except Exception as e:
        print(f"PDF Build Error: {e}")
        return None

def send_report_email(html_content, subject, unit, time_str, project, briefing, station, df_cmd, df_ptl):
    import urllib.parse as _ul
    try:
        sender   = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        receiver = sender 
        
        pdf_bytes = generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl)
        if pdf_bytes is None:
            return False, "PDF 生成失敗"

        msg = MIMEMultipart()
        msg["From"]    = sender
        msg["To"]      = receiver
        msg["Subject"] = subject
        msg.attach(MIMEText("請見附件 PDF 報表。\n\n本郵件由雲端勤務系統自動發送。", "plain", "utf-8"))
        
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        encoded_name = _ul.quote(f"{subject}.pdf", safe='')
        part.add_header(
            "Content-Disposition",
            f"attachment; filename=\"report.pdf\"; filename*=UTF-8''{encoded_name}"
        )
        msg.attach(part)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, receiver, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# --- 6. 主程式 (介面渲染) ---
# 讀取資料
df_set, df_cmd, df_ptl, error_msg = load_data()

if error_msg:
    st.error(f"❌ Google Sheets 讀取失敗：{error_msg}")
    st.warning("⚠️ 啟用離線範本模式")
    current_unit    = DEFAULT_UNIT
    current_time    = DEFAULT_TIME
    current_proj    = DEFAULT_PROJ
    current_brief   = DEFAULT_BRIEF
    current_station = DEFAULT_STATION
    df_command_edit = DEFAULT_CMD.copy()
    df_patrol_edit  = DEFAULT_PTL.copy()
elif df_set is None:
    current_unit    = DEFAULT_UNIT
    current_time    = DEFAULT_TIME
    current_proj    = DEFAULT_PROJ
    current_brief   = DEFAULT_BRIEF
    current_station = DEFAULT_STATION
    df_command_edit = DEFAULT_CMD.copy()
    df_patrol_edit  = DEFAULT_PTL.copy()
else:
    try:
        settings_dict = dict(zip(df_set.iloc[:, 0], df_set.iloc[:, 1]))
        current_unit    = settings_dict.get("unit_name",      DEFAULT_UNIT)
        current_time    = settings_dict.get("plan_full_time", DEFAULT_TIME)
        current_proj    = settings_dict.get("project_name",   DEFAULT_PROJ)
        current_brief   = settings_dict.get("briefing_info",  DEFAULT_BRIEF)
        current_station = settings_dict.get("check_station",  DEFAULT_STATION)
        df_command_edit = df_cmd if not df_cmd.empty else DEFAULT_CMD.copy()
        df_patrol_edit  = df_ptl if not df_ptl.empty else DEFAULT_PTL.copy()
    except Exception as e:
        st.error(f"資料解析失敗：{e}")
        st.stop()

# 介面
st.title("🚓 專案勤務規劃表 (雲端同步版)")
st.caption("資料與 Google Sheets 即時連線，手機、電腦皆可編輯")

st.subheader("1. 勤務基礎資訊")
c1, c2 = st.columns([1, 1])
project_name = c1.text_input("專案名稱", value=current_proj)
plan_time    = c2.text_input("勤務時間", value=current_time)

st.subheader("2. 指揮與幕僚編組")
with st.expander("編輯名單 (指揮組)", expanded=True):
    edited_cmd = st.data_editor(df_command_edit, num_rows="dynamic", use_container_width=True, key="editor_cmd")

c3, c4 = st.columns(2)
# 這裡定義了 brief_info 與 check_st，供後面的函數使用
brief_info = c3.text_area("📢 勤前教育",   value=current_brief,   height=100)
check_st   = c4.text_area("🚧 檢驗站資訊", value=current_station, height=100)

st.subheader("3. 執行勤務編組 (巡邏組)")
edited_ptl = st.data_editor(df_patrol_edit, num_rows="dynamic", use_container_width=True, key="editor_ptl")

# HTML 產生器 (修復變數名稱 NameError)
def generate_html_content():
    style = """
    <style>
        body { font-family: 'DFKai-SB', 'BiauKai', '標楷體', serif; color: #000; }
        .container { width: 100%; max-width: 800px; margin: 0 auto; padding: 20px; }
        h2 { text-align: left; margin-bottom: 5px; letter-spacing: 2px; }
        .info { text-align: right; font-weight: bold; margin-bottom: 15px; font-size: 14px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid black; padding: 5px; text-align: center; font-size: 14px; vertical-align: middle; }
        th { background-color: #f2f2f2; }
        .left-align { text-align: left; }
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2>{current_unit}執行{project_name}規劃表</h2>"
    html += f"<div class='info'>勤務時間：{plan_time}</div>"
    
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr>"
    html += "<tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, row in edited_cmd.iterrows():
        name = str(row.get('姓名', '')).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        html += f"<tr><td><b>{row.get('職稱','')}</b></td><td>{row.get('代號','')}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務','')}</td></tr>"
    html += "</table>"
    
    html += f"<div class='left-align' style='margin-bottom:20px;line-height:1.6'>"
    # 關鍵修正：使用正確的變數 brief_info 和 check_st
    html += f"<div><b>📢 勤前教育：</b>{brief_info}</div>"
    html += f"<div style='white-space:pre-wrap'><b>🚧 {check_st}</b></div></div>"
    
    html += "<table><tr><th width='10%'>編組</th><th width='8%'>代號</th><th width='12%'>單位</th><th width='18%'>服勤人員</th><th width='52%'>任務分工</th></tr>"
    for _, row in edited_ptl.iterrows():
        name = str(row.get('服勤人員', '')).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        unit_cell = str(row.get("單位","")).replace("、","<br>").replace(",","<br>")
        html += f"<tr><td>{row.get('編組','')}</td><td>{row.get('無線電','')}</td><td style='line-height:1.4'>{unit_cell}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務分工','')}<br><span style='color:blue;font-size:0.9em'>*雨備方案：轄區治安要點巡邏。</span></td></tr>"
    html += "</table></div></body></html>"
    return html

html_out = generate_html_content()

st.markdown("---")
col_view, col_dl = st.columns([3, 1])
with col_view:
    st.subheader("📄 即時預覽")
    st.components.v1.html(html_out, height=800, scrolling=True)

with col_dl:
    st.subheader("📥 存檔與輸出")
    file_name_date = datetime.now().strftime('%Y%m%d')
    if st.download_button(
        label="下載報表並同步雲端 💾",
        data=html_out.encode("utf-8"),
        file_name=f"勤務表_{file_name_date}.html",
        mime="text/html; charset=utf-8",
        type="primary"
    ):
        save_success = save_data(current_unit, plan_time, project_name, brief_info, check_st, edited_cmd, edited_ptl)
        if save_success:
            subject = f"噪音車勤務規劃表_{file_name_date}"
            ok, err = send_report_email(html_out, subject, current_unit, plan_time, project_name, brief_info, check_st, edited_cmd, edited_ptl)
            if ok:
                st.toast("📧 報表已寄出至信箱！", icon="✉️")
            else:
                st.error(f"❌ 寄信失敗：{err}")
