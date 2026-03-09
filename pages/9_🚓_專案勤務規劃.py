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
import re  # 將 re 移到全域引用比較乾淨

# --- 1. 頁面設定 ---
st.set_page_config(page_title="雲端勤務規劃", layout="wide", page_icon="🚓")
st.title("🚓 專案勤務規劃表 (雲端同步版)")
st.caption("資料與 Google Sheets 即時連線，手機、電腦皆可編輯")

# 常數設定
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# --- 預設範本資料 (保持不變) ---
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

# --- 2. 建立 gspread 連線 (加上快取，避免重複連線) ---
# ttl=600 表示憑證載入後快取 10 分鐘，通常憑證不會一直變，可以設久一點
@st.cache_resource
def get_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

# --- 3. 讀取函數 (關鍵修正：加上 st.cache_data) ---
# ttl=10 表示資料快取 10 秒。這樣使用者連續輸入時不會一直讀取 Google，但又能保持相對即時。
@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        # 一次讀取所有工作表，減少 API 呼叫
        ws_list = sh.worksheets()
        
        # 根據名稱找工作表，比較安全
        ws_set = next((w for w in ws_list if w.title == "設定"), None)
        ws_cmd = next((w for w in ws_list if w.title == "指揮組"), None)
        ws_ptl = next((w for w in ws_list if w.title == "巡邏組"), None)

        if not all([ws_set, ws_cmd, ws_ptl]):
            return None, None, None, "找不到指定的工作表 (設定, 指揮組, 巡邏組)"

        df_settings = pd.DataFrame(ws_set.get_all_records())
        df_command  = pd.DataFrame(ws_cmd.get_all_records())
        df_patrol   = pd.DataFrame(ws_ptl.get_all_records())
        return df_settings, df_command, df_patrol, None
    except Exception as e:
        return None, None, None, str(e)

# --- 4. 寫入函數 (寫入後清除快取) ---
def save_data(unit, time_str, project, briefing, station, df_cmd, df_ptl):
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)

        ws_set = sh.worksheet("設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"],
                       ["unit_name", unit], # 修正：原本好像漏了存 unit
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
        
        # 寫入成功後，清除讀取快取，確保下次讀取是新的
        load_data.clear()
        
        st.toast("✅ 雲端存檔成功！", icon="☁️")
        return True
    except Exception as e:
        st.error(f"❌ 存檔失敗：{e}")
        return False

# --- 字型 & PDF & 寄信函數 ---
def _get_font():
    fname = "kaiu"
    # 檢查是否已註冊
    if fname in pdfmetrics.getRegisteredFontNames():
        return fname
        
    # 嘗試載入字型
    font_paths = ["kaiu.ttf", "./kaiu.ttf", "font/kaiu.ttf"] # 多增加幾個搜尋路徑
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
            
    # 如果真的找不到中文字型，回傳 Helvetica 避免程式崩潰，但中文會變亂碼
    print("Warning: 找不到 kaiu.ttf，中文PDF將無法正常顯示。")
    return "Helvetica"

def _parse_html_to_pdf(html_content, page_title):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=12*mm, rightMargin=12*mm,
        topMargin=12*mm, bottomMargin=12*mm,
        title=page_title) # 加入 PDF 標題
        
    W = A4[0] - 24*mm
    # 定義樣式，確保使用正確字型
    title_s = ParagraphStyle("t",   fontName=font, fontSize=16, alignment=1, spaceAfter=10, leading=20)
    info_s  = ParagraphStyle("inf", fontName=font, fontSize=12, alignment=2, spaceAfter=10)
    cell_s  = ParagraphStyle("c",   fontName=font, fontSize=10, leading=14)
    note_s  = ParagraphStyle("n",   fontName=font, fontSize=10, leading=16, spaceAfter=5)

    def strip_tags(txt):
        txt = re.sub(r'<br\s*/?>', '\n', str(txt))
        txt = re.sub(r'<[^>]+>', '', txt).strip()
        return txt

    def cell(txt):
        return Paragraph(strip_tags(txt).replace('\n', '<br/>'), cell_s)

    # 簡化原本的 HTML 解析邏輯，直接抓重點，提高容錯率
    body = html_content
    
    story = []

    # 1. 標題
    h2 = re.search(r'<h2[^>]*>(.*?)</h2>', body, re.DOTALL|re.IGNORECASE)
    if h2:
        story.append(Paragraph(strip_tags(h2.group(1)), title_s))
        story.append(Spacer(1, 2*mm))

    # 2. 時間資訊
    info = re.search(r"<div class='info'>(.*?)</div>", body, re.DOTALL|re.IGNORECASE)
    if info:
        story.append(Paragraph(strip_tags(info.group(1)), info_s))
        story.append(Spacer(1, 2*mm))

    # 3. 表格處理
    tables = re.findall(r'<table[^>]*>(.*?)</table>', body, re.DOTALL|re.IGNORECASE)
    
    # 處理表格樣式
    ts = TableStyle([
            ('FONTNAME',      (0,0),(-1,-1), font),
            ('FONTSIZE',      (0,0),(-1,-1), 10),
            ('GRID',          (0,0),(-1,-1), 0.5, colors.black),
            ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
            ('ALIGN',         (0,0),(-1,0), 'CENTER'), # 表頭置中
            ('BACKGROUND',    (0,0),(-1, 0), colors.HexColor('#f2f2f2')),
            ('TOPPADDING',    (0,0),(-1,-1), 4),
            ('BOTTOMPADDING', (0,0),(-1,-1), 4),
        ])

    for idx, tbl_html in enumerate(tables):
        rows_raw = re.findall(r'<tr[^>]*>(.*?)</tr>', tbl_html, re.DOTALL|re.IGNORECASE)
        data = []
        for row_html in rows_raw:
            # 抓取 th 或 td
            cells = re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', row_html, re.DOTALL|re.IGNORECASE)
            # 過濾掉 colspan (這是一個簡單處理，如果表格很複雜可能要調整)
            if cells:
                # 簡單過濾掉 "任務編組" 這種大標題行，避免欄位數不對齊導致報錯
                if "任務　編　組" in cells[0]: 
                    story.append(Paragraph(strip_tags(cells[0]), ParagraphStyle("th", fontName=font, fontSize=12, alignment=1, spaceAfter=2)))
                    continue 
                data.append([cell(c) for c in cells])
        
        if not data:
            continue
            
        # 計算欄位數
        col_n = len(data[0]) if data else 1
        # 建立表格
        t = Table(data, colWidths=[W/col_n]*col_n, repeatRows=1)
        t.setStyle(ts)
        story.append(t)
        story.append(Spacer(1, 4*mm))
        
        # 在第一個表格(指揮組)後插入勤教資料
        if idx == 0:
             # 尋找勤前教育區塊
            note_match = re.search(r"📢 勤前教育：</b>(.*?)</div>", body, re.DOTALL)
            station_match = re.search(r"🚧 (.*?)</div></div>", body, re.DOTALL)
            
            if note_match:
                story.append(Paragraph(f"<b>📢 勤前教育：</b>{note_match.group(1)}", note_s))
            if station_match:
                # 清理一下多餘標籤
                st_text = station_match.group(1).replace("<b>", "").replace("</b>", "") 
                story.append(Paragraph(f"<b>🚧 檢驗站：</b>{st_text}", note_s))
            
            story.append(Spacer(1, 4*mm))

    try:
        doc.build(story)
    except Exception as e:
        # 如果 PDF 生成失敗，印出錯誤但不讓程式掛掉
        print(f"PDF Build Error: {e}")
        return None

    return buf.getvalue()

def send_report_email(html_content, subject):
    import urllib.parse as _ul
    try:
        sender   = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        receiver = sender # 寄給自己
        
        pdf_bytes = _parse_html_to_pdf(html_content, subject)
        if pdf_bytes is None:
            return False, "PDF 生成失敗 (可能是字型問題)"

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


# --- 5. 初始化資料 (主程式邏輯) ---
df_set, df_cmd, df_ptl, error_msg = load_data()

# 處理資料讀取邏輯
if error_msg:
    st.error(f"❌ 無法讀取 Google Sheets：\n{error_msg}")
    st.warning("⚠️ 目前使用預設範本模式。")
    # 使用預設值
    current_unit    = DEFAULT_UNIT
    current_time    = DEFAULT_TIME
    current_proj    = DEFAULT_PROJ
    current_brief   = DEFAULT_BRIEF
    current_station = DEFAULT_STATION
    df_command_edit = DEFAULT_CMD.copy()
    df_patrol_edit  = DEFAULT_PTL.copy()
elif df_set is None: # 剛開始可能是空的
    st.info("💡 資料庫為空，載入範本。")
    current_unit    = DEFAULT_UNIT
    current_time    = DEFAULT_TIME
    current_proj    = DEFAULT_PROJ
    current_brief   = DEFAULT_BRIEF
    current_station = DEFAULT_STATION
    df_command_edit = DEFAULT_CMD.copy()
    df_patrol_edit  = DEFAULT_PTL.copy()
else:
    try:
        # 將設定轉為字典
        settings_dict = dict(zip(df_set.iloc[:, 0], df_set.iloc[:, 1]))
        current_unit    = settings_dict.get("unit_name",      DEFAULT_UNIT)
        current_time    = settings_dict.get("plan_full_time", DEFAULT_TIME)
        current_proj    = settings_dict.get("project_name",   DEFAULT_PROJ)
        current_brief   = settings_dict.get("briefing_info",  DEFAULT_BRIEF)
        current_station = settings_dict.get("check_station",  DEFAULT_STATION)
        # 如果 Sheet 是空的，就用預設值，否則用 Sheet 的值
        df_command_edit = df_cmd if not df_cmd.empty else DEFAULT_CMD.copy()
        df_patrol_edit  = df_ptl if not df_ptl.empty else DEFAULT_PTL.copy()
    except Exception as e:
        st.error(f"資料解析錯誤：{e}")
        st.stop()


# --- 6. 介面呈現 ---
st.subheader("1. 勤務基礎資訊")
c1, c2 = st.columns([1, 1]) # 調整比例
project_name = c1.text_input("專案名稱", value=current_proj)
plan_time    = c2.text_input("勤務時間", value=current_time)

st.subheader("2. 指揮與幕僚編組")
with st.expander("編輯名單 (指揮組)", expanded=True):
    # 注意：這裡不隱藏任務欄位，以免資料對應錯誤
    edited_cmd = st.data_editor(
        df_command_edit,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_cmd" # 加上 key 避免重整時狀態跑掉
    )

c3, c4 = st.columns(2)
brief_info = c3.text_area("📢 勤前教育",   value=current_brief,   height=100)
check_st   = c4.text_area("🚧 檢驗站資訊", value=current_station, height=100)

st.subheader("3. 執行勤務編組 (巡邏組)")
edited_ptl = st.data_editor(
    df_patrol_edit, 
    num_rows="dynamic", 
    use_container_width=True,
    key="editor_ptl"
)

# --- 7. 輸出 HTML 報表邏輯 (產生預覽用) ---
# 將產生 HTML 的邏輯封裝，方便維護
def generate_html_content():
    # 內嵌 CSS 樣式
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
        .rain-plan { color: blue; font-size: 0.9em; display: block; margin-top: 4px; }
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2>{current_unit}執行{project_name}規劃表</h2>"
    html += f"<div class='info'>勤務時間：{plan_time}</div>"
    
    # 指揮組表格
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr>"
    html += "<tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, row in edited_cmd.iterrows():
        name = str(row.get('姓名', '')).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        html += f"<tr><td><b>{row.get('職稱','')}</b></td><td>{row.get('代號','')}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務','')}</td></tr>"
    html += "</table>"
    
    # 勤教與檢驗站
    html += f"<div class='left-align' style='margin-bottom:20px;line-height:1.6'>"
    html += f"<div><b>📢 勤前教育：</b>{brief_info}</div>"
    html += f"<div style='white-space:pre-wrap'><b>🚧 {check_st}</b></div></div>"
    
    # 巡邏組表格
    html += "<table><tr><th width='10%'>編組</th><th width='8%'>代號</th><th width='12%'>單位</th><th width='18%'>服勤人員</th><th width='52%'>任務分工</th></tr>"
    for _, row in edited_ptl.iterrows():
        name = str(row.get('服勤人員', '')).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        unit_cell = str(row.get("單位","")).replace("、","<br>").replace(",","<br>")
        html += f"<tr><td>{row.get('編組','')}</td><td>{row.get('無線電','')}</td><td style='line-height:1.4'>{unit_cell}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務分工','')}<br><span style='color:blue;font-size:0.9em'>*雨備方案：轄區治安要點巡邏。</span></td></tr>"
    html += "</table></div></body></html>"
    return html

html_out = generate_html_content()

# --- 8. 輸出區域 ---
st.markdown("---")
col_view, col_dl = st.columns([3, 1])

with col_view:
    st.subheader("📄 即時預覽")
    st.components.v1.html(html_out, height=800, scrolling=True)

with col_dl:
    st.subheader("📥 存檔與輸出")
    
    # 下載按鈕邏輯
    # 注意：這裡使用 plan_time 等變數是介面上最新輸入的值
    file_name_date = datetime.now().strftime('%Y%m%d')
    
    if st.download_button(
        label="下載報表並同步雲端 💾",
        data=html_out.encode("utf-8"),
        file_name=f"勤務表_{file_name_date}.html",
        mime="text/html; charset=utf-8",
        type="primary"
    ):
        # 1. 先存檔
        save_success = save_data(
            current_unit, # 使用變數
            plan_time, 
            project_name, 
            brief_info, 
            check_st, 
            edited_cmd, 
            edited_ptl
        )
        
        # 2. 存檔成功才寄信
        if save_success:
            subject = f"噪音車勤務規劃表_{file_name_date}"
            ok, err = send_report_email(html_out, subject)
            if ok:
                st.toast("📧 報表已寄出至信箱！", icon="✉️")
            else:
                st.error(f"❌ 寄信失敗：{err}")

    st.info("💡 提示：請確保專案目錄下有 `kaiu.ttf` 字型檔，否則 PDF 中文會顯示異常。")
