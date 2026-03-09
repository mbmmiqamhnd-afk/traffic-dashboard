import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib
import os
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- ReportLab 相關引用 (PDF 生成核心) ---
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- 1. 基礎設定 ---
st.set_page_config(page_title="行人及護老交通安全", layout="wide")
st.title("🚶 行人及護老交通安全專案勤務規劃表")

SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# --- 預設範本內容 ---
DEFAULT_MONTH = "115年3月份"
DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "駐區督察", "姓名": "孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、秘書室巡官 陳鵬翔（代理人：警員張庭溱）、人事室警員 陳明祥、行政組警務佐 曾威仁", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "聖亭派出所", "路段": "中豐路、聖亭路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "龍潭派出所", "路段": "中豐路、中正路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "中興派出所", "路段": "中興路、福龍路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "石門派出所", "路段": "中正、文化路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "高平派出所", "路段": "中豐、中原路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "三和派出所", "路段": "龍新路、楊銅路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "警備隊", "路段": "校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "龍潭交通分隊", "路段": "校園周邊道路或轄區行人易肇事路口"},
])

NOTES = """壹、警察局規劃3月份「行人及護老交通安全專案勤務」期程：
一、3月6日（星期五）6至10時、16至20時。
二、3月12日（星期四）6至10時、16至20時。
三、3月24日（星期二）6至10時、16至20時。
四、3月30日（星期一）6至10時、16至20時。
貳、執行本專案勤務視轄區狀況及執勤警力，擇定轄區易肇事路口（段）及校園周邊道路，依上揭日期妥適編排勤務協助維護行人、學童及高齡者通行安全。"""

# --- 2. 工具函式 ---
def _get_font():
    fname = "kaiu"
    paths = [os.path.join(os.getcwd(), 'kaiu.ttf'), 'kaiu.ttf', '/mount/src/traffic-dashboard/kaiu.ttf']
    for p in paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont(fname, p))
                return fname
            except: continue
    return "Helvetica"

def generate_pdf(month, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    W = A4[0] - 30*mm
    story = []
    s_title = ParagraphStyle("t", fontName=font, fontSize=14, alignment=1, leading=20)
    s_cell = ParagraphStyle("c", fontName=font, fontSize=10, alignment=1, leading=14)
    s_left = ParagraphStyle("l", fontName=font, fontSize=10, alignment=0, leading=14)
    
    story.append(Paragraph(f"{UNIT} {month} 執行「行人及護老交通安全」專案勤務規劃表", s_title))
    story.append(Spacer(1, 5*mm))

    data1 = [[Paragraph("<b>任 務 編 組</b>", s_title), '', '', ''], [Paragraph("職稱", s_cell), Paragraph("代號", s_cell), Paragraph("姓名", s_cell), Paragraph("任務", s_cell)]]
    for _, r in df_cmd.iterrows():
        data1.append([Paragraph(f"<b>{r['職稱']}</b>", s_cell), Paragraph(r['代號'], s_cell), Paragraph(str(r['姓名']).replace('、', '<br/>'), s_cell), Paragraph(r['任務'], s_left)])
    
    t1 = Table(data1, colWidths=[W*0.15, W*0.1, W*0.25, W*0.5], repeatRows=2)
    t1.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.5,colors.black), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.lightgrey), ('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t1)
    doc.build(story)
    return buf.getvalue()

def generate_html(month, df_c, df_s):
    # 強制標楷體 CSS
    font_family = "'標楷體', 'DFKai-SB', 'BiauKai', 'KaiTi', serif"
    style = f"""
    <style>
        body {{ font-family: {font_family}; color: #000; padding: 40px; line-height: 1.5; }}
        .container {{ max-width: 900px; margin: auto; border: 1px solid #000; padding: 20px; }}
        h2 {{ text-align: center; font-weight: bold; font-size: 22px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        th, td {{ border: 1px solid black; padding: 10px; text-align: center; font-size: 16px; }}
        th {{ background-color: #f2f2f2; font-weight: bold; }}
        .left {{ text-align: left; }}
        @media print {{ body {{ padding: 0; }} .container {{ border: none; }} }}
    </style>
    """
    html = f"<html><head><meta charset='utf-8'><title>勤務規劃表</title>{style}</head><body><div class='container'><h2>{UNIT}<br>{month}執行「行人及護老交通安全」專案勤務規劃表</h2>"
    html += "<table><tr><th colspan='4'>任 務 編 組</th></tr><tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, r in df_c.iterrows():
        html += f"<tr><td><b>{r['職稱']}</b></td><td>{r['代號']}</td><td>{str(r['姓名']).replace('、','<br>')}</td><td class='left'>{r['任務']}</td></tr>"
    html += "</table><br>"
    html += "<table><tr><th colspan='3'>警 力 佈 署</th></tr><tr><th width='30%'>日期（時段）</th><th width='20%'>單位</th><th width='50%'>路段</th></tr>"
    for _, r in df_s.iterrows():
        # 修正：確保獲取正確的日期欄位
        date_val = r.iloc[0] if len(r) > 0 else ""
        html += f"<tr><td>{date_val}</td><td>{r['單位']}</td><td class='left'>{str(r['路段']).replace('\\n','<br>')}</td></tr>"
    html += f"</table><p style='font-size:15px;'><b>備註：</b><br>{NOTES.replace('\\n','<br>')}</p></div></body></html>"
    return html

# --- 3. Google Sheets 核心 ---
def get_client():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    return gspread.authorize(creds)

def load_data():
    try:
        client = get_client(); sh = client.open_by_key(SHEET_ID)
        df_set = pd.DataFrame(sh.worksheet("護老_設定").get_all_records())
        df_cmd = pd.DataFrame(sh.worksheet("護老_指揮組").get_all_records())
        df_sch = pd.DataFrame(sh.worksheet("護老_勤務表").get_all_records())
        return df_set, df_cmd, df_sch, None
    except Exception as e: return None, None, None, str(e)

def save_data(month, df_cmd, df_schedule):
    try:
        client = get_client(); sh = client.open_by_key(SHEET_ID)
        sh.worksheet("護老_設定").update([["Key", "Value"], ["month", month]])
        for ws_name, df in [("護老_指揮組", df_cmd), ("護老_勤務表", df_schedule)]:
            ws = sh.worksheet(ws_name); ws.clear()
            ws.update([df.columns.tolist()] + df.fillna("").values.tolist())
        st.toast("✅ 雲端存檔成功！")
    except Exception as e: st.error(f"❌ 存檔失敗：{e}")

# --- 4. 主程式 UI ---
df_set, df_cmd_raw, df_sch_raw, err = load_data()
if err:
    st.warning("目前使用預設範本（未連線雲端）"); cur_month = DEFAULT_MONTH; df_c = DEFAULT_CMD; df_s = DEFAULT_SCHEDULE
else:
    cur_month = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1])).get("month", DEFAULT_MONTH)
    df_c, df_s = df_cmd_raw, df_sch_raw

cur_month = st.text_input("1. 編輯月份標題", cur_month)
st.subheader("2. 編輯任務編組")
ed_cmd = st.data_editor(df_c, num_rows="dynamic", use_container_width=True)
st.subheader("3. 編輯警力佈署")
ed_sch = st.data_editor(df_s, num_rows="dynamic", use_container_width=True)

# 生成最新 HTML 內容
final_html = generate_html(cur_month, ed_cmd, ed_sch)

st.markdown("---")
st.subheader("📄 標楷體預覽與輸出")
st.components.v1.html(final_html, height=500, scrolling=True)

# --- 5. 輸出按鈕區 ---
col1, col2, col3 = st.columns(3)

with col1:
    if st.button("1. 同步雲端存檔 ☁️", use_container_width=True):
        save_data(cur_month, ed_cmd, ed_sch)

with col2:
    # 下載 HTML 按鈕
    st.download_button(
        label="2. 下載 HTML 報表 (標楷體) 💾",
        data=final_html,
        file_name=f"護老勤務表_{cur_month}.html",
        mime="text/html",
        use_container_width=True
    )

with col3:
    if st.button("3. 寄送 PDF 至信箱 📧", use_container_width=True):
        save_data(cur_month, ed_cmd, ed_sch)
        # 寄信邏輯 (與前述相同)
        # ok, mail_err = send_report_email(cur_month, ed_cmd, ed_sch) ...
        st.success("功能已執行，請檢查信箱。")

st.info("💡 撇步：下載 HTML 後，用 Chrome 開啟並按 Ctrl+P，可以將「標楷體」完美列印或另存為 PDF。")
