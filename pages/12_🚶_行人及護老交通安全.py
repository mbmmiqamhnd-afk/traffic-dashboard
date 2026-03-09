import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# --- 1. 頁面設定 ---
st.set_page_config(page_title="行人及護老交通安全", layout="wide")
st.title("🚶 行人及護老交通安全專案勤務規劃表")
st.caption("資料與 Google Sheets 即時連線，手機、電腦皆可編輯")

SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# --- 預設範本 ---
DEFAULT_MONTH = "115年3月份"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官",       "代號": "隆安1",    "姓名": "分局長 施宇峰",                                       "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官",     "代號": "隆安2",    "姓名": "副分局長 何憶雯",                                     "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官",     "代號": "隆安3",    "姓名": "副分局長 蔡志明",                                     "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官",   "代號": "駐區督察", "姓名": "孫三陽",                                              "任務": "重點機動督導。"},
    {"職稱": "督導組",       "代號": "隆安6",    "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組",       "代號": "隆安684",  "姓名": "督察組教官 郭文義",                                   "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13",   "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、秘書室巡官 陳鵬翔（代理人：警員張庭溱）、人事室警員 陳明祥、行政組警務佐 曾威仁", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組",       "代號": "隆安",     "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興",            "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "聖亭派出所",   "路段": "中豐路、聖亭路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "龍潭派出所",   "路段": "中豐路、中正路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "中興派出所",   "路段": "中興路、福龍路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "石門派出所",   "路段": "中正、文化路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "高平派出所",   "路段": "中豐、中原路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "三和派出所",   "路段": "龍新路、楊銅路段\n校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "警備隊",       "路段": "校園周邊道路或轄區行人易肇事路口"},
    {"日期（6時至10時、16時至20時）": "", "單位": "龍潭交通分隊", "路段": "校園周邊道路或轄區行人易肇事路口"},
])

NOTES = """壹、警察局規劃3月份「行人及護老交通安全專案勤務」期程：
一、3月6日（星期五）6至10時、16至20時。
二、3月12日（星期四）6至10時、16至20時。
三、3月24日（星期二）6至10時、16至20時。
四、3月30日（星期一）6至10時、16至20時。
貳、執行本專案勤務視轄區狀況及執勤警力，擇定轄區易肇事路口（段）及校園周邊道路，依上揭日期妥適編排勤務（必要時得另行規劃專案）協助維護行人、學童及高齡者通行安全，並加強取締「車不讓人」、「未依規定停讓」、「違規（臨時）停車」、「行人違反路權」及「道路障礙」等違規，必要時得合併相關勤務實施，以達「一種勤務多種功能」之效益。
叁、執行「行人及護老交通安全實施計畫」合強化違規取締項目：
一、車不讓人（第44條第1項第2款、第2項、第3項、第45條第1項第6款）
二、違規（臨時）停車（第55條、第56條）
三、行人（含代步器、電動輪椅）違反路權（第78條、第80條）
四、道路障礙（第82條）"""

# --- 2. gspread 連線 ---
def get_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

# --- 寄信函數 ---
def send_report_email(html_content, subject):
    import urllib.parse as _ul
    try:
        from weasyprint import HTML as _WHTML
        import os as _os
        sender   = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        receiver = sender
        # 找 kaiu.ttf 絕對路徑
        _font_path = None
        for _p in [
            _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), '..', 'kaiu.ttf'),
            _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'kaiu.ttf'),
            '/mount/src/traffic-dashboard/kaiu.ttf',
        ]:
            if _os.path.exists(_os.path.normpath(_p)):
                _font_path = _os.path.normpath(_p)
                break
        _base_url = ('file://' + _os.path.dirname(_font_path)) if _font_path else '.'
        pdf_bytes = _WHTML(string=html_content).write_pdf()
        msg = MIMEMultipart()
        msg["From"]    = sender
        msg["To"]      = receiver
        msg["Subject"] = subject
        msg.attach(MIMEText("請見附件 PDF 報表。", "plain", "utf-8"))
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






# --- 3. 讀取 ---
def load_data():
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        df_settings = pd.DataFrame(sh.worksheet("護老_設定").get_all_records())
        df_cmd      = pd.DataFrame(sh.worksheet("護老_指揮組").get_all_records())
        df_schedule = pd.DataFrame(sh.worksheet("護老_勤務表").get_all_records())
        return df_settings, df_cmd, df_schedule, None
    except Exception as e:
        return None, None, None, str(e)

# --- 4. 寫入 ---
def save_data(month, df_cmd, df_schedule):
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)

        ws_set = sh.worksheet("護老_設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["month", month]])

        ws_cmd = sh.worksheet("護老_指揮組")
        ws_cmd.clear()
        df_cmd = df_cmd.fillna("")
        ws_cmd.update([df_cmd.columns.tolist()] + df_cmd.values.tolist())

        ws_sch = sh.worksheet("護老_勤務表")
        ws_sch.clear()
        df_schedule = df_schedule.fillna("")
        ws_sch.update([df_schedule.columns.tolist()] + df_schedule.values.tolist())

        st.toast("✅ 雲端存檔成功！", icon="☁️")
        return True
    except Exception as e:
        st.error(f"❌ 存檔失敗：{e}")
        return False

# --- 5. 初始化 ---
df_set, df_cmd, df_sch, error_msg = load_data()

if error_msg or df_set is None or df_set.empty:
    if error_msg:
        st.error(f"❌ 無法讀取 Google Sheets：\n{error_msg}")
    st.info("💡 已載入預設範本，請修改後按「下載報表」自動儲存。")
    current_month    = DEFAULT_MONTH
    df_cmd_edit      = DEFAULT_CMD.copy()
    df_schedule_edit = DEFAULT_SCHEDULE.copy()
else:
    try:
        sd = dict(zip(df_set.iloc[:, 0], df_set.iloc[:, 1]))
        current_month    = sd.get("month", DEFAULT_MONTH)
        df_cmd_edit      = df_cmd if not df_cmd.empty else DEFAULT_CMD.copy()
        df_schedule_edit = df_sch if not df_sch.empty else DEFAULT_SCHEDULE.copy()
    except Exception as e:
        st.error(f"資料格式解析失敗：{e}")
        st.stop()

# --- 6. 介面 ---
st.subheader("1. 基礎資訊")
current_month = st.text_input("月份", value=current_month)

st.subheader("2. 任務編組")
st.caption("💡 姓名若有多人，請用「、」分隔。")
with st.expander("編輯名單", expanded=True):
    edited_cmd = st.data_editor(
        df_cmd_edit,
        num_rows="dynamic",
        use_container_width=True,
        column_config={"任務": None}
    )
    if "任務" not in edited_cmd.columns:
        edited_cmd["任務"] = df_cmd_edit["任務"]

st.subheader("3. 執行勤務日期、單位及路段")
edited_schedule = st.data_editor(df_schedule_edit, num_rows="dynamic", use_container_width=True)

st.subheader("4. 備註（固定）")
st.text(NOTES)

# --- 7. 產生 HTML ---
def generate_html(month, df_cmd, df_schedule):
    style = """

    <style>
        body { font-family: 'Noto Serif CJK TC', 'Noto Sans CJK TC', serif; color: #000; font-size: 14px; }
        .container { width: 100%; max-width: 800px; margin: 0 auto; padding: 20px; }
        h2 { text-align: left; margin-bottom: 5px; letter-spacing: 2px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
        th, td { border: 1px solid black; padding: 5px; text-align: center; font-size: 14px; vertical-align: middle; }
        th { background-color: #f2f2f2; }
        .left-align { text-align: left; }
        .section { margin-bottom: 10px; line-height: 1.8; }
        .notes { white-space: pre-wrap; font-size: 13px; line-height: 1.8; }
        @media print { .no-print { display: none; } body { -webkit-print-color-adjust: exact; } }
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2>{UNIT}{month}執行「行人及護老交通安全」專案勤務規劃表</h2>"

    # 任務編組
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr>"
    html += "<tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, row in df_cmd.iterrows():
        name = str(row.get('姓名', '')).replace("、", "<br>").replace(",", "<br>")
        html += f"<tr><td><b>{row.get('職稱','')}</b></td><td>{row.get('代號','')}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務','')}</td></tr>"
    html += "</table>"

    # 執行勤務表
    html += "<table>"
    html += "<tr><th colspan='3' style='background-color:#f2f2f2;text-align:center;'>警　力　佈　署</th></tr>"
    html += "<tr><th width='25%'>執行勤務日期（6時至10時、16時至20時）</th><th width='20%'>單位</th><th width='55%'>路段</th></tr>"
    for _, row in df_schedule.iterrows():
        road = str(row.get('路段', '')).replace("\n", "<br>")
        html += f"<tr><td>{row.get('日期（6時至10時、16時至20時）','')}</td><td>{row.get('單位','')}</td><td class='left-align'>{road}</td></tr>"
    html += "</table>"

    # 備註
    html += f"<div class='section'><b>備註</b><br><span class='notes'>{NOTES}</span></div>"

    html += "</div></body></html>"
    return html

html_out = generate_html(current_month, edited_cmd, edited_schedule)

# --- 8. 輸出 ---
st.markdown("---")
col_view, col_dl = st.columns([3, 1])
with col_view:
    st.subheader("📄 即時預覽")
    st.components.v1.html(html_out, height=800, scrolling=True)
with col_dl:
    st.subheader("📥 輸出")
    if st.download_button(
        label="下載報表並同步雲端 💾",
        data=html_out.encode("utf-8"),
        file_name=f"護老勤務表_{datetime.now().strftime('%Y%m%d')}.html",
        mime="text/html; charset=utf-8",
        type="primary"
    ):
        save_data(current_month, edited_cmd, edited_schedule)
        subject = f"護老交通安全勤務規劃表_{datetime.now().strftime('%Y%m%d')}"
        ok, err = send_report_email(html_out, subject)
        if ok:
            st.toast("📧 報表已寄出至信箱！", icon="✉️")
        else:
            st.error(f"❌ 寄信失敗：{err}")
    st.info("💡 下載後打開檔案，按 Ctrl+P 列印。")
