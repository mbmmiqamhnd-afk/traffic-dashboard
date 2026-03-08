import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# --- 1. 頁面設定 ---
st.set_page_config(page_title="雲端勤務規劃", layout="wide")
st.title("🚓 專案勤務規劃表 (雲端同步版)")
st.caption("資料與 Google Sheets 即時連線，手機、電腦皆可編輯")

SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# --- 預設範本資料 ---
DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "115年3月20日19至23時"
DEFAULT_PROJ    = "0320「取締改裝(噪音)車輛專案監、警、環聯合稽查勤務」"
DEFAULT_BRIEF   = "19時30分於分局二樓會議室召開"
DEFAULT_STATION = "環保局臨時檢驗站開設時間：20時至23時\n地點：桃園市龍潭區大昌路一段277號（龍潭區警政聯合辦公大樓）廣場"

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官",       "代號": "隆安1",    "姓名": "分局長 施宇峰",                                      "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官",     "代號": "隆安2",    "姓名": "副分局長 何憶雯",                                    "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官",     "代號": "隆安3",    "姓名": "副分局長 蔡志明",                                    "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官",   "代號": "駐區督察", "姓名": "孫三陽",                                             "任務": "重點機動督導。"},
    {"職稱": "督導組",       "代號": "隆安6",    "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組",       "代號": "隆安684",  "姓名": "督察組教官 郭文義",                                  "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13",   "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、勤指中心警員 張庭溱（代理人：巡官陳鵬翔）、行政組警務佐 曾威仁、人事室警員 陳明祥", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"},
    {"職稱": "通訊組",       "代號": "隆安",     "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興",           "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_PTL = pd.DataFrame([
    {"編組": "第一巡邏組", "無線電": "隆安54",  "單位": "聖亭所",      "服勤人員": "巡佐傅錫城、警員曾建凱",      "任務分工": "於大昌路一段周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第二巡邏組", "無線電": "隆安62",  "單位": "龍潭所",      "服勤人員": "副所長全楚文、警員龔品璇",    "任務分工": "於大昌路二段周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第三巡邏組", "無線電": "隆安72",  "單位": "中興所",      "服勤人員": "副所長薛德祥、警員冷柔萱",    "任務分工": "於中興路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第四巡邏組", "無線電": "隆安83",  "單位": "石門所",      "服勤人員": "巡佐林偉政、警員盧瑾瑤",      "任務分工": "於北龍路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第五巡邏組", "無線電": "隆安33",  "單位": "三和所、高平所","服勤人員": "警員唐銘聰、警員張湃柏",     "任務分工": "於大昌路一、二段、北龍路及中興路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
    {"編組": "第六巡邏組", "無線電": "隆安994", "單位": "龍潭交通分隊", "服勤人員": "小隊長林振生、警員吳沛軒",   "任務分工": "於大昌路一、二段、北龍路及中興路周邊易有噪音車輛滋擾、聚集路段機動巡查改裝噪音車輛。"},
])

# --- 2. 建立 gspread 連線 ---
def get_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

# --- 3. 讀取函數 ---
def load_data():
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)
        df_settings = pd.DataFrame(sh.worksheet("設定").get_all_records())
        df_command  = pd.DataFrame(sh.worksheet("指揮組").get_all_records())
        df_patrol   = pd.DataFrame(sh.worksheet("巡邏組").get_all_records())
        return df_settings, df_command, df_patrol, None
    except Exception as e:
        return None, None, None, str(e)

# --- 4. 寫入函數 ---
def save_data(unit, time_str, project, briefing, station, df_cmd, df_ptl):
    try:
        client = get_client()
        sh = client.open_by_key(SHEET_ID)

        ws_set = sh.worksheet("設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"],
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

        st.toast("✅ 雲端存檔成功！", icon="☁️")
        return True
    except Exception as e:
        st.error(f"❌ 存檔失敗：{e}")
        return False

# --- 5. 初始化資料 ---
df_set, df_cmd, df_ptl, error_msg = load_data()

if error_msg:
    st.error(f"❌ 無法讀取 Google Sheets：\n{error_msg}")
    st.warning("⚠️ 目前使用預設範本模式，請修改後按「下載報表」自動儲存。")
    current_unit    = DEFAULT_UNIT
    current_time    = DEFAULT_TIME
    current_proj    = DEFAULT_PROJ
    current_brief   = DEFAULT_BRIEF
    current_station = DEFAULT_STATION
    df_command_edit = DEFAULT_CMD.copy()
    df_patrol_edit  = DEFAULT_PTL.copy()

elif df_set is None or df_set.empty:
    st.info("💡 尚無雲端資料，已載入預設範本，請修改後按「下載報表」自動儲存。")
    current_unit    = DEFAULT_UNIT
    current_time    = DEFAULT_TIME
    current_proj    = DEFAULT_PROJ
    current_brief   = DEFAULT_BRIEF
    current_station = DEFAULT_STATION
    df_command_edit = DEFAULT_CMD.copy()
    df_patrol_edit  = DEFAULT_PTL.copy()

else:
    try:
        settings_dict   = dict(zip(df_set.iloc[:, 0], df_set.iloc[:, 1]))
        current_unit    = settings_dict.get("unit_name",      DEFAULT_UNIT)
        current_time    = settings_dict.get("plan_full_time", DEFAULT_TIME)
        current_proj    = settings_dict.get("project_name",   DEFAULT_PROJ)
        current_brief   = settings_dict.get("briefing_info",  DEFAULT_BRIEF)
        current_station = settings_dict.get("check_station",  DEFAULT_STATION)
        df_command_edit = df_cmd if not df_cmd.empty else DEFAULT_CMD.copy()
        df_patrol_edit  = df_ptl if not df_ptl.empty else DEFAULT_PTL.copy()
    except Exception as e:
        st.error(f"資料格式解析失敗：{e}")
        st.stop()



unit_name = "桃園市政府警察局龍潭分局"
st.subheader("1. 勤務基礎資訊")
c1, c2 = st.columns(2)
project_name = c1.text_input("專案名稱", value=current_proj)
plan_time    = c2.text_input("勤務時間 (完整顯示文字)", value=current_time)

st.subheader("2. 指揮與幕僚編組")
st.caption("💡 姓名若有多人，請用「、」分隔，報表輸出時會自動變為「上下並列」。")
with st.expander("編輯名單", expanded=True):
    edited_cmd = st.data_editor(
        df_command_edit,
        num_rows="dynamic",
        use_container_width=True,
        column_config={"任務": None}  # 隱藏任務欄
    )
    # 確保任務欄資料不因隱藏而遺失
    if "任務" not in edited_cmd.columns:
        edited_cmd["任務"] = df_command_edit["任務"]

c3, c4 = st.columns(2)
brief_info = c3.text_area("📢 勤前教育",   value=current_brief,   height=100)
check_st   = c4.text_area("🚧 檢驗站資訊", value=current_station, height=100)

st.subheader("3. 執行勤務編組 (巡邏組)")
edited_ptl = st.data_editor(df_patrol_edit, num_rows="dynamic", use_container_width=True)

# --- 7. 輸出 HTML 報表 ---
def generate_html(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    style = """
    <style>
        body { font-family: 'DFKai-SB', 'BiauKai', '標楷體', serif; color: #000; }
        .container { width: 100%; max-width: 800px; margin: 0 auto; padding: 20px; }
        h2 { text-align: center; margin-bottom: 5px; letter-spacing: 2px; }
        .info { text-align: right; font-weight: bold; margin-bottom: 15px; font-size: 14px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid black; padding: 5px; text-align: center; font-size: 14px; vertical-align: middle; }
        th { background-color: #f2f2f2; }
        .left-align { text-align: left; }
        .rain-plan { color: blue; font-size: 0.9em; display: block; margin-top: 4px; }
        @media print { .no-print { display: none; } body { -webkit-print-color-adjust: exact; } }
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2>{unit}執行{project}規劃表</h2>"
    html += f"<div class='info'>勤務時間：{time_str}</div>"
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr>"
    html += "<tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, row in df_cmd.iterrows():
        name = str(row.get('姓名', '')).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        html += f"<tr><td><b>{row.get('職稱','')}</b></td><td>{row.get('代號','')}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務','')}</td></tr>"
    html += f"</table><div class='left-align' style='margin-bottom:20px;line-height:1.6'>"
    html += f"<div><b>📢 勤前教育：</b>{briefing}</div>"
    html += f"<div style='white-space:pre-wrap'><b>🚧 {station}</b></div></div>"
    html += "<table><tr><th width='10%'>編組</th><th width='8%'>代號</th><th width='12%'>單位</th><th width='18%'>服勤人員</th><th width='52%'>任務分工</th></tr>"
    for _, row in df_ptl.iterrows():
        name = str(row.get('服勤人員', '')).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        unit_cell = str(row.get("單位","")).replace("、","<br>").replace(",","<br>")
        html += f"<tr><td>{row.get('編組','')}</td><td>{row.get('無線電','')}</td><td style='line-height:1.4'>{unit_cell}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務分工','')}<br><span style='color:blue;font-size:0.9em'>*雨備方案：轄區治安要點巡邏。</span></td></tr>"
    html += "</table></div></body></html>"
    return html

html_out = generate_html(unit_name, project_name, plan_time, brief_info, check_st, edited_cmd, edited_ptl)

# --- 8. 輸出區域 ---
st.markdown("---")
col_view, col_dl = st.columns([3, 1])
with col_view:
    st.subheader("📄 即時預覽")
    st.components.v1.html(html_out, height=800, scrolling=True)
with col_dl:
    st.subheader("📥 輸出")
    # 按下下載時自動儲存到雲端
    if st.download_button(
        label="下載報表並同步雲端 💾",
        data=html_out.encode("utf-8"),
        file_name=f"勤務表_{datetime.now().strftime('%Y%m%d')}.html",
        mime="text/html; charset=utf-8",
        type="primary"
    ):
        save_data(unit_name, plan_time, project_name, brief_info, check_st, edited_cmd, edited_ptl)
    st.info("💡 下載後打開檔案，按 Ctrl+P 列印，網頁會自動隱藏選單，只印出表格。")
