import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# --- 1. 設定頁面與連線 ---
st.set_page_config(page_title="雲端勤務規劃", layout="wide")
st.title("🚓 專案勤務規劃表 (雲端同步版)")
st.caption("資料與 Google Sheets 即時連線，手機、電腦皆可編輯")

# 建立連線
conn = st.connection("gsheets", type=GSheetsConnection)

# --- 2. 讀取與寫入函數 ---

def load_data():
    """從 Google Sheets 讀取資料，若失敗則回傳空值"""
    try:
        # ttl=0 代表不快取，每次都抓最新的；使用 worksheet 指定分頁名稱
        df_settings = conn.read(worksheet="設定", ttl=0)
        df_command = conn.read(worksheet="指揮組", ttl=0)
        df_patrol = conn.read(worksheet="巡邏組", ttl=0)
        return df_settings, df_command, df_patrol
    except Exception as e:
        # 若發生錯誤 (例如還沒設定 secrets)，回傳 None
        return None, None, None

def save_data(unit, time_str, project, briefing, station, df_cmd, df_ptl):
    """將資料寫回 Google Sheets"""
    try:
        # 準備要寫入「設定」分頁的資料
        settings_data = [
            {"Key": "unit_name", "Value": unit},
            {"Key": "plan_full_time", "Value": time_str},
            {"Key": "project_name", "Value": project},
            {"Key": "briefing_info", "Value": briefing},
            {"Key": "check_station", "Value": station}
        ]
        df_settings_new = pd.DataFrame(settings_data)

        # 寫入三個分頁
        conn.update(worksheet="設定", data=df_settings_new)
        conn.update(worksheet="指揮組", data=df_cmd)
        conn.update(worksheet="巡邏組", data=df_ptl)
        
        st.success("✅ 雲端存檔成功！Google Sheets 已更新。")
        st.cache_data.clear() # 清除快取
    except Exception as e:
        st.error(f"❌ 存檔失敗，請檢查網路或權限。錯誤訊息：{e}")

# --- 3. 初始化資料邏輯 ---
df_set, df_cmd, df_ptl = load_data()

# 如果連線失敗或試算表是空的，使用預設值 (避免程式崩潰)
if df_set is None or df_set.empty:
    st.warning("⚠️ 尚未連接 Google Sheets 或讀取失敗，目前使用「本機預設值」。(請檢查 .streamlit/secrets.toml)")
    current_unit = "桃園市政府警察局龍潭分局"
    current_time = "115年2月26日19至23時"
    current_proj = "0226「取締改裝(噪音)車輛專案」"
    current_brief = "19時30分於分局二樓會議室召開"
    current_station = "時間：20時至23時\n地點：龍潭區大昌路一段277號"
    
    # 預設空表格
    df_command_edit = pd.DataFrame([{"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長", "任務": "督導"}])
    df_patrol_edit = pd.DataFrame([{"編組": "第一組", "無線電": "隆安54", "單位": "聖亭所", "服勤人員": "警員", "任務分工": "巡邏", "雨備": "巡邏"}])
else:
    # 解析設定檔 (將 Key-Value 轉回變數)
    try:
        # 確保欄位名稱正確，防止 CSV 標題錯誤
        settings_dict = dict(zip(df_set.iloc[:, 0], df_set.iloc[:, 1]))
        current_unit = settings_dict.get("unit_name", "")
        current_time = settings_dict.get("plan_full_time", "")
        current_proj = settings_dict.get("project_name", "")
        current_brief = settings_dict.get("briefing_info", "")
        current_station = settings_dict.get("check_station", "")
        df_command_edit = df_cmd
        df_patrol_edit = df_ptl
    except:
        st.error("試算表格式有誤，請確認「設定」分頁有 Key 與 Value 兩欄。")
        st.stop()

# --- 4. 介面編輯區 ---

# 側邊欄：儲存按鈕
with st.sidebar:
    st.header("💾 雲端操作")
    st.info("修改完後請按下方按鈕，資料才會同步到 Google Sheets。")
    save_btn = st.button("儲存並同步到雲端", type="primary")

# 主畫面輸入
st.subheader("1. 勤務基礎資訊")
c1, c2 = st.columns(2)
unit_name = c1.text_input("執行單位", value=current_unit)
plan_time = c1.text_input("勤務時間", value=current_time)
project_name = c2.text_input("專案名稱", value=current_proj)

st.subheader("2. 指揮與幕僚編組")
with st.expander("編輯名單 (支援多人姓名用「、」分隔)", expanded=True):
    edited_cmd = st.data_editor(df_command_edit, num_rows="dynamic", use_container_width=True)

c3, c4 = st.columns(2)
brief_info = c3.text_area("勤前教育", value=current_brief, height=80)
check_st = c4.text_area("檢驗站", value=current_station, height=80)

st.subheader("3. 執行勤務編組")
edited_ptl = st.data_editor(df_patrol_edit, num_rows="dynamic", use_container_width=True)

# 執行儲存
if save_btn:
    save_data(unit_name, plan_time, project_name, brief_info, check_st, edited_cmd, edited_ptl)
    st.rerun() # 重新整理頁面

# --- 5. 輸出 HTML 報表 (保持原本漂亮的格式) ---
def generate_html(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    style = """
    <style>
        body { font-family: 'DFKai-SB', 'BiauKai', '標楷體', serif; color: #000; }
        .container { width: 100%; max-width: 800px; margin: 0 auto; padding: 20px; }
        h2 { text-align: center; margin-bottom: 5px; letter-spacing: 2px; }
        .info { text-align: center; font-weight: bold; margin-bottom: 15px; font-size: 16px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid black; padding: 8px; text-align: center; font-size: 14px; vertical-align: middle; }
        th { background-color: #f2f2f2; }
        .left-align { text-align: left; }
        .rain-plan { color: blue; font-size: 0.9em; display: block; margin-top: 4px; }
    </style>
    """
    
    html = f"""
    <html><head>{style}</head><body><div class="container">
        <h2>{unit}執行{project}規劃表</h2>
        <div class="info">勤務時間：{time_str}</div>
        <table>
            <tr><th colspan="4">任　務　編　組</th></tr>
            <tr><th width="15%">職稱</th><th width="10%">代號</th><th width="25%">姓名</th><th width="50%">任務</th></tr>
    """
    for _, row in df_cmd.iterrows():
        formatted_name = str(row.get('姓名', '')).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        html += f"""<tr><td><b>{row.get('職稱','')}</b></td><td>{row.get('代號','')}</td>
                    <td style="line-height: 1.4;">{formatted_name}</td><td class="left-align">{row.get('任務','')}</td></tr>"""
    
    html += f"""</table>
        <div class="left-align" style="margin-bottom: 20px; line-height: 1.6;">
            <div><b>📢 勤前教育：</b>{briefing}</div>
            <div style="white-space: pre-wrap;"><b>🚧 {station}</b></div>
        </div>
        <table>
            <tr><th width="12%">編組</th><th width="10%">代號</th><th width="15%">單位</th><th width="20%">服勤人員</th><th width="43%">任務分工 / 雨備方案</th></tr>
    """
    for _, row in df_ptl.iterrows():
        formatted_ptl_name = str(row.get('服勤人員','')).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        rain = row.get('雨備', '')
        rain_text = f"<span class='rain-plan'>*雨備：{rain}</span>" if rain and str(rain) != "nan" else ""
        html += f"""<tr><td>{row.get('編組','')}</td><td>{row.get('無線電','')}</td><td>{row.get('單位','')}</td>
                    <td style="line-height: 1.4;">{formatted_ptl_name}</td><td class="left-align">{row.get('任務分工','')}{rain_text}</td></tr>"""
    
    html += "</table></div></body></html>"
    return html

html_out = generate_html(unit_name, project_name, plan_time, brief_info, check_st, edited_cmd, edited_ptl)

st.markdown("---")
col_view, col_dl = st.columns([3, 1])
with col_view:
    st.subheader("📄 預覽")
    st.components.v1.html(html_out, height=600, scrolling=True)
with col_dl:
    st.subheader("📥 輸出")
    st.download_button("下載報表 (.html)", html_out, f"勤務表_{datetime.now().strftime('%Y%m%d')}.html", "text/html")
