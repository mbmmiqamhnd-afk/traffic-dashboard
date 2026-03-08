# -*- coding: utf-8 -*-
import sys
import io

# 強制設定 UTF-8 編碼，避免中文字元引發 ASCII 錯誤
if hasattr(sys.stdout, 'buffer') and sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
if hasattr(sys.stderr, 'buffer') and sys.stderr.encoding != 'utf-8':
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# --- 1. 頁面設定 ---
st.set_page_config(page_title="雲端勤務規劃", layout="wide")
st.title("🚓 專案勤務規劃表 (雲端同步版)")
st.caption("資料與 Google Sheets 即時連線，手機、電腦皆可編輯")

# 建立 Google Sheets 連線
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"⚠️ 連線設定錯誤：請檢查 Secrets 設定。\n錯誤訊息：{e}")
    st.stop()

# --- 2. 讀取與寫入函數 ---

def load_data():
    """從 Google Sheets 讀取資料，若失敗則回傳詳細錯誤"""
    try:
        # ttl=0 代表不快取，每次都抓最新的
        # 使用 worksheet 指定分頁名稱，這三個名字必須跟 Google Sheet 一模一樣
        df_settings = conn.read(worksheet="設定", ttl=0)
        df_command = conn.read(worksheet="指揮組", ttl=0)
        df_patrol = conn.read(worksheet="巡邏組", ttl=0)
        return df_settings, df_command, df_patrol, None
    except Exception as e:
        # 回傳錯誤訊息
        return None, None, None, str(e)

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
        
        st.toast("✅ 雲端存檔成功！", icon="☁️")
        st.cache_data.clear() # 清除快取確保下次讀到新的
        return True
    except Exception as e:
        st.error(f"❌ 存檔失敗：{e}")
        return False

# --- 3. 初始化資料邏輯 ---
df_set, df_cmd, df_ptl, error_msg = load_data()

# 判斷讀取狀態
if error_msg:
    st.error(f"❌ 無法讀取 Google Sheets，原因如下：\n{error_msg}")
    st.info("💡 請檢查：\n1. Google Sheet 下方的分頁名稱是否已改為「設定」、「指揮組」、「巡邏組」？\n2. 試算表是否已開啟「知道連結者可編輯」權限？\n3. Secrets 的網址是否正確？")
    
    # 使用預設值讓程式繼續執行，不崩潰
    st.warning("⚠️ 目前暫時使用「本機預設範本」模式")
    current_unit = "桃園市政府警察局龍潭分局"
    current_time = "115年2月26日19至23時"
    current_proj = "0226「取締改裝(噪音)車輛專案監、警、環聯合稽查勤務」"
    current_brief = "19時30分於分局二樓會議室召開"
    current_station = "時間：20時至23時\n地點：龍潭區大昌路一段277號"
    
    # 預設範本
    df_command_edit = pd.DataFrame([
        {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
        {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"}
    ])
    df_patrol_edit = pd.DataFrame([
        {"編組": "第一巡邏組", "無線電": "隆安54", "單位": "聖亭所", "服勤人員": "巡佐傅錫城、警員曾建凱", "任務分工": "於大昌路一段周邊易有噪音車輛滋擾、聚集路段機動巡查。", "雨備": "轄區治安要點巡邏。"}
    ])

elif df_set is None or df_set.empty:
    # 連到了但裡面沒資料 (全空)
    st.warning("⚠️ Google Sheets 連線成功，但內容是空的。請先按一次「儲存」來初始化格式。")
    current_unit = "桃園市政府警察局龍潭分局"
    current_time = "115年2月26日19至23時"
    current_proj = "0226「取締改裝(噪音)車輛專案監、警、環聯合稽查勤務」"
    current_brief = "19時30分於分局二樓會議室召開"
    current_station = "時間：20時至23時\n地點：龍潭區大昌路一段277號"
    df_command_edit = pd.DataFrame([{"職稱": "職稱", "代號": "代號", "姓名": "姓名", "任務": "任務"}])
    df_patrol_edit = pd.DataFrame([{"編組": "編組", "無線電": "無線電", "單位": "單位", "服勤人員": "人員", "任務分工": "任務", "雨備": "雨備"}])

else:
    # --- 成功連線且有資料 ---
    try:
        # 將設定檔轉為字典
        settings_dict = dict(zip(df_set.iloc[:, 0], df_set.iloc[:, 1]))
        current_unit = settings_dict.get("unit_name", "")
        current_time = settings_dict.get("plan_full_time", "")
        current_proj = settings_dict.get("project_name", "")
        current_brief = settings_dict.get("briefing_info", "")
        current_station = settings_dict.get("check_station", "")
        df_command_edit = df_cmd
        df_patrol_edit = df_ptl
    except Exception as parse_error:
        st.error(f"資料格式解析失敗：{parse_error}")
        st.stop()

# --- 4. 介面編輯區 ---

# 側邊欄：儲存按鈕
with st.sidebar:
    st.header("💾 雲端操作")
    st.info("修改後請點擊下方按鈕，將資料同步回 Google Sheets。")
    if st.button("儲存並同步到雲端", type="primary"):
        # 設定 session_state 標記，在下方執行儲存
        st.session_state['do_save'] = True

# 主畫面輸入
st.subheader("1. 勤務基礎資訊")
c1, c2 = st.columns(2)
unit_name = c1.text_input("執行單位", value=current_unit)
plan_time = c1.text_input("勤務時間 (完整顯示文字)", value=current_time)
project_name = c2.text_input("專案名稱", value=current_proj)

st.subheader("2. 指揮與幕僚編組")
st.caption("💡 姓名若有多人，請用「、」或「,」分隔，報表輸出時會自動變為「上下並列」。")
with st.expander("編輯名單", expanded=True):
    edited_cmd = st.data_editor(df_command_edit, num_rows="dynamic", use_container_width=True)

c3, c4 = st.columns(2)
brief_info = c3.text_area("📢 勤前教育", value=current_brief, height=100)
check_st = c4.text_area("🚧 檢驗站資訊", value=current_station, height=100)

st.subheader("3. 執行勤務編組 (巡邏組)")
edited_ptl = st.data_editor(df_patrol_edit, num_rows="dynamic", use_container_width=True)

# 執行儲存邏輯
if st.session_state.get('do_save', False):
    success = save_data(unit_name, plan_time, project_name, brief_info, check_st, edited_cmd, edited_ptl)
    st.session_state['do_save'] = False # 重置按鈕狀態
    if success:
        st.rerun() # 重新整理頁面以顯示最新狀態

# --- 5. 輸出 HTML 報表邏輯 ---
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
        @media print {
            .no-print { display: none; }
            body { -webkit-print-color-adjust: exact; }
        }
    </style>
    """
    
    html = f"""
    <html><head><meta charset="utf-8">{style}</head><body><div class="container">
        <h2>{unit}執行{project}規劃表</h2>
        <div class="info">勤務時間：{time_str}</div>
        <table>
            <tr><th colspan="4">任　務　編　組</th></tr>
            <tr><th width="15%">職稱</th><th width="10%">代號</th><th width="25%">姓名</th><th width="50%">任務</th></tr>
    """
    for _, row in df_cmd.iterrows():
        # 處理多姓名換行
        formatted_name = str(row.get('姓名', '')).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        html += f"""<tr>
            <td><b>{row.get('職稱','')}</b></td>
            <td>{row.get('代號','')}</td>
            <td style="line-height: 1.4;">{formatted_name}</td>
            <td class="left-align">{row.get('任務','')}</td>
        </tr>"""
    
    html += f"""</table>
        <div class="left-align" style="margin-bottom: 20px; line-height: 1.6;">
            <div><b>📢 勤前教育：</b>{briefing}</div>
            <div style="white-space: pre-wrap;"><b>🚧 {station}</b></div>
        </div>
        <table>
            <tr><th width="12%">編組</th><th width="10%">代號</th><th width="15%">單位</th><th width="20%">服勤人員</th><th width="43%">任務分工 / 雨備方案</th></tr>
    """
    for _, row in df_ptl.iterrows():
        # 處理多姓名換行
        formatted_ptl_name = str(row.get('服勤人員','')).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        rain = row.get('雨備', '')
        rain_text = f"<span class='rain-plan'>*雨備：{rain}</span>" if rain and str(rain) != "nan" else ""
        html += f"""<tr>
            <td>{row.get('編組','')}</td>
            <td>{row.get('無線電','')}</td>
            <td>{row.get('單位','')}</td>
            <td style="line-height: 1.4;">{formatted_ptl_name}</td>
            <td class="left-align">{row.get('任務分工','')}{rain_text}</td>
        </tr>"""
    
    html += "</table></div></body></html>"
    return html

# 產生報表
html_out = generate_html(unit_name, plan_time, project_name, brief_info, check_st, edited_cmd, edited_ptl)

# --- 6. 輸出區域 ---
st.markdown("---")
col_view, col_dl = st.columns([3, 1])
with col_view:
    st.subheader("📄 即時預覽")
    st.components.v1.html(html_out, height=800, scrolling=True)
with col_dl:
    st.subheader("📥 輸出")
    st.download_button(
        label="下載報表 (.html)",
        data=html_out.encode("utf-8"),   # 明確指定 UTF-8 編碼輸出
        file_name=f"勤務表_{datetime.now().strftime('%Y%m%d')}.html",
        mime="text/html; charset=utf-8"
    )
    st.info("💡 下載後打開檔案，按 Ctrl+P (列印)，網頁會自動隱藏選單，只印出表格。")
