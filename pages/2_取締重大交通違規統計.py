import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import smtplib
import gspread
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# 強制清除快取
try:
    st.cache_data.clear()
    st.cache_resource.clear()
except: pass

st.set_page_config(page_title="取締重大交通違規統計", layout="wide", page_icon="🚔")
st.markdown("## 🚔 取締重大交通違規統計 (v71 標題空白修正版)")

# 初始化 Session State
if "sent_cache" not in st.session_state:
    st.session_state["sent_cache"] = set()

# --- 強制清除快取按鈕 ---
if st.button("🧹 清除快取 (若更新無效請按此)", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.session_state["sent_cache"] = set()
    st.success("快取已清除！請重新整理頁面 (F5) 並重新上傳檔案。")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所',
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所',
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '交通組': '科技執法'
}
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

TARGETS = {
    '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479,
    '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006
}

NOTE_TEXT = "重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項。"

# --- Google Sheets 工具 (省略重複的 get_precise_rich_text_req 與 get_color_only_req 以節省篇幅，請保留你原本代碼中的這兩段) ---
def get_precise_rich_text_req(sheet_id, row_idx, col_idx, text):
    text = str(text)
    tokens = re.split(r'([0-9\(\)\/\-\.\%\~\s:：\[\]]+)', text)
    runs = []
    current_pos = 0
    for token in tokens:
        if not token: continue
        color = {"red": 0, "green": 0, "blue": 0} 
        if re.match(r'^[0-9\(\)\/\-\.\%\~\s:：\[\]]+$', token):
            color = {"red": 1, "green": 0, "blue": 0}
        runs.append({"startIndex": current_pos, "format": {"foregroundColor": color, "bold": True}})
        current_pos += len(token)
    return {"updateCells": {"rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}], "fields": "userEnteredValue,textFormatRuns", "range": {"sheetId": sheet_id, "startRowIndex": row_idx, "endRowIndex": row_idx + 1, "startColumnIndex": col_idx, "endColumnIndex": col_idx + 1}}}

def get_color_only_req(sheet_id, row_index, col_index, is_red):
    color = {"red": 1.0, "green": 0.0, "blue": 0.0} if is_red else {"red": 0, "green": 0, "blue": 0}
    return {"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": row_index, "endRowIndex": row_index + 1, "startColumnIndex": col_index, "endColumnIndex": col_index + 1}, "cell": {"userEnteredFormat": {"textFormat": {"foregroundColor": color}}}, "fields": "userEnteredFormat.textFormat.foregroundColor"}}

def update_google_sheet(data_list, sheet_url):
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ 錯誤：未設定 Secrets！")
            return False
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(0)
        ws.update(range_name='A1', values=data_list)
        requests = []
        requests.append({"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 15, "startColumnIndex": 0, "endColumnIndex": 10}, "cell": {"userEnteredFormat": {"textFormat": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}}}}, "fields": "userEnteredFormat.textFormat.foregroundColor"}})
        requests.append(get_precise_rich_text_req(ws.id, 1, 1, data_list[1][1]))
        requests.append(get_precise_rich_text_req(ws.id, 1, 3, data_list[1][3]))
        requests.append(get_precise_rich_text_req(ws.id, 1, 5, data_list[1][5]))
        for i in range(3, len(data_list) - 1):
            row_data = data_list[i]
            unit_name = str(row_data[0]).strip()
            try:
                comp_val = float(str(row_data[7]).replace(',', ''))
            except: comp_val = 0
            if comp_val < 0:
                requests.append(get_color_only_req(ws.id, i, 7, True))
                if unit_name != "科技執法":
                    requests.append(get_color_only_req(ws.id, i, 0, True))
        sh.batch_update({'requests': requests})
        return True
    except Exception as e:
        st.error(f"❌ Google Sheet 錯誤: {e}")
        return False

# ==========================================
# 4. 核心解析函數 (v71 修改重點)
# ==========================================
def parse_focus_report(uploaded_file):
    if not uploaded_file: return None
    file_name = uploaded_file.name
    try:
        content = uploaded_file.getvalue()
        df_raw = pd.read_excel(io.BytesIO(content), header=None, nrows=25)
        
        start_date, end_date = "", ""
        header_idx = -1
        
        # 尋找日期與標題列
        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
            if not start_date:
                match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', row_str)
                if match: start_date, end_date = match.group(1), match.group(2)
            
            # 關鍵字判定：如果這列有「酒後」且不是標題所在的「第一列」
            if "酒後" in row_str or "闖紅燈" in row_str:
                header_idx = i
                if start_date: break
        
        if header_idx == -1:
            st.error(f"❌ {file_name}：找不到標題列 (請確認 Excel 內有「酒後」等字樣)")
            return None

        # 以 header_idx 讀取，並處理 A 欄空白標題
        df = pd.read_excel(io.BytesIO(content), header=header_idx)
        
        # 定義要抓的欄位
        keywords = ["酒後", "闖紅燈", "嚴重超速", "逆向", "轉彎", "蛇行", "不暫停讓行人", "機車"]
        stop_cols, cit_cols = [], []
        
        for i in range(len(df.columns)):
            col_name = str(df.columns[i])
            if any(k in col_name for k in keywords) and "路肩" not in col_name and "大型車" not in col_name:
                stop_cols.append(i)
                cit_cols.append(i + 1)
        
        unit_data = {}
        # 資料從 header_idx + 1 開始，對應 df 的內容
        for _, row in df.iterrows():
            # A 欄 (index 0) 是單位名稱
            raw_unit = str(row.iloc[0]).strip()
            if raw_unit in ['nan', 'None', '', '合計', '單位']: continue
            if "統計" in raw_unit: continue # 排除頁尾資訊
            
            unit_name = UNIT_MAP.get(raw_unit, raw_unit)
            s, c = 0, 0
            for col in stop_cols:
                try:
                    val = str(row.iloc[col]).replace(',', '')
                    s += float(val) if val != 'nan' else 0
                except: pass
            for col in cit_cols:
                try:
                    val = str(row.iloc[col]).replace(',', '')
                    c += float(val) if val != 'nan' else 0
                except: pass
            
            unit_data[unit_name] = {'stop': s, 'cit': c}

        # 計算天數
        duration = 0
        if start_date and end_date:
            try:
                s_d, e_d = re.sub(r'[^\d]', '', start_date), re.sub(r'[^\d]', '', end_date)
                d1 = date(int(s_d[:3])+1911, int(s_d[3:5]), int(s_d[5:]))
                d2 = date(int(e_d[:3])+1911, int(e_d[3:5]), int(e_d[5:]))
                duration = (d2 - d1).days
            except: duration = 0
            
        return {'data': unit_data, 'start': start_date, 'end': end_date, 'duration': duration, 'filename': file_name}
    except Exception as e:
        st.error(f"❌ {file_name} 解析失敗: {e}")
        return None

# --- 其餘發信、下載、與主程式邏輯與你原本相同 ---
# (為了縮短回應長度，我在此省略重複的 send_email, get_mmdd 和繪圖部分)
# 請務必保留你原本代碼中 5. 主程式之後的所有 HTML/Excel 產生邏輯。
