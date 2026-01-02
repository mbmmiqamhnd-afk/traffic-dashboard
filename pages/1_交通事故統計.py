import streamlit as st
import pandas as pd
import re
import io
import smtplib
import gspread
import calendar
import numpy as np
import traceback
import csv
from datetime import date, datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# --- 初始化配置 ---
st.set_page_config(page_title="交通事故統計", layout="wide", page_icon="💥")
st.title("💥 交通事故統計 (v86 縮排修正版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "[https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit](https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit)" 

# 單位對照表
UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', 
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', 
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '交通中隊': '交通分隊',
    '科技執法': '科技執法', '交通組': '交通組'
}

# 報表顯示順序
UNIT_ORDER = ['交通組', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. Google 試算表 API 輔助函式
# ==========================================
def get_merge_request(ws_id, start_col, end_col):
    """產生合併儲存格的 API 請求"""
    return {
        "mergeCells": {
            "range": {
                "sheetId": ws_id, 
                "startRowIndex": 1, "endRowIndex": 2, 
                "startColumnIndex": start_col, "endColumnIndex": end_col
            }, 
            "mergeType": "MERGE_ALL"
        }
    }

def get_center_align_request(ws_id, start_col, end_col):
    """產生置中對齊的 API 請求"""
    return {
        "repeatCell": {
            "range": {
                "sheetId": ws_id, 
                "startRowIndex": 1, "endRowIndex": 2, 
                "startColumnIndex": start_col, "endColumnIndex": end_col
            }, 
            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}}, 
            "fields": "userEnteredFormat.horizontalAlignment"
        }
    }

def get_header_red_req(ws_id, row_idx, col_idx, text):
    """產生紅字標題的 API 請求"""
    red_chars = set("0123456789~().%")
    runs = []
    text_str = str(text)
    last_is_red = None
    for i, char in enumerate(text_str):
        is_red = char in red_chars
        if is_red != last_is_red:
            color = {"red": 1.0, "green": 0, "blue": 0} if is_red else {"red": 0, "green": 0, "blue": 0}
            runs.append({"startIndex": i, "format": {"foregroundColor": color, "bold": is_red}})
            last_is_red = is_red
    return {
        "updateCells": {
            "rows": [{"values": [{"userEnteredValue": {"stringValue": text_str}, "textFormatRuns": runs}]}], 
            "fields": "userEnteredValue,textFormatRuns", 
            "range": {
                "sheetId": ws_id, 
                "startRowIndex": row_idx-1, "endRowIndex": row_idx, 
                "startColumnIndex": col_idx-1, "endColumnIndex": col_idx
            }
        }
    }

# ==========================================
# 2. 核心解析引擎 (針對派出所 CSV 結構)
# ==========================================
def clean_int(val):
    """將 CSV 中的空值、逗號轉為整數"""
    try:
        if pd.isna(val) or str(val).strip() in ['—', '', '-', 'nan', 'NaN']: return 0
        s = str(val).replace(',', '').strip()
        return int(float(s))
    except: return 0

def parse_police_station_csv_v86(file_obj):
    counts = {} 
    date_range_str = "0000~0000"
    
    try:
        file_obj.seek(0)
        content_lines = []
        try:
            content_str = file_obj.read().decode('utf-8')
            content_lines = content_str.splitlines()
        except:
            file_obj.seek(0)
            content_str = file_obj.read().decode('big5', errors='ignore')
            content_lines = content_str.splitlines()

        # 抓取統計日期
        for line in content_lines[:8]:
            m = re.search(r'統計日期[：:]\s*(\d+)/(\d+)/(\d+)\s*至\s*(\d+)/(\d+)/(\d+)', line)
            if m:
                s_m, s_d = m.group(2), m.group(3)
                e_m, e_d = m.group(5), m.group(6)
                date_range_str = f"{s_m}{s_d}~{e_m}{e_d}"
                break
        
        # 尋找標題列
        header_row_idx = -1
        for i, line in enumerate(content_lines):
            if "A 1 類" in line and "A 2 類" in line:
                header_row_idx = i
                break
        
        if header_row_idx != -1:
            try:
                df = pd.read_csv(io.StringIO("\n".join(content_lines)), skiprows=header_row_idx, header=None)
                
                # 鎖定欄位座標
                idx_unit = 0
                idx_a1 = 4
                idx_a2 = 7
                idx_a3 = 10
                
                for r in range(2, len(df)):
                    row = df.iloc[r]
                    if len(row) <= 10: continue
                    
                    unit_raw = str(row[idx_unit]).strip()
                    target_unit = None
                    
                    if "總計" in unit_raw or "合計" in unit_raw: 
                        target_unit = "合計"
                    else:
                        for full, short in UNIT_MAP.items():
                            if full in unit_raw or short in unit_raw:
                                target_unit = short; break
                    
                    if target_unit:
                        if target_unit in counts: continue
                        v_a1 = clean_int(row[idx_a1])
                        v_a2 = clean_int(row[idx_a2])
                        v_a3 = clean_int(row[idx_a3])
                        counts[target_unit] = [v_a1, v_a2, v_a3]
            except Exception as e:
                print(f"DataFrame error: {e}")

    except Exception as e:
        print(f"File error: {e}")

    return counts, date_range_str

# ==========================================
# 3. 畫面顯示與自動化邏輯
# ==========================================
files = st.file_uploader("請上傳 3 個派出所所轄案件統計表 (CSV)", accept_multiple_files=True)

if files and len(files) >= 3:
    try:
        # 1. 解析檔案
        parsed_data = []
        for f in files:
            d, d_str = parse_police_station_csv_v86(f)
            parsed_data.append({"file": f, "data": d, "date": d_str, "name": f.name})
        
        # 2. 檔案分類
        f_wk, f_yt, f_ly = None, None, None
        
        for item in parsed_data:
            nm = item['name']
            if "(2)" in nm: f_ly = item
            elif "(1)" in nm: f_yt = item
            else: f_wk = item
            
        if not f_wk or not f_yt or not f_ly:
            st.warning("⚠️ 檔名無法識別，依序排列。")
            f_wk = parsed_data[0]; f_yt = parsed_data[1]; f_ly = parsed_data[2]

        d_wk, title_wk = f_wk['data'], f"本期({f_wk['date']})"
        d_yt, title_yt = f_yt['data'], f"本年累計({f_yt['date']})"
        d_ly, title_ly = f_ly['data'], f"去年累計({f_ly['date']})"

        # HTML Header
        def red_h(t): return "".join([f"<span style='color:red; font-weight:bold;'>{c}</span>" if c in "0123456789~().%" else c for c in t])
        
        html_header = f"""
        <thead>
            <tr>
                <th>統計期間</th>
                <th colspan='3' style='text-align:center;'>{red_h(title_wk)}</th>
                <th colspan='3' style='text-align:center;'>{red_h(title_yt)}</th>
                <th colspan='3' style='text-align:center;'>{red_h(title_ly)}</th>
                <th>比較</th>
            </tr>
            <tr>
                <th>單位</th>
                <th>A1</th><th>A2</th><th>A3</th>
                <th>A1</th><th>A2</th><th>A3</th>
                <th>A1</th><th>A2</th><th>A3</th>
                <th>增減</th>
            </tr>
        </thead>
        """

        # 3. 數據組裝
        rows = []
        for u in UNIT_ORDER:
            wk = d_wk.get(u, [0, 0, 0])
            yt = d_yt.get(u, [0, 0, 0])
            ly = d_ly.get(u, [0, 0, 0])
            diff = sum(yt) - sum(ly)
            
            rows.append([
                u, 
                wk[0], wk[1], wk[2], 
                yt[0], yt[1], yt[2], 
                ly[0], ly[1], ly[2], 
                diff
            ])
            
        # 合計列
        total_row = ["合計"]
        for i in range(1, 11):
            col_sum = sum(r[i] for r in rows)
            total_row.append(col_sum)
            
        all_rows = [total_row] + rows
        
        st.success(f"✅ 解析成功！本期檔名: {f_wk['name']}")
        
        table_body = "".join([f"<tr>{''.join([f'<td>{x}</td>' for x in r])}</tr>" for r in all_rows])
        st.write(f"<table style='text-align:center; width:100%; border-collapse:collapse;' border='1'>{html_header}<tbody>{table_body}</tbody></table>", unsafe_allow_html=True)
        
        # 4. 自動寫入
        file_hash = "".join([f.name + str(f.size) for f in files])
        
        if st.session_state.get("v86_done") != file_hash:
            with st.status("🚀 正在自動寫入雲端...", expanded=True) as s:
                try:
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL); ws = sh.get_worksheet(0)
                    
                    clean_payload = [
                        ["統計期間", title_wk, "", "", title_yt, "", "", title_ly, "", "", "比較"],
                        ["單位", "A1", "A2", "A3", "A1", "A2", "A3", "A1", "A2", "A3", "增減"]
                    ]
                    
                    for r in all_rows:
                        clean_row = []
                        for cell in r:
                            if isinstance(cell, (int, float, np.integer)): clean_row.append(int(cell))
                            else: clean_row.append(str(cell))
                        clean_payload.append(clean_row)
                    
                    ws.update(range_name='A2', values=clean_payload)
                    
                    reqs = []
                    for s_col in [1, 4, 7]:
                         reqs.append(get_merge_request(ws.id, s_col, s_col+3))
                         reqs.append(get_center_align_request(ws.id, s_col, s_col+3))
                    
                    reqs.append(get_header_red_req(ws.id, 2, 2, title_wk))
                    reqs.append(get_header_red_req(ws.id, 2, 5, title_yt))
                    reqs.append(get_header_red_req(ws.id, 2, 8, title_ly))

                    sh.batch_update({"requests": reqs})
                    
                    st.session_state["v86_done"] = file_hash
                    s.update(label="✅ 數據已自動寫入 Google Sheets！", state="complete")
                    st.balloons()

                except Exception as e:
                    s.update(label="❌ 寫入失敗", state="error")
                    st.error(f"寫入錯誤詳情: {e}")
                    st.code(traceback.format_exc())

    except Exception as e:
        st.error(f"全域錯誤: {e}")
        st.code(traceback.format_exc())
