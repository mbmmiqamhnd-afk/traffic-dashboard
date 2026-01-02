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
st.title("💥 交通事故統計 (v83 完全對應版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 
ACCIDENT_TARGETS = {'A1': 0, 'A2': 0, 'A3': 0}

UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', 
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', 
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '交通中隊': '交通分隊',
    '科技執法': '科技執法', '交通組': '交通組'
}

UNIT_ORDER = ['交通組', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. Google 試算表格式指令
# ==========================================
def get_merge_request(ws_id, start_col, end_col):
    return {"mergeCells": {"range": {"sheetId": ws_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": start_col, "endColumnIndex": end_col}, "mergeType": "MERGE_ALL"}}

def get_center_align_request(ws_id, start_col, end_col):
    return {"repeatCell": {"range": {"sheetId": ws_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": start_col, "endColumnIndex": end_col}, "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}}, "fields": "userEnteredFormat.horizontalAlignment"}}

def get_header_red_req(ws_id, row_idx, col_idx, text):
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
    return {"updateCells": {"rows": [{"values": [{"userEnteredValue": {"stringValue": text_str}, "textFormatRuns": runs}]}], "fields": "userEnteredValue,textFormatRuns", "range": {"sheetId": ws_id, "startRowIndex": row_idx-1, "endRowIndex": row_idx, "startColumnIndex": col_idx-1, "endColumnIndex": col_idx}}}

# ==========================================
# 2. 核心解析引擎 (派出所統計表專用)
# ==========================================
def clean_int(val):
    try:
        if pd.isna(val) or str(val).strip() in ['—', '', '-', 'nan', 'NaN']: return 0
        s = str(val).replace(',', '').strip()
        return int(float(s))
    except: return 0

def parse_police_station_report(file_obj):
    counts = {} # Format: {Unit: [A1_cnt, A2_cnt, A3_cnt]}
    date_range_str = "0000~0000"
    
    try:
        # 1. 讀取 CSV (文字模式)
        file_obj.seek(0)
        content_lines = []
        try:
            content_str = file_obj.read().decode('utf-8')
            content_lines = content_str.splitlines()
        except:
            file_obj.seek(0)
            content_str = file_obj.read().decode('big5', errors='ignore')
            content_lines = content_str.splitlines()

        # 2. 抓取日期 (Row 1 附近)
        # 格式：統計日期：114/12/26 至 115/01/01
        for line in content_lines[:5]:
            m = re.search(r'統計日期[：:]\s*(\d+)/(\d+)/(\d+)\s*至\s*(\d+)/(\d+)/(\d+)', line)
            if m:
                # 提取月日: 114/12/26 -> 1226
                s_m, s_d = m.group(2), m.group(3)
                e_m, e_d = m.group(5), m.group(6)
                date_range_str = f"{s_m}{s_d}~{e_m}{e_d}"
                break
        
        # 3. 定位資料列
        # 標題通常是: ,總計,,,A 1 類,,,A 2 類,,,A 3 類
        # 下一行: ,件數,死亡,受傷,件數,死亡,受傷,件數,死亡,受傷,件數
        # 我們直接用 Pandas 讀取，並跳過前面幾行
        
        header_row_idx = -1
        for i, line in enumerate(content_lines):
            if "A 1 類" in line and "A 2 類" in line:
                header_row_idx = i
                break
        
        if header_row_idx != -1:
            try:
                # 從 Header Row 讀取
                # 注意：Header Row 下一行才是欄位名稱，數據在更下面
                # 我們讀取整個 DataFrame，Header 為 header_row_idx
                df = pd.read_csv(io.StringIO("\n".join(content_lines)), skiprows=header_row_idx, header=None)
                
                # 欄位索引 (0-based) 根據您的檔案結構：
                # Col 0: 單位名稱
                # Col 4: A1 件數
                # Col 7: A2 件數
                # Col 10: A3 件數
                
                idx_unit = 0
                idx_a1 = 4
                idx_a2 = 7
                idx_a3 = 10
                
                # 從第 2 列開始讀取數據 (跳過 'A1類' 和 '件數' 兩列標題)
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
                print(f"Pandas 讀取失敗: {e}")

    except Exception as e:
        print(f"File Error: {e}")

    return counts, date_range_str

# ==========================================
# 3. 畫面顯示與自動化
# ==========================================
files = st.file_uploader("請上傳 3 個派出所案件統計表 (CSV/Excel)", accept_multiple_files=True)

if files and len(files) >= 3:
    try:
        # 1. 解析所有檔案
        parsed_data = []
        for f in files:
            d, d_str = parse_police_station_report(f)
            parsed_data.append({"file": f, "data": d, "date": d_str, "name": f.name})
        
        # 2. 檔名鎖定排序
        f_wk, f_yt, f_ly = None, None, None
        
        for item in parsed_data:
            nm = item['name']
            if "(2)" in nm: f_ly = item     # 去年
            elif "(1)" in nm: f_yt = item   # 本年
            else: f_wk = item               # 本期
            
        if not f_wk or not f_yt or not f_ly:
             st.warning("⚠️ 檔名無法完全識別，請確認檔名包含 (1) 與 (2)。暫依上傳順序排列。")
             # Fallback
             used = []
             if f_wk: used.append(f_wk)
             if f_yt: used.append(f_yt)
             if f_ly: used.append(f_ly)
             rem = [x for x in parsed_data if x not in used]
             if not f_wk and rem: f_wk = rem.pop(0)
             if not f_yt and rem: f_yt = rem.pop(0)
             if not f_ly and rem: f_ly = rem.pop(0)

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
            # [A1, A2, A3]
            wk = d_wk.get(u, [0, 0, 0])
            yt = d_yt.get(u, [0, 0, 0])
            ly = d_ly.get(u, [0, 0, 0])
            
            # 比較: 本年總件數 - 去年總件數
            diff = sum(yt) - sum(ly)
            
            rows.append([u, 
                         wk[0], wk[1], wk[2], 
                         yt[0], yt[1], yt[2], 
                         ly[0], ly[1], ly[2], 
                         diff])
        
        # 合計列
        total_row = ["合計"]
        for i in range(1, 11):
            col_sum = sum(r[i] for r in rows)
            total_row.append(col_sum)
            
        all_rows = [total_row] + rows
        
        st.success(f"✅ 交通事故解析成功！(本期:{f_wk['name']} / 本年:{f_yt['name']} / 去年:{f_ly['name']})")
        
        # 渲染
        table_body = "".join([f"<tr>{''.join([f'<td>{x}</td>' for x in r])}</tr>" for r in all_rows])
        st.write(f"<table style='text-align:center; width:100%; border-collapse:collapse;' border='1'>{html_header}<tbody>{table_body}</tbody></table>", unsafe_allow_html=True)
        
        # 說明
        st.markdown(f"<br>#### 說明：本表統計各派出所轄區內 A1、A2、A3 類交通事故發生件數。", unsafe_allow_html=True)

        # 寫入功能
        file_hash = "".join([f.name + str(f.size) for f in files])
        if st.session_state.get("v83_done") != file_hash:
            if st.button("🚀 寫入 Google Sheets"):
                 try:
                     gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                     sh = gc.open_by_url(GOOGLE_SHEET_URL); ws = sh.get_worksheet(0)
                     
                     # 準備寫入資料
                     clean_payload = [["統計期間", title_wk, "", "", title_yt, "", "", title_ly, "", "", "比較"],
                                      ["單位", "A1", "A2", "A3", "A1", "A2", "A3", "A1", "A2", "A3", "增減"]]
                     
                     for r in all_rows:
                         clean_row = []
                         for cell in r:
                             if isinstance(cell, (int, float, np.integer)): clean_row.append(int(cell))
                             else: clean_row.append(str(cell))
                         clean_payload.append(clean_row)
                     
                     ws.update(range_name='A2', values=clean_payload)
                     
                     # 格式化
                     reqs = []
                     # 合併標題列
                     for s_col in [1, 4, 7]:
                         reqs.append(get_merge_request(ws.id, s_col, s_col+3))
                         reqs.append(get_center_align_request(ws.id, s_col, s_col+3))
                     
                     sh.batch_update({"requests": reqs})
                     st.session_state["v83_done"] = file_hash
                     st.success("✅ 寫入完成！")
                     st.balloons()
                 except Exception as e:
                     st.error(f"寫入失敗: {e}")

    except Exception as e:
        st.error(f"錯誤: {e}")
        st.code(traceback.format_exc())
