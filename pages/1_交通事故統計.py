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
st.title("💥 交通事故統計 (v85 派出所專用精確版)")

# ==========================================
# 0. 設定區
# ==========================================
# 請確認您的 secrets.toml 中有設定 gcp_service_account
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

# 單位對照表 (Key: 檔案內的名稱, Value: 報表顯示的簡稱)
UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', 
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', 
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '交通中隊': '交通分隊',
    '科技執法': '科技執法', '交通組': '交通組'
}

# 報表顯示順序 (若檔案中沒有該單位，數值會補0)
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

def parse_police_station_csv_v85(file_obj):
    """
    針對 '派出所所轄案件統計表' 的 CSV 解析
    """
    counts = {} # 格式: {單位簡稱: [A1, A2, A3]}
    date_range_str = "0000~0000"
    
    try:
        # 1. 以文字模式讀取檔案 (避免 Pandas header 誤判)
        file_obj.seek(0)
        content_lines = []
        try:
            content_str = file_obj.read().decode('utf-8')
            content_lines = content_str.splitlines()
        except:
            file_obj.seek(0)
            content_str = file_obj.read().decode('big5', errors='ignore')
            content_lines = content_str.splitlines()

        # 2. 抓取統計日期 (通常在第 2 行)
        # 格式：統計日期：114/12/26 至 115/01/01
        for line in content_lines[:8]:
            m = re.search(r'統計日期[：:]\s*(\d+)/(\d+)/(\d+)\s*至\s*(\d+)/(\d+)/(\d+)', line)
            if m:
                s_m, s_d = m.group(2), m.group(3)
                e_m, e_d = m.group(5), m.group(6)
                date_range_str = f"{s_m}{s_d}~{e_m}{e_d}"
                break
        
        # 3. 尋找標題列 (包含 'A 1 類' 與 'A 2 類')
        header_row_idx = -1
        for i, line in enumerate(content_lines):
            if "A 1 類" in line and "A 2 類" in line:
                header_row_idx = i
                break
        
        if header_row_idx != -1:
            try:
                # 重新讀取 DataFrame (跳過直到標題列)
                # 注意：因為 CSV 前幾欄可能有空值，header=None 比較保險
                df = pd.read_csv(io.StringIO("\n".join(content_lines)), skiprows=header_row_idx, header=None)
                
                # 4. 鎖定欄位座標 (基於您的檔案 Snippet)
                # Col 0: 單位名稱 (如: 三和派出所)
                # Col 4: A1 件數 (Index 4)
                # Col 7: A2 件數 (Index 7)
                # Col 10: A3 件數 (Index 10)
                
                idx_unit = 0
                idx_a1 = 4
                idx_a2 = 7
                idx_a3 = 10
                
                # 從第 2 列開始讀數據 (Row 0=類別標題, Row 1=細項標題)
                for r in range(2, len(df)):
                    row = df.iloc[r]
                    if len(row) <= 10: continue # 確保欄位足夠
                    
                    unit_raw = str(row[idx_unit]).strip()
                    target_unit = None
                    
                    # 辨識單位
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
                print(f"DataFrame 解析錯誤: {e}")

    except Exception as e:
        print(f"檔案讀取錯誤: {e}")

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
            d, d_str = parse_police_station_csv_v85(f)
            parsed_data.append({"file": f, "data": d, "date": d_str, "name": f.name})
        
        # 2. 檔案分類 (依檔名特徵)
        # (2) -> 去年累計
        # (1) -> 本年累計
        # 無括號 -> 本期
        f_wk, f_yt, f_ly = None, None, None
        
        for item in parsed_data:
            nm = item['name']
            if "(2)" in nm: f_ly = item
            elif "(1)" in nm: f_yt = item
            else: f_wk = item
            
        # 防呆 fallback
        if not f_wk or not f_yt or not f_ly:
             st.warning("⚠️ 檔名無法識別，依序排列。請確認檔名包含 (1) 與 (2)。")
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
            # 取得各時期的 [A1, A2, A3]
            wk = d_wk.get(u, [0, 0, 0])
            yt = d_yt.get(u, [0, 0, 0])
            ly = d_ly.get(u, [0, 0, 0])
            
            # 比較: 本年總件數 - 去年總件數
            diff = sum(yt) - sum(ly)
            
            rows.append([
                u, 
                wk[0], wk[1], wk[2], 
                yt[0], yt[1], yt[2], 
                ly[0], ly[1], ly[2], 
                diff
            ])
            
        # 計算合計列 (Row Total)
        total_row = ["合計"]
        for i in range(1, 11): # 加總第 1~10 欄
            col_sum = sum(r[i] for r in rows)
            total_row.append(col_sum)
            
        all_rows = [total_row] + rows
        
        st.success(f"✅ 解析
