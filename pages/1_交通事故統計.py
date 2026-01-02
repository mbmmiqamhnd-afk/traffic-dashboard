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
st.title("💥 交通事故統計 (v80 事故專用版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 
# 交通事故目標值 (範例，可依需求修改)
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
# 2. 核心解析引擎 (交通事故專用)
# ==========================================
def clean_int(val):
    try:
        if pd.isna(val) or str(val).strip() in ['—', '', '-', 'nan', 'NaN']: return 0
        s = str(val).replace(',', '').strip()
        return int(float(s))
    except: return 0

def parse_accident_file(file_obj):
    counts = {}
    date_range_str = "0000~0000"
    
    try:
        file_obj.seek(0)
        try: df = pd.read_csv(file_obj, header=None, encoding='utf-8', on_bad_lines='skip', engine='python')
        except: 
            file_obj.seek(0)
            df = pd.read_csv(file_obj, header=None, encoding='big5', on_bad_lines='skip', engine='python')

        # 1. 抓取日期
        top_txt = df.iloc[:10].astype(str).to_string()
        m = re.search(r'日期[：:]\s*(\d+)\s*至\s*(\d+)', top_txt)
        if m:
            s_d, e_d = m.group(1), m.group(2)
            date_range_str = f"{s_d[-4:]}~{e_d[-4:]}"

        # 2. 定位標題列 (尋找 A1, A2, A3 或 死亡, 受傷)
        header_idx = -1
        col_unit = -1
        # 假設欄位結構: [單位, A1件數, A1死亡, A1受傷, A2件數, A2受傷, A3件數...]
        # 需根據實際報表微調
        
        for r in range(min(20, len(df))):
            row = df.iloc[r]
            row_str = " ".join(row.astype(str))
            if "單位" in row_str and ("A1" in row_str or "死亡" in row_str):
                header_idx = r
                # 找單位欄
                for c, v in enumerate(row):
                    if "單位" in str(v): col_unit = c
                break
        
        if header_idx != -1:
            # 這裡假設 A1, A2, A3 數據緊跟在單位後
            # 您需要根據實際報表告訴我第幾欄是 A1, A2, A3
            # 目前暫定:
            # col_unit + 1 = A1件數
            # col_unit + 2 = A1死亡
            # col_unit + 3 = A1受傷
            # col_unit + 4 = A2件數...
            
            base_col = col_unit + 1
            
            for r in range(header_idx + 1, len(df)):
                row = df.iloc[r]
                if len(row) <= base_col + 5: continue
                
                unit_raw = str(row[col_unit]).strip()
                target_unit = None
                
                if "總計" in unit_raw or "合計" in unit_raw: target_unit = "合計"
                else:
                    for full, short in UNIT_MAP.items():
                        if short in unit_raw: 
                            target_unit = short; break
                
                if target_unit:
                    # 抓取 A1, A2, A3 數據 (範例索引，需修正)
                    a1 = clean_int(row[base_col])     # A1件數
                    a2 = clean_int(row[base_col + 3]) # A2件數 (假設A1佔3欄)
                    a3 = clean_int(row[base_col + 5]) # A3件數 (假設A2佔2欄)
                    
                    counts[target_unit] = [a1, a2, a3]

    except Exception as e:
        print(f"Error: {e}")
        
    return counts, date_range_str

# ==========================================
# 3. 畫面顯示與自動化
# ==========================================
files = st.file_uploader("請上傳 3 個交通事故統計表 (CSV)", accept_multiple_files=True)

if files and len(files) >= 3:
    try:
        # 1. 解析
        parsed_data = []
        for f in files:
            d, d_str = parse_accident_file(f)
            parsed_data.append({"file": f, "data": d, "date": d_str, "name": f.name})
        
        # 2. 排序 (依檔名)
        f_wk, f_yt, f_ly = None, None, None
        for item in parsed_data:
            if "(2)" in item['name']: f_ly = item
            elif "(1)" in item['name']: f_yt = item
            else: f_wk = item
            
        if not f_wk or not f_yt or not f_ly:
             f_wk = parsed_data[0]; f_yt = parsed_data[1]; f_ly = parsed_data[2]

        d_wk, title_wk = f_wk['data'], f"本期({f_wk['date']})"
        d_yt, title_yt = f_yt['data'], f"本年累計({f_yt['date']})"
        d_ly, title_ly = f_ly['data'], f"去年累計({f_ly['date']})"

        # HTML Header (A1/A2/A3)
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

        # 3. 組裝資料
        rows = []
        for u in UNIT_ORDER:
            wk = d_wk.get(u, [0, 0, 0])
            yt = d_yt.get(u, [0, 0, 0])
            ly = d_ly.get(u, [0, 0, 0])
            diff = sum(yt) - sum(ly)
            rows.append([u, wk[0], wk[1], wk[2], yt[0], yt[1], yt[2], ly[0], ly[1], ly[2], diff])
            
        # 合計列
        total_row = ["合計"]
        for i in range(1, 11):
            total_row.append(sum(r[i] for r in rows))
            
        all_rows = [total_row] + rows
        
        st.success("✅ 交通事故報表解析完成！")
        
        # 渲染
        table_body = "".join([f"<tr>{''.join([f'<td>{x}</td>' for x in r])}</tr>" for r in all_rows])
        st.write(f"<table style='text-align:center; width:100%; border-collapse:collapse;' border='1'>{html_header}<tbody>{table_body}</tbody></table>", unsafe_allow_html=True)
        
        # (此處省略寫入 Google Sheets 代碼，若確認格式正確後再補上)

    except Exception as e:
        st.error(f"錯誤: {e}")
        st.code(traceback.format_exc())
