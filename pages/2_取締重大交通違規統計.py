import streamlit as st
import pandas as pd
import re
import io
import smtplib
import gspread
import calendar
import pypdf
import numpy as np
import traceback
import csv
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# --- 初始化配置 ---
st.set_page_config(page_title="重大交通違規統計", layout="wide", page_icon="🚦")
st.title("🚦 重大交通違規統計 (v73 座標鎖定版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 
VIOLATION_TARGETS = {'合計': 11817, '科技執法': 0, '聖亭所': 1200, '龍潭所': 1500, '中興所': 1200, '石門所': 1000, '高平所': 800, '三和所': 500, '警備隊': 0, '交通分隊': 1000}

# 單位對照表
UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', 
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', 
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '交通中隊': '交通分隊',
    '科技執法': '科技執法', '交通組': '科技執法'
}

UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

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

def get_footer_percent_red_req(ws_id, row_idx, col_idx, text):
    runs = [{"startIndex": 0, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}}]
    text_str = str(text)
    match = re.search(r'(\d+\.?\d*%)', text_str)
    if match:
        start, end = match.start(), match.end()
        runs.append({"startIndex": start, "format": {"foregroundColor": {"red": 1.0, "green": 0, "blue": 0}, "bold": True}})
        if end < len(text_str): runs.append({"startIndex": end, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}})
    return {"updateCells": {"rows": [{"values": [{"userEnteredValue": {"stringValue": text_str}, "textFormatRuns": runs}]}], "fields": "userEnteredValue,textFormatRuns", "range": {"sheetId": ws_id, "startRowIndex": row_idx-1, "endRowIndex": row_idx, "startColumnIndex": col_idx-1, "endColumnIndex": col_idx}}}

# ==========================================
# 2. 核心解析引擎 (Focus CSV 專用)
# ==========================================
def clean_int(val):
    try:
        if pd.isna(val) or str(val).strip() in ['—', '', '-', 'nan', 'NaN']: return 0
        s = str(val).replace(',', '').strip()
        return int(float(s))
    except: return 0

def extract_focus_csv_data(file_obj):
    counts = {}
    date_str = "0000~0000"
    
    try:
        file_obj.seek(0)
        # 嘗試讀取 CSV
        try: df = pd.read_csv(file_obj, header=None, encoding='utf-8', on_bad_lines='skip')
        except: 
            file_obj.seek(0)
            df = pd.read_csv(file_obj, header=None, encoding='big5', on_bad_lines='skip')
        
        # 1. 抓取日期 (從 Row 4: "入案日期：1141223 至 1141229")
        # 掃描前 10 行尋找日期
        top_txt = df.iloc[:10].astype(str).to_string()
        m = re.search(r'入案日期：(\d+)\s*至\s*(\d+)', top_txt)
        if m: 
            s_d, e_d = m.group(1), m.group(2)
            date_str = f"{s_d[-4:]}~{e_d[-4:]}" # 取後4碼 MMDD

        # 2. 定位「合計」欄位
        # 根據分析，Row 5 (index 5) 包含 "合計"
        # 我們搜尋包含 "單位" 的那一行作為 Header Row
        header_idx = -1
        for idx, row in df.iterrows():
            if "單位" in str(row.values) and "合計" in str(row.values):
                header_idx = idx
                break
        
        if header_idx != -1:
            # 在 Header Row 尋找 "合計" 的索引
            row_vals = df.iloc[header_idx].tolist()
            # 找最後一個 "合計" (通常在最右邊)
            total_indices = [i for i, x in enumerate(row_vals) if "合計" in str(x)]
            if total_indices:
                col_total_int = total_indices[-1] # 合計欄位 (對應下一列的 "現場攔停")
                col_total_rem = col_total_int + 1 # 合計欄位 (對應下一列的 "逕行舉發")

                # 3. 抓取數據 (從 Header Row + 2 開始)
                for r_idx in range(header_idx + 2, len(df)):
                    row = df.iloc[r_idx]
                    raw_unit = str(row[2]).strip() # 單位通常在第 2 欄 (index 2), 視 CSV 而定
                    # 根據 Snippet: 單位在 col 2 (index 2)
                    
                    # 防呆: 如果 col 2 不是單位，嘗試 col 0, 1
                    target_unit = None
                    # 檢查 raw_unit 是否在 map 中
                    for full, short in UNIT_MAP.items():
                        if short in raw_unit: 
                            target_unit = short
                            break
                    if not target_unit:
                         # 嘗試 col 0, 1, 3
                         for c in [0, 1, 3]:
                             if c < len(row):
                                 val = str(row[c]).strip()
                                 for full, short in UNIT_MAP.items():
                                     if short in val: target_unit = short; break
                                 if target_unit: break
                    
                    if target_unit:
                        v_int = clean_int(row[col_total_int])
                        v_rem = clean_int(row[col_total_rem])
                        counts[target_unit] = [v_int, v_rem]

    except Exception as e:
        print(f"解析錯誤 {file_obj.name}: {e}")
        
    return counts, date_str

# ==========================================
# 3. 畫面顯示與自動化
# ==========================================
files = st.file_uploader("請上傳 3 個 Focus 報表 (focus114.csv)", accept_multiple_files=True)

if files and len(files) >= 3:
    try:
        # 1. 解析
        parsed_results = []
        for f in files:
            d, date_rng = extract_focus_csv_data(f)
            parsed_results.append({"file": f, "data": d, "date": date_rng})
        
        # 2. 排序 (依檔名)
        f_wk, f_yt, f_ly = None, None, None
        for item in parsed_results:
            nm = item['file'].name
            if "(2)" in nm: f_ly = item     # (2) 通常是去年
            elif "(1)" in nm: f_yt = item   # (1) 通常是本年
            else: f_wk = item               # 無括號是本期
        
        # 防呆排序
        if not f_yt: f_yt = parsed_results[1]
        if not f_ly: f_ly = parsed_results[2]
        if not f_wk: f_wk = parsed_results[0]

        d_wk, title_wk = f_wk['data'], f"本期({f_wk['date']})"
        d_yt, title_yt = f_yt['data'], f"本年累計({f_yt['date']})"
        d_ly, title_ly = f_ly['data'], f"去年累計({f_ly['date']})"

        # HTML Header
        def red_h(t): return "".join([f"<span style='color:red; font-weight:bold;'>{c}</span>" if c in "0123456789~().%" else c for c in t])
        html_header = f"""
        <thead>
            <tr>
                <th>統計期間</th>
                <th colspan='2' style='text-align:center;'>{red_h(title_wk)}</th>
                <th colspan='2' style='text-align:center;'>{red_h(title_yt)}</th>
                <th colspan='2' style='text-align:center;'>{red_h(title_ly)}</th>
                <th>同期比較</th>
                <th>目標值</th>
                <th>達成率</th>
            </tr>
        </thead>
        """

        # 3. 數據組裝
        rows = []
        for u in UNIT_ORDER:
            wk = d_wk.get(u, [0, 0]); yt = d_yt.get(u, [0, 0]); ly = d_ly.get(u, [0, 0])
            yt_tot = sum(yt); ly_tot = sum(ly); target = VIOLATION_TARGETS.get(u, 0)
            rows.append([u, wk[0], wk[1], yt[0], yt[1], ly[0], ly[1], yt_tot - ly_tot, target, f"{yt_tot/target:.0%}" if target > 0 else "—"])
        
        # 合計列
        sum_wk0 = sum(r[1] for r in rows); sum_wk1 = sum(r[2] for r in rows)
        sum_yt0 = sum(r[3] for r in rows); sum_yt1 = sum(r[4] for r in rows)
        sum_ly0 = sum(r[5] for r in rows); sum_ly1 = sum(r[6] for r in rows)
        sum_diff = (sum_yt0 + sum_yt1) - (sum_ly0 + sum_ly1)
        total_target = VIOLATION_TARGETS.get('合計', 11817)
        total_acc = f"{(sum_yt0+sum_yt1)/total_target:.0%}" if total_target > 0 else "0%"
        
        total_row = ["合計", sum_wk0, sum_wk1, sum_yt0, sum_yt1, sum_ly0, sum_ly1, sum_diff, total_target, total_acc]
        method_row = ["取締方式", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "", "", ""]
        all_rows = [method_row, total_row] + rows
        
        st.success("✅ Focus 報表解析完成！")
        
        # 顯示
        table_body = "".join([f"<tr>{''.join([f'<td>{x}</td>' for x in r])}</tr>" for r in all_rows])
        st.write(f"<table style='text-align:center; width:100%; border-collapse:collapse;' border='1'>{html_header}<tbody>{table_body}</tbody></table>", unsafe_allow_html=True)

        # 說明
        try:
            curr_year = date.today().year
            d_str = f_yt['date'].split('~')[1]
            mon = int(d_str[:2]); day = int(d_str[2:])
            prog = f"{((date(curr_year, mon, day) - date(curr_year, 1, 1)).days + 1) / (366 if calendar.isleap(curr_year) else 365):.1%}"
            e_yt_str = f"{curr_year-1911}年{mon}月{day}日"
        except: prog = "98.0%"; e_yt_str = "114年12月XX日"

        f1 = f"一、本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt_str} (入案日期)應達成率為{prog}。"
        f2 = "二、重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項。"
        st.markdown(f"<br>#### {f1.replace(prog, f':red[{prog}]')}\n#### {f2}", unsafe_allow_html=True)

        # 寫入 & 寄信
        file_hash = "".join([f.name + str(f.size) for f in files])
        if st.session_state.get("v73_done") != file_hash:
            with st.status("🚀 寫入 Google Sheets...") as s:
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL); ws = sh.get_worksheet(0)
                
                h1_raw = ["統計期間", title_wk, "", title_yt, "", title_ly, "", "同期比較", "目標值", "達成率"]
                # 強制清洗
                clean_payload = [h1_raw]
                for r in all_rows:
                    clean_row = []
                    for cell in r:
                        if isinstance(cell, (int, float, np.integer)): clean_row.append(int(cell))
                        else: clean_row.append(str(cell))
                    clean_payload.append(clean_row)
                
                ws.update(range_name='A2', values=clean_payload)
                
                reqs = []
                for col_p in [(1,3), (3,5), (5,7)]:
                    reqs.append(get_merge_request(ws.id, col_p[0], col_p[1]))
                    reqs.append(get_center_align_request(ws.id, col_p[0], col_p[1]))
                for i, txt in [(2, title_wk), (4, title_yt), (6, title_ly)]:
                    reqs.append(get_header_red_req(ws.id, 2, i, txt))
                idx_f = 2 + len(clean_payload) + 1
                ws.update_cell(idx_f, 1, f1); ws.update_cell(idx_f+1, 1, f2)
                reqs.append(get_footer_percent_red_req(ws.id, idx_f, 1, f1))
                sh.batch_update({"requests": reqs})
                
                if "email" in st.secrets:
                    sender = st.secrets["email"]["user"]
                    receiver = st.secrets.get("email", {}).get("to", sender)
                    out = io.BytesIO(); pd.DataFrame(clean_payload).to_excel(out, index=False)
                    server = smtplib.SMTP('smtp.gmail.com', 587); server.starttls()
                    server.login(sender, st.secrets["email"]["password"])
                    msg = MIMEMultipart(); msg['From'] = sender; msg['To'] = receiver
                    msg['Subject'] = Header(f"🚦 Focus 報表 - {e_yt_str}", "utf-8").encode()
                    msg.attach(MIMEText(f"{f1}\n{f2}", "plain"))
                    part = MIMEBase("application", "octet-stream"); part.set_payload(out.getvalue())
                    encoders.encode_base64(part); part.add_header("Content-Disposition", 'attachment; filename="Report.xlsx"')
                    msg.attach(part); server.send_message(msg); server.quit()
                
                st.session_state["v73_done"] = file_hash
                st.balloons(); s.update(label="完成", state="complete")

    except Exception as e:
        st.error(f"錯誤: {e}")
        st.code(traceback.format_exc())
