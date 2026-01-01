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
st.title("🚦 重大交通違規統計 (v70 精準對位版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

# 修改：科技執法 即 交通組
# 這裡的 Key 是我們報表上要顯示的名稱，Value 是 CSV 檔裡可能出現的名稱
UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', 
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', 
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '交通中隊': '交通分隊',
    '科技執法': '科技執法', '交通組': '科技執法' # 若 CSV 出現交通組，視為科技執法
}

# 報表顯示順序
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

VIOLATION_TARGETS = {'合計': 11817, '科技執法': 0, '聖亭所': 1200, '龍潭所': 1500, '中興所': 1200, '石門所': 1000, '高平所': 800, '三和所': 500, '警備隊': 0, '交通分隊': 1000}

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
# 2. 核心解析引擎 (針對 CSV/整合報表優化)
# ==========================================
def clean_int(val):
    try:
        if pd.isna(val) or str(val).strip() in ['—', '', '-', 'nan']: return 0
        s = str(val).replace(',', '').strip()
        return int(float(s))
    except: return 0

def parse_integrated_csv(file_obj):
    """解析整合型 CSV 報表"""
    counts = {} # {unit: [wk_int, wk_rem, yt_int, yt_rem, ly_int, ly_rem]}
    dates = {"wk": "0000~0000", "yt": "0000~0000", "ly": "0000~0000"}
    
    try:
        file_obj.seek(0)
        # 嘗試讀取 CSV，忽略錯誤行
        try:
            df = pd.read_csv(file_obj, header=None, on_bad_lines='skip', encoding='utf-8')
        except:
            file_obj.seek(0)
            df = pd.read_csv(file_obj, header=None, on_bad_lines='skip', encoding='big5')

        # 1. 抓取日期 (通常在第 2 列)
        header_text = df.iloc[:5].astype(str).to_string()
        m_wk = re.search(r'本期\((\d+~\d+)\)', header_text)
        m_yt = re.search(r'本年累計\((\d+~\d+)\)', header_text)
        m_ly = re.search(r'去年累計\((\d+~\d+)\)', header_text)
        if m_wk: dates["wk"] = m_wk.group(1)
        if m_yt: dates["yt"] = m_yt.group(1)
        if m_ly: dates["ly"] = m_ly.group(1)
        
        # 2. 定位「取締方式」那一列
        # 根據您的 CSV snippet，數據通常在「取締方式」列的下方
        start_row = -1
        for idx, row in df.iterrows():
            if "取締方式" in str(row.values) and "現場攔停" in str(row.values):
                start_row = idx
                break
        
        if start_row == -1: return {}, dates # 找不到定位點

        # 3. 抓取數據 (假設攔停與逕行交錯排列)
        # CSV 欄位順序預測: [單位, 本期攔, 本期逕, 本年攔, 本年逕, 去年攔, 去年逕...]
        # 根據您的 snippet: 
        # 合計, 18, 297, 2787, 15180, 2327, 15738...
        # 索引: 0, 1, 2, 3, 4, 5, 6
        
        for idx in range(start_row + 1, len(df)):
            row = df.iloc[idx].tolist()
            row_str = "".join([str(x) for x in row])
            
            # 辨識單位
            found_unit = None
            if "合計" in str(row[0]): found_unit = "合計"
            elif "科技執法" in str(row[0]) or "交通組" in str(row[0]): found_unit = "科技執法"
            else:
                for full, short in UNIT_MAP.items():
                    if short in str(row[0]): found_unit = short; break
            
            if found_unit:
                # 依序抓取 6 個數字
                # 注意：CSV 讀進來可能會有空欄位，需過濾
                # 這裡假設第 1 欄是單位名稱，接著就是數據
                nums = []
                for cell in row[1:]:
                    if pd.notna(cell) and str(cell).strip() != '':
                        nums.append(clean_int(cell))
                
                # 補齊至 6 個
                while len(nums) < 6: nums.append(0)
                
                counts[found_unit] = nums[:6] # 取前 6 個: wk_i, wk_r, yt_i, yt_r, ly_i, ly_r

    except Exception as e:
        print(f"CSV 解析失敗: {e}")

    return counts, dates

# ==========================================
# 3. 畫面顯示與自動化
# ==========================================
files = st.file_uploader("請上傳整合型報表 (CSV/Excel)", accept_multiple_files=True)

if files:
    try:
        # 只處理那個整合 CSV
        target_file = files[0] # 假設使用者只上傳該檔案
        
        data_map, date_map = parse_integrated_csv(target_file)
        
        if not data_map:
            st.error("❌ 無法抓取數據，請確認檔案格式是否為整合型 CSV。")
            st.stop()

        # 標題日期
        title_wk = f"本期({date_map['wk']})"
        title_yt = f"本年累計({date_map['yt']})"
        title_ly = f"去年累計({date_map['ly']})"
        
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

        rows = []
        for u in UNIT_ORDER:
            # vals: [wk_int, wk_rem, yt_int, yt_rem, ly_int, ly_rem]
            vals = data_map.get(u, [0, 0, 0, 0, 0, 0])
            yt_tot = vals[2] + vals[3]
            ly_tot = vals[4] + vals[5]
            target = VIOLATION_TARGETS.get(u, 0)
            rows.append([u, vals[0], vals[1], vals[2], vals[3], vals[4], vals[5], yt_tot - ly_tot, target, f"{yt_tot/target:.0%}" if target > 0 else "—"])
        
        # 合計列
        if '合計' in data_map:
            s = data_map['合計']
            s_yt = s[2] + s[3]; s_ly = s[4] + s[5]
            total_target = VIOLATION_TARGETS.get('合計', 11817)
            total_row = ["合計", s[0], s[1], s[2], s[3], s[4], s[5], s_yt - s_ly, total_target, f"{s_yt/total_target:.0%}" if total_target > 0 else "0%"]
        else:
            # 自動計算
            df_tmp = pd.DataFrame(rows)
            sums = df_tmp.iloc[:, 1:7].sum().tolist()
            s_yt = sums[2] + sums[3]; s_ly = sums[4] + sums[5]
            total_target = 11817
            total_row = ["合計", sums[0], sums[1], sums[2], sums[3], sums[4], sums[5], s_yt - s_ly, total_target, f"{s_yt/total_target:.0%}"]

        method_row = ["取締方式", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "", "", ""]
        all_rows = [method_row, total_row] + rows
        
        st.success("✅ 數據抓取成功！(v70 精準對位)")
        
        # 渲染
        table_body = "".join([f"<tr>{''.join([f'<td>{x}</td>' for x in r])}</tr>" for r in all_rows])
        st.write(f"<table style='text-align:center; width:100%; border-collapse:collapse;' border='1'>{html_header}<tbody>{table_body}</tbody></table>", unsafe_allow_html=True)

        # 說明
        try:
            curr_year = date.today().year
            d_str = date_map['yt'].split('~')[1]
            mon = int(d_str[:2]); day = int(d_str[2:])
            prog = f"{((date(curr_year, mon, day) - date(curr_year, 1, 1)).days + 1) / (366 if calendar.isleap(curr_year) else 365):.1%}"
            e_yt_str = f"{curr_year-1911}年{mon}月{day}日"
        except: prog = "98.0%"; e_yt_str = "114年12月XX日"

        f1 = f"一、本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt_str} (入案日期)應達成率為{prog}。"
        f2 = "二、重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項。"
        st.markdown(f"<br>#### {f1.replace(prog, f':red[{prog}]')}\n#### {f2}", unsafe_allow_html=True)

        # 寫入與寄信
        file_hash = target_file.name + str(target_file.size)
        if st.session_state.get("v70_done") != file_hash:
            with st.status("🚀 執行寫入與寄信...") as s:
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL); ws = sh.get_worksheet(0)
                
                h1_raw = ["統計期間", title_wk, "", title_yt, "", title_ly, "", "同期比較", "目標值", "達成率"]
                
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
                    
                    msg = MIMEMultipart()
                    msg['From'] = sender; msg['To'] = receiver
                    msg['Subject'] = Header(f"🚦 Focus 報表 - {e_yt_str}", "utf-8").encode()
                    msg.attach(MIMEText(f"{f1}\n{f2}", "plain"))
                    part = MIMEBase("application", "octet-stream"); part.set_payload(out.getvalue())
                    encoders.encode_base64(part); part.add_header("Content-Disposition", 'attachment; filename="Report.xlsx"')
                    msg.attach(part)
                    
                    server.send_message(msg); server.quit()
                
                st.session_state["v70_done"] = file_hash
                st.balloons(); s.update(label="完成", state="complete")

    except Exception as e:
        st.error(f"錯誤: {e}")
        st.code(traceback.format_exc())
