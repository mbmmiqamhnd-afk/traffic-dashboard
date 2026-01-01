import streamlit as st
import pandas as pd
import re
import io
import smtplib
import gspread
import calendar
import pypdf
import numpy as np
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# --- 初始化配置 ---
st.set_page_config(page_title="重大交通違規統計", layout="wide", page_icon="🚦")
st.title("🚦 重大交通違規統計 (v67 強制寫入版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 
VIOLATION_TARGETS = {'合計': 11817, '科技執法': 0, '聖亭所': 1200, '龍潭所': 1500, '中興所': 1200, '石門所': 1000, '高平所': 800, '三和所': 500, '警備隊': 0, '交通分隊': 1000}
UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '科技執法': '科技執法'}
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
    last_is_red = None
    for i, char in enumerate(text):
        is_red = char in red_chars
        if is_red != last_is_red:
            color = {"red": 1.0, "green": 0, "blue": 0} if is_red else {"red": 0, "green": 0, "blue": 0}
            runs.append({"startIndex": i, "format": {"foregroundColor": color, "bold": is_red}})
            last_is_red = is_red
    return {"updateCells": {"rows": [{"values": [{"userEnteredValue": {"stringValue": str(text)}, "textFormatRuns": runs}]}], "fields": "userEnteredValue,textFormatRuns", "range": {"sheetId": ws_id, "startRowIndex": row_idx-1, "endRowIndex": row_idx, "startColumnIndex": col_idx-1, "endColumnIndex": col_idx}}}

def get_footer_percent_red_req(ws_id, row_idx, col_idx, text):
    runs = [{"startIndex": 0, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}}]
    match = re.search(r'(\d+\.?\d*%)', str(text))
    if match:
        start, end = match.start(), match.end()
        runs.append({"startIndex": start, "format": {"foregroundColor": {"red": 1.0, "green": 0, "blue": 0}, "bold": True}})
        if end < len(text): runs.append({"startIndex": end, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}})
    return {"updateCells": {"rows": [{"values": [{"userEnteredValue": {"stringValue": str(text)}, "textFormatRuns": runs}]}], "fields": "userEnteredValue,textFormatRuns", "range": {"sheetId": ws_id, "startRowIndex": row_idx-1, "endRowIndex": row_idx, "startColumnIndex": col_idx-1, "endColumnIndex": col_idx}}}

# ==========================================
# 2. 核心解析引擎 (強化數值清洗)
# ==========================================
def clean_int(val):
    """強力清洗並轉為 int"""
    try:
        if pd.isna(val) or val == '—' or val == '': return 0
        s = str(val).replace(',', '').strip()
        return int(float(s))
    except: return 0

def extract_single_report_data(file_obj):
    counts = {}
    date_str = "0000~0000"
    is_pdf = file_obj.name.lower().endswith('.pdf')
    
    try:
        if is_pdf:
            reader = pypdf.PdfReader(file_obj)
            text = ""
            for page in reader.pages: text += page.extract_text() + "\n"
            m = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', text)
            if m: date_str = f"{m.group(1)}~{m.group(2)}"
            
            clean_text = text.replace(')', ' ').replace('(', ' ').replace('%', ' ').replace('\n', ' ')
            for unit in UNIT_ORDER + ['合計']:
                try:
                    start = clean_text.find(unit)
                    if start != -1:
                        sub = clean_text[start+len(unit):start+150]
                        tokens = [t.replace(',','') for t in sub.split() if t.replace(',','').replace('-','',1).isdigit()]
                        if len(tokens) >= 3: counts[unit] = [clean_int(tokens[1]), clean_int(tokens[2])]
                        elif len(tokens) >= 2: counts[unit] = [clean_int(tokens[0]), clean_int(tokens[1])]
                except: continue
        else:
            try: df = pd.read_excel(file_obj, header=None)
            except: 
                file_obj.seek(0); df = pd.read_csv(file_obj, header=None, encoding='utf-8', on_bad_lines='skip')
            
            top_txt = df.iloc[:10].astype(str).to_string()
            m = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', top_txt)
            if m: date_str = f"{m.group(1)}~{m.group(2)}"
            
            idx_int, idx_rem = 1, 2
            for r in range(min(30, len(df))):
                row_vals = df.iloc[r].astype(str).tolist()
                for c, val in enumerate(row_vals):
                    v = val.replace('\n', '').replace(' ', '')
                    if "攔停" in v: idx_int = c
                    if "逕行" in v: idx_rem = c
            
            active_unit = None
            for _, row in df.iterrows():
                row_s = " ".join(row.astype(str))
                if "合計" in str(row[0]): active_unit = "合計"
                elif "科技執法" in str(row[0]): active_unit = "科技執法"
                else:
                    for full, short in UNIT_MAP.items():
                        if short in str(row[0]): active_unit = short; break
                
                if active_unit:
                    counts[active_unit] = [clean_int(row[idx_int]), clean_int(row[idx_rem])]
                    active_unit = None

    except Exception as e: print(f"Error: {e}")
    return counts, date_str

# ==========================================
# 3. 畫面顯示與自動化
# ==========================================
files = st.file_uploader("請上傳 3 個 Focus 報表 (Excel/CSV/PDF)", accept_multiple_files=True)

if files and len(files) >= 3:
    try:
        parsed_results = []
        for f in files:
            d, date_rng = extract_single_report_data(f)
            parsed_results.append({"file": f, "data": d, "date": date_rng})
        
        f_wk, f_yt, f_ly = None, None, None
        for item in parsed_results:
            nm = item['file'].name
            if "(1)" in nm: f_yt = item
            elif "(2)" in nm: f_ly = item
            else: f_wk = item
            
        if not f_yt: f_yt = parsed_results[1]
        if not f_ly: f_ly = parsed_results[2]
        if not f_wk: f_wk = parsed_results[0]

        d_wk = f_wk['data']; title_wk = f"本期({f_wk['date']})"
        d_yt = f_yt['data']; title_yt = f"本年累計({f_yt['date']})"
        d_ly = f_ly['data']; title_ly = f"去年累計({f_ly['date']})"
        
        def red_h(t): return "".join([f"<span style='color:red; font-weight:bold;'>{c}</span>" if c in "0123456789~().%" else c for c in t])
        
        # HTML 表頭 (單層)
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
            wk = d_wk.get(u, [0, 0]); yt = d_yt.get(u, [0, 0]); ly = d_ly.get(u, [0, 0])
            yt_tot = sum(yt); ly_tot = sum(ly); target = VIOLATION_TARGETS.get(u, 0)
            rows.append([u, wk[0], wk[1], yt[0], yt[1], ly[0], ly[1], yt_tot - ly_tot, target, f"{yt_tot/target:.0%}" if target > 0 else "—"])
        
        # 合計列 (強制純 Python int 計算)
        # 用來寫入 Google Sheets 的數據必須是純 Python 類型
        sum_wk_int = sum([r[1] for r in rows])
        sum_wk_rem = sum([r[2] for r in rows])
        sum_yt_int = sum([r[3] for r in rows])
        sum_yt_rem = sum([r[4] for r in rows])
        sum_ly_int = sum([r[5] for r in rows])
        sum_ly_rem = sum([r[6] for r in rows])
        sum_diff   = sum([r[7] for r in rows])
        
        total_target = VIOLATION_TARGETS.get('合計', 11817)
        total_acc = f"{(sum_yt_int+sum_yt_rem)/total_target:.0%}" if total_target > 0 else "0%"
        
        total_row = ["合計", sum_wk_int, sum_wk_rem, sum_yt_int, sum_yt_rem, sum_ly_int, sum_ly_rem, sum_diff, total_target, total_acc]
        method_row = ["取締方式", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "", "", ""]
        
        all_rows = [method_row, total_row] + rows
        
        st.success("✅ 解析成功！正在強制寫入資料...")
        
        # 渲染網頁
        table_body = "".join([f"<tr>{''.join([f'<td>{x}</td>' for x in r])}</tr>" for r in all_rows])
        st.write(f"<table style='text-align:center; width:100%;'>{html_header}<tbody>{table_body}</tbody></table>", unsafe_allow_html=True)

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

        # 強制寫入流程
        file_hash = "".join([f.name + str(f.size) for f in files])
        if st.session_state.get("v67_done") != file_hash:
            with st.status("🚀 執行強制寫入...") as s:
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL); ws = sh.get_worksheet(0)
                
                # 1. 準備 Payload
                h1_raw = ["統計期間", title_wk, "", title_yt, "", title_ly, "", "同期比較", "目標值", "達成率"]
                
                # 2. 強制型別清洗 (確保每個格子都是 str 或 int)
                # Google Sheets API 很挑食，numpy.int64 會導致寫入失敗
                clean_payload = [h1_raw]
                for r in all_rows:
                    clean_r = []
                    for cell in r:
                        if isinstance(cell, (np.integer, np.int64)): clean_r.append(int(cell))
                        elif isinstance(cell, (np.floating, float)): clean_r.append(int(cell))
                        else: clean_r.append(str(cell))
                    clean_payload.append(clean_r)
                
                # 3. 執行寫入
                ws.update(range_name='A2', values=clean_payload)
                
                # 4. 格式化請求
                reqs = []
                for col_p in [(1,3), (3,5), (5,7)]: # 合併標題
                    reqs.append(get_merge_request(ws.id, col_p[0], col_p[1]))
                    reqs.append(get_center_align_request(ws.id, col_p[0], col_p[1]))
                
                for i, txt in [(2, title_wk), (4, title_yt), (6, title_ly)]: # 紅字日期
                    reqs.append(get_header_red_req(ws.id, 2, i, txt))
                
                idx_f = 2 + len(clean_payload) + 1
                ws.update_cell(idx_f, 1, f1); ws.update_cell(idx_f+1, 1, f2)
                reqs.append(get_footer_percent_red_req(ws.id, idx_f, 1, f1))
                
                sh.batch_update({"requests": reqs})
                
                # 5. 寄信
                if "email" in st.secrets:
                    out = io.BytesIO(); pd.DataFrame(clean_payload).to_excel(out, index=False)
                    server = smtplib.SMTP('smtp.gmail.com', 587); server.starttls()
                    server.login(st.secrets["email"]["user"], st.secrets["email"]["password"])
                    msg = MIMEMultipart(); msg['Subject'] = Header(f"🚦 Focus 報表 - {e_yt_str}", "utf-8").encode()
                    msg.attach(MIMEText(f"{f1}\n{f2}", "plain"))
                    part = MIMEBase("application", "octet-stream"); part.set_payload(out.getvalue())
                    encoders.encode_base64(part); part.add_header("Content-Disposition", 'attachment; filename="Report.xlsx"')
                    msg.attach(part); server.send_message(msg); server.quit()
                
                st.session_state["v67_done"] = file_hash
                st.balloons(); s.update(label="數據寫入完成", state="complete")
    except Exception as e: st.error(f"嚴重錯誤: {e}")
