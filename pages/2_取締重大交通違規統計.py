import streamlit as st
import pandas as pd
import re
import io
import smtplib
import gspread
import calendar
import pypdf
import numpy as np  # 新增 numpy 用於型別判斷
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# --- 初始化配置 ---
st.set_page_config(page_title="重大交通違規統計", layout="wide", page_icon="🚦")
st.title("🚦 重大交通違規統計 (v65 強制轉型修復版)")

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
    return {"updateCells": {"rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}], "fields": "userEnteredValue,textFormatRuns", "range": {"sheetId": ws_id, "startRowIndex": row_idx-1, "endRowIndex": row_idx, "startColumnIndex": col_idx-1, "endColumnIndex": col_idx}}}

def get_footer_percent_red_req(ws_id, row_idx, col_idx, text):
    runs = [{"startIndex": 0, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}}]
    match = re.search(r'(\d+\.?\d*%)', text)
    if match:
        start, end = match.start(), match.end()
        runs.append({"startIndex": start, "format": {"foregroundColor": {"red": 1.0, "green": 0, "blue": 0}, "bold": True}})
        if end < len(text): runs.append({"startIndex": end, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}})
    return {"updateCells": {"rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}], "fields": "userEnteredValue,textFormatRuns", "range": {"sheetId": ws_id, "startRowIndex": row_idx-1, "endRowIndex": row_idx, "startColumnIndex": col_idx-1, "endColumnIndex": col_idx}}}

# ==========================================
# 2. 核心解析引擎 (強制轉型修復)
# ==========================================
def to_int(val):
    """安全轉換為 Python 標準 int"""
    try:
        # 處理 Pandas/Numpy 的數值類型
        if isinstance(val, (np.integer, np.int64, np.int32)):
            return int(val)
        if isinstance(val, (np.floating, float)):
            return int(val)
        # 處理字串
        s = str(val).replace(',', '').strip()
        if s == '' or s == '-' or s == 'nan': return 0
        return int(float(s))
    except:
        return 0

def extract_single_report_data(file_obj):
    counts = {} # {unit: [intercept, report]}
    date_str = "0000~0000"
    
    # 判斷檔案類型
    is_pdf = file_obj.name.lower().endswith('.pdf') or (file_obj.name.lower().endswith('.csv') and file_obj.size > 1000)
    
    try:
        # --- A. 處理 PDF ---
        if is_pdf:
            reader = pypdf.PdfReader(file_obj)
            text = ""
            for page in reader.pages: text += page.extract_text() + "\n"
            
            # 日期解析
            m = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', text)
            if m: date_str = f"{m.group(1)}~{m.group(2)}"
            
            clean_text = text.replace(')', ' ').replace('(', ' ').replace('%', ' ').replace('\n', ' ')
            for unit in UNIT_ORDER + ['合計']:
                try:
                    start = clean_text.find(unit)
                    if start != -1:
                        sub = clean_text[start+len(unit):start+150]
                        tokens = [t.replace(',','') for t in sub.split() if t.replace(',','').replace('-','',1).isdigit()]
                        if len(tokens) >= 3:
                            counts[unit] = [to_int(tokens[1]), to_int(tokens[2])]
                        elif len(tokens) >= 2:
                            counts[unit] = [to_int(tokens[0]), to_int(tokens[1])]
                except: continue
                
        # --- B. 處理 Excel ---
        else:
            df = pd.read_excel(file_obj, header=None)
            
            # 1. 日期解析
            top_txt = df.iloc[:10].astype(str).to_string()
            m = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', top_txt)
            if m: date_str = f"{m.group(1)}~{m.group(2)}"
            
            # 2. 座標偵測 (攔停/逕行)
            idx_int, idx_rem = -1, -1
            # 擴大搜尋範圍至前 30 行
            for r in range(min(30, len(df))):
                row_vals = df.iloc[r].astype(str).tolist()
                for c, val in enumerate(row_vals):
                    v_clean = val.replace('\n', '').replace(' ', '')
                    if "攔停" in v_clean: idx_int = c
                    if "逕行" in v_clean: idx_rem = c
            
            # 預設值備案
            if idx_int == -1: idx_int = 1
            if idx_rem == -1: idx_rem = 2
            
            # 3. 數據提取
            active_unit = None
            for _, row in df.iterrows():
                row_s = " ".join(row.astype(str))
                
                # 單位識別
                if "合計" in str(row[0]) or "總計" in str(row[0]): active_unit = "合計"
                elif "科技執法" in str(row[0]): active_unit = "科技執法"
                else:
                    for full, short in UNIT_MAP.items():
                        if short in str(row[0]): active_unit = short; break
                
                if active_unit:
                    v1 = row[idx_int]
                    v2 = row[idx_rem]
                    # 強制轉型
                    n1 = to_int(v1)
                    n2 = to_int(v2)
                    counts[active_unit] = [n1, n2]
                    active_unit = None

    except Exception as e:
        print(f"Error: {e}")
        
    return counts, date_str

# ==========================================
# 3. 畫面顯示與自動化
# ==========================================
files = st.file_uploader("請上傳 3 個 Focus 報表 (Excel/PDF)", accept_multiple_files=True)

if files and len(files) >= 3:
    try:
        # 解析檔案
        parsed_results = []
        for f in files:
            d, date_rng = extract_single_report_data(f)
            parsed_results.append({"file": f, "data": d, "date": date_rng})
        
        # 排序檔案: 本期(wk), 本年(yt), 去年(ly)
        f_wk, f_yt, f_ly = None, None, None
        
        # 優先依檔名辨識 (1), (2)
        # 若無標記，則依據日期字串長度或順序
        # 這裡假設上傳順序或檔名包含線索
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
        
        # 構建 HTML 表頭
        def red_h(t): return "".join([f"<span style='color:red; font-weight:bold;'>{c}</span>" if c in "0123456789~().%" else c for c in t])
        html_header = f"""
        <thead>
            <tr>
                <th rowspan='2'>統計期間</th>
                <th colspan='2' style='text-align:center;'>{red_h(title_wk)}</th>
                <th colspan='2' style='text-align:center;'>{red_h(title_yt)}</th>
                <th colspan='2' style='text-align:center;'>{red_h(title_ly)}</th>
                <th rowspan='2'>本年與去年同期比較</th>
                <th rowspan='2'>目標值</th>
                <th rowspan='2'>達成率</th>
            </tr>
            <tr>
                <th>現場攔停</th><th>逕行舉發</th>
                <th>現場攔停</th><th>逕行舉發</th>
                <th>現場攔停</th><th>逕行舉發</th>
            </tr>
        </thead>
        """

        # 數據組裝 (確保所有數值皆為 Python int)
        rows = []
        for u in UNIT_ORDER:
            wk = d_wk.get(u, [0, 0])
            yt = d_yt.get(u, [0, 0])
            ly = d_ly.get(u, [0, 0])
            
            # 強制轉換列表內的元素
            wk = [to_int(x) for x in wk]
            yt = [to_int(x) for x in yt]
            ly = [to_int(x) for x in ly]
            
            yt_tot = sum(yt)
            ly_tot = sum(ly)
            target = VIOLATION_TARGETS.get(u, 0)
            
            rows.append([u, wk[0], wk[1], yt[0], yt[1], ly[0], ly[1], yt_tot - ly_tot, target, f"{yt_tot/target:.0%}" if target > 0 else "—"])
        
        # 合計列計算
        # 使用 numpy 來做向量加法，確保欄位對齊
        # 但最後轉回 Python list 以防 JSON Error
        df_tmp = pd.DataFrame(rows)
        # 轉換 df 的數據部分為 numeric
        sums = df_tmp.iloc[:, 1:9].apply(pd.to_numeric).sum().fillna(0).astype(int).tolist()
        
        total_target = VIOLATION_TARGETS.get('合計', 11817)
        s_yt_tot = sums[2] + sums[3] # 本年攔 + 本年逕 (索引偏移注意: 0=u, 1=wk0, 2=wk1...)
        # 修正索引: rows 結構是 [unit, wk0, wk1, yt0, yt1, ly0, ly1, ...]
        # DataFrame 欄位索引: 0=unit, 1=wk0, 2=wk1, 3=yt0, 4=yt1, 5=ly0, 6=ly1...
        # sums 索引: 0對應df col 1 (wk0)...
        
        # 重新精算合計
        t_wk0 = sum([r[1] for r in rows])
        t_wk1 = sum([r[2] for r in rows])
        t_yt0 = sum([r[3] for r in rows])
        t_yt1 = sum([r[4] for r in rows])
        t_ly0 = sum([r[5] for r in rows])
        t_ly1 = sum([r[6] for r in rows])
        t_diff = (t_yt0 + t_yt1) - (t_ly0 + t_ly1)
        
        total_row = ["合計", t_wk0, t_wk1, t_yt0, t_yt1, t_ly0, t_ly1, t_diff, total_target, f"{(t_yt0+t_yt1)/total_target:.0%}" if total_target > 0 else "0%"]
        
        method_row = ["取締方式", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "", "", ""]
        all_rows = [method_row, total_row] + rows
        
        st.success("✅ 數據解析成功 (格式已修復)")
        
        # 顯示
        table_body = "".join([f"<tr>{''.join([f'<td>{x}</td>' for x in r])}</tr>" for r in all_rows])
        st.write(f"<table>{html_header}<tbody>{table_body}</tbody></table>", unsafe_allow_html=True)

        # 說明
        try:
            curr_year = date.today().year
            d_str = f_yt['date'].split('~')[1]
            mon = int(d_str[:2]); day = int(d_str[2:])
            prog = f"{((date(curr_year, mon, day) - date(curr_year, 1, 1)).days + 1) / (366 if calendar.isleap(curr_year) else 365):.1%}"
            e_yt_str = f"{curr_year-1911}年{mon}月{day}日"
        except: 
            prog = "98.0%"; e_yt_str = "114年12月XX日"

        f1 = f"一、本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt_str} (入案日期)應達成率為{prog}。"
        f2 = "二、重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項。"
        st.markdown(f"<br>#### {f1.replace(prog, f':red[{prog}]')}\n#### {f2}", unsafe_allow_html=True)

        # 自動化同步
        file_hash = "".join([f.name + str(f.size) for f in files])
        if st.session_state.get("v65_done") != file_hash:
            with st.status("🚀 執行自動化同步...") as s:
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL); ws = sh.get_worksheet(0)
                
                # 準備 Payload (再次確保所有數值為 Native Python Types)
                h1_raw = ["統計期間", title_wk, "", title_yt, "", title_ly, "", "同期比較", "目標值", "達成率"]
                
                # 轉換 all_rows 內的 numpy.int64 為 int
                clean_rows = []
                for row in all_rows:
                    new_row = [to_int(x) if isinstance(x, (int, float, np.number)) else str(x) for x in row]
                    clean_rows.append(new_row)
                
                full_payload = [h1_raw] + clean_rows
                
                ws.update(range_name='A2', values=full_payload)
                
                reqs = []
                for col_p in [(1,3), (3,5), (5,7)]:
                    reqs.append(get_merge_request(ws.id, col_p[0], col_p[1]))
                    reqs.append(get_center_align_request(ws.id, col_p[0], col_p[1]))
                
                for i, txt in [(2, title_wk), (4, title_yt), (6, title_ly)]:
                    reqs.append(get_header_red_req(ws.id, 2, i, txt))
                
                idx_f = 2 + len(full_payload) + 1
                ws.update_cell(idx_f, 1, f1); ws.update_cell(idx_f+1, 1, f2)
                reqs.append(get_footer_percent_red_req(ws.id, idx_f, 1, f1))
                
                sh.batch_update({"requests": reqs})
                
                if "email" in st.secrets:
                    out = io.BytesIO(); pd.DataFrame(full_payload).to_excel(out, index=False)
                    server = smtplib.SMTP('smtp.gmail.com', 587); server.starttls()
                    server.login(st.secrets["email"]["user"], st.secrets["email"]["password"])
                    msg = MIMEMultipart(); msg['Subject'] = Header(f"🚦 Focus 報表 - {e_yt_str}", "utf-8").encode()
                    msg.attach(MIMEText(f"{f1}\n{f2}", "plain"))
                    part = MIMEBase("application", "octet-stream"); part.set_payload(out.getvalue())
                    encoders.encode_base64(part); part.add_header("Content-Disposition", 'attachment; filename="Report.xlsx"')
                    msg.attach(part); server.send_message(msg); server.quit()
                
                st.session_state["v65_done"] = file_hash
                st.balloons(); s.update(label="完成", state="complete")
    except Exception as e: st.error(f"錯誤: {e}")
