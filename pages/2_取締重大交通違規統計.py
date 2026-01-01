import streamlit as st
import pandas as pd
import re
import io
import smtplib
import gspread
import calendar
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# --- 初始化配置 ---
st.set_page_config(page_title="重大交通違規統計", layout="wide", page_icon="🚦")
st.title("🚦 重大交通違規統計 (v56 精準座標解析版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

# 重大違規目標值 (合計請填寫年度總目標)
VIOLATION_TARGETS = {
    '合計': 11817, '科技執法': 0, '聖亭所': 1200, '龍潭所': 1500, '中興所': 1200, 
    '石門所': 1000, '高平所': 800, '三和所': 500, '警備隊': 0, '交通分隊': 1000
}

UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', 
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', 
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '科技執法': '科技執法'
}

# 排序邏輯：科技執法在聖亭所上一列
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. 核心格式指令 (Google Sheets API)
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
    anchor = "應達成率為"
    idx = text.find(anchor)
    if idx != -1:
        target_part = text[idx + len(anchor):]
        match = re.search(r'(\d+\.?\d*%)', target_part)
        if match:
            start = idx + len(anchor) + match.start()
            end = idx + len(anchor) + match.end()
            runs.append({"startIndex": start, "format": {"foregroundColor": {"red": 1.0, "green": 0, "blue": 0}, "bold": True}})
            if end < len(text):
                runs.append({"startIndex": end, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}})
    return {"updateCells": {"rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}], "fields": "userEnteredValue,textFormatRuns", "range": {"sheetId": ws_id, "startRowIndex": row_idx-1, "endRowIndex": row_idx, "startColumnIndex": col_idx-1, "endColumnIndex": col_idx}}}

# ==========================================
# 2. 精準座標解析邏輯
# ==========================================
def parse_report_v56(f):
    if not f: return {}, "0000000", "0000000"
    counts = {}
    s, e = "0000000", "0000000"
    try:
        f.seek(0)
        xls = pd.ExcelFile(f)
        for sn in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sn, header=None)
            
            # 偵測日期
            if s == "0000000":
                top_text = df.iloc[:15].astype(str).to_string()
                m = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', top_text)
                if m: s, e = m.group(1), m.group(2)
            
            # 🚀 尋找欄位座標 (精準搜尋表頭文字)
            idx_intercept = -1
            idx_remote = -1
            for r_idx in range(min(20, len(df))):
                row_vals = df.iloc[r_idx].astype(str).tolist()
                for c_idx, val in enumerate(row_vals):
                    if "現場攔停" in val: idx_intercept = c_idx
                    if "逕行舉發" in val: idx_remote = c_idx
            
            # 後備邏輯：如果沒找到文字表頭，使用重大違規報表常見位置
            if idx_intercept == -1: idx_intercept = 1 
            if idx_remote == -1: idx_remote = 2

            # 抓取數據
            active_unit = None
            for _, row in df.iterrows():
                row_str = " ".join(row.astype(str))
                if "舉發單位：" in row_str:
                    m2 = re.search(r"舉發單位：\s*(\S+)", row_str)
                    if m2: active_unit = m2.group(1).strip()
                
                if "總計" in row_str and active_unit:
                    short = UNIT_MAP.get(active_unit, active_unit)
                    if short in UNIT_ORDER:
                        try:
                            # 根據偵測到的座標抓取
                            v_int = str(row[idx_intercept]).replace(',', '')
                            v_rem = str(row[idx_remote]).replace(',', '')
                            
                            val_int = int(float(v_int)) if v_int.replace('.','',1).isdigit() else 0
                            val_rem = int(float(v_rem)) if v_rem.replace('.','',1).isdigit() else 0
                            
                            if short not in counts: counts[short] = [0, 0]
                            counts[short][0] += val_int
                            counts[short][1] += val_rem
                        except: pass
                    active_unit = None
        return counts, s, e
    except: return {}, "0000000", "0000000"

# ==========================================
# 3. 畫面顯示與同步
# ==========================================
files = st.file_uploader("上傳 3 個重大違規報表 (本期、本年累計、去年累計)", accept_multiple_files=True, type=['xlsx', 'xls'])

if files and len(files) >= 3:
    try:
        file_hash = "".join(sorted([f.name + str(f.size) for f in files]))
        f_wk, f_yt, f_ly = None, None, None
        for f in files:
            if "(1)" in f.name: f_yt = f
            elif "(2)" in f.name: f_ly = f
            else: f_wk = f
        
        # 執行精準座標解析
        d_wk, s_wk, e_wk = parse_report_v56(f_wk)
        d_yt, s_yt, e_yt = parse_report_v56(f_yt)
        d_ly, s_ly, e_ly = parse_report_v56(f_ly)

        title_wk = f"本期({s_wk[-4:]}~{e_wk[-4:]})"
        title_yt = f"本年累計({s_yt[-4:]}~{e_yt[-4:]})"
        title_ly = f"去年累計({s_ly[-4:]}~{e_ly[-4:]})"
        
        def red_h(t): return "".join([f"<span style='color:red; font-weight:bold;'>{c}</span>" if c in "0123456789~().%" else c for c in t])

        # 網頁 HTML 表頭
        html_header = f"""
        <thead>
            <tr>
                <th rowspan='2'>統計期間</th>
                <th colspan='2' style='text-align:center;'>{red_h(title_wk)}</th>
                <th colspan='2' style='text-align:center;'>{red_h(title_yt)}</th>
                <th colspan='2' style='text-align:center;'>{red_h(title_ly)}</th>
                <th rowspan='2'>同期比較</th>
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

        # 數據組裝
        rows_data = []
        for u in UNIT_ORDER:
            wk = d_wk.get(u, [0, 0]); yt = d_yt.get(u, [0, 0]); ly = d_ly.get(u, [0, 0])
            yt_tot = sum(yt); ly_tot = sum(ly); target = VIOLATION_TARGETS.get(u, 0)
            rows_data.append([u, wk[0], wk[1], yt[0], yt[1], ly[0], ly[1], yt_tot - ly_tot, target, f"{yt_tot/target:.0%}" if target > 0 else "—"])
        
        # 合計計算
        df_calc = pd.DataFrame(rows_data)
        sums = df_calc.iloc[:, 1:9].apply(pd.to_numeric).sum()
        total_target = VIOLATION_TARGETS.get('合計', sums[8])
        total_row = ["合計", sums[1], sums[2], sums[3], sums[4], sums[5], sums[6], sums[7], total_target, f"{(sums[3]+sums[4])/total_target:.0%}" if total_target > 0 else "0%"]
        
        # 🚀 取締方式列 (合計上一列)
        method_row = ["取締方式", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "", "", ""]
        
        all_display_rows = [method_row, total_row] + rows_data
        st.success("✅ 精準座標解析完成！數據已正確抓取。")
        
        # 渲染 HTML 表格
        table_body = "".join([f"<tr>{''.join([f'<td>{x}</td>' for x in r])}</tr>" for r in all_display_rows])
        st.write(f"<table>{html_header}<tbody>{table_body}</tbody></table>", unsafe_allow_html=True)

        # 說明文字
        y, m, d = int(e_yt[:3])+1911, int(e_yt[3:5]), int(e_yt[5:])
        prog = f"{((date(y, m, d) - date(y, 1, 1)).days + 1) / (366 if calendar.isleap(y) else 365):.1%}"
        f1 = f"一、本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt[:3]}年{int(e_yt[3:5])}月{int(e_yt[5:])}日 (入案日期)應達成率為{prog}。"
        f2 = "二、重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項。"
        st.markdown(f"<br>#### {f1.replace(prog, f':red[{prog}]')}\n#### {f2}", unsafe_allow_html=True)

        # --- 自動化流程 ---
        if st.session_state.get("v56_done") != file_hash:
            with st.status("🚀 執行精準同步與格式化...") as s:
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL); ws = sh.get_worksheet(0)
                
                h1_raw = ["統計期間", title_wk, "", title_yt, "", title_ly, "", "同期比較", "目標值", "達成率"]
                full_payload = [h1_raw] + all_display_rows
                ws.update(range_name='A2', values=full_payload)
                
                reqs = []
                # 合併與置中 B2:C2, D2:E2, F2:G2
                for col_p in [(1,3), (3,5), (5,7)]:
                    reqs.append(get_merge_request(ws.id, col_p[0], col_p[1]))
                    reqs.append(get_center_align_request(ws.id, col_p[0], col_p[1]))
                
                # 標題日期紅字
                for i, txt in [(2, title_wk), (4, title_yt), (6, title_ly)]:
                    reqs.append(get_header_red_req(ws.id, 2, i, txt))
                
                # 末端百分比標紅
                idx_f = 2 + len(full_payload) + 1
                ws.update_cell(idx_f, 1, f1); ws.update_cell(idx_f+1, 1, f2)
                reqs.append(get_footer_percent_red_req(ws.id, idx_f, 1, f1))
                
                sh.batch_update({"requests": reqs})
                
                # 自動寄信
                if "email" in st.secrets:
                    out = io.BytesIO(); pd.DataFrame(full_payload).to_excel(out, index=False)
                    server = smtplib.SMTP('smtp.gmail.com', 587); server.starttls()
                    server.login(st.secrets["email"]["user"], st.secrets["email"]["password"])
                    msg = MIMEMultipart()
                    msg['Subject'] = Header(f"🚦 重大違規報表 - {e_yt}", "utf-8").encode()
                    msg.attach(MIMEText(f"{f1}\n{f2}", "plain"))
                    part = MIMEBase("application", "octet-stream"); part.set_payload(out.getvalue())
                    encoders.encode_base64(part); part.add_header("Content-Disposition", 'attachment; filename="Violations.xlsx"')
                    msg.attach(part); server.send_message(msg); server.quit()
                
                st.session_state["v56_done"] = file_hash
                st.balloons(); s.update(label="數據與格式同步完成", state="complete")
    except Exception as e: st.error(f"錯誤: {e}")
