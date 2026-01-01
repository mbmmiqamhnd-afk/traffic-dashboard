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
st.title("🚦 重大交通違規統計 (v53 表頭合併版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 
VIOLATION_TARGETS = {'合計': 11817, '科技執法': 0, '聖亭所': 1200, '龍潭所': 1500, '中興所': 1200, '石門所': 1000, '高平所': 800, '三和所': 500, '警備隊': 0, '交通分隊': 1000}
UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '科技執法': '科技執法'}
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. 核心格式指令 (含合併儲存格)
# ==========================================
def get_merge_request(ws_id, start_col, end_col):
    """產生合併儲存格請求 (針對第 2 列)"""
    return {
        "mergeCells": {
            "range": {
                "sheetId": ws_id,
                "startRowIndex": 1, "endRowIndex": 2, # 指 A2 這一列
                "startColumnIndex": start_col, "endColumnIndex": end_col
            },
            "mergeType": "MERGE_ALL"
        }
    }

def get_center_align_request(ws_id, start_col, end_col):
    """標題文字置中"""
    return {
        "repeatCell": {
            "range": {"sheetId": ws_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": start_col, "endColumnIndex": end_col},
            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat.horizontalAlignment"
        }
    }

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
# 2. 解析邏輯
# ==========================================
def parse_report(f):
    if not f: return {}, "0000000", "0000000"
    counts = {}
    s, e = "0000000", "0000000"
    try:
        f.seek(0)
        df_top = pd.read_excel(f, header=None, nrows=10).to_string()
        m = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', df_top)
        if m: s, e = m.group(1), m.group(2)
        f.seek(0)
        xls = pd.ExcelFile(f)
        for sn in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sn, header=None)
            u = None
            for _, r in df.iterrows():
                rs = " ".join(r.astype(str))
                if "舉發單位：" in rs:
                    m2 = re.search(r"舉發單位：(\S+)", rs)
                    if m2: u = m2.group(1).strip()
                if "總計" in rs and u:
                    nums = [int(str(x).replace(',','')) for x in r if str(x).replace('.','',1).isdigit()]
                    if len(nums) >= 2:
                        short = UNIT_MAP.get(u, u)
                        if short in UNIT_ORDER:
                            if short not in counts: counts[short] = [0, 0]
                            counts[short][0] += nums[-2] # 攔停
                            counts[short][1] += nums[-1] # 逕行
                        u = None
        return counts, s, e
    except: return {}, "0000000", "0000000"

# ==========================================
# 3. 畫面顯示與自動化
# ==========================================
files = st.file_uploader("上傳 3 個重大違規報表檔案", accept_multiple_files=True, type=['xlsx', 'xls'])

if files and len(files) >= 3:
    try:
        file_hash = "".join(sorted([f.name + str(f.size) for f in files]))
        f_wk, f_yt, f_ly = None, None, None
        for f in files:
            if "(1)" in f.name: f_yt = f
            elif "(2)" in f.name: f_ly = f
            else: f_wk = f
        
        d_wk, s_wk, e_wk = parse_report(f_wk); d_yt, s_yt, e_yt = parse_report(f_yt); d_ly, s_ly, e_ly = parse_report(f_ly)

        # 🚀 準備表頭與合併資訊
        def red_h(t): return "".join([f"<span style='color:red; font-weight:bold;'>{c}</span>" if c in "0123456789~().%" else c for c in t])
        
        title_wk = f"本期({s_wk[-4:]}~{e_wk[-4:]})"
        title_yt = f"本年累計({s_yt[-4:]}~{e_yt[-4:]})"
        title_ly = f"去年累計({s_ly[-4:]}~{e_ly[-4:]})"
        
        # 網頁端 HTML 渲染 (使用 colspan 合併)
        html_header = f"""
        <thead>
            <tr>
                <th rowspan='2'>統計期間</th>
                <th colspan='2'>{red_h(title_wk)}</th>
                <th colspan='2'>{red_h(title_yt)}</th>
                <th colspan='2'>{red_h(title_ly)}</th>
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

        # 組裝單位數據
        rows = []
        for u in UNIT_ORDER:
            wk = d_wk.get(u, [0, 0]); yt = d_yt.get(u, [0, 0]); ly = d_ly.get(u, [0, 0])
            yt_total = sum(yt); ly_total = sum(ly); target = VIOLATION_TARGETS.get(u, 0)
            rows.append([u, wk[0], wk[1], yt[0], yt[1], ly[0], ly[1], yt_total - ly_total, target, f"{yt_total/target:.0%}" if target > 0 else "—"])
        
        df_temp = pd.DataFrame(rows)
        sums = df_temp.iloc[:, 1:9].apply(pd.to_numeric).sum()
        total_row = ["合計", sums[1], sums[2], sums[3], sums[4], sums[5], sums[6], sums[7], sums[8], f"{(sums[3]+sums[4])/sums[8]:.0%}" if sums[8]>0 else "0%"]
        
        # 最終顯示
        all_rows = [total_row] + rows
        st.success("✅ 解析成功！表頭已依要求合併。")
        
        # 渲染 HTML 表格
        table_body = "".join([f"<tr>{''.join([f'<td>{x}</td>' for x in r])}</tr>" for r in all_rows])
        st.write(f"<table>{html_header}<tbody>{table_body}</tbody></table>", unsafe_allow_html=True)

        # 說明文字
        y, m, d = int(e_yt[:3])+1911, int(e_yt[3:5]), int(e_yt[5:])
        prog = f"{((date(y, m, d) - date(y, 1, 1)).days + 1) / (366 if calendar.isleap(y) else 365):.1%}"
        f1 = f"一、本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，至本({e_yt[:3]})年{int(e_yt[3:5])}月{int(e_yt[5:])}日應達成率為{prog}。"
        f2 = "二、重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項。"
        st.markdown(f"<br>#### {f1.replace(prog, f':red[{prog}]')}\n#### {f2}", unsafe_allow_html=True)

        # --- 自動化流程 ---
        if st.session_state.get("v53_done") != file_hash:
            with st.status("🚀 執行雲端同步與合併...") as s:
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL); ws = sh.get_worksheet(0)
                
                # 寫入資料 (試算表端不能有 HTML)
                h1_raw = ["統計期間", title_wk, "", title_yt, "", title_ly, "", "本年與去年同期比較", "目標值", "達成率"]
                h2_raw = ["取締方式", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "", "", ""]
                full_payload = [h1_raw, h2_raw] + all_rows
                ws.update(range_name='A2', values=full_payload)
                
                # 發送批次格式請求 (合併 + 標紅 + 置中)
                reqs = []
                # 合併儲存格: B2:C2(1,3), D2:E2(3,5), F2:G2(5,7) - 索引皆為 0-based
                for col_pair in [(1,3), (3,5), (5,7)]:
                    reqs.append(get_merge_request(ws.id, col_pair[0], col_pair[1]))
                    reqs.append(get_center_align_request(ws.id, col_pair[0], col_pair[1]))
                
                # 標頭標紅
                for i, txt in [(2, title_wk), (4, title_yt), (6, title_ly)]:
                    reqs.append(get_header_red_req(ws.id, 2, i, txt))
                
                # 末端說明標紅
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
                
                st.session_state["v53_done"] = file_hash
                st.balloons(); s.update(label="雲端合併與同步完成", state="complete")
    except Exception as e: st.error(f"錯誤: {e}")
