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
st.title("🚦 重大交通違規自動統計 (v48 完整定義版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 
VIOLATION_TARGETS = {'聖亭所': 10, '龍潭所': 12, '中興所': 10, '石門所': 8, '高平所': 6, '三和所': 4, '警備隊': 0, '交通分隊': 15}
UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'}
UNIT_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. 核心格式指令 (Google Sheets API)
# ==========================================
def get_footer_precise_red_req(ws_id, row_idx, col_idx, text):
    """針對說明列：僅將應達成率後的百分比標紅"""
    runs = [{"startIndex": 0, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}}]
    anchor = "應達成率為"
    idx = text.find(anchor)
    if idx != -1:
        search_part = text[idx + len(anchor):]
        match = re.search(r'(\d+\.?\d*%)', search_part)
        if match:
            start_pos = idx + len(anchor) + match.start()
            end_pos = idx + len(anchor) + match.end()
            runs.append({"startIndex": start_pos, "format": {"foregroundColor": {"red": 1.0, "green": 0, "blue": 0}, "bold": True}})
            if end_pos < len(text):
                runs.append({"startIndex": end_pos, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}})
    return {"updateCells": {"rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}], "fields": "userEnteredValue,textFormatRuns", "range": {"sheetId": ws_id, "startRowIndex": row_idx-1, "endRowIndex": row_idx, "startColumnIndex": col_idx-1, "endColumnIndex": col_idx}}}

def get_header_num_red_req(ws_id, row_idx, col_idx, text):
    """標題列：數字與符號標紅"""
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

# ==========================================
# 2. 解析邏輯
# ==========================================
def parse_report(f):
    if not f: return {}, "0000000", "0000000"
    counts, s, e = {}, "0000000", "0000000"
    try:
        f.seek(0)
        top_txt = pd.read_excel(f, header=None, nrows=15).to_string()
        m = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', top_txt)
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
                    nums = [float(str(x).replace(',','')) for x in r if str(x).replace('.','',1).isdigit()]
                    if nums:
                        short = UNIT_MAP.get(u, u)
                        if short in UNIT_ORDER: counts[short] = counts.get(short, 0) + int(nums[-1])
                        u = None
        return counts, s, e
    except: return {}, "0000000", "0000000"

# ==========================================
# 3. 執行介面
# ==========================================
files = st.file_uploader("請上傳重大違規 stoneCnt 報表", accept_multiple_files=True, type=['xlsx', 'xls'])

if files and len(files) >= 3:
    try:
        file_hash = "".join(sorted([f.name + str(f.size) for f in files]))
        f_wk, f_yt, f_ly = None, None, None
        for f in files:
            if "(1)" in f.name: f_yt = f
            elif "(2)" in f.name: f_ly = f
            else: f_wk = f
        
        d_wk, s_wk, e_wk = parse_report(f_wk)
        d_yt, s_yt, e_yt = parse_report(f_yt)
        d_ly, s_ly, e_ly = parse_report(f_ly)

        # 欄位標題 (月日簡化版)
        raw_wk = f"本期 ({s_wk[-4:]}~{e_wk[-4:]})"
        raw_yt = f"本年累計 ({s_yt[-4:]}~{e_yt[-4:]})"
        raw_ly = f"去年累計 ({s_ly[-4:]}~{e_ly[-4:]})"

        def h_html(t): return "".join([f"<span style='color:red; font-weight:bold;'>{c}</span>" if c in "0123456789~().%" else c for c in t])
        h_wk, h_yt, h_ly = map(h_html, [raw_wk, raw_yt, raw_ly])

        # 組裝表格
        body = []
        for u in UNIT_ORDER:
            yv, tv = d_yt.get(u, 0), VIOLATION_TARGETS.get(u, 0)
            body.append({'統計期間': u, h_wk: d_wk.get(u, 0), h_yt: yv, h_ly: d_ly.get(u, 0), '同期比較': yv - d_ly.get(u, 0), '目標值': tv, '達成率': f"{yv/tv:.0%}" if tv > 0 else "—"})
        
        df_final = pd.DataFrame(body)

        # 🚀 全新兩段式說明文字 (動態計算)
        y, m, d = int(e_yt[:3])+1911, int(e_yt[3:5]), int(e_yt[5:])
        prog_str = f"{((date(y, m, d) - date(y, 1, 1)).days + 1) / (366 if calendar.isleap(y) else 365):.1%}"
        
        footer_line1 = f"一、本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，至本({e_yt[:3]})年{int(e_yt[3:5])}月{int(e_yt[5:])}日應達成率為{prog_str}。"
        footer_line2 = "二、重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項"
        
        # 網頁 HTML 標紅百分比
        f_html1 = footer_line1.replace(prog_str, f"<span style='color:red; font-weight:bold;'>{prog_str}</span>")

        st.success("✅ 報表解析成功！")
        st.write(df_final.to_html(escape=False, index=False), unsafe_allow_html=True)
        st.markdown(f"#### {f_html1}\n#### {footer_line2}", unsafe_allow_html=True)

        # --- 自動化流程 ---
        if st.session_state.get("v48_processed") != file_hash:
            with st.status("🚀 正在執行自動化同步...") as s:
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL)
                ws = sh.get_worksheet(0) # 假設重大違規在分頁 1
                
                clean_cols = ['統計期間', raw_wk, raw_yt, raw_ly, '同期比較', '目標值', '達成率']
                ws.update(range_name='A2', values=[clean_cols] + df_final.values.tolist())
                
                # 格式化標題 (B2:D2)
                reqs = [get_header_num_red_req(ws.id, 2, i, t) for i, t in enumerate(clean_cols[1:4], 2)]
                
                # 格式化末端說明 (兩列)
                idx_f1 = 2 + len(df_final) + 1
                idx_f2 = idx_f1 + 1
                
                # 寫入文字並標紅百分比
                ws.update_cell(idx_f1, 1, footer_line1)
                ws.update_cell(idx_f2, 1, footer_line2)
                reqs.append(get_footer_precise_red_req(ws.id, idx_f1, 1, footer_line1))
                
                sh.batch_update({"requests": reqs})
                
                # 自動寄信 (如有 Secrets)
                if "email" in st.secrets:
                    sender = st.secrets["email"]["user"]
                    out = io.BytesIO()
                    df_final.to_excel(out, index=False)
                    with smtplib.SMTP('smtp.gmail.com', 587) as server:
                        server.starttls()
                        server.login(sender, st.secrets["email"]["password"])
                        msg = MIMEMultipart()
                        msg['Subject'] = Header(f"🚦 重大違規報表 - {e_yt}", 'utf-8').encode()
                        msg.attach(MIMEText(f"{footer_line1}\n{footer_line2}", 'plain'))
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(out.getvalue()); encoders.encode_base64(part)
                        part.add_header('Content-Disposition', 'attachment; filename="Report.xlsx"'); msg.attach(part)
                        server.send_message(msg)
                
                st.session_state["v48_processed"] = file_hash
                st.balloons(); s.update(label="全部流程已完成", state="complete")
    except Exception as e: st.error(f"錯誤: {e}")
