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
st.title("🚦 重大交通違規統計 (v59 Focus 專用解析版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

# 重大違規目標值 (請依據您的公文數據微調)
VIOLATION_TARGETS = {
    '合計': 11817, '科技執法': 0, '聖亭所': 1200, '龍潭所': 1500, '中興所': 1200, 
    '石門所': 1000, '高平所': 800, '三和所': 500, '警備隊': 0, '交通分隊': 1000
}

UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', 
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', 
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '科技執法': '科技執法'
}

UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. Google 試算表格式化引擎
# ==========================================
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
# 2. Focus 報表解析邏輯
# ==========================================
def parse_focus_report(f):
    if not f: return {}, "0000000", "0000000"
    counts = {}
    s, e = "0000000", "0000000"
    try:
        f.seek(0)
        xls = pd.ExcelFile(f)
        for sn in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sn, header=None).astype(str)
            
            # 1. 偵測日期 (Focus 報表通常在 A1 或 B1)
            full_text = " ".join(df.iloc[:5, :].values.flatten())
            m = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', full_text)
            if m: s, e = m.group(1), m.group(2)
            
            # 2. 座標偵測：尋找「攔停」與「逕行」欄位位置
            idx_int, idx_rem = -1, -1
            for r_idx in range(min(30, len(df))):
                row_vals = df.iloc[r_idx].tolist()
                for c_idx, val in enumerate(row_vals):
                    if "攔停" in val: idx_int = c_idx
                    if "逕行" in val: idx_rem = c_idx
            
            # 若沒偵測到，使用預設值
            if idx_int == -1: idx_int = 1
            if idx_rem == -1: idx_rem = 2

            # 3. 數據抓取
            for _, row in df.iterrows():
                row_str = " ".join(row.tolist())
                # 比對單位名稱
                found_unit = None
                for full_name, short_name in UNIT_MAP.items():
                    if short_name in row_str:
                        found_unit = short_name
                        break
                
                # 如果該列包含「合計」或「總計」且已識別單位
                if ("合計" in row_str or "總計" in row_str) and found_unit:
                    try:
                        v_int = row[idx_int].replace(',', '')
                        v_rem = row[idx_rem].replace(',', '')
                        val_int = int(float(v_int)) if v_int.replace('.','',1).isdigit() else 0
                        val_rem = int(float(v_rem)) if v_rem.replace('.','',1).isdigit() else 0
                        
                        if found_unit not in counts: counts[found_unit] = [0, 0]
                        counts[found_unit][0] += val_int
                        counts[found_unit][1] += val_rem
                    except: pass
        return counts, s, e
    except Exception as ex:
        st.error(f"Focus 報表解析失敗: {ex}")
        return {}, "0000000", "0000000"

# ==========================================
# 3. 畫面與執行
# ==========================================
files = st.file_uploader("上傳 Focus 系列報表 (1.本期 2.本年 3.去年)", accept_multiple_files=True, type=['xlsx', 'xls'])

if files and len(files) >= 3:
    try:
        file_hash = "".join(sorted([f.name + str(f.size) for f in files]))
        f_wk, f_yt, f_ly = None, None, None
        for f in files:
            if "(1)" in f.name: f_yt = f
            elif "(2)" in f.name: f_ly = f
            else: f_wk = f
        
        # 執行解析
        d_wk, s_wk, e_wk = parse_focus_report(f_wk)
        d_yt, s_yt, e_yt = parse_focus_report(f_yt)
        d_ly, s_ly, e_ly = parse_focus_report(f_ly)

        # 標題日期
        title_wk = f"本期({s_wk[-4:]}~{e_wk[-4:]})"
        title_yt = f"本年累計({s_yt[-4:]}~{e_yt[-4:]})"
        title_ly = f"去年累計({s_ly[-4:]}~{e_ly[-4:]})"
        
        def red_h(t): return "".join([f"<span style='color:red; font-weight:bold;'>{c}</span>" if c in "0123456789~().%" else c for c in t])

        # 網頁預覽表頭
        h_html = ["統計期間", red_h(title_wk), "", red_h(title_yt), "", red_h(title_ly), "", "同期比較", "目標值", "達成率"]
        h_raw = ["統計期間", title_wk, "", title_yt, "", title_ly, "", "本年與去年同期比較", "目標值", "達成率"]

        # 數據計算與排序
        rows = []
        for u in UNIT_ORDER:
            wk = d_wk.get(u, [0, 0]); yt = d_yt.get(u, [0, 0]); ly = d_ly.get(u, [0, 0])
            yt_tot = sum(yt); ly_tot = sum(ly); target = VIOLATION_TARGETS.get(u, 0)
            rows.append([u, wk[0], wk[1], yt[0], yt[1], ly[0], ly[1], yt_tot - ly_tot, target, f"{yt_tot/target:.0%}" if target > 0 else "—"])
        
        df_calc = pd.DataFrame(rows)
        sums = df_calc.iloc[:, 1:9].apply(pd.to_numeric).sum()
        total_target = VIOLATION_TARGETS.get('合計', sums[8])
        total_row = ["合計", sums[1], sums[2], sums[3], sums[4], sums[5], sums[6], sums[7], total_target, f"{(sums[3]+sums[4])/total_target:.0%}" if total_target > 0 else "0%"]
        
        # 🚀 取締方式列
        method_row = ["取締方式", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "", "", ""]
        
        all_rows = [method_row, total_row] + rows
        st.success("✅ Focus 報表解析成功！")
        
        # 顯示網頁表格
        header_tr = "".join([f"<th>{x}</th>" for x in h_html])
        body_tr = "".join([f"<tr>{''.join([f'<td>{x}</td>' for x in r])}</tr>" for r in all_rows])
        st.write(f"<table><thead><tr>{header_tr}</tr></thead><tbody>{body_tr}</tbody></table>", unsafe_allow_html=True)

        # 備註說明
        y, m, d = int(e_yt[:3])+1911, int(e_yt[3:5]), int(e_yt[5:])
        prog = f"{((date(y, m, d) - date(y, 1, 1)).days + 1) / (366 if calendar.isleap(y) else 365):.1%}"
        f1 = f"一、本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt[:3]}年{int(e_yt[3:5])}月{int(e_yt[5:])}日 (入案日期)應達成率為{prog}。"
        f2 = "二、重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項。"
        st.markdown(f"<br>#### {f1.replace(prog, f':red[{prog}]')}\n#### {f2}", unsafe_allow_html=True)

        # --- 自動化流程 ---
        if st.session_state.get("v59_done") != file_hash:
            with st.status("🚀 執行 Focus 數據自動化處理...") as s:
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL); ws = sh.get_worksheet(0)
                
                full_payload = [h_raw] + all_rows
                ws.update(range_name='A2', values=full_payload)
                
                # 標題標紅
                reqs = [get_header_red_req(ws.id, 2, i, h_raw[i-1]) for i in [2, 4, 6]]
                # 備註標紅
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
                    msg['Subject'] = Header(f"🚦 重大違規(Focus)報表 - {e_yt}", "utf-8").encode()
                    msg.attach(MIMEText(f"{f1}\n{f2}", "plain"))
                    part = MIMEBase("application", "octet-stream"); part.set_payload(out.getvalue())
                    encoders.encode_base64(part); part.add_header("Content-Disposition", 'attachment; filename="Focus_Report.xlsx"')
                    msg.attach(part); server.send_message(msg); server.quit()
                
                st.session_state["v59_done"] = file_hash
                st.balloons(); s.update(label="數據解析與雲端同步全數完成", state="complete")
    except Exception as e: st.error(f"系統錯誤: {e}")
