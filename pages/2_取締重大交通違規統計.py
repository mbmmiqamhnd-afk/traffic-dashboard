import streamlit as st
import pandas as pd
import re
import io
import smtplib
import gspread
import calendar
import pypdf
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# --- 初始化配置 ---
st.set_page_config(page_title="重大交通違規統計", layout="wide", page_icon="🚦")
st.title("🚦 重大交通違規統計 (v60 Focus PDF 智能解析版)")

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
# 2. PDF 解析引擎 (針對 Focus 報表)
# ==========================================
def parse_focus_pdf(file_obj):
    counts = {} # {unit: [wk_int, wk_rem, yt_int, yt_rem, ly_int, ly_rem]}
    dates = {"wk": "0000~0000", "yt": "0000~0000", "ly": "0000~0000"}
    
    try:
        reader = pypdf.PdfReader(file_obj)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        
        # 1. 抓取日期
        # 格式範例: 本期(1217~1223)本年累計(0101~1223)去年累計(0101~1223)
        m_wk = re.search(r'本期\((\d+~\d+)\)', text)
        m_yt = re.search(r'本年累計\((\d+~\d+)\)', text)
        m_ly = re.search(r'去年累計\((\d+~\d+)\)', text)
        if m_wk: dates["wk"] = m_wk.group(1)
        if m_yt: dates["yt"] = m_yt.group(1)
        if m_ly: dates["ly"] = m_ly.group(1)

        # 2. 抓取數據行
        # 格式範例: 聖亭所 3) 0) 199) 1097) 171) 1863) ...
        # 注意: 科技執法有時會黏在一起如 "科技執法0)"
        
        for unit in UNIT_ORDER + ['合計']:
            # 建構 Regex: 單位名稱 + 接著一串數字與右括號
            # 容許單位後方可能沒有空格 (針對科技執法)
            pattern = re.compile(f"{unit}.*?([\d\-\)]+.*)")
            match = pattern.search(text)
            if match:
                data_str = match.group(1)
                # 清理數據: 移除 ')' 和 '%'，將負號保留
                # 分割邏輯: 數字可能黏著 ')'，如 "3) 0)"
                cleaned = data_str.replace(')', ' ').replace('%', ' ')
                tokens = [t for t in cleaned.split() if t.replace('-','').isdigit()]
                
                # 預期至少有 6 個數據: 本期(攔/逕), 本年(攔/逕), 去年(攔/逕)
                if len(tokens) >= 6:
                    nums = [int(t) for t in tokens[:6]]
                    counts[unit] = nums
    except Exception as e:
        st.error(f"PDF 解析錯誤: {e}")
        return {}, dates

    return counts, dates

# ==========================================
# 3. 畫面顯示與自動化
# ==========================================
files = st.file_uploader("上傳 Focus 系列報表 (支援 PDF/CSV格式)", accept_multiple_files=True)

if files:
    try:
        # 尋找並解析檔案
        target_file = files[0] # 假設只需一個 PDF 報表即可包含所有數據
        file_hash = target_file.name + str(target_file.size)
        
        # 執行解析
        data_map, date_map = parse_focus_pdf(target_file)
        
        if not data_map:
            st.warning("無法從檔案中讀取數據，請確認上傳的是 Focus 交通違規統計報表。")
            st.stop()

        # 標題日期
        title_wk = f"本期({date_map['wk']})"
        title_yt = f"本年累計({date_map['yt']})"
        title_ly = f"去年累計({date_map['ly']})"
        
        def red_h(t): return "".join([f"<span style='color:red; font-weight:bold;'>{c}</span>" if c in "0123456789~().%" else c for c in t])
        
        # 網頁 HTML 表頭
        h_html = ["統計期間", red_h(title_wk), "", red_h(title_yt), "", red_h(title_ly), "", "同期比較", "目標值", "達成率"]
        h_raw = ["統計期間", title_wk, "", title_yt, "", title_ly, "", "本年與去年同期比較", "目標值", "達成率"]

        # 數據組裝
        rows = []
        for u in UNIT_ORDER:
            # data: [wk_int, wk_rem, yt_int, yt_rem, ly_int, ly_rem]
            vals = data_map.get(u, [0, 0, 0, 0, 0, 0])
            yt_tot = vals[2] + vals[3]
            ly_tot = vals[4] + vals[5]
            target = VIOLATION_TARGETS.get(u, 0)
            
            # 列數據: 單位, 本期攔, 本期逕, 本年攔, 本年逕, 去年攔, 去年逕, 比較, 目標, 達成率
            row = [u, vals[0], vals[1], vals[2], vals[3], vals[4], vals[5], yt_tot - ly_tot, target, f"{yt_tot/target:.0%}" if target > 0 else "—"]
            rows.append(row)
        
        # 合計列 (直接從 PDF 抓取或重新計算)
        # 這裡選擇使用 PDF 抓取的合計值以確保一致性，若無則計算
        if '合計' in data_map:
            s_vals = data_map['合計']
            total_target = VIOLATION_TARGETS.get('合計', 11817)
            s_yt_tot = s_vals[2] + s_vals[3]
            s_ly_tot = s_vals[4] + s_vals[5]
            total_row = ["合計", s_vals[0], s_vals[1], s_vals[2], s_vals[3], s_vals[4], s_vals[5], s_yt_tot - s_ly_tot, total_target, f"{s_yt_tot/total_target:.0%}" if total_target > 0 else "0%"]
        else:
            # 備用計算
            df_calc = pd.DataFrame(rows)
            sums = df_calc.iloc[:, 1:7].apply(pd.to_numeric).sum()
            total_target = 11817
            total_row = ["合計", sums[1], sums[2], sums[3], sums[4], sums[5], sums[6], (sums[3]+sums[4])-(sums[5]+sums[6]), total_target, "0%"]

        # 取締方式列
        method_row = ["取締方式", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "現場攔停", "逕行舉發", "", "", ""]
        
        all_rows = [method_row, total_row] + rows
        st.success("✅ Focus PDF 報表解析成功！")
        
        # 網頁渲染
        header_row = "".join([f"<th>{x}</th>" for x in h_html])
        body_rows = "".join([f"<tr>{''.join([f'<td>{x}</td>' for x in r])}</tr>" for r in all_rows])
        st.write(f"<table><thead><tr>{header_row}</tr></thead><tbody>{body_rows}</tbody></table>", unsafe_allow_html=True)

        # 說明文字
        # 從日期字串解析年分
        try:
            # 假設日期格式為 MMDD，需要從 PDF 標題或其他地方推斷年份，這裡暫用今年
            curr_year = date.today().year
            # 如果是 12 月份的報表，可能需要注意
            d_str = date_map['yt'].split('~')[1] # 取結束日期 "1223"
            mon = int(d_str[:2])
            day = int(d_str[2:])
            # 計算天數比例
            prog = f"{((date(curr_year, mon, day) - date(curr_year, 1, 1)).days + 1) / (366 if calendar.isleap(curr_year) else 365):.1%}"
            e_yt_str = f"{curr_year-1911}年{mon}月{day}日"
        except:
            prog = "0.0%"
            e_yt_str = "114年12月XX日"

        f1 = f"一、本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt_str} (入案日期)應達成率為{prog}。"
        f2 = "二、重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項。"
        st.markdown(f"<br>#### {f1.replace(prog, f':red[{prog}]')}\n#### {f2}", unsafe_allow_html=True)

        # --- 自動化流程 ---
        if st.session_state.get("v60_done") != file_hash:
            with st.status("🚀 執行 Focus 數據同步...") as s:
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL); ws = sh.get_worksheet(0)
                
                full_payload = [h_raw] + all_rows
                ws.update(range_name='A2', values=full_payload)
                
                reqs = [get_header_red_req(ws.id, 2, i, h_raw[i-1]) for i in [2, 4, 6]]
                idx_f = 2 + len(full_payload) + 1
                ws.update_cell(idx_f, 1, f1); ws.update_cell(idx_f+1, 1, f2)
                reqs.append(get_footer_percent_red_req(ws.id, idx_f, 1, f1))
                sh.batch_update({"requests": reqs})
                
                if "email" in st.secrets:
                    out = io.BytesIO(); pd.DataFrame(full_payload).to_excel(out, index=False)
                    server = smtplib.SMTP('smtp.gmail.com', 587); server.starttls()
                    server.login(st.secrets["email"]["user"], st.secrets["email"]["password"])
                    msg = MIMEMultipart()
                    msg['Subject'] = Header(f"🚦 Focus 違規報表 - {e_yt_str}", "utf-8").encode()
                    msg.attach(MIMEText(f"{f1}\n{f2}", "plain"))
                    part = MIMEBase("application", "octet-stream"); part.set_payload(out.getvalue())
                    encoders.encode_base64(part); part.add_header("Content-Disposition", 'attachment; filename="Focus_Report.xlsx"')
                    msg.attach(part); server.send_message(msg); server.quit()
                
                st.session_state["v60_done"] = file_hash
                st.balloons(); s.update(label="完成", state="complete")
    except Exception as e: st.error(f"解析失敗: {e}")
