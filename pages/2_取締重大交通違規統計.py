import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import smtplib
import gspread
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# --- 1. 基礎設定 ---
st.set_page_config(page_title="取締重大交通違規統計", layout="wide", page_icon="🚔")

# 側邊欄清理功能
if st.sidebar.button("🧹 清除系統快取"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.session_state.clear()
    st.success("快取已清除！")

st.markdown("## 🚔 取締重大交通違規統計 (v75 行位自動適應版)")
st.info("💡 邏輯更新：自動偵測標題列位置，並掃描 A 欄所有列位以匹配單位名稱。")

# ==========================================
# 0. 常數設定
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

# 單位關鍵字比對表
UNIT_MAP = {
    '聖亭': '聖亭所', '龍潭派出所': '龍潭所', '龍潭所': '龍潭所', '中興': '中興所',
    '石門': '石門所', '高平': '高平所', '三和': '三和所',
    '警備隊': '警備隊', '交通分隊': '交通分隊', '交通組': '科技執法', '科技執法': '科技執法'
}
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

TARGETS = {
    '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479,
    '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006
}

NOTE_TEXT = "重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項。"

# --- 工具函數 ---
def update_google_sheet(data_list, sheet_url):
    try:
        if "gcp_service_account" not in st.secrets:
            st.warning("⚠️ Secrets 未設定，略過 Google Sheet 更新。")
            return False
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(0)
        ws.update(range_name='A1', values=data_list)
        return True
    except Exception as e:
        st.error(f"Google Sheets 更新失敗: {e}")
        return False

def send_email(recipient, subject, body, file_bytes, filename):
    try:
        if "email" not in st.secrets: return False
        conf = st.secrets["email"]
        msg = MIMEMultipart()
        msg['From'] = conf["user"]; msg['To'] = recipient; msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_bytes); encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename={Header(filename, 'utf-8').encode()}")
        msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com', 587); server.starttls()
        server.login(conf["user"], conf["password"])
        server.sendmail(conf["user"], recipient, msg.as_string()); server.quit()
        return True
    except: return False

# ==========================================
# 2. 核心解析函數 (v75 修改重點)
# ==========================================
def parse_focus_report(uploaded_file):
    if not uploaded_file: return None
    try:
        content = uploaded_file.getvalue()
        # 先讀取原始矩陣進行結構分析
        df_raw = pd.read_excel(io.BytesIO(content), header=None, nrows=40)
        
        start_date, end_date, header_idx = "", "", -1
        keywords = ["酒後", "闖紅燈", "嚴重超速", "逆向", "轉彎", "蛇行", "不暫停讓行人", "機車"]
        
        # 1. 自動尋找日期與標題列位置
        max_hits = 0
        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
            
            # 抓取日期
            if not start_date:
                match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', row_str)
                if match: start_date, end_date = match.group(1), match.group(2)
            
            # 判定標題列 (包含最多關鍵字的列即為標題)
            hits = sum(1 for k in keywords if k in row_str)
            if hits > max_hits:
                max_hits = hits
                header_idx = i
        
        if header_idx == -1:
            st.error(f"❌ {uploaded_file.name} 找不到標題列")
            return None

        # 2. 根據標題列正式讀取
        df = pd.read_excel(io.BytesIO(content), header=header_idx)
        
        # 3. 找出數據所在的欄位索引
        stop_cols, cit_cols = [], []
        for i in range(len(df.columns)):
            col_name = str(df.columns[i])
            if any(k in col_name for k in keywords) and "路肩" not in col_name:
                stop_cols.append(i)
                cit_cols.append(i + 1) # 假設 攔停 隔壁是 逕行
        
        # 4. 掃描 A 欄各列，抓取單位數據
        unit_data = {}
        for _, row in df.iterrows():
            # 我們鎖定 A 欄 (row.iloc[0]) 進行單位比對
            raw_val = str(row.iloc[0]).strip()
            if raw_val in ['nan', 'None', '', '合計', '單位'] or "統計" in raw_val:
                continue
            
            # 模糊比對單位名稱
            matched_name = None
            for key, short_name in UNIT_MAP.items():
                if key in raw_val:
                    matched_name = short_name
                    break
            
            if matched_name:
                # 數值加總 (處理逗號與空值)
                def clean_val(v):
                    v_str = str(v).replace(',', '').strip()
                    return float(v_str) if v_str not in ['', 'nan', 'None'] else 0.0
                
                s_sum = sum([clean_val(row.iloc[c]) for c in stop_cols if c < len(row)])
                c_sum = sum([clean_val(row.iloc[c]) for c in cit_cols if c < len(row)])
                
                # 若單位重複出現則累加
                if matched_name in unit_data:
                    unit_data[matched_name]['stop'] += s_sum
                    unit_data[matched_name]['cit'] += c_sum
                else:
                    unit_data[matched_name] = {'stop': s_sum, 'cit': c_sum}

        # 5. 計算統計天數
        dur = 0
        try:
            s_d, e_d = re.sub(r'[^\d]', '', start_date), re.sub(r'[^\d]', '', end_date)
            d1 = date(int(s_d[:3])+1911, int(s_d[3:5]), int(s_d[5:]))
            d2 = date(int(e_d[:3])+1911, int(e_d[3:5]), int(e_d[5:]))
            dur = (d2 - d1).days
        except: dur = 0
            
        return {'data': unit_data, 'start': start_date, 'end': end_date, 'duration': dur, 'filename': uploaded_file.name}
    except Exception as e:
        st.error(f"解析錯誤 ({uploaded_file.name}): {e}")
        return None

# ==========================================
# 3. 主程式介面
# ==========================================
uploaded_files = st.file_uploader("📂 請上傳 3 個 Focus 檔案 (去年同期、本年累計、本期)", accept_multiple_files=True, type=['xlsx', 'xls'])

if uploaded_files and len(uploaded_files) >= 3:
    parsed = []
    for f in uploaded_files:
        res = parse_focus_report(f)
        if res: parsed.append(res)
    
    if len(parsed) >= 3:
        # 檔案分類：去年(日期最舊)、本年累計(天數長)、本期(天數短)
        parsed.sort(key=lambda x: x['start'])
        file_last = parsed[0]
        others = sorted(parsed[1:], key=lambda x: x['duration'], reverse=True)
        file_year, file_week = others[0], others[1]

        # 構建結果表
        final_rows = []; acc = {'ws':0, 'wc':0, 'ys':0, 'yc':0, 'ls':0, 'lc':0}
        
        for u in UNIT_ORDER:
            w = file_week['data'].get(u, {'stop':0, 'cit':0})
            y = file_year['data'].get(u, {'stop':0, 'cit':0})
            l = file_last['data'].get(u, {'stop':0, 'cit':0})
            
            # 科技執法無攔停數據
            if u == '科技執法': w['stop'] = y['stop'] = l['stop'] = 0
            
            y_total = y['stop'] + y['cit']
            l_total = l['stop'] + l['cit']
            diff = int(y_total - l_total)
            tgt = TARGETS.get(u, 0)
            rate_val = (y_total / tgt) if tgt > 0 else 0
            
            row = [u, int(w['stop']), int(w['cit']), int(y['stop']), int(y['cit']), int(l['stop']), int(l['cit']), diff, tgt, f"{rate_val:.0%}"]
            
            # 警備隊不計目標與比較
            if u == '警備隊': row[7] = "—"; row[9] = "—"
            
            final_rows.append(row)
            for k, v in zip(['ws','wc','ys','yc','ls','lc'], row[1:7]): acc[k] += v

        # 計算合計列
        t_y, t_l = acc['ys'] + acc['yc'], acc['ls'] + acc['lc']
        t_tgt = sum([v for k,v in TARGETS.items() if k != '警備隊'])
        total_row = ['合計', acc['ws'], acc['wc'], acc['ys'], acc['yc'], acc['ls'], acc['lc'], t_y - t_l, t_tgt, f"{(t_y/t_tgt):.0%}"]
        final_rows.insert(0, total_row)

        # 表格呈現
        st.success(f"✅ 檔案解析成功！本期統計區間：{file_week['start']} ~ {file_week['end']}")
        df_display = pd.DataFrame(final_rows, columns=['單位', '本期攔停', '本期逕行', '本年累計攔停', '本年累計逕行', '去年同期攔停', '去年同期逕行', '比較', '目標值', '達成率'])
        st.dataframe(df_display, use_container_width=True)

        # 檔案匯出與自動化
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_display.to_excel(writer, index=False, sheet_name='Sheet1')
        excel_data = output.getvalue()

        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📥 下載 Excel 報表", data=excel_data, file_name=f"重大違規統計_{file_year['end']}.xlsx")
        with col2:
            if st.button("🚀 執行自動化同步 (Email & Sheet)", type="primary"):
                with st.status("正在同步數據...") as status:
                    # 準備寫入 Sheet 的資料
                    sheet_final = [df_display.columns.tolist()] + final_rows + [[NOTE_TEXT]+[""]*9]
                    update_google_sheet(sheet_final, GOOGLE_SHEET_URL)
                    
                    if "email" in st.secrets:
                        send_email(st.secrets["email"]["user"], f"📊 交通統計更新_{file_year['end']}", "報表如附件。", excel_data, f"報表_{file_year['end']}.xlsx")
                    
                    status.update(label="同步完畢！", state="complete")
                    st.balloons()
elif uploaded_files:
    st.warning("⚠️ 檔案數量不足。請確認上傳了 3 個檔案：1.去年同期、2.今年累計、3.本週(本期)資料。")
