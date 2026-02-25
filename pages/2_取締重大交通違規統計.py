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

# --- 1. 基礎設定與環境檢查 ---
st.set_page_config(page_title="取締重大交通違規統計", layout="wide", page_icon="🚔")

# 清除快取按鈕 (放在側邊欄)
if st.sidebar.button("🧹 清除系統快取"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.session_state.clear()
    st.success("快取已清除！")

st.markdown("## 🚔 取締重大交通違規統計 (v73 穩定修復版)")

# --- 2. 常數設定 ---
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所',
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所',
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '交通組': '科技執法'
}
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

TARGETS = {
    '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479,
    '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006
}

NOTE_TEXT = "重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項。"

# --- 3. 工具函數 ---
def update_google_sheet(data_list, sheet_url):
    try:
        if "gcp_service_account" not in st.secrets:
            st.warning("⚠️ 未設定 GCP Secrets，無法更新 Google Sheets。")
            return False
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(0)
        ws.update(range_name='A1', values=data_list)
        return True
    except Exception as e:
        st.error(f"Google Sheets 錯誤: {e}")
        return False

def send_email(recipient, subject, body, file_bytes, filename):
    try:
        if "email" not in st.secrets: return False
        conf = st.secrets["email"]
        msg = MIMEMultipart()
        msg['From'] = conf["user"]
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename={Header(filename, 'utf-8').encode()}")
        msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(conf["user"], conf["password"])
        server.sendmail(conf["user"], recipient, msg.as_string())
        server.quit()
        return True
    except: return False

def parse_focus_report(uploaded_file):
    if not uploaded_file: return None
    try:
        content = uploaded_file.getvalue()
        df_raw = pd.read_excel(io.BytesIO(content), header=None, nrows=25)
        start_date, end_date, header_idx = "", "", -1
        
        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
            if not start_date:
                match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', row_str)
                if match: start_date, end_date = match.group(1), match.group(2)
            if "酒後" in row_str or "闖紅燈" in row_str:
                header_idx = i
        
        if header_idx == -1: return None

        df = pd.read_excel(io.BytesIO(content), header=header_idx)
        keywords = ["酒後", "闖紅燈", "嚴重超速", "逆向", "轉彎", "蛇行", "不暫停讓行人", "機車"]
        stop_cols, cit_cols = [], []
        
        for i in range(len(df.columns)):
            col_name = str(df.columns[i])
            if any(k in col_name for k in keywords) and "路肩" not in col_name:
                stop_cols.append(i)
                cit_cols.append(i + 1)
        
        unit_data = {}
        for _, row in df.iterrows():
            raw_unit = str(row.iloc[0]).strip() # A 欄單位
            if raw_unit in ['nan', 'None', '', '合計', '單位'] or "統計" in raw_unit: continue
            
            unit_name = UNIT_MAP.get(raw_unit, raw_unit)
            s_val = sum([float(str(row.iloc[c]).replace(',', '')) for c in stop_cols if pd.notna(row.iloc[c]) and str(row.iloc[c]).strip() != ''])
            c_val = sum([float(str(row.iloc[c]).replace(',', '')) for c in cit_cols if pd.notna(row.iloc[c]) and str(row.iloc[c]).strip() != ''])
            unit_data[unit_name] = {'stop': s_val, 'cit': c_val}

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

# --- 4. 主程式流程 ---
uploaded_files = st.file_uploader("📂 請上傳 3 個 Focus Excel 檔案", accept_multiple_files=True, type=['xlsx', 'xls'])

if uploaded_files and len(uploaded_files) >= 3:
    parsed = []
    for f in uploaded_files:
        res = parse_focus_report(f)
        if res: parsed.append(res)
    
    if len(parsed) >= 3:
        # 排序：去年、本年累計(天數長)、本期(天數短)
        parsed.sort(key=lambda x: x['start'])
        file_last = parsed[0]
        others = sorted(parsed[1:], key=lambda x: x['duration'], reverse=True)
        file_year, file_week = others[0], others[1]

        final_rows = []
        acc = {'ws':0, 'wc':0, 'ys':0, 'yc':0, 'ls':0, 'lc':0}
        
        for u in UNIT_ORDER:
            w = file_week['data'].get(u, {'stop':0, 'cit':0})
            y = file_year['data'].get(u, {'stop':0, 'cit':0})
            l = file_last['data'].get(u, {'stop':0, 'cit':0})
            if u == '科技執法': w['stop'] = y['stop'] = l['stop'] = 0
            
            diff = int((y['stop']+y['cit']) - (l['stop']+l['cit']))
            tgt = TARGETS.get(u, 0)
            rate = f"{(y['stop']+y['cit'])/tgt:.0%}" if tgt > 0 else "0%"
            
            row = [u, int(w['stop']), int(w['cit']), int(y['stop']), int(y['cit']), int(l['stop']), int(l['cit']), diff, tgt, rate]
            if u == '警備隊': row[7] = "—"; row[9] = "—"
            
            final_rows.append(row)
            for k, v in zip(['ws','wc','ys','yc','ls','lc'], row[1:7]): acc[k] += v

        # 合計列
        t_y, t_l = acc['ys'] + acc['yc'], acc['ls'] + acc['lc']
        t_tgt = sum([v for k,v in TARGETS.items() if k != '警備隊'])
        total_row = ['合計', acc['ws'], acc['wc'], acc['ys'], acc['yc'], acc['ls'], acc['lc'], t_y - t_l, t_tgt, f"{t_y/t_tgt:.0%}"]
        final_rows.insert(0, total_row)

        # 顯示表格 (修正後的安全版本)
        st.success("✅ 分析成功")
        df_display = pd.DataFrame(final_rows, columns=['單位', '本期攔停', '本期逕行', '本年攔停', '本年逕行', '去年攔停', '去年逕行', '比較', '目標', '達成率'])
        st.dataframe(df_display, use_container_width=True)

        # 功能區域
        col1, col2 = st.columns(2)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_display.to_excel(writer, index=False, sheet_name='Sheet1')
        excel_data = output.getvalue()

        with col1:
            st.download_button("📥 下載 Excel", data=excel_data, file_name=f"統計_{file_year['end']}.xlsx")
        
        with col2:
            if st.button("🚀 執行自動化同步", type="primary"):
                with st.status("執行中...") as status:
                    sheet_data = [df_display.columns.tolist()] + final_rows + [[NOTE_TEXT]+[""]*9]
                    update_google_sheet(sheet_data, GOOGLE_SHEET_URL)
                    if "email" in st.secrets:
                        send_email(st.secrets["email"]["user"], f"🚔 交通統計更新_{file_year['end']}", "請查收附件。", excel_data, f"報表_{file_year['end']}.xlsx")
                    status.update(label="同步完畢！", state="complete")
                    st.balloons()

elif uploaded_files:
    st.warning("⚠️ 檔案數量不足，請至少上傳 3 個（去年同期、本年累計、本期）。")
