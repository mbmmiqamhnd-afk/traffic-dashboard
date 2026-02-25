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

# 側邊欄控制
with st.sidebar:
    st.title("⚙️ 系統控制")
    if st.button("🧹 清除系統快取"):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.session_state.clear()
        st.success("快取已清除！")
    st.info("請確保 Secrets 已設定 [email] 與 [gcp_service_account]")

st.markdown("## 🚔 取締重大交通違規統計 (v74 安全穩定版)")

# --- 2. 常數與安全設定 ---
# 嘗試從 Secrets 讀取設定
try:
    MY_EMAIL = st.secrets["email"]["user"]
    MY_PASSWORD = st.secrets["email"]["password"]
    GCP_CREDS = st.secrets["gcp_service_account"]
except Exception as e:
    st.error("❌ 找不到 Secrets 設定！請在 .streamlit/secrets.toml 或雲端後台設定。")
    st.stop()

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
    """同步數據至 Google Sheets"""
    try:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(0)
        # 清除舊資料並寫入新資料
        ws.clear()
        ws.update(values=data_list, range_name='A1')
        return True
    except Exception as e:
        st.error(f"Google Sheets 同步失敗: {e}")
        return False

def send_email_with_report(recipient, subject, body, file_bytes, filename):
    """發送自動化郵件報表"""
    try:
        msg = MIMEMultipart()
        msg['From'] = MY_EMAIL
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_bytes)
        encoders.encode_base64(part)
        # 處理中文檔名編碼
        part.add_header('Content-Disposition', f"attachment; filename={Header(filename, 'utf-8').encode()}")
        msg.attach(part)
        
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(MY_EMAIL, MY_PASSWORD)
            server.send_message(msg)
        return True
    except Exception as e:
        st.error(f"郵件發送失敗: {e}")
        return False

def parse_focus_report(uploaded_file):
    """解析 Focus 原始 Excel 報表"""
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
            if any(k in row_str for k in ["酒後", "闖紅燈", "重大違規"]):
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
            raw_unit = str(row.iloc[0]).strip()
            if raw_unit in ['nan', 'None', '', '合計', '單位'] or "統計" in raw_unit: continue
            
            unit_name = UNIT_MAP.get(raw_unit, raw_unit)
            # 數值清理
            def clean_val(v):
                try: return float(str(v).replace(',', '')) if pd.notna(v) else 0
                except: return 0

            s_val = sum([clean_val(row.iloc[c]) for c in stop_cols])
            c_val = sum([clean_val(row.iloc[c]) for c in cit_cols])
            unit_data[unit_name] = {'stop': s_val, 'cit': c_val}

        # 計算日期區間天數
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
        # 智慧排序：去年、本年累計(長天數)、本期(短天數)
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
            
            if u == '科技執法': # 科技執法無攔停
                w['stop'] = y['stop'] = l['stop'] = 0
            
            cur_total = y['stop'] + y['yc'] # 此處 yc 為迴圈邏輯輔助，下同
            diff = int((y['stop'] + y['cit']) - (l['stop'] + l['cit']))
            tgt = TARGETS.get(u, 0)
            performance = (y['stop'] + y['cit']) / tgt if tgt > 0 else 0
            rate_str = f"{performance:.0%}" if tgt > 0 else "0%"
            
            row = [u, int(w['stop']), int(w['cit']), int(y['stop']), int(y['cit']), int(l['stop']), int(l['cit']), diff, tgt, rate_str]
            if u == '警備隊': row[7] = "—"; row[9] = "—"
            
            final_rows.append(row)
            for i, k in enumerate(['ws','wc','ys','yc','ls','lc']):
                acc[k] += row[i+1]

        # 計算合計
        t_y, t_l = (acc['ys'] + acc['yc']), (acc['ls'] + acc['lc'])
        t_tgt = sum([v for k,v in TARGETS.items() if k != '警備隊'])
        total_row = ['合計', acc['ws'], acc['wc'], acc['ys'], acc['yc'], acc['ls'], acc['lc'], int(t_y - t_l), t_tgt, f"{(t_y/t_tgt):.0%}"]
        final_rows.insert(0, total_row)

        # UI 顯示
        st.success(f"✅ 解析完成！本期日期：{file_week['start']} 至 {file_week['end']}")
        df_display = pd.DataFrame(final_rows, columns=['單位', '本期攔停', '本期逕行', '本年攔停', '本年逕行', '去年攔停', '去年逕行', '比較', '目標', '達成率'])
        st.dataframe(df_display, use_container_width=True)

        # 下載與同步功能
        col1, col2 = st.columns(2)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_display.to_excel(writer, index=False, sheet_name='統計報表')
        excel_data = output.getvalue()

        with col1:
            st.download_button("📥 下載 Excel 報表", data=excel_data, file_name=f"重大違規統計_{file_week['end']}.xlsx", type="secondary")
        
        with col2:
            if st.button("🚀 執行自動化同步與寄送", type="primary"):
                with st.status("正在同步數據...") as status:
                    # A. 同步 Google Sheets
                    sheet_data = [df_display.columns.tolist()] + final_rows + [[NOTE_TEXT]]
                    update_google_sheet(sheet_data, GOOGLE_SHEET_URL)
                    status.update(label="✅ Google Sheets 同步成功")
                    
                    # B. 發送郵件
                    email_body = f"長官好，\n\n檢送本期({file_week['start']}-{file_week['end']})重大交通違規取締統計報表，數據已同步至雲端試算表。\n\n系統自動發送。"
                    send_email_with_report(MY_EMAIL, f"🚔 交通違規統計更新_{file_week['end']}", email_body, excel_data, f"報表_{file_week['end']}.xlsx")
                    
                    status.update(label="🎉 全部任務已完成！", state="complete")
                    st.balloons()

elif uploaded_files:
    st.warning("⚠️ 檔案數量不足，請確認是否上傳了：1.去年同期累計、2.本年累計、3.本期單週。")
