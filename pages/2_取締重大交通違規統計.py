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

if st.sidebar.button("🧹 清除系統快取"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.session_state.clear()
    st.success("快取已清除！")

st.markdown("## 🚔 取締重大交通違規統計 (v77 偵錯強化版)")

# ==========================================
# 0. 常數設定
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

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

# ==========================================
# 2. 核心解析函數 (v77 偵錯版)
# ==========================================
def parse_focus_report(uploaded_file):
    if not uploaded_file: return None
    try:
        content = uploaded_file.getvalue()
        df_raw = pd.read_excel(io.BytesIO(content), header=None, nrows=40)
        
        start_date, end_date, header_idx = "", "", -1
        keywords = ["酒後", "闖紅燈", "嚴重超速", "逆向", "轉彎", "蛇行", "不暫停讓行人", "機車"]
        
        # 1. 尋找日期與標題
        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
            if not start_date:
                match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', row_str)
                if match: start_date, end_date = match.group(1), match.group(2)
            if "酒後" in row_str or "闖紅燈" in row_str:
                header_idx = i
                break # 找到第一個標題列即停止
        
        if header_idx == -1: return None
        df = pd.read_excel(io.BytesIO(content), header=header_idx)
        
        # 2. 偵測數據欄位 (排除 P-U 欄，即 Index 15-20)
        stop_cols, cit_cols = [], []
        debug_info = []
        
        for i in range(len(df.columns)):
            # 排除 P(15) 至 U(20)
            if 15 <= i <= 20: continue
            
            col_name = str(df.columns[i])
            if any(k in col_name for k in keywords) and "路肩" not in col_name:
                # 攔停欄位在 i，逕行欄位在 i+1
                if i+1 < len(df.columns) and not (15 <= (i+1) <= 20):
                    stop_cols.append(i)
                    cit_cols.append(i + 1)
                    debug_info.append(f"• 項目: {col_name} (欄位 Index {i} & {i+1})")
        
        # 3. 抓取單位數據
        unit_data = {}
        for _, row in df.iterrows():
            raw_val = str(row.iloc[0]).strip()
            if raw_val in ['nan', 'None', '', '合計', '單位'] or "統計" in raw_val: continue
            
            matched_name = None
            for key, short_name in UNIT_MAP.items():
                if key in raw_val:
                    matched_name = short_name
                    break
            
            if matched_name:
                def clean_val(v):
                    v_str = str(v).replace(',', '').strip()
                    return float(v_str) if v_str not in ['', 'nan', 'None'] else 0.0
                
                s_sum = sum([clean_val(row.iloc[c]) for c in stop_cols if c < len(row)])
                c_sum = sum([clean_val(row.iloc[c]) for c in cit_cols if c < len(row)])
                
                if matched_name in unit_data:
                    unit_data[matched_name]['stop'] += s_sum
                    unit_data[matched_name]['cit'] += c_sum
                else:
                    unit_data[matched_name] = {'stop': s_sum, 'cit': c_sum}

        dur = 0
        try:
            s_d, e_d = re.sub(r'[^\d]', '', start_date), re.sub(r'[^\d]', '', end_date)
            d1 = date(int(s_d[:3])+1911, int(s_d[3:5]), int(s_d[5:]))
            d2 = date(int(e_d[:3])+1911, int(e_d[3:5]), int(e_d[5:]))
            dur = (d2 - d1).days
        except: dur = 0
            
        return {'data': unit_data, 'start': start_date, 'end': end_date, 'duration': dur, 'debug': debug_info, 'filename': uploaded_file.name}
    except Exception as e:
        st.error(f"解析失敗: {e}"); return None

# ==========================================
# 3. 主程式介面
# ==========================================
uploaded_files = st.file_uploader("📂 請上傳 3 個 Focus 檔案", accept_multiple_files=True, type=['xlsx', 'xls'])

if uploaded_files and len(uploaded_files) >= 3:
    parsed = []
    for f in uploaded_files:
        res = parse_focus_report(f)
        if res: parsed.append(res)
    
    if len(parsed) >= 3:
        # 顯示偵錯日誌
        with st.expander("🔍 欄位偵測日誌 (若統計錯誤請展開核對)"):
            for p in parsed:
                st.write(f"**檔案: {p['filename']}**")
                for info in p['debug']: st.write(info)

        # 排序與分類
        parsed.sort(key=lambda x: x['start'])
        file_last = parsed[0]
        others = sorted(parsed[1:], key=lambda x: x['duration'], reverse=True)
        file_year, file_week = others[0], others[1]

        final_rows = []; acc = {'ws':0, 'wc':0, 'ys':0, 'yc':0, 'ls':0, 'lc':0}
        for u in UNIT_ORDER:
            w = file_week['data'].get(u, {'stop':0, 'cit':0})
            y = file_year['data'].get(u, {'stop':0, 'cit':0})
            l = file_last['data'].get(u, {'stop':0, 'cit':0})
            if u == '科技執法': w['stop'] = y['stop'] = l['stop'] = 0
            
            y_tot = y['stop'] + y['cit']
            l_tot = l['stop'] + l['cit']
            row = [u, int(w['stop']), int(w['cit']), int(y['stop']), int(y['cit']), int(l['stop']), int(l['cit']), int(y_tot - l_tot), TARGETS.get(u, 0), ""]
            
            tgt = TARGETS.get(u, 0)
            row[9] = f"{y_tot/tgt:.0%}" if tgt > 0 else "0%"
            if u == '警備隊': row[7] = "—"; row[9] = "—"
            
            final_rows.append(row)
            for k, v in zip(['ws','wc','ys','yc','ls','lc'], row[1:7]): acc[k] += v

        # 合計列
        t_y, t_l = acc['ys'] + acc['yc'], acc['ls'] + acc['lc']
        t_tgt = sum([v for k,v in TARGETS.items() if k != '警備隊'])
        total_row = ['合計', acc['ws'], acc['wc'], acc['ys'], acc['yc'], acc['ls'], acc['lc'], t_y - t_l, t_tgt, f"{(t_y/t_tgt):.0%}"]
        final_rows.insert(0, total_row)

        st.dataframe(pd.DataFrame(final_rows, columns=['單位', '本期攔停', '本期逕行', '本年攔停', '本年逕行', '去年攔停', '去年逕行', '比較', '目標', '達成率']), use_container_width=True)
        
        if st.button("🚀 同步至 Google Sheets", type="primary"):
            update_google_sheet([['單位', '本期攔停', '本期逕行', '本年攔停', '本年逕行', '去年攔停', '去年逕行', '比較', '目標', '達成率']] + final_rows, GOOGLE_SHEET_URL)
            st.success("同步成功！")
elif uploaded_files:
    st.warning("⚠️ 需上傳 3 個檔案。")
