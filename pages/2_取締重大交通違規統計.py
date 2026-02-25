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

st.markdown("## 🚔 取締重大交通違規統計 (v79 數值精準修正版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

# 單位對照表：增加精確度以防止誤抓
UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所',
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所',
    '警備隊': '警備隊', '交通分隊': '交通分隊', '科技執法': '科技執法', '交通組': '科技執法'
}
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

TARGETS = {
    '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479,
    '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006
}

NOTE_TEXT = "重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項。"

# --- Google Sheet 更新工具 ---
def update_google_sheet(data_list, sheet_url):
    try:
        if "gcp_service_account" not in st.secrets:
            return False
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(0)
        ws.update(range_name='A1', values=data_list)
        return True
    except: return False

# ==========================================
# 2. 核心解析函數 (v79 修復數值重複問題)
# ==========================================
def parse_focus_report(uploaded_file):
    if not uploaded_file: return None
    try:
        content = uploaded_file.getvalue()
        # 讀取前 40 列尋找標題列
        df_raw = pd.read_excel(io.BytesIO(content), header=None, nrows=40)
        
        start_date, end_date, header_idx = "", "", -1
        keywords = ["酒後", "闖紅燈", "嚴重超速", "逆向", "轉彎", "蛇行", "不暫停讓行人", "機車"]
        
        for i, row in df_raw.iterrows():
            row_str = "".join([str(x) for x in row.values if pd.notna(x)])
            if not start_date:
                match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', row_str)
                if match: start_date, end_date = match.group(1), match.group(2)
            # 必須同時包含多個關鍵字才認定為標題列
            if sum(1 for k in keywords if k in row_str) >= 3:
                header_idx = i
                break
        
        if header_idx == -1: return None
        df = pd.read_excel(io.BytesIO(content), header=header_idx)
        
        # 1. 找出有效數據欄位 (嚴格排除 P-U 欄位 Index 15~20)
        stop_cols, cit_cols = [], []
        for i in range(len(df.columns)):
            if 15 <= i <= 20: continue # 封鎖 P-U 欄
            
            col_name = str(df.columns[i])
            if any(k in col_name for k in keywords) and "路肩" not in col_name:
                # 確保 逕行 欄位 (i+1) 也不在封鎖區
                target_cit = i + 1
                if target_cit < len(df.columns) and not (15 <= target_cit <= 20):
                    stop_cols.append(i)
                    cit_cols.append(target_cit)
        
        # 2. 讀取數據
        unit_data = {}
        for _, row in df.iterrows():
            raw_val = str(row.iloc[0]).strip()
            if raw_val in ['nan', 'None', '', '合計', '單位'] or "統計" in raw_val: continue
            
            # 單位精確匹配 (防止交通分隊抓到科技執法)
            matched_name = None
            if "科技" in raw_val: 
                matched_name = "科技執法"
            else:
                for key, short_name in UNIT_MAP.items():
                    if key in raw_val:
                        matched_name = short_name
                        break
            
            if matched_name:
                def clean(v):
                    try:
                        s = str(v).replace(',', '').strip()
                        return float(s) if s not in ['', 'nan', 'None'] else 0.0
                    except: return 0.0

                s_val = sum([clean(row.iloc[c]) for c in stop_cols])
                c_val = sum([clean(row.iloc[c]) for c in cit_cols])
                
                # 若該單位已存在，僅取第一筆有效數據 (通常是主表列)，避免重複加總
                if matched_name not in unit_data:
                    unit_data[matched_name] = {'stop': s_val, 'cit': c_val}
                elif s_val > 0 or c_val > 0:
                    # 如果龍潭交通分隊已經有值，且目前這一行是 0，就不覆蓋
                    # 反之，若目前這一行有值，才考慮是否累加（但通常不建議累加，Excel 裡重複出現通常是小計）
                    pass 

        dur = 0
        try:
            s_d, e_d = re.sub(r'[^\d]', '', start_date), re.sub(r'[^\d]', '', end_date)
            d1 = date(int(s_d[:3])+1911, int(s_d[3:5]), int(s_d[5:]))
            d2 = date(int(e_d[:3])+1911, int(e_d[3:5]), int(e_d[5:]))
            dur = (d2 - d1).days
        except: dur = 0
            
        return {'data': unit_data, 'start': start_date, 'end': end_date, 'duration': dur, 'filename': uploaded_file.name}
    except: return None

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
        # 分類：去年、本年累計(長)、本期(短)
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
            tgt = TARGETS.get(u, 0)
            diff = int(y_tot - l_tot)
            rate = f"{y_tot/tgt:.0%}" if tgt > 0 else "0%"
            
            if u == '警備隊': diff = "—"; rate = "—"
            
            row = [u, int(w['stop']), int(w['cit']), int(y['stop']), int(y['cit']), int(l['stop']), int(l['cit']), diff, tgt, rate]
            final_rows.append(row)
            for k, v in zip(['ws','wc','ys','yc','ls','lc'], row[1:7]): acc[k] += v

        # 合計列
        t_y, t_l = acc['ys'] + acc['yc'], acc['ls'] + acc['lc']
        t_tgt = sum([v for k,v in TARGETS.items() if k != '警備隊'])
        total_row = ['合計', acc['ws'], acc['wc'], acc['ys'], acc['yc'], acc['ls'], acc['lc'], t_y - t_l, t_tgt, f"{(t_y/t_tgt):.0%}"]
        final_rows.insert(0, total_row)

        st.success(f"✅ 解析完成！(本期區間: {file_week['start']} ~ {file_week['end']})")
        df_display = pd.DataFrame(final_rows, columns=['單位', '本期攔停', '本期逕行', '本年攔停', '本年逕行', '去年攔停', '去年逕行', '比較', '目標', '達成率'])
        st.dataframe(df_display, use_container_width=True)
        
        if st.button("🚀 同步至 Google Sheets", type="primary"):
            if update_google_sheet([df_display.columns.tolist()] + final_rows, GOOGLE_SHEET_URL):
                st.success("同步成功！")
                st.balloons()
            else: st.error("同步失敗，請檢查 Secrets。")

elif uploaded_files:
    st.warning("⚠️ 需上傳 3 個檔案。")
