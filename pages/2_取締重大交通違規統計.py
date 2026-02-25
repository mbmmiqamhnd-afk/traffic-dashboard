import streamlit as st
import pandas as pd
import re
import io
from datetime import date

# --- 1. 基礎設定 ---
st.set_page_config(page_title="交通違規統計 (修正版)", layout="wide", page_icon="🚔")

UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所',
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所',
    '警備隊': '警備隊', '交通分隊': '交通分隊', '科技執法': '科技執法'
}
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
TARGETS = {
    '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479,
    '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006
}

# --- 2. 工具函數：自動尋找工作表 ---
def get_sheet_name(all_sheets, keyword, default):
    """在所有工作表名稱中尋找包含關鍵字的名稱"""
    for s in all_sheets:
        if keyword in s:
            return s
    return default

# --- 3. 核心解析函數 ---
def parse_excel_data(uploaded_file, keyword, col_indices):
    try:
        content = uploaded_file.getvalue()
        xl = pd.ExcelFile(io.BytesIO(content))
        
        # 自動尋找名稱包含關鍵字的工作表
        target_sheet = get_sheet_name(xl.sheet_names, keyword, "重點違規統計表")
        
        # 讀取資料
        df = pd.read_excel(xl, sheet_name=target_sheet, header=None)
        
        # 提取日期
        info_text = "".join(df.iloc[:5].astype(str).values.flatten())
        match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', info_text)
        start_str = match.group(1) if match else "0000000"
        end_str = match.group(2) if match else "0000000"
        
        unit_results = {}
        for i in range(len(df)):
            row = df.iloc[i]
            unit_name_raw = str(row[0]).strip()
            
            matched_name = None
            for key, val in UNIT_MAP.items():
                if key in unit_name_raw:
                    matched_name = val
                    break
            
            if matched_name:
                def clean_val(v):
                    try:
                        s = str(v).replace(',', '').strip()
                        # 處理 Excel 中的負數或括號
                        if '(' in s and ')' in s: s = '-' + s.replace('(','').replace(')','')
                        return float(s) if s not in ['', 'nan', 'None', '-'] else 0.0
                    except: return 0.0
                
                total_val = sum([clean_val(row[c]) for c in col_indices])
                unit_results[matched_name] = unit_results.get(matched_name, 0) + total_val
        
        return {'data': unit_results, 'start': start_str, 'end': end_str, 'sheet': target_sheet}
    except Exception as e:
        st.error(f"解析檔案 {uploaded_file.name} 時出錯: {e}")
        return None

# --- 4. 主介面 ---
st.markdown("## 🚔 取締重大交通違規統計 (自動匹配版)")

files = st.file_uploader("📂 請上傳 2 個 Focus 檔案", accept_multiple_files=True, type=['xlsx', 'xls'])

if files and len(files) == 2:
    parsed_meta = []
    for f in files:
        # 先讀取第一個工作表判斷日期
        m = parse_excel_data(f, "重點違規統計表", [15, 16, 17])
        if m:
            try:
                s, e = m['start'], m['end']
                d1 = date(int(s[:3])+1911, int(s[3:5]), int(s[5:]))
                d2 = date(int(e[:3])+1911, int(e[3:5]), int(e[5:]))
                m['duration'] = (d2 - d1).days
            except: m['duration'] = 0
            m['file_obj'] = f
            parsed_meta.append(m)
    
    if len(parsed_meta) == 2:
        # 天數長的是「累計檔」(含本年、去年)
        parsed_meta.sort(key=lambda x: x['duration'], reverse=True)
        file_long = parsed_meta[0]['file_obj']
        file_short = parsed_meta[1]['file_obj']
        
        # 提取數據
        # 1. 本期：短檔案 (關鍵字：重點違規統計表) P-R (15-17)
        res_week_all = parse_excel_data(file_short, "重點違規統計表", [15, 16, 17])
        
        # 2. 本年與去年：長檔案 (關鍵字：(1))
        # 先偵測長檔案中是否有名稱為 (1) 的表，如果找不到，退而求其次找包含 "統計" 的表
        res_year_all = parse_excel_data(file_long, "(1)", [15, 16, 17])
        res_last_all = parse_excel_data(file_long, "(1)", [18, 19, 20])
        
        if res_week_all and res_year_all and res_last_all:
            final_table = []
            for u in UNIT_ORDER:
                w = res_week_all['data'].get(u, 0)
                y = res_year_all['data'].get(u, 0)
                l = res_last_all['data'].get(u, 0)
                
                tgt = TARGETS.get(u, 0)
                diff = y - l
                rate = f"{(y/tgt):.1%}" if tgt > 0 else "0%"
                
                final_table.append([u, int(w), int(y), int(l), int(diff), tgt, rate])
                
            df_display = pd.DataFrame(final_table, columns=['單位', '本期數值', '本年累計', '去年同期', '增減比較', '目標值', '達成率'])
            st.success(f"✅ 解析完成！")
            st.info(f"📋 使用工作表：本期({res_week_all['sheet']}) / 歷史({res_year_all['sheet']})")
            st.dataframe(df_display, use_container_width=True)
else:
    st.info("💡 請同時上傳兩個檔案：一個是本期(週/月)報表，一個是累計報表。")
