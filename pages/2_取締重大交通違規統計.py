import streamlit as st
import pandas as pd
import re
import io
from datetime import date

# --- 1. 基礎設定 ---
st.set_page_config(page_title="交通違規統計 (2檔精準版)", layout="wide", page_icon="🚔")

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

# --- 2. 核心解析函數 ---
def parse_excel_data(uploaded_file, sheet_name, col_indices):
    try:
        content = uploaded_file.getvalue()
        # 讀取指定工作表，header 設為 None 方便手動定位
        df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
        # 提取日期 (用於識別檔案天數)
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
                        return float(s) if s not in ['', 'nan', 'None'] else 0.0
                    except: return 0.0
                
                # 加總指定欄位
                total_val = sum([clean_val(row[c]) for c in col_indices])
                unit_results[matched_name] = unit_results.get(matched_name, 0) + total_val
        
        return {'data': unit_results, 'start': start_str, 'end': end_str}
    except Exception as e:
        st.error(f"解析失敗: {e}")
        return None

# --- 3. 主介面 ---
st.markdown("## 🚔 取締重大交通違規統計 (2 檔案精準版)")

files = st.file_uploader("📂 請上傳 2 個 Focus 檔案", accept_multiple_files=True, type=['xlsx', 'xls'])

if files and len(files) == 2:
    # 步驟 A: 識別哪個是「長區間檔案(本年/去年)」，哪個是「短區間檔案(本期)」
    # 這裡簡單判斷：讀取第一個工作表，看日期區間長度
    parsed_meta = []
    for f in files:
        m = parse_excel_data(f, "重點違規統計表", [15, 16, 17]) # 暫讀 P-R 測天數
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
        # 按天數排序：天數長的是「本年/去年檔案」，短的是「本期檔案」
        parsed_meta.sort(key=lambda x: x['duration'], reverse=True)
        file_long = parsed_meta[0]['file_obj']
        file_short = parsed_meta[1]['file_obj']
        
        # 步驟 B: 依照規則提取數據
        # 1. 本期：短檔案的 '重點違規統計表' P-R (15-17)
        res_week = parse_excel_data(file_short, "重點違規統計表", [15, 16, 17])['data']
        # 2. 本年：長檔案的 '重點違規統計表 (1)' P-R (15-17)
        res_year = parse_excel_data(file_long, "重點違規統計表 (1)", [15, 16, 17])['data']
        # 3. 去年：長檔案的 '重點違規統計表 (1)' S-U (18-20)
        res_last = parse_excel_data(file_long, "重點違規統計表 (1)", [18, 19, 20])['data']
        
        # 步驟 C: 組合報表
        final_table = []
        for u in UNIT_ORDER:
            w = res_week.get(u, 0)
            y = res_year.get(u, 0)
            l = res_last.get(u, 0)
            
            tgt = TARGETS.get(u, 0)
            diff = y - l
            rate = f"{(y/tgt):.1%}" if tgt > 0 else "0%"
            
            final_table.append([u, int(w), int(y), int(l), int(diff), tgt, rate])
            
        df_display = pd.DataFrame(final_table, columns=['單位', '本期數值', '本年累計', '去年同期', '增減比較', '目標值', '達成率'])
        st.success(f"✅ 解析完成！(長區間檔案：{parsed_meta[0]['start']}~{parsed_meta[0]['end']})")
        st.dataframe(df_display, use_container_width=True)

elif files:
    st.warning("⚠️ 請確認上傳數量為 2 個檔案（本期檔案 + 包含本去年之累計檔案）。")
