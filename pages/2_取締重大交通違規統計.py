import streamlit as st
import pandas as pd
import numpy as np
import re
import io
from datetime import date

# --- 1. 基礎設定 ---
st.set_page_config(page_title="取締重大交通違規統計", layout="wide", page_icon="🚔")

# --- 0. 設定區 ---
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

# --- 2. 核心解析函數 (依照您的 P-R, S-U 欄位需求) ---
def parse_focus_report(uploaded_file, mode="week"):
    """
    mode: 
    - "week": 讀取 '重點違規統計表', 統計 P-R 欄 (index 15-17)
    - "year": 讀取 '重點違規統計表 (1)', 統計 P-R 欄 (index 15-17)
    - "last": 讀取 '重點違規統計表 (1)', 統計 S-U 欄 (index 18-20)
    """
    try:
        content = uploaded_file.getvalue()
        sheet_name = "重點違規統計表" if mode == "week" else "重點違規統計表 (1)"
        
        # 讀取指定工作表
        df = pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, header=None)
        
        # 尋找日期 (假設在報表前幾列)
        info_text = "".join(df.iloc[:5].astype(str).values.flatten())
        match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', info_text)
        start_date = match.group(1) if match else "0000000"
        end_date = match.group(2) if match else "0000000"
        
        # 確定欄位範圍
        col_indices = [15, 16, 17] if mode in ["week", "year"] else [18, 19, 20]
        
        unit_results = {}
        # 從第5列開始遍歷 (避開標題)
        for i in range(len(df)):
            row = df.iloc[i]
            unit_name_raw = str(row[0]).strip()
            
            # 單位匹配
            matched_name = None
            for key, val in UNIT_MAP.items():
                if key in unit_name_raw:
                    matched_name = val
                    break
            
            if matched_name:
                # 數值清洗加總
                def clean_val(v):
                    try:
                        s = str(v).replace(',', '').strip()
                        return float(s) if s not in ['', 'nan', 'None'] else 0.0
                    except: return 0.0
                
                total_val = sum([clean_val(row[c]) for c in col_indices])
                
                # 存入結果
                if matched_name not in unit_results:
                    unit_results[matched_name] = total_val
        
        # 計算天數
        try:
            d1 = date(int(start_date[:3])+1911, int(start_date[3:5]), int(start_date[5:]))
            d2 = date(int(end_date[:3])+1911, int(end_date[3:5]), int(end_date[5:]))
            duration = (d2 - d1).days
        except: duration = 0
            
        return {'data': unit_results, 'start': start_date, 'end': end_date, 'duration': duration}
    except Exception as e:
        st.error(f"解析 {uploaded_file.name} 失敗: {str(e)}")
        return None

# --- 3. 主介面 ---
st.markdown("## 🚔 取締重大交通違規統計 (欄位精準版)")

# 確保 uploaded_files 在此處被定義
uploaded_files = st.file_uploader("📂 請上傳 3 個 Focus 檔案", accept_multiple_files=True, type=['xlsx', 'xls'])

if uploaded_files and len(uploaded_files) >= 3:
    all_res = []
    for f in uploaded_files:
        # 先以 week 模式讀取來獲取日期與天數
        res = parse_focus_report(f, mode="week")
        if res:
            res['file_obj'] = f
            all_res.append(res)
    
    if len(all_res) >= 3:
        # 排序邏輯：日期最早的是去年，剩下兩個中天數較長的是本年累計
        all_res.sort(key=lambda x: x['start'])
        f_last_raw = all_res[0]
        others = sorted(all_res[1:], key=lambda x: x['duration'], reverse=True)
        f_year_raw, f_week_raw = others[0], others[1]
        
        # 重新根據您的規則進行精確解析
        data_week = parse_focus_report(f_week_raw['file_obj'], mode="week")['data']
        data_year = parse_focus_report(f_year_raw['file_obj'], mode="year")['data']
        data_last = parse_focus_report(f_last_raw['file_obj'], mode="last")['data']
        
        # 組合表格
        final_table = []
        for u in UNIT_ORDER:
            w_val = data_week.get(u, 0)
            y_val = data_year.get(u, 0)
            l_val = data_last.get(u, 0)
            
            # 科技執法通常不計入某些攔停數值，若需歸零可在此處理
            # if u == '科技執法': w_val = 0 ... 
            
            tgt = TARGETS.get(u, 0)
            diff = y_val - l_val
            rate = f"{(y_val/tgt):.1%}" if tgt > 0 else "0%"
            
            final_table.append([u, int(w_val), int(y_val), int(l_val), int(diff), tgt, rate])
            
        df_display = pd.DataFrame(final_table, columns=['單位', '本期數值', '本年累計', '去年同期', '增減比較', '目標值', '達成率'])
        st.success("✅ 報表解析成功！")
        st.dataframe(df_display, use_container_width=True)
else:
    st.info("💡 請上傳三個檔案：分別代表「去年累計」、「本年累計」與「本期(週/月)」。")
