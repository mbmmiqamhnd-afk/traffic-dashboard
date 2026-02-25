import streamlit as st
import pandas as pd
import re
import io
from datetime import date

# --- 1. 定義關鍵字與目標 ---
# 用來在「變動的列」中捕捉正確的單位
UNIT_MAPPING = {
    '科技': '科技執法', '聖亭': '聖亭所', '龍潭': '龍潭所', '中興': '中興所',
    '石門': '石門所', '高平': '高平所', '三和': '三和所', '警備': '警備隊', '分隊': '交通分隊'
}
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
TARGETS = {
    '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479,
    '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006
}

# --- 2. 核心解析函數：動態掃描列 ---
def parse_excel_dynamic(uploaded_file, keyword_sheet, col_indices):
    try:
        content = uploaded_file.getvalue()
        xl = pd.ExcelFile(io.BytesIO(content))
        
        # A. 尋找正確的工作表
        target_sheet = xl.sheet_names[0]
        for s in xl.sheet_names:
            if keyword_sheet in s:
                target_sheet = s
                break
        
        df = pd.read_excel(xl, sheet_name=target_sheet, header=None)
        
        # B. 取得日期區間 (通常在表頭前幾列)
        info_text = "".join(df.iloc[:10].astype(str).values.flatten())
        match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', info_text)
        start_str = match.group(1) if match else "0000000"
        end_str = match.group(2) if match else "0000000"
        
        unit_results = {}
        
        # C. 核心：遍歷每一列，動態尋找單位名稱
        for _, row in df.iterrows():
            first_cell = str(row.iloc[0]).strip() # 檢查第一欄 (A欄)
            
            # 對於這一列，檢查是否命中我們的關鍵字
            matched_unit = None
            for key, standard_name in UNIT_MAPPING.items():
                if key in first_cell:
                    matched_unit = standard_name
                    break
            
            # 如果命中了單位，且這一列不是標題列（排除包含 "單位" 或 "合計" 字眼的列）
            if matched_unit and "合計" not in first_cell:
                def clean_val(v):
                    try:
                        # 處理 Excel 中的逗號與空值
                        s = str(v).replace(',', '').strip()
                        return float(s) if s not in ['', 'nan', 'None', '-'] else 0.0
                    except: return 0.0
                
                # 根據傳入的索引 (P-R=15-17 或 S-U=18-20) 進行加總
                row_sum = sum([clean_val(row.iloc[c]) for c in col_indices])
                
                # 將數值存入字典 (若同一單位在同表出現多次則累加，例如科技執法有多筆時)
                unit_results[matched_unit] = unit_results.get(matched_unit, 0) + row_sum
                
        return {'data': unit_results, 'start': start_str, 'end': end_str, 'sheet': target_sheet}
    except Exception as e:
        st.error(f"解析失敗: {e}")
        return None

# --- 3. Streamlit 主介面 ---
st.markdown("## 🚔 交通違規統計 (列位動態掃描版)")

files = st.file_uploader("📂 請上傳 2 個檔案 (本期檔 + 累計檔)", accept_multiple_files=True)

if files and len(files) == 2:
    # 判斷檔案天數以區分本期/累計
    meta_list = []
    for f in files:
        res = parse_excel_dynamic(f, "重點違規統計表", [15, 16, 17])
        if res:
            try:
                s, e = res['start'], res['end']
                d_days = (date(int(e[:3])+1911, int(e[3:5]), int(e[5:])) - 
                          date(int(s[:3])+1911, int(s[3:5]), int(s[5:]))).days
                res['duration'] = d_days
            except: res['duration'] = 0
            res['file_obj'] = f
            meta_list.append(res)
            
    if len(meta_list) == 2:
        # 天數長 = 累計檔；天數短 = 本期檔
        meta_list.sort(key=lambda x: x['duration'], reverse=True)
        f_long, f_short = meta_list[0]['file_obj'], meta_list[1]['file_obj']
        
        # 1. 本期數據：短檔 -> 第一張表 -> P-R (15, 16, 17)
        d_week = parse_excel_dynamic(f_short, "重點違規統計表", [15, 16, 17])['data']
        # 2. 本年數據：長檔 -> 表名含(1) -> P-R (15, 16, 17)
        d_year = parse_excel_dynamic(f_long, "(1)", [15, 16, 17])['data']
        # 3. 去年數據：長檔 -> 表名含(1) -> S-U (18, 19, 20)
        d_last = parse_excel_dynamic(f_long, "(1)", [18, 19, 20])['data']
        
        # 4. 組合最終表格
        final_rows = []
        for u in UNIT_ORDER:
            w_val = d_week.get(u, 0)
            y_val = d_year.get(u, 0)
            l_val = d_last.get(u, 0)
            diff = y_val - l_val
            tgt = TARGETS.get(u, 0)
            rate = f"{(y_val/tgt):.1%}" if tgt > 0 else "0%"
            
            final_rows.append([u, int(w_val), int(y_val), int(l_val), int(diff), tgt, rate])
            
        df_final = pd.DataFrame(final_rows, columns=['單位', '本期', '本年累計', '去年同期', '增減比較', '目標值', '達成率'])
        st.success("✅ 解析成功！單位已跨行自動對齊。")
        st.dataframe(df_final, use_container_width=True)
