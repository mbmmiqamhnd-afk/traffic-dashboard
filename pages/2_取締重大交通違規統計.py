import streamlit as st
import pandas as pd
import re
import io
from datetime import date

# --- 1. 標準單位與識別邏輯 ---
# 依照分隊優先、其後所別的邏輯，避免「龍潭」與「交通分隊」混淆
def get_standard_unit(raw_name):
    name = str(raw_name).strip()
    if '分隊' in name: return '交通分隊'
    if '科技' in name or '交通組' in name: return '科技執法'
    if '警備' in name: return '警備隊'
    if '聖亭' in name: return '聖亭所'
    if '龍潭' in name: return '龍潭所'
    if '中興' in name: return '中興所'
    if '石門' in name: return '石門所'
    if '高平' in name: return '高平所'
    if '三和' in name: return '三和所'
    return None

UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
TARGETS = {
    '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479,
    '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006
}

# --- 2. 核心解析函數 ---
def parse_excel_precision(uploaded_file, sheet_keyword, col_indices):
    """
    col_indices 只應包含 [攔停欄位, 逕行欄位]，不包含總計欄位，避免重複計算。
    本年/本期 (P-R): 使用 [15, 16] (即 P, Q)
    去年 (S-U): 使用 [18, 19] (即 S, T)
    """
    try:
        content = uploaded_file.getvalue()
        xl = pd.ExcelFile(io.BytesIO(content))
        
        # 尋找工作表
        target_sheet = xl.sheet_names[0]
        for s in xl.sheet_names:
            if sheet_keyword in s:
                target_sheet = s
                break
        
        df = pd.read_excel(xl, sheet_name=target_sheet, header=None)
        
        # 提取日期資訊
        info_text = "".join(df.iloc[:8].astype(str).values.flatten())
        match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', info_text)
        start_date = match.group(1) if match else "0000000"
        end_date = match.group(2) if match else "0000000"
        
        unit_data = {}
        for _, row in df.iterrows():
            unit_name = get_standard_unit(row.iloc[0])
            if unit_name and "合計" not in str(row.iloc[0]):
                def clean(v):
                    try:
                        s = str(v).replace(',', '').strip()
                        return float(s) if s not in ['', 'nan', 'None', '-'] else 0.0
                    except: return 0.0
                
                # 只加總攔停與逕行 (例如 P+Q 或 S+T)
                val = sum([clean(row.iloc[c]) for c in col_indices])
                unit_data[unit_name] = unit_data.get(unit_name, 0) + val
                
        return {'data': unit_data, 'start': start_date, 'end': end_date}
    except Exception as e:
        st.error(f"解析 {uploaded_file.name} 失敗: {e}")
        return None

# --- 3. 主程式介面 ---
st.markdown("## 🚔 交通違規統計 (數值精準修正版)")

files = st.file_uploader("📂 請上傳 2 個檔案 (重點違規統計表.xlsx & 重點違規統計表 (1).xlsx)", accept_multiple_files=True)

if files and len(files) == 2:
    # A. 識別檔案類型 (天數判斷)
    parsed_files = []
    for f in files:
        res = parse_excel_precision(f, "重點違規統計表", [15, 16])
        if res:
            try:
                s, e = res['start'], res['end']
                d_days = (date(int(e[:3])+1911, int(e[3:5]), int(e[5:])) - 
                          date(int(s[:3])+1911, int(s[3:5]), int(s[5:]))).days
                res['duration'] = d_days
            except: res['duration'] = 0
            res['file_obj'] = f
            parsed_files.append(res)
            
    if len(parsed_files) == 2:
        # 長天數為累計檔
        parsed_files.sort(key=lambda x: x['duration'], reverse=True)
        f_long = parsed_files[0]['file_obj']
        f_short = parsed_files[1]['file_obj']
        
        # B. 根據您的精確要求抓取欄位
        # 1. 本期：短檔 -> 重點違規統計表 -> P+Q (Index 15, 16)
        data_week = parse_excel_precision(f_short, "重點違規統計表", [15, 16])['data']
        # 2. 本年：長檔 -> 重點違規統計表 (1) -> P+Q (Index 15, 16)
        data_year = parse_excel_precision(f_long, "(1)", [15, 16])['data']
        # 3. 去年：長檔 -> 重點違規統計表 (1) -> S+T (Index 18, 19)
        data_last = parse_excel_precision(f_long, "(1)", [18, 19])['data']
        
        # C. 組合與計算
        final_rows = []
        for u in UNIT_ORDER:
            w, y, l = data_week.get(u,0), data_year.get(u,0), data_last.get(u,0)
            tgt = TARGETS.get(u, 0)
            diff = y - l
            rate = f"{(y/tgt):.1%}" if tgt > 0 else "0%"
            
            final_rows.append([u, int(w), int(y), int(l), int(diff), tgt, rate])
            
        df_final = pd.DataFrame(final_rows, columns=['單位', '本期', '本年累計', '去年同期', '增減比較', '目標值', '達成率'])
        st.success("✅ 數據解析完成！已修正重複計數問題。")
        st.dataframe(df_final, use_container_width=True)

        # 顯示日期區間資訊確認
        st.info(f"📊 本期區間：{parsed_files[1]['start']} ~ {parsed_files[1]['end']} \n\n"
                f"📅 累計區間：{parsed_files[0]['start']} ~ {parsed_files[0]['end']}")
