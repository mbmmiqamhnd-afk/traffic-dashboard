import streamlit as st
import pandas as pd
import re
import io
from datetime import date

# --- 1. 定義單位識別 ---
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
TARGETS = {'聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479, '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006}

# --- 2. 核心解析函數 ---
def parse_report_with_methods(uploaded_file, sheet_keyword, col_indices):
    """
    col_indices: 傳入 [攔停Index, 逕行Index]
    """
    try:
        content = uploaded_file.getvalue()
        xl = pd.ExcelFile(io.BytesIO(content))
        target_sheet = next((s for s in xl.sheet_names if sheet_keyword in s), xl.sheet_names[0])
        df = pd.read_excel(xl, sheet_name=target_sheet, header=None)
        
        # 提取日期
        info_text = "".join(df.iloc[:5].astype(str).values.flatten())
        match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', info_text)
        start_date = match.group(1) if match else "0000000"
        end_date = match.group(2) if match else "0000000"
        
        unit_results = {}
        for _, row in df.iterrows():
            u = get_standard_unit(row.iloc[0])
            if u and "合計" not in str(row.iloc[0]):
                def clean(v):
                    try:
                        s = str(v).replace(',', '').strip()
                        return int(float(s)) if s not in ['', 'nan', 'None', '-'] else 0
                    except: return 0
                
                # 分別存儲攔停與逕行
                unit_results[u] = {
                    'stop': clean(row.iloc[col_indices[0]]),
                    'cit': clean(row.iloc[col_indices[1]])
                }
        return {'data': unit_results, 'start': start_date, 'end': end_date}
    except Exception as e:
        st.error(f"解析失敗: {e}")
        return None

# --- 3. 主程式介面 ---
st.markdown("## 🚔 交通違規統計 (攔停/逕行細分版)")

files = st.file_uploader("📂 上傳 2 個檔案 (本期檔 + 累計檔)", accept_multiple_files=True)

if files and len(files) == 2:
    meta = []
    for f in files:
        res = parse_report_with_methods(f, "重點違規統計表", [15, 16])
        if res:
            try:
                s, e = res['start'], res['end']
                d = (date(int(e[:3])+1911, int(e[3:5]), int(e[5:])) - date(int(s[:3])+1911, int(s[3:5]), int(s[5:]))).days
                res['duration'] = d
            except: res['duration'] = 0
            res['file_obj'] = f
            meta.append(res)
            
    if len(meta) == 2:
        meta.sort(key=lambda x: x['duration'], reverse=True)
        f_long, f_short = meta[0]['file_obj'], meta[1]['file_obj']
        
        # 抓取數據：本期 (短檔 P, Q), 本年 (長檔 (1) P, Q), 去年 (長檔 (1) S, T)
        d_week = parse_report_with_methods(f_short, "重點違規統計表", [15, 16])['data']
        d_year = parse_report_with_methods(f_long, "(1)", [15, 16])['data']
        d_last = parse_report_with_methods(f_long, "(1)", [18, 19])['data']
        
        rows = []
        for u in UNIT_ORDER:
            w = d_week.get(u, {'stop':0, 'cit':0})
            y = d_year.get(u, {'stop':0, 'cit':0})
            l = d_last.get(u, {'stop':0, 'cit':0})
            
            y_total = y['stop'] + y['cit']
            l_total = l['stop'] + l['cit']
            tgt = TARGETS.get(u, 0)
            
            rows.append([
                u, 
                w['stop'], w['cit'],    # 本期
                y['stop'], y['cit'],    # 本年
                l['stop'], l['cit'],    # 去年
                y_total - l_total,      # 比較
                tgt, 
                f"{(y_total/tgt):.1%}" if tgt > 0 else "0%"
            ])
            
        columns = [
            '單位', 
            '本期攔停', '本期逕行', 
            '本年攔停', '本年逕行', 
            '去年攔停', '去年逕行', 
            '增減比較', '目標值', '達成率'
        ]
        df_final = pd.DataFrame(rows, columns=columns)
        st.success("✅ 解析成功！已按攔停/逕行分類統計。")
        st.dataframe(df_final, use_container_width=True)
