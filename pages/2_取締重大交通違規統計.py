import streamlit as st
import pandas as pd
import re
import io

# --- 1. 定義識別與目標 ---
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
def parse_excel_with_cols(uploaded_file, sheet_keyword, col_indices):
    """
    col_indices: [攔停Index, 逕行Index]
    """
    try:
        content = uploaded_file.getvalue()
        xl = pd.ExcelFile(io.BytesIO(content))
        # 尋找指定工作表，若找不到則取第一個
        target_sheet = next((s for s in xl.sheet_names if sheet_keyword in s), xl.sheet_names[0])
        df = pd.read_excel(xl, sheet_name=target_sheet, header=None)
        
        unit_data = {}
        for _, row in df.iterrows():
            u = get_standard_unit(row.iloc[0])
            if u and "合計" not in str(row.iloc[0]):
                def clean(v):
                    try:
                        s = str(v).replace(',', '').strip()
                        return int(float(s)) if s not in ['', 'nan', 'None', '-'] else 0
                    except: return 0
                
                # 科技執法的攔停數強制歸零
                stop_val = 0 if u == '科技執法' else clean(row.iloc[col_indices[0]])
                cit_val = clean(row.iloc[col_indices[1]])
                
                # 若同一單位出現多次則累加
                if u not in unit_data:
                    unit_data[u] = {'stop': stop_val, 'cit': cit_val}
                else:
                    unit_data[u]['stop'] += stop_val
                    unit_data[u]['cit'] += cit_val
        return unit_data
    except Exception as e:
        st.error(f"解析失敗: {e}")
        return None

# --- 3. Streamlit 介面 ---
st.markdown("## 🚔 交通違規統計 (攔停/逕行細分修正版)")

col1, col2 = st.columns(2)
with col1:
    file_period = st.file_uploader("📂 上傳「本期」檔案 (重點違規統計表)", type=['xlsx'])
with col2:
    file_year = st.file_uploader("📂 上傳「累計」檔案 (重點違規統計表 (1))", type=['xlsx'])

if file_period and file_year:
    # 1. 抓取本期數據 (來自短檔, 工作表不含(1), P&Q欄)
    data_week = parse_excel_with_cols(file_period, "重點違規統計表", [15, 16])
    
    # 2. 抓取本年數據 (來自長檔, 工作表含(1), P&Q欄)
    data_year = parse_excel_with_cols(file_year, "(1)", [15, 16])
    
    # 3. 抓取去年數據 (來自長檔, 工作表含(1), S&T欄)
    data_last = parse_excel_with_cols(file_year, "(1)", [18, 19])
    
    if data_week and data_year and data_last:
        final_rows = []
        # 初始化合計數值
        total_vals = {k: 0 for k in ['ws', 'wc', 'ys', 'yc', 'ls', 'lc', 'diff', 'tgt']}
        
        for u in UNIT_ORDER:
            w = data_week.get(u, {'stop':0, 'cit':0})
            y = data_year.get(u, {'stop':0, 'cit':0})
            l = data_last.get(u, {'stop':0, 'cit':0})
            
            y_sum = y['stop'] + y['cit']
            l_sum = l['stop'] + l['cit']
            tgt = TARGETS.get(u, 0)
            diff = y_sum - l_sum
            rate = f"{(y_sum/tgt):.1%}" if tgt > 0 else "0%"
            
            # 加入表格
            final_rows.append([
                u, w['stop'], w['cit'], y['stop'], y['cit'], l['stop'], l['cit'], 
                diff, tgt, rate
            ])
            
            # 累計合計列
            total_vals['ws'] += w['stop']; total_vals['wc'] += w['cit']
            total_vals['ys'] += y['stop']; total_vals['yc'] += y['cit']
            total_vals['ls'] += l['stop']; total_vals['lc'] += l['cit']
            total_vals['diff'] += diff; total_vals['tgt'] += tgt

        # 計算合計列的達成率
        total_rate = f"{( (total_vals['ys'] + total_vals['yc']) / total_vals['tgt']):.1%}" if total_vals['tgt'] > 0 else "0%"
        
        # 插入合計列到第一列
        total_row = [
            '合計', total_vals['ws'], total_vals['wc'], total_vals['ys'], total_vals['yc'], 
            total_vals['ls'], total_vals['lc'], total_vals['diff'], total_vals['tgt'], total_rate
        ]
        final_rows.insert(0, total_row)
            
        columns = [
            '單位', '本期攔停', '本期逕行', '本年攔停', '本年逕行', '去年攔停', '去年逕行', 
            '增減比較', '目標值', '達成率'
        ]
        
        st.success("✅ 數據統計完成，已依據攔停/逕行分類。")
        st.dataframe(pd.DataFrame(final_rows, columns=columns), use_container_width=True)
else:
    st.info("💡 請分別上傳「本期週報」與「年度累計」兩個 Excel 檔案以進行對比。")
