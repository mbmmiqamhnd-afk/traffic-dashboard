# --- 修正後的核心計算邏輯 ---

# 1. 定義 A1 計算函式 (傳入 3 個資料集與派出所清單)
def build_a1_final(wk, cur, lst, stations):
    col_name = 'A1_Deaths'
    # 合併：本期 + 本年累計 + 去年累計
    m = pd.merge(wk[['Station_Short', col_name]], 
                 cur[['Station_Short', col_name]], on='Station_Short', suffixes=('_wk', '_cur'))
    m = pd.merge(m, lst[['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_lst'})
    
    # 篩選派出所
    m = m[m['Station_Short'].isin(stations)].copy()
    
    # 計算合計列
    total_row = m.select_dtypes(include='number').sum().to_dict()
    total_row['Station_Short'] = '合計'
    m = pd.concat([pd.DataFrame([total_row]), m], ignore_index=True)
    
    # 計算比較 (本年累計 - 去年累計)
    m['Diff'] = m[col_name+'_cur'] - m[col_name+'_lst']
    
    # 整理欄位順序 (確保共 5 欄)
    m = m[['Station_Short', col_name+'_wk', col_name+'_cur', col_name+'_lst', 'Diff']]
    return m

# 2. 定義 A2 計算函式 (傳入 3 個資料集與派出所清單)
def build_a2_final(wk, cur, lst, stations):
    col_name = 'A2_Injuries'
    # 合併
    m = pd.merge(wk[['Station_Short', col_name]], 
                 cur[['Station_Short', col_name]], on='Station_Short', suffixes=('_wk', '_cur'))
    m = pd.merge(m, lst[['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_lst'})
    
    # 篩選派出所
    m = m[m['Station_Short'].isin(stations)].copy()
    
    # 計算合計列
    total_row = m.select_dtypes(include='number').sum().to_dict()
    total_row['Station_Short'] = '合計'
    m = pd.concat([pd.DataFrame([total_row]), m], ignore_index=True)
    
    # 計算比較 (Diff) 與 百分比 (Pct)
    m['Diff'] = m[col_name+'_cur'] - m[col_name+'_lst']
    m['Pct'] = m.apply(lambda x: f"{(x['Diff']/x[col_name+'_lst']):.2%}" if x[col_name+'_lst'] != 0 else "0.00%", axis=1)
    
    # 插入「前期」占位符，確保數值不偏移 (現在欄位順序：0:單位, 1:本期, 2:前期, 3:本年, 4:去年, 5:比較, 6:比例)
    m.insert(2, 'Prev', '-') 
    
    # 重新整理順序 (確保共 7 欄)
    m = m[['Station_Short', col_name+'_wk', 'Prev', col_name+'_cur', col_name+'_lst', 'Diff', 'Pct']]
    return m

# --- 在主程式邏輯中使用 ---
# 當檔案解析完成後，呼叫方式如下：
# (假設 df_wk, df_cur, df_lst 已經解析為含有 'df' 鍵值的字典)

stations = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']

a1_res = build_a1_final(df_wk['df'], df_cur['df'], df_lst['df'], stations)
a1_res.columns = ['統計期間', f'本期({df_wk["date_range"]})', f'本年累計({df_cur["date_range"]})', f'去年累計({df_lst["date_range"]})', '比較']

a2_res = build_a2_final(df_wk['df'], df_cur['df'], df_lst['df'], stations)
a2_res.columns = ['統計期間', f'本期({df_wk["date_range"]})', '前期', f'本年累計({df_cur["date_range"]})', f'去年累計({df_lst["date_range"]})', '比較', '增減比例']
