# ==========================================
# 核心計算邏輯修正版 (處理 A1 & A2 數值對齊)
# ==========================================

# B. 計算合併 (A1 部分：共 5 欄)
def build_a1_final():
    col_name = 'A1_Deaths'
    # 合併三份資料
    m = pd.merge(df_wk['df'][['Station_Short', col_name]], 
                 df_cur['df'][['Station_Short', col_name]], on='Station_Short', suffixes=('_wk', '_cur'))
    m = pd.merge(m, df_lst['df'][['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_lst'})
    
    # 只取特定派出所
    m = m[m['Station_Short'].isin(stations)].copy()
    
    # 計算合計列
    total_row = m.select_dtypes(include='number').sum().to_dict()
    total_row['Station_Short'] = '合計'
    m = pd.concat([pd.DataFrame([total_row]), m], ignore_index=True)
    
    # 計算比較 (本年累計 - 去年累計)
    m['Diff'] = m[col_name+'_cur'] - m[col_name+'_lst']
    
    # 最終欄位排序
    m = m[['Station_Short', col_name+'_wk', col_name+'_cur', col_name+'_lst', 'Diff']]
    m.columns = ['統計期間', f'本期({df_wk["date_range"]})', f'本年累計({df_cur["date_range"]})', f'去年累計({df_lst["date_range"]})', '比較']
    return m

# B. 計算合併 (A2 部分：共 7 欄)
def build_a2_final():
    col_name = 'A2_Injuries'
    # 合併三份資料
    m = pd.merge(df_wk['df'][['Station_Short', col_name]], 
                 df_cur['df'][['Station_Short', col_name]], on='Station_Short', suffixes=('_wk', '_cur'))
    m = pd.merge(m, df_lst['df'][['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_lst'})
    
    # 只取特定派出所
    m = m[m['Station_Short'].isin(stations)].copy()
    
    # 計算合計列
    total_row = m.select_dtypes(include='number').sum().to_dict()
    total_row['Station_Short'] = '合計'
    m = pd.concat([pd.DataFrame([total_row]), m], ignore_index=True)
    
    # 計算比較 (本年累計 - 去年累計)
    m['Diff'] = m[col_name+'_cur'] - m[col_name+'_lst']
    
    # 計算百分比
    m['Pct'] = m.apply(lambda x: f"{(x['Diff']/x[col_name+'_lst']):.2%}" if x[col_name+'_lst'] != 0 else "0.00%", axis=1)
    
    # 插入「前期」占位符 (這就是造成錯位的主因，現在我們精確指定位置)
    # 順序：統計期間(0), 本期(1), 前期(2), 本年累計(3), 去年累計(4), 比較(5), 增減比例(6)
    m.insert(2, 'Prev', '-') 
    
    # 重新整理順序並命名
    m = m[['Station_Short', col_name+'_wk', 'Prev', col_name+'_cur', col_name+'_lst', 'Diff', 'Pct']]
    m.columns = ['統計期間', f'本期({df_wk["date_range"]})', '前期', f'本年累計({df_cur["date_range"]})', f'去年累計({df_lst["date_range"]})', '比較', '增減比例']
    return m

# 呼叫函數生成結果
a1_res = build_a1_final()
a2_res = build_a2_final()
