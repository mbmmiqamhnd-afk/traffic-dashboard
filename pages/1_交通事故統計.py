def process_files(uploaded_files):
    # --- 2. 智慧辨識檔案身分 (終極版：依起始日期判斷) ---
    file_data_list = []
    
    for file_obj in uploaded_files:
        file_obj.seek(0)
        df = parse_police_stats_raw(file_obj)
        
        try:
            # 抓取日期字串
            date_str = df.iloc[1, 0].replace("統計日期：", "").strip()
            dates = re.findall(r'(\d{3})/(\d{2})/(\d{2})', date_str)
            
            if not dates:
                st.warning(f"無法識別日期：{file_obj.name}")
                continue
            
            start_y, start_m, start_d = map(int, dates[0])
            end_y, end_m, end_d = map(int, dates[1])
            
            # 計算天數 (輔助判斷用)
            dt_start = datetime(start_y + 1911, start_m, start_d)
            dt_end = datetime(end_y + 1911, end_m, end_d)
            delta_days = (dt_end - dt_start).days
            
            file_data_list.append({
                'df': df,
                'date_str': date_str,
                'delta_days': delta_days,
                'start_date': (start_y, start_m, start_d),
                'filename': file_obj.name
            })
        except Exception as e:
            st.error(f"檔案解析失敗 {file_obj.name}: {e}")
            return

    if len(file_data_list) != 3:
        st.error(f"解析失敗：只成功讀取了 {len(file_data_list)} 個有效檔案，請確認檔案數量。")
        return

    # --- 核心判斷邏輯 (修正版) ---
    # 先將所有檔案分類
    df_wk, df_cur, df_lst = None, None, None
    d_wk, d_cur, d_lst = "", "", ""
    
    # 排序方便處理：依起始年份由小到大
    file_data_list.sort(key=lambda x: x['start_date'])
    
    # 1. 找出「去年累計」：起始月日為 01/01 且 年份最小
    # (通常是排序後的第一個，但為了保險我們檢查 01/01)
    last_candidates = [f for f in file_data_list if f['start_date'][1] == 1 and f['start_date'][2] == 1]
    
    if last_candidates:
        # 年份最小的 01/01 是去年累計
        last_candidates.sort(key=lambda x: x['start_date'][0])
        lst_data = last_candidates[0]
        
        # 從清單中移除已找到的
        file_data_list.remove(lst_data)
        
        # 2. 找出「今年累計」：剩下的檔案中，起始為 01/01 的 (年份較大)
        cur_candidates = [f for f in file_data_list if f['start_date'][1] == 1 and f['start_date'][2] == 1]
        
        if cur_candidates:
            cur_data = cur_candidates[0] # 應該只剩一個
            file_data_list.remove(cur_data)
            
            # 3. 剩下的就是「週報表」
            if file_data_list:
                wk_data = file_data_list[0]
            else:
                st.error("邏輯錯誤：找不到週報表")
                return
        else:
            # 如果剩下的沒有 01/01 開頭，代表今年累計可能還沒開始?? (不合理)
            # 或者週報表也是 01/01 開頭 (如年初第一週)
            # 這時候依天數判斷：天數長的是累計，短的是週
            if len(file_data_list) == 2:
                file_data_list.sort(key=lambda x: x['delta_days'], reverse=True)
                cur_data = file_data_list[0] # 天數長 -> 今年累計
                wk_data = file_data_list[1]  # 天數短 -> 週報表
            else:
                st.error("無法識別今年累計與週報表")
                return
    else:
        st.error("無法識別去年累計檔案 (找不到 01/01 開頭的檔案)")
        return

    # 分配資料
    df_wk, d_wk = wk_data['df'], wk_data['date_str']
    df_cur, d_cur = cur_data['df'], cur_data['date_str']
    df_lst, d_lst = lst_data['df'], lst_data['date_str']

    st.success(f"✅ 成功辨識：\n- **本期**: {d_wk}\n- **今年**: {d_cur}\n- **去年**: {d_lst}")

    # --- 3. 資料清理與計算 ---
    df_wk_clean = process_data(df_wk)
    df_cur_clean = process_data(df_cur)
    df_lst_clean = process_data(df_lst)

    # 準備標題日期
    h_wk = format_date(d_wk)
    h_cur = format_date(d_cur)
    h_lst = format_date(d_lst)

    # --- 合併 A1 ---
    a1_wk = df_wk_clean[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'wk'})
    a1_cur = df_cur_clean[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'cur'})
    a1_lst = df_lst_clean[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'last'})
    
    m_a1 = pd.merge(a1_wk, a1_cur, on='Station_Short', how='outer')
    m_a1 = pd.merge(m_a1, a1_lst, on='Station_Short', how='outer').fillna(0)
    m_a1['Diff'] = m_a1['cur'] - m_a1['last']

    # --- 合併 A2 ---
    a2_wk = df_wk_clean[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'wk'})
    a2_cur = df_cur_clean[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'cur'})
    a2_lst = df_lst_clean[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'last'})
    
    m_a2 = pd.merge(a2_wk, a2_cur, on='Station_Short', how='outer')
    m_a2 = pd.merge(m_a2, a2_lst, on='Station_Short', how='outer').fillna(0)
    m_a2['Diff'] = m_a2['cur'] - m_a2['last']
    m_a2['Pct'] = m_a2.apply(lambda x: (x['Diff']/x['last']) if x['last']!=0 else 0, axis=1)
    m_a2['Pct_Str'] = m_a2['Pct'].apply(lambda x: f"{x:.2%}")
    m_a2['Prev'] = "-"

    # 排序
    m_a1 = sort_stations(m_a1)
    m_a2 = sort_stations(m_a2)

    # 整理最終表格
    a1_final = m_a1[['Station_Short', 'wk', 'cur', 'last', 'Diff']].copy()
    a1_final.columns = ['單位', f'本期({h_wk})', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較']
    
    a2_display = m_a2[['Station_Short', 'wk', 'Prev', 'cur', 'last', 'Diff', 'Pct_Str']].copy()
    a2_display.columns = ['單位', f'本期({h_wk})', '前期', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較', '本年較去年增減比例']

    a2_download = m_a2[['Station_Short', 'wk', 'Prev', 'cur', 'last', 'Diff', 'Pct']].copy()
    a2_download.columns = ['單位', f'本期({h_wk})', '前期', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較', '本年較去年增減比例']

    # --- 4. 顯示結果與下載按鈕 ---
    st.markdown("### 📊 統計結果")
    
    st.subheader("1. A1 類交通事故死亡人數")
    st.dataframe(a1_final, use_container_width=True)
    
    st.subheader("2. A2 類交通事故受傷人數")
    st.dataframe(a2_display, use_container_width=True)

    # 產出 Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        a1_final.to_excel(writer, sheet_name='A1死亡人數', index=False)
        a2_download.to_excel(writer, sheet_name='A2受傷人數', index=False)
        
        workbook  = writer.book
        worksheet = writer.sheets['A2受傷人數']
        percent_fmt = workbook.add_format({'num_format': '0.00%'})
        worksheet.set_column(6, 6, None, percent_fmt)
        
    output.seek(0)
    filename = f'交通事故統計表_{datetime.now().strftime("%Y%m%d")}.xlsx'
    
    st.download_button(
        label="📥 下載 Excel 報表",
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
