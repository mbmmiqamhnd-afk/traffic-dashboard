try:
            # --- 1. 定義讀取與清理函數 (保持不變) ---
            def parse_raw(file_obj):
                try: return pd.read_csv(file_obj, header=None)
                except: file_obj.seek(0); return pd.read_excel(file_obj, header=None)

            def clean_data(df_raw):
                # 先把第一欄轉字串，避免讀成數字造成錯誤
                df_raw[0] = df_raw[0].astype(str)
                df_data = df_raw[df_raw[0].notna()].copy()
                df_data = df_data[df_data[0].str.contains("總計|派出所|合計")].copy()
                df_data = df_data.reset_index(drop=True)
                columns_map = {
                    0: "Station", 1: "Total_Cases", 2: "Total_Deaths", 3: "Total_Injuries",
                    4: "A1_Cases", 5: "A1_Deaths", 6: "A1_Injuries",
                    7: "A2_Cases", 8: "A2_Deaths", 9: "A2_Injuries", 10: "A3_Cases"
                }
                for i in range(11):
                    if i not in df_data.columns: df_data[i] = 0
                df_data = df_data.rename(columns=columns_map)
                df_data = df_data[list(columns_map.values())]
                for col in list(columns_map.values())[1:]:
                    df_data[col] = pd.to_numeric(df_data[col].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
                df_data['Station_Short'] = df_data['Station'].astype(str).str.replace('派出所', '所').str.replace('總計', '合計')
                return df_data

            # --- 2. 智慧辨識檔案日期 (增強版) ---
            file_data_map = {}
            debug_info = []  # 儲存偵測資訊，若失敗時顯示給使用者看

            for uploaded_file in uploaded_files:
                uploaded_file.seek(0)
                df = parse_raw(uploaded_file)
                
                found_dates = []
                date_str_found = "未找到日期"
                
                # 策略：掃描前 5 列、前 3 欄，尋找日期格式
                for r in range(min(5, len(df))):
                    for c in range(min(3, len(df.columns))):
                        val = str(df.iloc[r, c])
                        # 尋找民國年格式 (e.g., 113/01/01 或 113.1.1)
                        dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', val)
                        if len(dates) >= 2: # 至少要找到 起、迄 兩個日期
                            found_dates = dates
                            date_str_found = val
                            break
                    if found_dates: break

                if found_dates:
                    start_y, start_m, start_d = map(int, found_dates[0])
                    end_y, end_m, end_d = map(int, found_dates[1])
                    
                    # 判斷邏輯
                    month_diff = (end_y - start_y) * 12 + (end_m - start_m)
                    days_diff = end_d - start_d # 簡易判斷
                    
                    # 如果跨度小於 1 個月且天數少於 20 天 -> 視為本期週報
                    if month_diff == 0 and days_diff < 20:
                        category = 'weekly'
                    else:
                        category = 'cumulative'
                        
                    file_data_map[uploaded_file.name] = {
                        'df': df, 
                        'category': category, 
                        'year': start_y, 
                        'raw_date': f"{start_y}/{start_m:02d}/{start_d:02d}-{end_y}/{end_m:02d}/{end_d:02d}"
                    }
                    debug_info.append(f"✅ {uploaded_file.name}: 判斷為 [{category}], 日期: {found_dates[0]}~{found_dates[1]}")
                else:
                    debug_info.append(f"❌ {uploaded_file.name}: 無法識別日期 (程式看到的文字: {str(df.iloc[0:2, 0].values)})")

            # --- 3. 分配 DataFrame ---
            df_wk = None; df_cur = None; df_lst = None
            h_wk = ""; h_cur = ""; h_lst = ""

            for fname, data in file_data_map.items():
                if data['category'] == 'weekly':
                    df_wk = clean_data(data['df']); h_wk = data['raw_date']

            cumu_files = [d for d in file_data_map.values() if d['category'] == 'cumulative']
            if len(cumu_files) >= 2:
                cumu_files.sort(key=lambda x: x['year'], reverse=True) # 年份大的是今年
                df_cur = clean_data(cumu_files[0]['df']); h_cur = cumu_files[0]['raw_date']
                df_lst = clean_data(cumu_files[1]['df']); h_lst = cumu_files[1]['raw_date']

            # --- 4. 錯誤檢核與顯示 ---
            if df_wk is None or df_cur is None or df_lst is None:
                st.error("❌ 無法識別完整的 3 份檔案。")
                with st.expander("🕵️‍♂️ 點擊查看偵測細節 (除錯用)"):
                    for info in debug_info:
                        st.write(info)
                    st.write("---")
                    st.write("請確認：")
                    st.write("1. 報表內是否有類似 `113/01/01` 的日期格式？")
                    st.write("2. 是否上傳了兩份年度累計(不同年) + 一份週報表？")
                st.stop()

            # --- (以下接續原本的計算邏輯: A1, A2 合併計算...) ---
            # ... 請將原本程式碼的 # A1 ... 開始的部分接在這邊 ...
