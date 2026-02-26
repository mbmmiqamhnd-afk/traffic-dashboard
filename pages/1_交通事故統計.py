# --- 主流程開始 ---
uploaded_files = st.file_uploader("請上傳 3 個報表檔案", accept_multiple_files=True)

if uploaded_files and len(uploaded_files) == 3:
    with st.spinner("⚡ 處理中..."):
        try:
            # 1. 解析檔案與日期 (維持原本邏輯)
            files_meta = []
            for f in uploaded_files:
                f.seek(0)
                df = parse_raw(f)
                # 偵測日期 (範例: 112/01/01)
                dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', str(df.iloc[:5, :3].values))
                if len(dates) >= 2:
                    d_str = f"{int(dates[0][1]):02d}{int(dates[1][1]):02d}-{int(dates[1][1]):02d}{int(dates[1][2]):02d}"
                    files_meta.append({
                        'df': clean_data(df), 
                        'year': int(dates[1][0]), 
                        'date_range': d_str, 
                        'is_cumu': (int(dates[0][1]) == 1)
                    })

            # 🛑 關鍵檢查：確保有 3 個成功解析的檔案
            if len(files_meta) < 3:
                st.error(f"❌ 解析失敗：僅偵測到 {len(files_meta)} 個有效日期區間。請確認報表內含民國年月日格式。")
                st.stop()

            # 2. 變數分配 (先初始化為 None)
            df_wk = df_cur = df_lst = None
            
            # 排序：年份大到小
            files_meta.sort(key=lambda x: x['year'], reverse=True)
            cur_year = files_meta[0]['year']
            
            # 分配邏輯
            try:
                df_wk = [f for f in files_meta if f['year'] == cur_year and not f['is_cumu']][0]
                df_cur = [f for f in files_meta if f['year'] == cur_year and f['is_cumu']][0]
                df_lst = [f for f in files_meta if f['year'] < cur_year][0]
            except IndexError:
                st.error("❌ 檔案分類失敗：需包含「今年本期」、「今年累計」與「去年同期累計」各一份。")
                st.stop()

            # 3. 呼叫函式生成結果 (傳入 df_wk['df'] 等)
            stations = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
            
            a1_res = build_a1_final(df_wk['df'], df_cur['df'], df_lst['df'], stations)
            a1_res.columns = ['統計期間', f'本期({df_wk["date_range"]})', f'本年累計({df_cur["date_range"]})', f'去年累計({df_lst["date_range"]})', '比較']
            
            a2_res = build_a2_final(df_wk['df'], df_cur['df'], df_lst['df'], stations)
            a2_res.columns = ['統計期間', f'本期({df_wk["date_range"]})', '前期', f'本年累計({df_cur["date_range"]})', f'去年累計({df_lst["date_range"]})', '比較', '增減比例']

            # --- 後續 Excel 產製與同步邏輯 ---
            # ... (略)

        except Exception as e:
            st.error(f"分析失敗，詳細錯誤：{e}")
