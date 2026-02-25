def parse_focus_report(uploaded_file, file_type="week"):
    """
    file_type 選項: 
    - "week": 對應本期，讀取「重點違規統計表」，P-R欄
    - "year": 對應本年，讀取「重點違規統計表 (1)」，P-R欄
    - "last": 對應去年，讀取「重點違規統計表 (1)」，S-U欄
    """
    if not uploaded_file: return None
    try:
        content = uploaded_file.getvalue()
        
        # 根據檔案類型決定工作表與欄位索引
        if file_type == "week":
            target_sheet = "重點違規統計表"
            col_range = slice(15, 18) # P, Q, R 欄 (Index 15, 16, 17)
        elif file_type == "year":
            target_sheet = "重點違規統計表 (1)"
            col_range = slice(15, 18) # P, Q, R 欄
        else: # last
            target_sheet = "重點違規統計表 (1)"
            col_range = slice(18, 21) # S, T, U 欄 (Index 18, 19, 20)

        # 讀取 Excel，跳過前幾列直到抓到單位 (通常 header 在第 3 或 4 列，這裡設為 3)
        df = pd.read_excel(io.BytesIO(content), sheet_name=target_sheet, header=3)
        
        # 取得日期區間 (從第一列的文字中擷取)
        df_info = pd.read_excel(io.BytesIO(content), sheet_name=target_sheet, header=None, nrows=5)
        info_str = "".join(df_info.astype(str).values.flatten())
        match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', info_str)
        start_date = match.group(1) if match else ""
        end_date = match.group(2) if match else ""

        unit_data = {}
        for _, row in df.iterrows():
            raw_unit = str(row.iloc[0]).strip()
            if raw_unit in ['nan', 'None', '', '合計', '單位'] or "統計" in raw_unit:
                continue
            
            # 單位匹配
            matched_name = None
            if "科技" in raw_unit: matched_name = "科技執法"
            else:
                for key, short_name in UNIT_MAP.items():
                    if key in raw_unit:
                        matched_name = short_name
                        break
            
            if matched_name:
                # 數值清洗與加總指定欄位 (P-R 或 S-U)
                def clean_sum(series):
                    return series.apply(lambda v: float(str(v).replace(',', '').strip()) if str(v).strip() not in ['', 'nan', 'None'] else 0.0).sum()
                
                # 計算指定範圍的總和
                total_val = clean_sum(row.iloc[col_range])
                
                # 由於你的需求沒有區分攔停/逕行，這裡統一放入數據
                # 若需要區分，請再告訴我 P, Q, R 各代表什麼
                if matched_name not in unit_data:
                    unit_data[matched_name] = {'total': total_val}

        # 計算天數
        dur = 0
        try:
            s_d, e_d = re.sub(r'[^\d]', '', start_date), re.sub(r'[^\d]', '', end_date)
            d1 = date(int(s_d[:3])+1911, int(s_d[3:5]), int(s_d[5:]))
            d2 = date(int(e_d[:3])+1911, int(e_d[3:5]), int(e_d[5:]))
            dur = (d2 - d1).days
        except: dur = 0

        return {'data': unit_data, 'start': start_date, 'end': end_date, 'duration': dur}
    except Exception as e:
        st.error(f"解析錯誤 ({uploaded_file.name}): {e}")
        return None

# --- 主程式邏輯部分需配合修改 ---
if uploaded_files and len(uploaded_files) >= 3:
    # 這裡需要一個邏輯來判定哪個檔案是「去年」、「本年」、「本期」
    # 建議可以依據檔案名稱或日期來決定傳入的 file_type
    # 以下為示範性排序與分配
    raw_parsed = []
    for f in uploaded_files:
        # 先初步讀取判斷天數
        res = parse_focus_report(f, "week") # 先用預設讀取
        if res:
            res['file_obj'] = f
            raw_parsed.append(res)
            
    if len(raw_parsed) >= 3:
        raw_parsed.sort(key=lambda x: x['start'])
        f_last_obj = raw_parsed[0]['file_obj'] # 日期最早的是去年
        
        others = sorted(raw_parsed[1:], key=lambda x: x['duration'], reverse=True)
        f_year_obj = others[0]['file_obj'] # 天數長的是本年累計
        f_week_obj = others[1]['file_obj'] # 天數短的是本期

        # 重新依照正確的欄位規則解析
        file_last = parse_focus_report(f_last_obj, "last")
        file_year = parse_focus_report(f_year_obj, "year")
        file_week = parse_focus_report(f_week_obj, "week")
        
        # ... 後續 final_rows 處理邏輯 ...
        # 注意：現在 data 裡面只有 'total'，需對應修改你的表格組合邏輯
