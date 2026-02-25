def parse_focus_report(uploaded_file):
    if not uploaded_file: return None
    file_name = uploaded_file.name
    try:
        content = uploaded_file.getvalue()
        start_date, end_date = "", ""
        df = None
        header_idx = -1
        
        # 1. 先讀取前 25 列來找日期和標題列位置
        df_raw = pd.read_excel(io.BytesIO(content), header=None, nrows=25)
        
        # 定義偵測標題的關鍵字
        target_keywords = ["酒後", "闖紅燈", "嚴重超速"] 

        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
            
            # 偵測日期 (格式: 入案日期：1120101 至 1120131)
            if not start_date:
                match = re.search(r'入案日期[：:]?\s*(\d{3,7}).*至\s*(\d{3,7})', row_str)
                if match: 
                    start_date, end_date = match.group(1), match.group(2)
            
            # 偵測標題列：如果這一列包含多個違規關鍵字，就認定它是標題列
            if any(k in row_str for k in target_keywords):
                header_idx = i
                if start_date: break

        if header_idx == -1:
            st.warning(f"⚠️ 檔案 {file_name} 解析失敗：找不到包含違規項目的標題列。")
            return None

        # 2. 以找到的 header_idx 正式讀取資料
        df = pd.read_excel(io.BytesIO(content), header=header_idx)
        
        # 定義我們要抓取的欄位關鍵字
        keywords = ["酒後", "闖紅燈", "嚴重超速", "逆向", "轉彎", "蛇行", "不暫停讓行人", "機車"]
        stop_cols = []
        cit_cols = []
        
        # 尋找關鍵字所在的欄位索引
        for i in range(len(df.columns)):
            col_str = str(df.columns[i])
            if any(k in col_str for k in keywords) and "路肩" not in col_str and "大型車" not in col_str:
                # 假設：當場攔停在該關鍵字欄位，逕行舉發在下一欄
                stop_cols.append(i)
                cit_cols.append(i + 1)
        
        unit_data = {}
        for _, row in df.iterrows():
            # ★ 關鍵修改：因為 A 欄標題空白，我們直接用索引 0 抓取第一欄的值
            raw_unit = str(row.iloc[0]).strip()
            
            # 跳過空白、合計列或非單位的文字
            if raw_unit == 'nan' or not raw_unit or "合計" in raw_unit or "單位" in raw_unit:
                continue
            
            unit_name = UNIT_MAP.get(raw_unit, raw_unit)
            s, c = 0, 0
            
            # 加總攔停數
            for col in stop_cols:
                try:
                    val = row.iloc[col]
                    if pd.isna(val) or str(val).strip() == "": val = 0
                    s += float(str(val).replace(',', ''))
                except: pass
            
            # 加總逕行數
            for col in cit_cols:
                try:
                    val = row.iloc[col]
                    if pd.isna(val) or str(val).strip() == "": val = 0
                    c += float(str(val).replace(',', ''))
                except: pass

            unit_data[unit_name] = {'stop': s, 'cit': c}

        # 計算統計天數
        duration = 0
        try:
            if start_date and end_date:
                s_d = re.sub(r'[^\d]', '', start_date)
                e_d = re.sub(r'[^\d]', '', end_date)
                d1 = date(int(s_d[:3])+1911, int(s_d[3:5]), int(s_d[5:]))
                d2 = date(int(e_d[:3])+1911, int(e_d[3:5]), int(e_d[5:]))
                duration = (d2 - d1).days
        except: duration = 0

        return {
            'data': unit_data, 
            'start': start_date or "0000000", 
            'end': end_date or "0000000", 
            'duration': duration, 
            'filename': file_name
        }
    except Exception as e:
        st.warning(f"⚠️ 檔案 {file_name} 錯誤: {e}")
        return None
