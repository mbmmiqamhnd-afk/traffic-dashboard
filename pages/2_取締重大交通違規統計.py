# ==========================================
# 2. 核心解析函數 (v78 數值校正版)
# ==========================================
def parse_focus_report(uploaded_file):
    if not uploaded_file: return None
    try:
        content = uploaded_file.getvalue()
        df_raw = pd.read_excel(io.BytesIO(content), header=None, nrows=40)
        
        start_date, end_date, header_idx = "", "", -1
        keywords = ["酒後", "闖紅燈", "嚴重超速", "逆向", "轉彎", "蛇行", "不暫停讓行人", "機車"]
        
        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
            if not start_date:
                match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', row_str)
                if match: start_date, end_date = match.group(1), match.group(2)
            hits = sum(1 for k in keywords if k in row_str)
            if hits >= 4: # 增加標題判定門檻，避免誤判
                header_idx = i
                break 
        
        if header_idx == -1: return None
        df = pd.read_excel(io.BytesIO(content), header=header_idx)
        
        # 排除 P-U 欄 (Index 15-20)
        stop_cols, cit_cols = [], []
        for i in range(len(df.columns)):
            if 15 <= i <= 20: continue
            col_name = str(df.columns[i])
            if any(k in col_name for k in keywords) and "路肩" not in col_name:
                if i+1 < len(df.columns) and not (15 <= (i+1) <= 20):
                    stop_cols.append(i)
                    cit_cols.append(i + 1)
        
        unit_data = {}
        unit_debug = [] # 追蹤數據來源

        for idx, row in df.iterrows():
            raw_val = str(row.iloc[0]).strip()
            if raw_val in ['nan', 'None', '', '合計', '單位'] or "統計" in raw_val: continue
            
            # 精確匹配邏輯：避免「交通分隊」抓到「科技執法-交通分隊」
            matched_name = None
            # 先檢查是否為科技執法
            if "科技" in raw_val or "交通組" in raw_val:
                matched_name = "科技執法"
            else:
                for key, short_name in UNIT_MAP.items():
                    if key != "科技執法" and key in raw_val:
                        matched_name = short_name
                        break
            
            if matched_name:
                def clean_val(v):
                    try:
                        v_str = str(v).replace(',', '').strip()
                        return float(v_str) if v_str not in ['', 'nan', 'None'] else 0.0
                    except: return 0.0
                
                s_sum = sum([clean_val(row.iloc[c]) for c in stop_cols if c < len(row)])
                c_sum = sum([clean_val(row.iloc[c]) for c in cit_cols if c < len(row)])
                
                if s_sum > 0 or c_sum > 0:
                    if matched_name not in unit_data:
                        unit_data[matched_name] = {'stop': s_sum, 'cit': c_sum}
                    else:
                        # 如果該單位已存在，僅在數值不同時累加，或記錄下來
                        unit_data[matched_name]['stop'] += s_sum
                        unit_data[matched_name]['cit'] += c_sum
                    
                    unit_debug.append(f"📍 {matched_name} 於第 {idx+header_idx+2} 行抓到數據: 攔停={s_sum}, 逕行={c_sum} (原始文字: {raw_val})")

        dur = 0
        try:
            s_d, e_d = re.sub(r'[^\d]', '', start_date), re.sub(r'[^\d]', '', end_date)
            d1 = date(int(s_d[:3])+1911, int(s_d[3:5]), int(s_d[5:]))
            d2 = date(int(e_d[:3])+1911, int(e_d[3:5]), int(e_d[5:]))
            dur = (d2 - d1).days
        except: dur = 0
            
        return {'data': unit_data, 'start': start_date, 'end': end_date, 'duration': dur, 'debug': unit_debug, 'filename': uploaded_file.name}
    except Exception as e:
        st.error(f"解析失敗: {e}"); return None

# 主程式中的表格生成部分請維持 v77 結構...
