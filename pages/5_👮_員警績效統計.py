import streamlit as st
import pandas as pd
import io

# 設定頁面資訊
st.set_page_config(page_title="員警績效統計系統", layout="wide", page_icon="👮")

st.title("👮 員警交通執法績效統計系統")
st.markdown("""
### 📝 使用說明
1. **上傳配分表**：請上傳「常用交通執法重點工作配分表」(.xlsx 或 .csv)。
2. **上傳績效檔**：請一次選取所有「PoliceResult...」系列的檔案 (支援 .csv 與 .xlsx)。
3. **系統運算**：系統自動清洗資料、移除千分位符號、計算分數並彙整排名。
""")

# ==========================================
# 1. 檔案上傳區
# ==========================================
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 上傳配分表")
    uploaded_score_file = st.file_uploader("請上傳配分表", type=['xlsx', 'xls', 'csv'], key="score_uploader")

with col2:
    st.subheader("2. 上傳績效報表")
    uploaded_result_files = st.file_uploader("請上傳 PoliceResult 檔案 (可多選)", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True, key="result_uploader")

# ==========================================
# 2. 核心邏輯函數
# ==========================================

def load_data(file_obj, header=0):
    """
    通用讀取函數：自動處理 csv 編碼 (utf-8/cp950) 與 excel
    """
    file_obj.seek(0)
    filename = file_obj.name.lower()
    
    try:
        if filename.endswith('.csv'):
            # 優先嘗試台灣常用編碼 cp950 (Big5)，失敗則用 utf-8
            try:
                return pd.read_csv(file_obj, header=header, encoding='cp950')
            except UnicodeDecodeError:
                file_obj.seek(0)
                return pd.read_csv(file_obj, header=header, encoding='utf-8')
        else:
            return pd.read_excel(file_obj, header=header)
    except Exception as e:
        return None

def parse_score_table(file_obj):
    """解析配分表，回傳字典 Code -> {stop, report}"""
    score_map = {}
    
    # 讀取檔案
    df = load_data(file_obj)
    if df is None:
        st.error("配分表讀取失敗，請確認格式。")
        return None

    # 欄位檢查 (至少要有5欄)
    if df.shape[1] < 5:
        st.error("配分表欄位不足，請確認是否為標準格式。")
        return None

    # 建立配分字典
    for index, row in df.iterrows():
        try:
            # 假設欄位順序固定：[1]=違規代碼, [3]=攔停配分, [4]=逕舉配分
            code = str(row.iloc[1]).strip()
            
            # 排除空值
            if not code or code.lower() == 'nan': continue

            # 轉數字 (失敗歸0)
            stop_pt = pd.to_numeric(row.iloc[3], errors='coerce')
            report_pt = pd.to_numeric(row.iloc[4], errors='coerce')
            
            score_map[code] = {
                'stop': 0 if pd.isna(stop_pt) else stop_pt,
                'report': 0 if pd.isna(report_pt) else report_pt
            }
        except:
            continue
            
    return score_map

def extract_metadata_from_lines(df_head):
    """
    從檔案的前幾列 (DataFrame) 中尋找「舉發單位」與「舉發員警」
    """
    unit_name = "未知單位"
    officer_name = "未知員警"
    header_idx = 0
    
    # 轉成字串搜尋
    # 只需要看前 20 列，避免效能浪費
    search_range = df_head.head(20).astype(str)
    
    for idx, row in search_range.iterrows():
        row_str = " ".join(row.values) # 將整列合併成字串搜尋
        
        if "舉發單位" in row_str:
            # 簡易切割邏輯，視實際檔案格式可能需微調
            try:
                # 尋找冒號後的內容
                parts = row_str.split("舉發單位")
                if len(parts) > 1:
                    target = parts[1].replace("：", "").replace(":", "").strip()
                    unit_name = target.split()[0].split(',')[0] # 取第一個空白或逗號前的字
            except: pass
            
        if "舉發員警" in row_str:
            try:
                parts = row_str.split("舉發員警")
                if len(parts) > 1:
                    target = parts[1].replace("：", "").replace(":", "").strip()
                    officer_name = target.split()[0].split(',')[0]
            except: pass
            
        if "違規條款" in row_str and "攔停數" in row_str:
            header_idx = idx
            break # 找到表頭就可以停了
            
    return unit_name, officer_name, header_idx

def clean_number(x):
    """將含有逗號的字串轉為 float"""
    if isinstance(x, str):
        x = x.replace(',', '').strip()
    return pd.to_numeric(x, errors='coerce')

def process_police_results(files, score_map):
    """處理績效檔案 (優化版)"""
    officer_stats = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, file_obj in enumerate(files):
        progress_bar.progress((i + 1) / len(files))
        status_text.text(f"正在處理: {file_obj.name} ...")
        
        try:
            # 1. 先不帶 header 讀取整個檔案，為了抓取上方的單位資訊
            df_raw = load_data(file_obj, header=None)
            if df_raw is None: continue

            # 2. 提取 Metadata 與 真實 Header 位置
            unit_name, officer_name, header_idx = extract_metadata_from_lines(df_raw)
            
            # 3. 重新整理 DataFrame，設定正確的 Header
            # 取出 header_idx 之後的資料
            df_data = df_raw.iloc[header_idx+1:].copy()
            # 設定欄位名稱
            df_data.columns = df_raw.iloc[header_idx].astype(str).str.strip()
            df_data.reset_index(drop=True, inplace=True)
            
            # 4. 欄位識別
            code_col = next((c for c in df_data.columns if "違規條款" in c), None)
            stop_col = next((c for c in df_data.columns if "攔停數" in c), None)
            report_col = next((c for c in df_data.columns if "逕舉數" in c), None)
            
            if not (code_col and stop_col and report_col):
                continue # 找不到關鍵欄位則跳過

            # 5. 資料清洗與計算 (使用向量化運算取代迴圈，速度更快)
            # 移除不需要的列 (如 合計、nan)
            df_calc = df_data[df_data[code_col].notna()].copy()
            df_calc = df_calc[~df_calc[code_col].astype(str).str.contains("合計|舉發單張數|nan", case=False)]
            
            # 處理數字 (移除逗號)
            df_calc['clean_stop'] = df_calc[stop_col].apply(clean_number).fillna(0)
            df_calc['clean_report'] = df_calc[report_col].apply(clean_number).fillna(0)
            
            # 6. 計算分數 (Mapping)
            # 建立分數查找表
            df_calc['stop_score'] = df_calc[code_col].map(lambda x: score_map.get(str(x).strip(), {}).get('stop', 0))
            df_calc['report_score'] = df_calc[code_col].map(lambda x: score_map.get(str(x).strip(), {}).get('report', 0))
            
            # 總分運算
            total_score = (df_calc['clean_stop'] * df_calc['stop_score'] + 
                           df_calc['clean_report'] * df_calc['report_score']).sum()
            
            officer_stats.append({
                '單位': unit_name,
                '員警': officer_name,
                '檔案': file_obj.name,
                '積分': total_score
            })
            
        except Exception as e:
            st.error(f"處理檔案 {file_obj.name} 時發生錯誤: {e}")
            continue
            
    progress_bar.empty()
    status_text.empty()
    return pd.DataFrame(officer_stats)

# ==========================================
# 3. 主執行區
# ==========================================

if st.button("🚀 開始統計", type="primary"):
    if not uploaded_score_file:
        st.warning("⚠️ 請先上傳「配分表」")
    elif not uploaded_result_files:
        st.warning("⚠️ 請上傳至少一個「績效報表」")
    else:
        # 1. 解析配分表
        with st.spinner("正在解析配分表..."):
            score_map = parse_score_table(uploaded_score_file)
        
        if score_map:
            st.success(f"✅ 配分表載入成功！共 {len(score_map)} 條規則。")
            
            # 2. 計算績效
            df_summary = pd.DataFrame() # 預設為空
            with st.spinner("正在計算員警積分..."):
                df_raw = process_police_results(uploaded_result_files, score_map)
            
            if not df_raw.empty:
                # 3. 彙整與排序
                # 排除未知資料
                df_clean = df_raw[(df_raw['單位'] != '未知單位') & (df_raw['員警'] != '未知員警')]
                
                # 同一位員警若有多個檔案，將積分加總
                df_summary = df_clean.groupby(['單位', '員警'])['積分'].sum().reset_index()
                df_summary = df_summary.sort_values(by=['積分', '單位'], ascending=[False, True])
                
                # 格式化積分 (若為整數則顯示整數)
                df_summary['積分'] = df_summary['積分'].apply(lambda x: int(x) if x % 1 == 0 else round(x, 1))

                st.divider()
                col_res1, col_res2 = st.columns([2, 1])
                
                with col_res1:
                    st.subheader("📊 員警績效排行榜")
                    st.dataframe(
                        df_summary,
                        column_config={
                            "單位": st.column_config.TextColumn("單位", width="medium"),
                            "員警": st.column_config.TextColumn("員警", width="medium"),
                            "積分": st.column_config.ProgressColumn(
                                "總積分", 
                                format="%d", 
                                min_value=0, 
                                max_value=int(df_summary['積分'].max()) if not df_summary.empty else 100
                            ),
                        },
                        use_container_width=True,
                        height=600
                    )

                with col_res2:
                    st.subheader("📥 匯出報告")
                    # 下載按鈕
                    csv_data = df_summary.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        label="下載統計結果 (CSV)",
                        data=csv_data,
                        file_name="員警績效統計排名.csv",
                        mime="text/csv",
                        type="primary"
                    )
                    
                    st.info(f"本次共統計 {len(df_summary)} 位員警，\n處理了 {len(uploaded_result_files)} 個檔案。")
                    
                    with st.expander("查看詳細除錯資料"):
                        st.dataframe(df_raw)
            else:
                st.error("❌ 未提取到有效數據，請檢查上傳檔案的格式是否正確。")
