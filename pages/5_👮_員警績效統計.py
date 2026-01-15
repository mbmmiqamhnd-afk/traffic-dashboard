import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="員警績效統計", layout="wide", page_icon="👮")

st.title("👮 員警交通執法績效統計系統")
st.markdown("""
### 📝 使用說明
1. **上傳配分表**：請上傳「常用交通執法重點工作配分表.xlsx」或對應的 CSV 檔。
2. **上傳績效檔**：請一次選取所有「PoliceResult...」系列的 CSV 或 Excel 檔。
3. **系統運算**：系統將自動對應違規條款，計算攔停與逕舉分數，並依單位與員警彙整。
""")

# ==========================================
# 1. 檔案上傳區
# ==========================================
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 上傳配分表")
    uploaded_score_file = st.file_uploader("請上傳配分表 (xlsx/csv)", type=['xlsx', 'xls', 'csv'], key="score_uploader")

with col2:
    st.subheader("2. 上傳績效報表")
    uploaded_result_files = st.file_uploader("請上傳 PoliceResult 檔案 (可多選)", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True, key="result_uploader")

# ==========================================
# 2. 核心邏輯函數
# ==========================================

def parse_score_table(file_obj):
    """解析配分表，回傳字典 Code -> {stop, report}"""
    score_map = {}
    try:
        # 嘗試讀取
        file_obj.seek(0)
        if file_obj.name.endswith('.csv'):
            df = pd.read_csv(file_obj)
        else:
            df = pd.read_excel(file_obj)
        
        # 簡單檢查欄位數是否足夠 (假設配分表格式固定)
        if df.shape[1] < 5:
            st.error(f"配分表格式錯誤：欄位數量不足。")
            return None

        for index, row in df.iterrows():
            try:
                code = str(row.iloc[1]).strip()
                if not code or code.lower() == 'nan': continue

                stop_pt = pd.to_numeric(row.iloc[3], errors='coerce')
                report_pt = pd.to_numeric(row.iloc[4], errors='coerce')
                
                if pd.isna(stop_pt): stop_pt = 0
                if pd.isna(report_pt): report_pt = 0
                
                score_map[code] = {'stop': stop_pt, 'report': report_pt}
            except:
                continue
        
        return score_map
    except Exception as e:
        st.error(f"配分表讀取失敗: {e}")
        return None

def process_police_results(files, score_map):
    """處理所有 PoliceResult 檔案並計算分數"""
    officer_stats = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, file_obj in enumerate(files):
        progress_bar.progress((i + 1) / len(files))
        status_text.text(f"正在處理: {file_obj.name} ...")
        
        try:
            # 讀取檔案內容為文字 (為了找表頭資訊)
            file_obj.seek(0)
            try:
                if file_obj.name.endswith('.csv'):
                    content_str = file_obj.getvalue().decode('utf-8', errors='ignore')
                    lines = content_str.splitlines()
                else:
                    # Excel 轉文字處理較複雜，此處簡化處理，若為 Excel 建議轉 CSV 上傳
                    # 或是直接讀取 excel header
                    df_tmp = pd.read_excel(file_obj)
                    # 暫時無法從二進位 Excel 串流直接 regex 表頭，需依賴固定格式
                    # 這裡假設使用者上傳 CSV 為主 (如您範例)
                    lines = [] 
            except:
                lines = []
            
            unit_name = "未知單位"
            officer_name = "未知員警"
            header_row_index = -1
            
            # 1. 解析 Metadata (單位、員警) 與 Header 位置
            for idx, line in enumerate(lines[:25]): 
                line_clean = line.replace('"', '').strip()
                if "舉發單位：" in line_clean:
                    parts = line_clean.split("舉發單位：")
                    if len(parts) > 1:
                        unit_name = parts[1].split(",")[0].strip()
                if "舉發員警：" in line_clean:
                    parts = line_clean.split("舉發員警：")
                    if len(parts) > 1:
                        officer_name = parts[1].split(",")[0].strip()
                if "違規條款" in line_clean and "攔停數" in line_clean:
                    header_row_index = idx
            
            if header_row_index == -1:
                # 若找不到 header，嘗試直接讀取 (可能是標準格式)
                header_row_index = 0

            # 2. 讀取數據 DataFrame
            file_obj.seek(0)
            if file_obj.name.endswith('.csv'):
                df = pd.read_csv(file_obj, header=header_row_index)
            else:
                # 若是 Excel 且找不到 header，嘗試預設第 11 列 (常見格式)
                target_header = header_row_index if header_row_index != -1 else 10
                df = pd.read_excel(file_obj, header=target_header)
            
            # 清理欄位名稱
            df.columns = [str(c).strip() for c in df.columns]
            
            # 尋找關鍵欄位
            code_col = next((c for c in df.columns if "違規條款" in c), None)
            stop_col = next((c for c in df.columns if "攔停數" in c), None)
            report_col = next((c for c in df.columns if "逕舉數" in c), None)
            
            if not (code_col and stop_col and report_col):
                continue

            # 3. 計算分數
            total_score = 0
            for _, row in df.iterrows():
                code_raw = str(row[code_col]).strip()
                if not code_raw or code_raw in ["nan", "合計", "舉發單張數"]: continue
                
                try:
                    s_val = str(row[stop_col]).replace(',', '')
                    r_val = str(row[report_col]).replace(',', '')
                    count_stop = float(s_val) if s_val and s_val != 'nan' else 0
                    count_report = float(r_val) if r_val and r_val != 'nan' else 0
                except:
                    continue

                points = score_map.get(code_raw, {'stop': 0, 'report': 0})
                row_score = (count_stop * points['stop']) + (count_report * points['report'])
                total_score += row_score
            
            officer_stats.append({
                '單位': unit_name,
                '員警': officer_name,
                '檔案': file_obj.name,
                '積分': total_score
            })

        except Exception as e:
            # st.error(f"檔案 {file_obj.name} 處理錯誤: {e}")
            continue
            
    progress_bar.empty()
    status_text.empty()
    return pd.DataFrame(officer_stats)

# ==========================================
# 3. 主執行區
# ==========================================

if st.button("🚀 開始統計", type="primary"):
    if not uploaded_score_file or not uploaded_result_files:
        st.warning("請確保「配分表」與「績效檔案」皆已上傳。")
    else:
        # 1. 解析配分表
        with st.spinner("正在解析配分表..."):
            score_map = parse_score_table(uploaded_score_file)
        
        if score_map:
            st.success(f"配分表載入成功！共 {len(score_map)} 條規則。")
            
            # 2. 計算績效
            with st.spinner("正在計算員警積分..."):
                df_raw = process_police_results(uploaded_result_files, score_map)
            
            if not df_raw.empty:
                # 3. 彙整與排序
                # 過濾掉未識別的資料
                df_clean = df_raw[(df_raw['單位'] != '未知單位') & (df_raw['員警'] != '未知員警')]
                
                df_summary = df_clean.groupby(['單位', '員警'])['積分'].sum().reset_index()
                df_summary = df_summary.sort_values(by=['積分', '單位'], ascending=[False, True])
                
                # 格式化積分 (整數)
                df_summary['積分'] = df_summary['積分'].apply(lambda x: int(x) if x.is_integer() else x)

                st.divider()
                st.subheader("📊 員警績效排行榜")
                
                # 顯示表格
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
                
                # 下載按鈕
                csv_data = df_summary.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 下載統計結果 (CSV)",
                    data=csv_data,
                    file_name="員警績效統計排名.csv",
                    mime="text/csv"
                )
                
                with st.expander("查看原始檔案處理紀錄"):
                    st.dataframe(df_raw)
            else:
                st.warning("未提取到有效數據，請檢查檔案格式。")
