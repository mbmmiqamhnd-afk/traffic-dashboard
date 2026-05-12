import streamlit as st
import pandas as pd
import io
import re

# 1. 頁面配置
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def run_app():
    st.title("⏱️ 交通疏導勤務時數彙整 (番號過濾版)")
    st.markdown("---")

    # --- 側邊欄設定 ---
    st.sidebar.header("⚙️ 篩選規則設定")
    
    # 設定尖峰時段所在的欄位 (C欄=2, D欄=3...)
    am_cols_input = st.sidebar.text_input("上午尖峰欄位索引 (C欄起, 逗號隔開)", value="2, 3")
    pm_cols_input = st.sidebar.text_input("下午尖峰欄位索引 (逗號隔開)", value="12, 13")
    
    # 設定要排除的番號
    exclude_codes = st.sidebar.text_input("要排除的番號 (A欄內容)", value="A, B, C")
    exclude_list = [i.strip().upper() for i in exclude_codes.split(',') if i.strip()]

    try:
        peak_col_indices = [int(i.strip()) for i in (am_cols_input + "," + pm_cols_input).split(',') if i.strip()]
    except:
        st.sidebar.error("請輸入正確的數字格式")
        peak_col_indices = [2, 3, 12, 13]

    # --- 主介面 ---
    uploaded_files = st.file_uploader("請上傳勤務明細檔 (CSV/Excel)", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files:
        all_records = []
        
        for file in uploaded_files:
            try:
                # 讀取檔案 (header=None 確保不漏掉資料)
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=None, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, header=None, encoding='cp950')
                else:
                    df = pd.read_excel(file, header=None)

                unit_name = re.split(r'\d+', file.name)[0]

                # 從第 3 列 (Index 2) 開始掃描
                for r_idx in range(2, len(df)):
                    row = df.iloc[r_idx]
                    
                    # --- 核心過濾 1：排除番號為 A, B, C 的人員 ---
                    # 假設番號在 A 欄 (Index 0)
                    shift_code = str(row[0]).strip().upper()
                    if shift_code in exclude_list:
                        continue
                    
                    # 姓名通常在 B 欄 (Index 1)
                    name = str(row[1]).replace('\n', '').replace(' ', '')
                    if name in ['nan', 'None', '', '姓名', '合計', '總計']: 
                        continue
                    
                    watch_hours = 0
                    # --- 核心過濾 2：只檢查設定的尖峰欄位 ---
                    for c_idx in peak_col_indices:
                        if c_idx < len(row):
                            content = str(row[c_idx]).replace('\n', '').replace(' ', '')
                            if "守望" in content:
                                watch_hours += 1
                    
                    if watch_hours > 0:
                        all_records.append({
                            "單位": unit_name,
                            "姓名": name,
                            "時數": watch_hours,
                            "番號": shift_code,
                            "來源檔案": file.name
                        })
            except Exception as e:
                st.error(f"解析 {file.name} 失敗：{e}")

        # --- 顯示結果 ---
        if all_records:
            full_df = pd.DataFrame(all_records)
            summary = full_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
            summary = summary.sort_values(['單位', '時數'], ascending=[True, False])
            
            tab1, tab2 = st.tabs(["🏆 月總計結果", "🔍 明細對帳 (已排除 ABC)"])
            with tab1:
                st.subheader("📊 彙整統計結果")
                st.dataframe(summary, use_container_width=True, hide_index=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月彙整')
                    full_df.to_excel(writer, index=False, sheet_name='對帳明細')
                st.download_button("📥 下載 Excel 報表", output.getvalue(), "尖峰時數統計.xlsx")
                
            with tab2:
                st.subheader("📋 每日詳細紀錄")
                st.dataframe(full_df, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 找不到符合條件的『守望』紀錄。可能是番號皆為 A/B/C 或欄位索引設定錯誤。")

if __name__ == "__main__":
    run_app()
