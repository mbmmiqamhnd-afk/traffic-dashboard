import streamlit as st
import pandas as pd
import io
import re

# 1. 頁面配置
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def run_app():
    st.title("⏱️ 交通疏導勤務時數彙整 (座標精確比對版)")
    st.markdown("---")

    # --- 側邊欄：根據座標設定尖峰時段 ---
    st.sidebar.header("⚙️ 座標與時段設定")
    st.sidebar.info("根據您的說明，資料從 C 欄 (索引 2) 開始。")
    
    # 請根據您的 Excel 實際欄位順序調整（C欄=2, D欄=3, E欄=4...以此類推）
    am_cols_input = st.sidebar.text_input("上午尖峰所在的欄位索引 (C欄起, 逗號隔開)", value="2, 3")
    pm_cols_input = st.sidebar.text_input("下午尖峰所在的欄位索引 (逗號隔開)", value="12, 13")
    
    try:
        peak_col_indices = [int(i.strip()) for i in (am_cols_input + "," + pm_cols_input).split(',') if i.strip()]
    except:
        st.sidebar.error("請輸入正確的數字格式（例如: 2, 3, 12）")
        peak_col_indices = [2, 3, 12, 13]

    # --- 主介面 ---
    uploaded_files = st.file_uploader("請選取當月勤務明細檔 (CSV/Excel)", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files:
        all_records = []
        
        for file in uploaded_files:
            try:
                # 讀取檔案：不設 header，完全手動控制座標
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=None, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, header=None, encoding='cp950')
                else:
                    df = pd.read_excel(file, header=None)

                # 單位名稱
                unit_name = re.split(r'\d+', file.name)[0]

                # 從第 3 列 (Index 2) 開始掃描，因為您說守望在 C3 以後
                for r_idx in range(2, len(df)):
                    row = df.iloc[r_idx]
                    
                    # 姓名通常在 B 欄 (Index 1)
                    name = str(row[1]).replace('\n', '').replace(' ', '')
                    if name in ['nan', 'None', '', '姓名', '合計', '總計']: 
                        continue
                    
                    watch_hours = 0
                    # 只檢查我們設定的「尖峰欄位索引」
                    for c_idx in peak_col_indices:
                        if c_idx < len(row):
                            content = str(row[c_idx]).replace('\n', '').replace(' ', '')
                            if "守望" in content:
                                watch_hours += 1
                    
                    if watch_hours > 0:
                        all_records.append({
                            "單位": unit_name,
                            "姓名": name,
                            "守望時數": watch_hours,
                            "來源日期": file.name
                        })
            except Exception as e:
                st.error(f"解析 {file.name} 失敗：{e}")

        # --- 顯示結果 ---
        if all_records:
            full_df = pd.DataFrame(all_records)
            summary = full_df.groupby(['單位', '姓名'])['守望時數'].sum().reset_index()
            summary = summary.sort_values(['單位', '守望時數'], ascending=[True, False])
            
            tab1, tab2 = st.tabs(["🏆 月總計結果", "🔍 明細對帳"])
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
            st.warning("⚠️ 依然找不到『守望』。")
            st.write("請對照下方預覽，確認您在左側設定的「欄位索引」是否正確：")
            # 顯示前幾筆資料並標註索引，方便用戶校對
            preview_df = df.head(10).copy()
            st.write("目前讀取到的前 10 列資料 (欄位標題為索引數字)：")
            st.dataframe(preview_df, use_container_width=True)

if __name__ == "__main__":
    run_app()
