import streamlit as st
import pandas as pd
import io
import re

# 1. 頁面配置
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def run_app():
    st.title("⏱️ 交通疏導勤務時數彙整 (自定義排除版)")
    st.markdown("---")

    # --- 側邊欄：動態規則設定 ---
    st.sidebar.header("⚙️ 篩選規則設定")
    
    # A. 設定要排除的番號 (自由輸入)
    exclude_input = st.sidebar.text_input(
        "要排除的番號 (A欄內容)", 
        value="A, B, C", 
        help="請以逗號分隔多個代號。例如：A, B, 專案, 休息"
    )
    # 將輸入轉為大寫清單，並移除空格
    exclude_list = [i.strip().upper() for i in exclude_input.split(',') if i.strip()]
    
    st.sidebar.divider()

    # B. 設定尖峰時段所在的欄位 (C欄=2, D欄=3...)
    am_cols_input = st.sidebar.text_input("上午尖峰欄位索引 (C欄起, 逗號隔開)", value="2, 3")
    pm_cols_input = st.sidebar.text_input("下午尖峰欄位索引 (逗號隔開)", value="12, 13")
    
    try:
        # 整合所有要檢查的尖峰欄位索引
        peak_col_indices = [int(i.strip()) for i in (am_cols_input + "," + pm_cols_input).split(',') if i.strip()]
    except:
        st.sidebar.error("❌ 欄位索引請輸入數字 (例如: 2, 3)")
        peak_col_indices = [2, 3, 12, 13]

    # --- 主介面：檔案處理 ---
    uploaded_files = st.file_uploader("請上傳勤務明細檔 (CSV/Excel)", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files:
        all_records = []
        
        for file in uploaded_files:
            try:
                # 讀取檔案 (header=None 確保不漏掉任何原始資料)
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=None, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, header=None, encoding='cp950')
                else:
                    df = pd.read_excel(file, header=None)

                # 解析單位名稱 (檔名前段文字)
                unit_name = re.split(r'\d+', file.name)[0]

                # 從第 3 列 (Index 2) 開始掃描資料
                for r_idx in range(2, len(df)):
                    row = df.iloc[r_idx]
                    
                    # 1. 取得 A 欄番號並判斷是否排除
                    shift_code = str(row[0]).strip().upper()
                    if shift_code in exclude_list:
                        continue
                    
                    # 2. 取得 B 欄姓名
                    name = str(row[1]).replace('\n', '').replace(' ', '')
                    # 跳過無效名稱或合計行
                    if name in ['nan', 'None', '', '姓名', '合計', '總計']: 
                        continue
                    
                    # 3. 掃描設定的尖峰時段欄位
                    watch_hours = 0
                    for c_idx in peak_col_indices:
                        if c_idx < len(row):
                            cell_content = str(row[c_idx]).replace('\n', '').replace(' ', '')
                            if "守望" in cell_content:
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

        # --- 結果呈現 ---
        if all_records:
            full_df = pd.DataFrame(all_records)
            
            # 按單位與姓名匯總總時數
            summary = full_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
            summary = summary.sort_values(['單位', '時數'], ascending=[True, False])
            
            tab1, tab2 = st.tabs(["🏆 月總計結果", "📋 每日明細 (已過濾排除對象)"])
            
            with tab1:
                st.subheader("📊 彙整統計結果")
                st.dataframe(summary, use_container_width=True, hide_index=True)
                
                # Excel 下載功能
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月彙整')
                    full_df.to_excel(writer, index=False, sheet_name='對帳明細')
                
                st.download_button(
                    label="📥 下載 Excel 報表",
                    data=output.getvalue(),
                    file_name=f"交通疏導統計_{pd.Timestamp.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            with tab2:
                st.subheader("📋 每日詳細紀錄")
                st.info(f"當前排除之番號清單：{', '.join(exclude_list)}")
                st.dataframe(full_df, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 找不到符合條件的資料。請檢查：1. 番號是否被過濾？ 2. 欄位索引設定是否正確？")
            # 顯示預覽輔助對齊
            with st.expander("查看原始資料預覽 (協助確認欄位索引)"):
                st.write(df.head(10))

if __name__ == "__main__":
    run_app()
