import streamlit as st
import pandas as pd
import io
import re

# 1. 頁面配置
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def run_app():
    st.title("⏱️ 交通疏導勤務時數彙整 (明細可刪除版)")
    st.markdown("---")

    # --- 側邊欄：動態規則設定 ---
    st.sidebar.header("⚙️ 篩選規則設定")
    
    # A. 設定要排除的番號
    exclude_input = st.sidebar.text_input(
        "要排除的番號 (A欄內容)", 
        value="A, B, C", 
        help="請以逗號分隔多個代號。例如：A, B, 專案, 休息"
    )
    exclude_list = [i.strip().upper() for i in exclude_input.split(',') if i.strip()]
    
    st.sidebar.divider()

    # B. 設定尖峰時段所在的欄位 (C欄=2, D欄=3...)
    am_cols_input = st.sidebar.text_input("上午尖峰欄位索引 (C欄起, 逗號隔開)", value="2, 3")
    pm_cols_input = st.sidebar.text_input("下午尖峰欄位索引 (逗號隔開)", value="12, 13")
    
    try:
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
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=None, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, header=None, encoding='cp950')
                else:
                    df = pd.read_excel(file, header=None)

                unit_name = re.split(r'\d+', file.name)[0]

                # 從第 3 列 (Index 2) 開始掃描資料
                for r_idx in range(2, len(df)):
                    row = df.iloc[r_idx]
                    
                    shift_code = str(row[0]).strip().upper()
                    if shift_code in exclude_list:
                        continue
                    
                    name = str(row[1]).replace('\n', '').replace(' ', '')
                    if name in ['nan', 'None', '', '姓名', '合計', '總計']: 
                        continue
                    
                    # 只要在尖峰時段內發現「守望」，就記錄一筆明細
                    for c_idx in peak_col_indices:
                        if c_idx < len(row):
                            cell_content = str(row[c_idx]).replace('\n', '').replace(' ', '')
                            if "守望" in cell_content:
                                all_records.append({
                                    "單位": unit_name,
                                    "姓名": name,
                                    "時數": 1,
                                    "番號": shift_code,
                                    "來源檔案": file.name,
                                    "原始座標": f"列{r_idx+1}-欄{c_idx+1}"
                                })
            except Exception as e:
                st.error(f"解析 {file.name} 失敗：{e}")

        # --- 結果呈現 ---
        if all_records:
            raw_detail_df = pd.DataFrame(all_records)
            
            st.divider()
            st.subheader("📋 第一步：每日明細編輯區 (可手動刪除列)")
            st.info("💡 **操作方式**：點擊列最左側選取後按鍵盤 `Delete` 即可刪除。刪除後下方的匯總結果會自動更新。")
            
            # --- 核心功能：明細層級刪除 ---
            edited_detail_df = st.data_editor(
                raw_detail_df,
                use_container_width=True,
                num_rows="dynamic",  # 允許使用者刪除行
                key="detail_editor_v2",
                hide_index=False
            )

            if not edited_detail_df.empty:
                # --- 自動根據刪除後的明細重新彙整 ---
                summary = edited_detail_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
                summary = summary.sort_values(['單位', '時數'], ascending=[True, False])
                
                st.divider()
                st.subheader("🏆 第二步：自動更新之彙整結果")
                st.dataframe(summary, use_container_width=True, hide_index=True)
                
                # Excel 下載功能 (下載編輯後的結果)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月彙整總表')
                    edited_detail_df.to_excel(writer, index=False, sheet_name='修正後明細對帳')
                
                st.download_button(
                    label="📥 下載最終修正版 Excel 報表",
                    data=output.getvalue(),
                    file_name=f"交通疏導統計_修正版_{pd.Timestamp.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ 明細已被全數刪除，無資料可供統計。")
        else:
            st.warning("⚠️ 找不到符合條件的資料。")

if __name__ == "__main__":
    run_app()
