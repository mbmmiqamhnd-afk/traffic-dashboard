import streamlit as st
import pandas as pd
import io
import re

# 1. 頁面配置
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def run_app():
    st.title("⏱️ 交通疏導勤務時數彙整 (可編輯下載版)")
    st.markdown("---")

    # --- 側邊欄設定 ---
    st.sidebar.header("⚙️ 篩選規則設定")
    exclude_input = st.sidebar.text_input("要排除的番號 (A欄內容)", value="A, B, C")
    exclude_list = [i.strip().upper() for i in exclude_input.split(',') if i.strip()]
    
    st.sidebar.divider()
    am_cols_input = st.sidebar.text_input("上午尖峰欄位索引 (C欄起, 逗號隔開)", value="2, 3")
    pm_cols_input = st.sidebar.text_input("下午尖峰欄位索引 (逗號隔開)", value="12, 13")
    
    try:
        peak_col_indices = [int(i.strip()) for i in (am_cols_input + "," + pm_cols_input).split(',') if i.strip()]
    except:
        peak_col_indices = [2, 3, 12, 13]

    # --- 主介面 ---
    uploaded_files = st.file_uploader("請上傳勤務明細檔", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files:
        all_records = []
        for file in uploaded_files:
            try:
                df = pd.read_csv(file, header=None, encoding='utf-8-sig') if file.name.endswith('.csv') else pd.read_excel(file, header=None)
                unit_name = re.split(r'\d+', file.name)[0]

                for r_idx in range(2, len(df)):
                    row = df.iloc[r_idx]
                    shift_code = str(row[0]).strip().upper()
                    if shift_code in exclude_list: continue
                    
                    name = str(row[1]).replace('\n', '').replace(' ', '')
                    if name in ['nan', 'None', '', '姓名', '合計', '總計']: continue
                    
                    watch_hours = 0
                    for c_idx in peak_col_indices:
                        if c_idx < len(row):
                            if "守望" in str(row[c_idx]).replace('\n', ''):
                                watch_hours += 1
                    
                    if watch_hours > 0:
                        all_records.append({"單位": unit_name, "姓名": name, "時數": watch_hours, "番號": shift_code, "來源檔案": file.name})
            except Exception as e:
                st.error(f"解析 {file.name} 失敗：{e}")

        if all_records:
            full_df = pd.DataFrame(all_records)
            # 初始加總彙整
            summary_df = full_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
            summary_df = summary_df.sort_values(['單位', '時數'], ascending=[True, False])

            st.divider()
            st.subheader("📝 統計結果編輯區")
            st.info("💡 **如何刪除？** 點擊表格左側的小方塊選取該列，按鍵盤 `Backspace` 或 `Delete`；或直接修改儲存格內容。")

            # --- 核心功能：可編輯表格 ---
            # 使用 num_rows="dynamic" 允許使用者刪除整列
            edited_df = st.data_editor(
                summary_df, 
                use_container_width=True, 
                num_rows="dynamic",
                key="data_editor_summary",
                hide_index=True
            )

            # --- 下載區：使用編輯後的 edited_df ---
            st.subheader("📥 下載最終報表")
            col1, col2 = st.columns(2)
            with col1:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    edited_df.to_excel(writer, index=False, sheet_name='月彙整(已手動修正)')
                    full_df.to_excel(writer, index=False, sheet_name='原始明細對帳')
                
                st.download_button(
                    label="📥 下載修正後的 Excel 報表",
                    data=output.getvalue(),
                    file_name=f"交通疏導統計_修正版_{pd.Timestamp.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with st.expander("🔍 查看原始過濾明細"):
                st.dataframe(full_df, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 找不到符合條件的資料。")

if __name__ == "__main__":
    run_app()
