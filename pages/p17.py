import streamlit as st
import pandas as pd
import io
import re

# 1. 頁面配置
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def run_app():
    st.title("⏱️ 交通疏導勤務時數彙整 (明細編輯版)")
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
    uploaded_files = st.file_uploader("請選取並上傳當月勤務明細檔", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files:
        all_records = []
        for file in uploaded_files:
            try:
                # 讀取檔案
                df = pd.read_csv(file, header=None, encoding='utf-8-sig') if file.name.endswith('.csv') else pd.read_excel(file, header=None)
                unit_name = re.split(r'\d+', file.name)[0]

                for r_idx in range(2, len(df)):
                    row = df.iloc[r_idx]
                    shift_code = str(row[0]).strip().upper()
                    if shift_code in exclude_list: continue
                    
                    name = str(row[1]).replace('\n', '').replace(' ', '')
                    if name in ['nan', 'None', '', '姓名', '合計', '總計']: continue
                    
                    # 統計該行在尖峰時段是否有守望
                    for c_idx in peak_col_indices:
                        if c_idx < len(row):
                            cell_val = str(row[c_idx]).replace('\n', '')
                            if "守望" in cell_val:
                                all_records.append({
                                    "單位": unit_name,
                                    "姓名": name,
                                    "時數": 1,
                                    "番號": shift_code,
                                    "來源檔案": file.name,
                                    "原始位置": f"列{r_idx+1}-欄{c_idx+1}"
                                })
            except Exception as e:
                st.error(f"解析 {file.name} 失敗：{e}")

        if all_records:
            initial_df = pd.DataFrame(all_records)

            st.divider()
            st.subheader("📝 第一步：編輯/刪除原始明細")
            st.info("💡 **如何刪除明細？** 點擊最左側勾選想要刪除的列，按鍵盤 `Delete`；或點擊表格左上角全選後取消勾選特定列。")

            # --- 核心功能：編輯原始明細 ---
            edited_detail_df = st.data_editor(
                initial_df,
                use_container_width=True,
                num_rows="dynamic", # 允許刪除列
                key="detail_editor",
                hide_index=False
            )

            # --- 第二步：自動根據編輯後的明細進行加總 ---
            st.divider()
            st.subheader("📊 第二步：自動加總結果 (根據上述明細)")
            
            if not edited_detail_df.empty:
                # 重新計算總額
                final_summary = edited_detail_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
                final_summary = final_summary.sort_values(['單位', '時數'], ascending=[True, False])
                st.dataframe(final_summary, use_container_width=True, hide_index=True)

                # --- 下載區 ---
                st.subheader("📥 第三步：下載報表")
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_summary.to_excel(writer, index=False, sheet_name='月彙整總表')
                    edited_detail_df.to_excel(writer, index=False, sheet_name='修正後明細備查')
                
                st.download_button(
                    label="📥 下載最終修正版報表",
                    data=output.getvalue(),
                    file_name=f"交通疏導統計_明細修正版_{pd.Timestamp.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("明細已被全數刪除，無資料可供統計。")
        else:
            st.warning("⚠️ 找不到符合條件的守望紀錄。")

if __name__ == "__main__":
    run_app()
