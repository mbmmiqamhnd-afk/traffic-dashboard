import streamlit as st
import pandas as pd
import io
import re

# 1. 頁面配置
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def run_app():
    st.title("⏱️ 交通疏導勤務時數彙整 (明細編輯版)")
    st.markdown("---")

    # --- 側邊欄：動態規則設定 ---
    st.sidebar.header("⚙️ 篩選規則設定")
    
    exclude_input = st.sidebar.text_input(
        "要排除的番號 (A欄內容)", 
        value="A, B, C", 
        help="請以逗號分隔多個代號。例如：A, B, 專案"
    )
    exclude_list = [i.strip().upper() for i in exclude_input.split(',') if i.strip()]
    
    st.sidebar.divider()

    am_cols_input = st.sidebar.text_input("上午尖峰欄位索引 (C欄起, 逗號隔開)", value="2, 3")
    pm_cols_input = st.sidebar.text_input("下午尖峰欄位索引 (逗號隔開)", value="12, 13")
    
    try:
        peak_col_indices = [int(i.strip()) for i in (am_cols_input + "," + pm_cols_input).split(',') if i.strip()]
    except:
        st.sidebar.error("❌ 欄位索引請輸入數字")
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

                for r_idx in range(2, len(df)):
                    row = df.iloc[r_idx]
                    
                    # 1. 番號過濾
                    shift_code = str(row[0]).strip().upper()
                    if shift_code in exclude_list:
                        continue
                    
                    # 2. 姓名取得
                    name = str(row[1]).replace('\n', '').replace(' ', '')
                    if name in ['nan', 'None', '', '姓名', '合計', '總計']: 
                        continue
                    
                    # 3. 掃描尖峰欄位 (每一小時拆成一列明細)
                    for c_idx in peak_col_indices:
                        if c_idx < len(row):
                            cell_content = str(row[c_idx]).replace('\n', '').replace(' ', '')
                            if "守望" in cell_content:
                                all_records.append({
                                    "單位": unit_name,
                                    "姓名": name,
                                    "時數": 1,  # 每筆明細代表1小時
                                    "番號": shift_code,
                                    "來源檔案": file.name,
                                    "原始座標": f"列{r_idx+1}-欄{c_idx+1}"
                                })
            except Exception as e:
                st.error(f"解析 {file.name} 失敗：{e}")

        if all_records:
            # 轉換為初始 DataFrame
            raw_detail_df = pd.DataFrame(all_records)
            
            st.divider()
            
            # --- 第一區：明細編輯區 ---
            st.subheader("📝 第一步：編輯原始明細紀錄")
            st.info("💡 **如何刪除？** 點擊最左側序號框選取列後，按鍵盤 `Delete`。您刪除的每一列都會即時反映在下方的統計中。")
            
            # 使用 data_editor 讓使用者可以刪除特定明細
            edited_df = st.data_editor(
                raw_detail_df,
                use_container_width=True,
                num_rows="dynamic",  # 關鍵：允許動態刪除行
                key="detail_editor",
                hide_index=False
            )

            st.divider()

            # --- 第二區：即時加總呈現 ---
            st.subheader("📊 第二步：月總計結果 (根據上方編輯內容自動計算)")
            
            if not edited_df.empty:
                # 根據編輯後的明細重新加總
                summary = edited_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
                summary = summary.sort_values(['單位', '時數'], ascending=[True, False])
                
                st.dataframe(summary, use_container_width=True, hide_index=True)

                # --- 第三區：下載功能 ---
                st.subheader("📥 第三步：匯出報表")
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月彙整總表')
                    edited_df.to_excel(writer, index=False, sheet_name='修正後明細')
                
                st.download_button(
                    label="📥 下載最終修正版 Excel",
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
