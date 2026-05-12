import streamlit as st
import pandas as pd
import io
import re

# 1. 頁面配置
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def run_app():
    st.title("⏱️ 交通疏導勤務時數彙整 (檔名優化版)")
    st.markdown("---")

    # --- 側邊欄：規則設定 ---
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

    # --- 主介面：檔案處理 ---
    uploaded_files = st.file_uploader("請上傳勤務明細檔 (CSV/Excel)", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files:
        all_records = []
        detected_units = set() # 用來記錄偵測到的所有單位名稱
        
        for file in uploaded_files:
            try:
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=None, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, header=None, encoding='cp950')
                else:
                    df = pd.read_excel(file, header=None)

                # 解析單位名稱 (檔名前段文字)
                unit_name = re.split(r'\d+', file.name)[0].strip()
                if unit_name:
                    detected_units.add(unit_name)

                for r_idx in range(2, len(df)):
                    row = df.iloc[r_idx]
                    shift_code = str(row[0]).strip().upper()
                    if shift_code in exclude_list: continue
                    
                    name = str(row[1]).replace('\n', '').replace(' ', '')
                    if name in ['nan', 'None', '', '姓名', '合計', '總計']: continue
                    
                    daily_watch_hours = 0
                    for c_idx in peak_col_indices:
                        if c_idx < len(row):
                            cell_content = str(row[c_idx]).replace('\n', '').replace(' ', '')
                            if "守望" in cell_content:
                                daily_watch_hours += 1
                    
                    if daily_watch_hours > 0:
                        all_records.append({
                            "單位": unit_name,
                            "姓名": name,
                            "當日尖峰時數": daily_watch_hours,
                            "番號": shift_code,
                            "日期來源": file.name
                        })
            except Exception as e:
                st.error(f"解析 {file.name} 失敗：{e}")

        if all_records:
            raw_person_detail_df = pd.DataFrame(all_records)
            
            st.divider()
            st.subheader("📝 第一步：確認每日人員名單 (可整列刪除)")
            
            edited_detail_df = st.data_editor(
                raw_person_detail_df,
                use_container_width=True,
                num_rows="dynamic",
                key="person_detail_editor",
                hide_index=False
            )

            if not edited_detail_df.empty:
                summary = edited_detail_df.groupby(['單位', '姓名'])['當日尖峰時數'].sum().reset_index()
                summary.columns = ['單位', '姓名', '總計尖峰時數']
                summary = summary.sort_values(['單位', '總計尖峰時數'], ascending=[True, False])
                
                st.divider()
                st.subheader("📊 第二步：自動加總結果")
                st.dataframe(summary, use_container_width=True, hide_index=True)

                # --- 產生下載檔名邏輯 ---
                # 如果有多個單位，顯示前兩個並加上「等」；如果只有一個就顯示該單位
                unit_list = sorted(list(detected_units))
                if len(unit_list) > 1:
                    filename_prefix = f"{unit_list[0]}_{unit_list[1]}等{len(unit_list)}單位"
                elif len(unit_list) == 1:
                    filename_prefix = unit_list[0]
                else:
                    filename_prefix = "交通疏導"

                today_str = pd.Timestamp.now().strftime('%m%d')
                final_filename = f"{filename_prefix}_交通疏導統計_{today_str}.xlsx"

                # --- 下載區 ---
                st.subheader("📥 第三步：下載最終報表")
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月彙整總表')
                    edited_detail_df.to_excel(writer, index=False, sheet_name='人員核銷明細')
                
                st.download_button(
                    label=f"📥 下載 {final_filename}",
                    data=output.getvalue(),
                    file_name=final_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ 名單已被全數刪除。")
        else:
            st.warning("⚠️ 找不到符合條件的人員紀錄。")

if __name__ == "__main__":
    run_app()
