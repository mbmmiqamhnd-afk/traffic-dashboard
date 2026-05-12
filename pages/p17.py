import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def main():
    st.title("⏱️ 交通疏導勤務時數彙整系統")
    st.info("💡 **座標模式啟動**：已針對標題列偵測失敗進行修正，現在將掃描全表格儲存格。")

    uploaded_files = st.file_uploader("請上傳勤務明細檔", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files:
        all_records = []
        progress_bar = st.progress(0)
        
        for i, file in enumerate(uploaded_files):
            try:
                # 1. 讀取檔案 (不設標題， header=None，確保所有資料都在 DataFrame 內)
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=None, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, header=None, encoding='cp950')
                else:
                    df = pd.read_excel(file, header=None)

                # 2. 決定「姓名」所在的欄位索引 (通常是第 1 或 2 欄)
                # 根據您的除錯資訊，姓名出現在標題列的第 2 個位置 (index 1)
                name_col_idx = 1 
                
                # 3. 找出單位名稱 (從檔名)
                unit_name = re.split(r'\d+', file.name)[0]

                # 4. 從第 3 列 (index 2) 開始掃描資料
                # 因為 C3 之後才是守望，代表我們從 index 2 的列與 index 2 的欄之後開始找
                for r_idx in range(2, len(df)):
                    row_data = df.iloc[r_idx]
                    name = str(row_data[name_col_idx]).strip()
                    
                    # 排除空行或合計行
                    if not name or name in ['nan', 'None', '', '合計', '總計', '姓名']: 
                        continue
                    
                    # 5. 掃描該行從 C 欄 (index 2) 以後的所有儲存格
                    watch_hours = 0
                    for c_idx in range(2, len(row_data)):
                        cell_content = str(row_data[c_idx]).replace('\n', '').replace(' ', '')
                        if "守望" in cell_content:
                            watch_hours += 1
                    
                    if watch_hours > 0:
                        all_records.append({
                            "單位": unit_name,
                            "姓名": name,
                            "時數": watch_hours,
                            "來源檔案": file.name
                        })
            except Exception as e:
                st.error(f"解析 {file.name} 失敗：{e}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))

        if all_records:
            full_df = pd.DataFrame(all_records)
            summary = full_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
            summary = summary.sort_values(['單位', '時數'], ascending=[True, False])
            
            tab1, tab2 = st.tabs(["🏆 月總計結果", "🔍 明細核對"])
            with tab1:
                st.dataframe(summary, use_container_width=True, hide_index=True)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月總計')
                    full_df.to_excel(writer, index=False, sheet_name='明細紀錄')
                st.download_button("📥 下載統計報表", output.getvalue(), f"交通統計_{pd.Timestamp.now().strftime('%m%d')}.xlsx")
            with tab2:
                st.dataframe(full_df, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 依然找不到「守望」。請檢查：1. 檔案內是否真的有這兩個字？ 2. 檔案是否受到加密保護？")
            st.write("檔案內容預覽（前 5 列）：")
            st.write(df.head(10)) # 顯示更多行數供您確認位置

if __name__ == "__main__":
    main()
