import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️")

def run_hour_stats():
    st.title("⏱️ 交通疏導勤務時數彙整")
    st.markdown("---")
    st.info("💡 **使用秘訣**：一次選取全月 (30~31份) 的 Excel/CSV 檔案拖入，系統會自動按姓名加總『守望』時數。")

    # 檔案上傳
    uploaded_files = st.file_uploader(
        "請上傳勤務明細表", 
        accept_multiple_files=True, 
        type=['csv', 'xlsx'],
        key="traffic_stat_uploader"
    )

    if uploaded_files:
        all_data = []
        
        for file in uploaded_files:
            try:
                # 自動判斷編碼讀取 CSV 或 Excel
                if file.name.endswith('.csv'):
                    df = pd.read_csv(file, encoding='utf-8-sig') 
                else:
                    df = pd.read_excel(file)

                # 1. 自動抓取時段欄位 (符合 08-09 這種格式)
                time_cols = [c for c in df.columns if re.search(r'\d{1,2}.?\d{1,2}', str(c))]
                
                # 2. 從檔名抓單位 (例如: 龍潭所1150401 -> 龍潭所)
                unit_name = re.split(r'\d+', file.name)[0]

                for _, row in df.iterrows():
                    name = str(row.get('姓名', '')).strip()
                    if not name or name in ['nan', 'None', '', '姓名']: continue
                    
                    # 3. 核心計數：掃描該行所有時段欄位中「守望」出現次數
                    watch_count = sum(1 for col in time_cols if "守望" in str(row[col]))
                    
                    if watch_count > 0:
                        all_data.append({
                            "單位": unit_name,
                            "姓名": name,
                            "守望時數": watch_count,
                            "日期日期來源": file.name
                        })
            except Exception as e:
                st.error(f"檔案 {file.name} 讀取失敗: {e}")

        if all_data:
            full_df = pd.DataFrame(all_data)
            
            tab1, tab2 = st.tabs(["🏆 個人總計結果", "📋 每日明細對帳"])
            
            with tab1:
                # 彙整與排序
                summary = full_df.groupby(['單位', '姓名'])['守望時數'].sum().reset_index()
                summary = summary.sort_values(by='守望時數', ascending=False)
                
                st.dataframe(summary, use_container_width=True, hide_index=True)
                
                # 下載 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月總計')
                    full_df.to_excel(writer, index=False, sheet_name='明細備查')
                
                st.download_button(
                    "📥 下載彙整報表 (Excel)", 
                    output.getvalue(), 
                    "交通疏導守望時數統計.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with tab2:
                st.write("可搜尋姓名核對該員在哪幾天、哪些時段執行守望：")
                q_name = st.text_input("搜尋姓名：")
                if q_name:
                    st.dataframe(full_df[full_df['姓名'].str.contains(q_name)], use_container_width=True)
                else:
                    st.dataframe(full_df, use_container_width=True)
        else:
            st.warning("上傳的檔案中找不到任何含有『守望』關鍵字的資料。")

if __name__ == "__main__":
    run_hour_stats()
