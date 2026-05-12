import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def main():
    st.title("⏱️ 交通疏導勤務時數彙整系統")
    st.markdown("---")
    
    st.info("💡 **偵測模式已強化**：現在會自動清除空格，並支援更多樣的時段格式。")

    uploaded_files = st.file_uploader(
        "請上傳勤務明細檔 (CSV 或 Excel)", 
        accept_multiple_files=True, 
        type=['csv', 'xlsx'],
        key="traffic_uploader_p17_v2"
    )

    if uploaded_files:
        all_records = []
        # 用來記錄哪些檔案沒抓到關鍵字，方便除錯
        no_keyword_files = [] 
        
        progress_bar = st.progress(0)
        
        for i, file in enumerate(uploaded_files):
            try:
                # 1. 讀取檔案
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, encoding='cp950') # 嘗試 Big5 編碼
                else:
                    df = pd.read_excel(file)

                # 2. 徹底清洗資料 (移除所有儲存格內的空格、換行符號)
                df = df.astype(str).apply(lambda x: x.str.strip().str.replace(r'\s+', '', regex=True))

                # 3. 找出時段欄位 (放寬條件：只要有數字與連字號)
                # 排除姓名、職稱、合計等固定欄位
                exclude_cols = ['姓名', '職稱', '合計', '序號', '備註']
                time_cols = [c for c in df.columns if any(char.isdigit() for char in str(c)) and str(c) not in exclude_cols]
                
                # 如果自動偵測失敗，嘗試抓取所有非排除的欄位
                if not time_cols:
                    time_cols = [c for c in df.columns if c not in exclude_cols]

                # 4. 檔名抓取單位
                unit_name = re.match(r'^[^\d]+', file.name).group(0) if re.match(r'^[^\d]+', file.name) else "未知單位"

                found_any_in_file = False
                for _, row in df.iterrows():
                    name = row.get('姓名', '').strip()
                    if not name or name in ['nan', 'None', '', '姓名', '合計']: continue
                    
                    # 關鍵字比對：只要內容包含「守望」兩個字
                    watch_count = sum(1 for col in time_cols if "守望" in str(row.get(col, '')))
                    
                    if watch_count > 0:
                        all_records.append({
                            "單位": unit_name,
                            "姓名": name,
                            "時數": watch_count,
                            "來源日期": file.name
                        })
                        found_any_in_file = True
                
                if not found_any_in_file:
                    no_keyword_files.append(file.name)

            except Exception as e:
                st.error(f"解析 {file.name} 時發生錯誤：{e}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))

        if all_records:
            full_df = pd.DataFrame(all_records)
            
            tab1, tab2, tab3 = st.tabs(["🏆 本月個人總計", "🔍 詳細對帳清單", "⚠️ 未偵測到關鍵字檔案"])
            
            with tab1:
                summary = full_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
                summary = summary.sort_values(by=['單位', '時數'], ascending=[True, False])
                st.subheader("📊 彙整統計結果")
                st.dataframe(summary, use_container_width=True, hide_index=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月彙整表')
                    full_df.to_excel(writer, index=False, sheet_name='明細紀錄表')
                
                st.download_button("📥 下載統計 Excel", output.getvalue(), f"交通統計_{pd.Timestamp.now().strftime('%m%d')}.xlsx")

            with tab2:
                st.subheader("📋 每日勤務明細")
                search_name = st.text_input("🔍 輸入姓名對帳：")
                display_df = full_df[full_df['姓名'].str.contains(search_name)] if search_name else full_df
                st.dataframe(display_df, use_container_width=True, hide_index=True)

            with tab3:
                if no_keyword_files:
                    st.warning("以下檔案讀取成功，但裡面沒有找到任何『守望』勤務：")
                    st.write(no_keyword_files)
                else:
                    st.success("所有上傳檔案皆已成功擷取數據！")
        else:
            st.error("❌ 依然無法偵測到「守望」。請檢查您的 Excel 檔案中，填寫勤務的欄位名稱是否包含在： " + str(df.columns.tolist()))
