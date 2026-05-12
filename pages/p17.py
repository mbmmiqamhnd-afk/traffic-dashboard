import streamlit as st
import pandas as pd
import re
import io

# 設定頁面配置
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def main():
    st.title("⏱️ 交通疏導勤務時數彙整系統")
    st.markdown("---")
    
    st.info("💡 **操作引導**：您可以一次將全月（約30~31份）的勤務明細表拖入下方上傳區，系統將自動進行交叉彙整。")

    # 檔案上傳
    uploaded_files = st.file_uploader(
        "請上傳勤務明細檔 (支援 CSV 或 Excel)", 
        accept_multiple_files=True, 
        type=['csv', 'xlsx'],
        key="traffic_uploader_p17"
    )

    if uploaded_files:
        all_records = []
        progress_bar = st.progress(0, text="正在解析檔案數據...")
        
        for i, file in enumerate(uploaded_files):
            try:
                # 讀取檔案
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, encoding='big5')
                else:
                    df = pd.read_excel(file)

                # 1. 自動辨識時段欄位 (找標題格式如 08-09 或 08:00-09:00)
                time_cols = [c for c in df.columns if re.search(r'\d{1,2}.?\d{1,2}', str(c))]
                
                # 2. 檔名抓取單位 (抓取數字前的文字)
                unit_name = re.match(r'^[^\d]+', file.name).group(0) if re.match(r'^[^\d]+', file.name) else "未知單位"

                for _, row in df.iterrows():
                    name = str(row.get('姓名', '')).strip()
                    if not name or name in ['nan', 'None', '', '姓名']: continue
                    
                    # 3. 關鍵字「守望」統計
                    watch_count = sum(1 for col in time_cols if "守望" in str(row[col]))
                    
                    if watch_count > 0:
                        all_records.append({
                            "單位": unit_name,
                            "姓名": name,
                            "時數": watch_count,
                            "來源日期": file.name
                        })
            except Exception as e:
                st.error(f"解析 {file.name} 時發生錯誤：{e}")
            
            progress_bar.progress((i + 1) / len(uploaded_files), text=f"已完成 {i+1}/{len(uploaded_files)} 個檔案")

        if all_records:
            full_df = pd.DataFrame(all_records)
            
            st.divider()
            
            # 分頁顯示結果
            tab1, tab2 = st.tabs(["🏆 本月個人總計", "🔍 詳細對帳清單"])
            
            with tab1:
                # 數據彙整
                summary = full_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
                summary = summary.sort_values(by=['單位', '時數'], ascending=[True, False])
                
                st.subheader("📊 彙整統計結果")
                st.dataframe(summary, use_container_width=True, hide_index=True)
                
                # 下載 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月彙整表')
                    full_df.to_excel(writer, index=False, sheet_name='明細紀錄表')
                
                st.download_button(
                    label="📥 下載完整統計 Excel",
                    data=output.getvalue(),
                    file_name=f"交通疏導統計_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with tab2:
                st.subheader("📋 每日勤務明細")
                search_name = st.text_input("🔍 輸入同仁姓名快速對帳：")
                if search_name:
                    st.dataframe(full_df[full_df['姓名'].str.contains(search_name)], use_container_width=True, hide_index=True)
                else:
                    st.dataframe(full_df, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 檔案讀取成功，但未在內容中偵測到「守望」關鍵字。")

if __name__ == "__main__":
    main()
