import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def main():
    st.title("⏱️ 交通疏導勤務時數彙整系統")
    st.info("💡 **偵測模式更新**：已針對標題『直列顯示 (換行)』進行特殊解析。")

    uploaded_files = st.file_uploader("請上傳勤務明細檔", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files:
        all_records = []
        progress_bar = st.progress(0)
        
        for i, file in enumerate(uploaded_files):
            try:
                # 1. 讀取檔案 (跳過前兩行無關標題)
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, skiprows=2, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, skiprows=2, encoding='cp950')
                else:
                    df = pd.read_excel(file, skiprows=2)

                # 2. 【關鍵修正】清洗欄位名稱：移除所有換行符號 (\n, \r) 與空格
                # 這會把直列的 08-09 轉回正常的橫列 08-09
                df.columns = [re.sub(r'[\s\n\r]', '', str(c)) for c in df.columns]

                # 3. 找出時段欄位 (找包含 00-01, 08-09 等數字組合的欄位)
                time_cols = [c for c in df.columns if re.search(r'\d+', c) and '-' in c]
                
                # 如果還是抓不到，就抓所有包含數字的欄位 (排除職稱、序號等)
                if not time_cols:
                    time_cols = [c for c in df.columns if any(char.isdigit() for char in c) 
                                 and not any(x in c for x in ['職稱', '序號', '合計'])]

                # 4. 抓取單位 (從檔名)
                unit_name = re.split(r'\d+', file.name)[0]

                for _, row in df.iterrows():
                    # 姓名欄位也可能因為直列而有換行，一併清洗
                    name = str(row.get('姓名', '')).replace('\n', '').strip()
                    if not name or name in ['nan', 'None', '', '合計', '總計']: 
                        continue
                    
                    # 5. 關鍵字統計 (一小時一格)
                    # 先清洗儲存格內容，移除換行再比對
                    watch_hours = 0
                    for col in time_cols:
                        cell_content = str(row.get(col, '')).replace('\n', '')
                        if "守望" in cell_content:
                            watch_hours += 1  # 一小時一格計為 1
                    
                    if watch_hours > 0:
                        all_records.append({
                            "單位": unit_name,
                            "姓名": name,
                            "時數": watch_hours,
                            "來源日期": file.name
                        })
            except Exception as e:
                st.error(f"解析 {file.name} 失敗：{e}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))

        if all_records:
            full_df = pd.DataFrame(all_records)
            
            # 按單位與姓名匯總
            summary = full_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
            summary = summary.sort_values(['單位', '時數'], ascending=[True, False])
            
            tab1, tab2 = st.tabs(["🏆 月總計結果", "🔍 明細核對"])
            with tab1:
                st.subheader("📊 本月彙整統計")
                st.dataframe(summary, use_container_width=True, hide_index=True)
                
                # 下載 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月總計')
                    full_df.to_excel(writer, index=False, sheet_name='明細對帳')
                
                st.download_button(
                    label="📥 下載統計報表 (Excel)",
                    data=output.getvalue(),
                    file_name=f"交通疏導統計_{pd.Timestamp.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with tab2:
                st.subheader("📋 每日詳細紀錄")
                st.dataframe(full_df, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 依然找不到關鍵字。請確認表格中「守望」這兩個字是否正確。")
            st.write("目前偵測到的欄位標題：", df.columns.tolist())
            st.write("第一行資料預覽：", df.iloc[0].to_dict() if not df.empty else "無資料")

if __name__ == "__main__":
    main()
