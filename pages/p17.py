import streamlit as st
import pandas as pd
import re
import io

# --- 1. 頁面配置 (必須是第一個 Streamlit 指令) ---
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def run_hour_stats():
    # --- 2. 介面標題 ---
    st.title("⏱️ 交通疏導勤務時數彙整系統")
    st.markdown("---")
    
    st.info("💡 **操作引導**：請一次選取全月（約30~31份）的勤務明細表拖入，系統將自動按姓名加總『守望』時數。")

    # --- 3. 檔案上傳 ---
    uploaded_files = st.file_uploader(
        "請上傳勤務明細檔 (支援 CSV 或 Excel)", 
        accept_multiple_files=True, 
        type=['csv', 'xlsx'],
        key="traffic_uploader_p17_final"
    )

    if uploaded_files:
        all_records = []
        progress_bar = st.progress(0)
        
        for i, file in enumerate(uploaded_files):
            try:
                # A. 彈性讀取不同編碼與格式
                if file.name.endswith('.csv'):
                    try:
                        # 優先嘗試 utf-8 (包含 sig)
                        df = pd.read_csv(file, encoding='utf-8-sig')
                    except:
                        # 若失敗則嘗試警用系統常見的 Big5 (cp950)
                        df = pd.read_csv(file, encoding='cp950')
                else:
                    df = pd.read_excel(file)

                # B. 資料清洗：強制轉字串、移除所有儲存格內的空格與換行
                df = df.astype(str).apply(lambda x: x.str.strip().str.replace(r'\s+', '', regex=True))

                # C. 自動辨識時段欄位 (標題包含數字且不屬於姓名職稱等排除名單)
                exclude_keywords = ['姓名', '職稱', '合計', '序號', '備註', '單位', '總計']
                time_cols = [c for c in df.columns if any(char.isdigit() for char in str(c)) 
                             and not any(k in str(c) for k in exclude_keywords)]
                
                # 若完全找不到含數字欄位，則退而求其次抓取所有欄位（扣除姓名）
                if not time_cols:
                    time_cols = [c for c in df.columns if c not in exclude_keywords]

                # D. 從檔名抓單位 (例如: 龍潭所1150401... -> 龍潭所)
                unit_name = re.match(r'^[^\d]+', file.name).group(0) if re.match(r'^[^\d]+', file.name) else "其他單位"

                # E. 逐行掃描
                for _, row in df.iterrows():
                    name = row.get('姓名', '').strip()
                    # 過濾無效行
                    if not name or name in ['nan', 'None', '', '姓名', '合計', '總計']: 
                        continue
                    
                    # 精確計算：該員在時段欄位中，出現「守望」的總次數
                    # (已在前面將資料清洗過，所以能對付「守 望」或「守望 」)
                    watch_count = sum(1 for col in time_cols if "守望" in str(row.get(col, '')))
                    
                    if watch_count > 0:
                        all_records.append({
                            "單位": unit_name,
                            "姓名": name,
                            "時數": watch_count,
                            "來源日期": file.name
                        })
            except Exception as e:
                st.error(f"解析檔案 `{file.name}` 時出錯：{e}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))

        # --- 4. 結果顯示區 ---
        if all_records:
            full_df = pd.DataFrame(all_records)
            
            st.divider()
            tab1, tab2 = st.tabs(["🏆 個人總計結果", "🔍 詳細對帳清單"])
            
            with tab1:
                # 數據彙整加總
                summary = full_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
                summary = summary.sort_values(by=['單位', '時數'], ascending=[True, False])
                
                st.subheader("📊 本月彙整統計")
                st.dataframe(summary, use_container_width=True, hide_index=True)
                
                # Excel 匯出功能
                output = io.BytesIO()
                try:
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        summary.to_excel(writer, index=False, sheet_name='月總計')
                        full_df.to_excel(writer, index=False, sheet_name='明細對帳表')
                    
                    st.download_button(
                        label="📥 下載完整統計 Excel",
                        data=output.getvalue(),
                        file_name=f"交通疏導時數統計_{pd.Timestamp.now().strftime('%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as ex:
                    st.error(f"匯出 Excel 失敗，請確認是否安裝 openpyxl：{ex}")

            with tab2:
                st.subheader("📋 每日詳細紀錄")
                search_name = st.text_input("🔍 搜尋同仁姓名核對明細：")
                if search_name:
                    st.dataframe(full_df[full_df['姓名'].str.contains(search_name)], use_container_width=True, hide_index=True)
                else:
                    st.dataframe(full_df, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 檔案讀取成功，但內容中沒有找到任何『守望』關鍵字。請檢查檔案內容或確認欄位名稱。")

# 執行主程式
if __name__ == "__main__":
    run_hour_stats()
