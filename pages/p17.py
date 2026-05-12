import streamlit as st
import pandas as pd
import io
import re

# 1. 頁面配置 (必須在最頂端)
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def run_app():
    st.title("⏱️ 交通疏導勤務時數彙整 (精確崗位版)")
    st.markdown("---")

    # --- 側邊欄：設定合法崗位關鍵字 ---
    # 考慮到手動解析配置表容易報錯，建議直接在此輸入或由系統預設
    st.sidebar.header("⚙️ 崗位設定")
    default_posts = "大昌與五福,大昌與健行,中正與北龍,中正與中豐,中正與大昌,中正與新龍,大昌與復興,大昌與九龍"
    posts_input = st.sidebar.text_area("合法尖峰崗位關鍵字 (請用逗號隔開)", value=default_posts)
    valid_posts = [p.strip() for p in posts_input.split(',') if p.strip()]

    # --- 主介面：上傳勤務明細 ---
    st.subheader("1️⃣ 上傳每日勤務明細")
    st.info(f"當前比對崗位：{', '.join(valid_posts[:5])} ...等 {len(valid_posts)} 個")
    
    uploaded_files = st.file_uploader(
        "請選取當月所有勤務明細檔 (CSV/Excel)", 
        accept_multiple_files=True, 
        type=['csv', 'xlsx'],
        key="traffic_final_uploader"
    )

    if uploaded_files:
        all_records = []
        progress_bar = st.progress(0)
        
        for i, file in enumerate(uploaded_files):
            try:
                # A. 讀取檔案
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=None, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, header=None, encoding='cp950')
                else:
                    df = pd.read_excel(file, header=None)

                # B. 解析單位 (從檔名)
                unit_name = re.split(r'\d+', file.name)[0]

                # C. 掃描資料列 (從第 3 列開始)
                for r_idx in range(2, len(df)):
                    row_data = df.iloc[r_idx]
                    # 姓名通常在 B 欄 (index 1)
                    name = str(row_data[1]).replace('\n', '').strip()
                    
                    if not name or name in ['nan', 'None', '', '合計', '總計', '姓名']: 
                        continue
                    
                    # D. 掃描時段格 (C欄 index 2 之後)
                    watch_hours = 0
                    for cell in row_data[2:]:
                        cell_content = str(cell).replace('\n', '').replace(' ', '')
                        
                        # 核心邏輯：儲存格需包含「守望」且包含任何一個「合法崗位」
                        if "守望" in cell_content:
                            if any(post in cell_content for post in valid_posts):
                                watch_hours += 1
                    
                    if watch_hours > 0:
                        all_records.append({
                            "單位": unit_name,
                            "姓名": name,
                            "時數": watch_hours,
                            "來源日期": file.name
                        })
            except Exception as e:
                st.warning(f"解析檔案 {file.name} 時發生錯誤，已跳過。錯誤訊息：{e}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))

        # --- 顯示結果 ---
        if all_records:
            full_df = pd.DataFrame(all_records)
            summary = full_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
            summary = summary.sort_values(['單位', '時數'], ascending=[True, False])
            
            tab1, tab2 = st.tabs(["🏆 月總計報表", "🔍 詳細明細核對"])
            
            with tab1:
                st.subheader("📊 彙整統計結果")
                st.dataframe(summary, use_container_width=True, hide_index=True)
                
                # Excel 匯出
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月彙整')
                    full_df.to_excel(writer, index=False, sheet_name='明細')
                
                st.download_button(
                    label="📥 下載 Excel 統計表",
                    data=output.getvalue(),
                    file_name=f"交通疏導統計_{pd.Timestamp.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with tab2:
                st.subheader("📋 每日詳細紀錄")
                st.dataframe(full_df, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 已讀取檔案，但在設定的尖峰崗位中未發現『守望』紀錄。")

# 執行 App
if __name__ == "__main__":
    try:
        run_app()
    except Exception as e:
        st.error(f"頁面加載失敗：{e}")
