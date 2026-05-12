import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def main():
    st.title("⏱️ 交通疏導勤務時數彙整系統 (精確崗位版)")
    st.info("💡 **運作邏輯**：系統將依據您上傳的『崗位配置表』，僅針對尖峰時段的崗位計算時數。")

    # --- 第一部分：上傳配置表 ---
    st.subheader("1️⃣ 上傳崗位配置白名單")
    col_a, col_b = st.columns(2)
    with col_a:
        weekday_config = st.file_uploader("上傳『平日』崗位配置表", type=['csv', 'xlsx'])
    with col_b:
        weekend_config = st.file_uploader("上傳『假日』崗位配置表", type=['csv', 'xlsx'])

    # 解析白名單
    valid_posts = set()
    if weekday_config and weekend_config:
        try:
            # 讀取平日與假日配置
            def get_peak_posts(file):
                df_cfg = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
                # 篩選「時段」欄位包含「尖峰」字眼的崗位名稱
                # 假設崗位名稱在第2欄，時段在第3欄（請依實際調整）
                peaks = df_cfg[df_cfg.iloc[:, 2].str.contains('尖峰', na=False)].iloc[:, 1].unique()
                return [str(p).strip() for p in peaks]

            valid_posts.update(get_peak_posts(weekday_config))
            valid_posts.update(get_peak_posts(weekend_config))
            st.success(f"✅ 已載入 {len(valid_posts)} 個尖峰時段合法崗位。")
            with st.expander("查看合法崗位清單"):
                st.write(list(valid_posts))
        except Exception as e:
            st.error(f"配置表解析失敗，請檢查欄位格式：{e}")

    # --- 第二部分：上傳每日勤務表 ---
    st.subheader("2️⃣ 上傳每日勤務明細")
    uploaded_files = st.file_uploader("請選取當月所有勤務明細檔", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files and valid_posts:
        all_records = []
        progress_bar = st.progress(0)
        
        for i, file in enumerate(uploaded_files):
            try:
                # 讀取檔案 (不設標題， header=None 處理複雜格式)
                df = pd.read_csv(file, header=None, encoding='utf-8-sig') if file.name.endswith('.csv') else pd.read_excel(file, header=None)

                unit_name = re.split(r'\d+', file.name)[0]

                # 遍歷資料 (從第3列開始，即 index 2)
                for r_idx in range(2, len(df)):
                    row_data = df.iloc[r_idx]
                    name = str(row_data[1]).strip() # 姓名通常在 B 欄 (index 1)
                    
                    if not name or name in ['nan', 'None', '', '合計', '總計', '姓名']: 
                        continue
                    
                    watch_hours = 0
                    # 掃描時段格 (C欄之後)
                    for c_idx in range(2, len(row_data)):
                        cell_content = str(row_data[c_idx]).replace('\n', '').replace(' ', '')
                        
                        # 核心判斷：內容包含「守望」且 崗位名稱位於白名單中
                        # 這裡假設您的勤務表儲存格內會寫「[崗位名]守望」或類似文字
                        if "守望" in cell_content:
                            # 檢查該儲存格內容是否包含任何一個合法崗位名稱
                            if any(post in cell_content for post in valid_posts):
                                watch_hours += 1
                    
                    if watch_hours > 0:
                        all_records.append({
                            "單位": unit_name, "姓名": name, "時數": watch_hours, "來源": file.name
                        })
            except Exception as e:
                st.error(f"解析 {file.name} 失敗：{e}")
            progress_bar.progress((i + 1) / len(uploaded_files))

        # --- 第三部分：顯示結果 ---
        if all_records:
            full_df = pd.DataFrame(all_records)
            summary = full_df.groupby(['單位', '姓名'])['時數'].sum().reset_index()
            
            tab1, tab2 = st.tabs(["🏆 月總計結果", "🔍 明細核對"])
            with tab1:
                st.dataframe(summary.sort_values('時數', ascending=False), use_container_width=True, hide_index=True)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False)
                st.download_button("📥 下載統計報表", output.getvalue(), "尖峰時數統計.xlsx")
            with tab2:
                st.dataframe(full_df, use_container_width=True, hide_index=True)
    elif not valid_posts and uploaded_files:
        st.warning("請先上傳配置表以建立合法崗位名單。")
