import streamlit as st
import pandas as pd
from datetime import datetime

# 頁面配置
st.set_page_config(page_title="科技執法成效統計", layout="wide", page_icon="📸")

st.title("📸 科技執法成效分析系統 (穩定版)")
st.markdown("使用系統內建圖表，確保相容性。")

# ==========================================
# 1. 檔案上傳
# ==========================================
uploaded_file = st.file_uploader("請上傳科技執法清冊 (list2.csv)", type=['csv', 'xlsx'], key="tech_uploader_v2")

if uploaded_file:
    try:
        # 讀取檔案
        if uploaded_file.name.endswith('.csv'):
            try:
                df = pd.read_csv(uploaded_file)
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding='cp950')
        else:
            df = pd.read_excel(uploaded_file)
        
        df.columns = [str(c).strip() for c in df.columns]

        # 檢查必要欄位
        required_cols = ['違規日期', '違規時間', '違規地點', '車種', '違規事實1']
        
        if not all(col in df.columns for col in required_cols):
            st.error(f"❌ 檔案格式不符！請確保包含：{required_cols}")
        else:
            # --- 資料前處理 ---
            # 民國轉西元
            def parse_roc_date(val):
                try:
                    s = str(int(val)).zfill(7)
                    year = int(s[:-4]) + 1911
                    month = int(s[-4:-2])
                    day = int(s[-2:])
                    return datetime(year, month, day)
                except:
                    return None
            
            df['日期_dt'] = df['違規日期'].apply(parse_roc_date)
            df['小時'] = df['違規時間'].apply(lambda x: int(str(int(x)).zfill(4)[:2]) if pd.notna(x) else 0)

            # --- 儀表板呈現 ---
            total_count = len(df)
            st.metric("📸 舉發總件數", f"{total_count:,} 件")

            col1, col2 = st.columns(2)

            with col1:
                st.subheader("📍 十大違規路段排行")
                loc_df = df['違規地點'].value_counts().head(10)
                # 使用 Streamlit 內建長條圖
                st.bar_chart(loc_df)

            with col2:
                st.subheader("⏰ 違規高峰時段 (0-23時)")
                hour_counts = df['小時'].value_counts().sort_index()
                # 補足 24 小時確保圖表美觀
                full_hours = pd.Series(0, index=range(24))
                hour_counts = hour_counts.combine_first(full_hours)
                st.bar_chart(hour_counts)

            st.divider()
            
            st.subheader("📅 執法成效趨勢")
            if not df['日期_dt'].isnull().all():
                trend_df = df.groupby('日期_dt').size()
                # 使用 Streamlit 內建折線圖
                st.line_chart(trend_df)

            with st.expander("🔍 查看詳細資料表"):
                st.dataframe(df)

    except Exception as e:
        st.error(f"處理出錯：{e}")
else:
    st.info("💡 請上傳 list2.csv 檔案。")
