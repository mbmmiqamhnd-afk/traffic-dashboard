import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# 頁面配置
st.set_page_config(page_title="科技執法成效統計", layout="wide", page_icon="📸")

st.title("📸 科技執法成效分析系統")
st.markdown("""
### 📝 功能說明
本頁面專門分析 **科技執法系統** 匯出的逕行舉發清冊（如 `list2.csv`）。
1. **數據視覺化**：自動統計違規熱點、時段、車種及趨勢。
2. **格式支援**：支援包含「違規日期、時間、地點、車種」等欄位的 CSV 或 Excel 檔。
""")

# ==========================================
# 1. 檔案上傳
# ==========================================
uploaded_file = st.file_uploader("請上傳科技執法清冊 (CSV 或 Excel)", type=['csv', 'xlsx'], key="tech_uploader")

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
        
        # 清理欄位名稱
        df.columns = [str(c).strip() for c in df.columns]

        # 檢查必要欄位
        required_cols = ['違規日期', '違規時間', '違規地點', '車種', '違規事實1']
        
        if not all(col in df.columns for col in required_cols):
            st.error(f"❌ 檔案格式不符！請確保包含以下欄位：{required_cols}")
            st.info("目前的欄位有：" + ", ".join(df.columns.tolist()))
        else:
            # --- 資料前處理 ---
            # A. 日期處理 (民國轉西元)
            def parse_roc_date(val):
                try:
                    s = str(int(val))
                    if len(s) == 6: s = '0' + s # 處理 990101
                    year = int(s[:-4]) + 1911
                    month = int(s[-4:-2])
                    day = int(s[-2:])
                    return datetime(year, month, day)
                except:
                    return None
            
            df['日期_dt'] = df['違規日期'].apply(parse_roc_date)
            
            # B. 時間處理 (HHMM 轉 小時)
            def parse_hour(val):
                try:
                    s = str(int(val)).zfill(4)
                    return int(s[:2])
                except:
                    return 0
            df['小時'] = df['違規時間'].apply(parse_hour)

            # --- 頁面呈現 ---
            
            # 1. KPI 數據指標
            st.divider()
            total_count = len(df)
            top_loc = df['違規地點'].mode()[0]
            top_v = df['違規事實1'].mode()[0]
            
            kpi1, kpi2, kpi3 = st.columns(3)
            kpi1.metric("📸 舉發總件數", f"{total_count:,} 件")
            kpi2.metric("📍 違規熱點", top_loc)
            kpi3.metric("⚠️ 主要違規行為", top_v)

            # 2. 圖表分析區 - 第一排 (地點與車種)
            st.divider()
            row1_col1, row1_col2 = st.columns(2)
            
            with row1_col1:
                st.subheader("📍 十大違規路段排行")
                loc_df = df['違規地點'].value_counts().reset_index().head(10)
                loc_df.columns = ['地點', '件數']
                fig_loc = px.bar(loc_df.sort_values('件數'), x='件數', y='地點', orientation='h',
                                 text_auto=True, color='件數', color_continuous_scale='Reds')
                st.plotly_chart(fig_loc, use_container_width=True)

            with row1_col2:
                st.subheader("🚙 違規車種組成")
                type_df = df['車種'].value_counts().reset_index()
                type_df.columns = ['車種', '件數']
                fig_type = px.pie(type_df, values='件數', names='車種', hole=0.4,
                                  color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_type, use_container_width=True)

            # 3. 圖表分析區 - 第二排 (趨勢與時段)
            st.divider()
            row2_col1, row2_col2 = st.columns([2, 1])
            
            with row2_col1:
                st.subheader("📅 執法成效趨勢 (每日)")
                if not df['日期_dt'].isnull().all():
                    trend_df = df.groupby('日期_dt').size().reset_index(name='件數')
                    fig_trend = px.line(trend_df, x='日期_dt', y='件數', markers=True, title='每日件數變化')
                    fig_trend.update_xaxes(title="日期", tickformat="%m/%d")
                    st.plotly_chart(fig_trend, use_container_width=True)
                else:
                    st.warning("無法解析日期格式。")

            with row2_col2:
                st.subheader("⏰ 違規高峰時段")
                hour_df = df['小時'].value_counts().sort_index().reset_index()
                hour_df.columns = ['小時', '件數']
                # 補齊 24 小時
                full_hours = pd.DataFrame({'小時': range(24)})
                hour_df = pd.merge(full_hours, hour_df, on='小時', how='left').fillna(0)
                
                fig_hour = px.bar(hour_df, x='小時', y='件數', color='件數',
                                  labels={'小時': '24小時制', '件數': '違規量'})
                st.plotly_chart(fig_hour, use_container_width=True)

            # 4. 原始資料預覽
            st.divider()
            with st.expander("🔍 查看詳細資料表"):
                st.dataframe(df, use_container_width=True)
                
            # 5. 下載統計報表
            csv = df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 下載本次統計資料 (CSV)", csv, "科技執法統計結果.csv", "text/csv")

    except Exception as e:
        st.error(f"檔案讀取錯誤：{e}")

else:
    st.info("💡 請在上方上傳科技執法清冊檔案以開始分析。")
