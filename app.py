import streamlit as st
import pandas as pd
import gspread
import io
from datetime import datetime

# ==========================================
# 1. 頁面基本設定
# ==========================================
st.set_page_config(page_title="龍潭分局交通戰情室", page_icon="🚓", layout="wide")

# 設定區
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"

# 單位映射表
u_map = {
    '龍潭交通分隊': '交通分隊', '交通分隊': '交通分隊', '交通組': '科技執法',
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所',
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所'
}

def map_unit_name(raw_name):
    for key, val in u_map.items():
        if key in str(raw_name): return val
    return None

# ==========================================
# 2. 側邊欄選單 (整合原有與新功能)
# ==========================================
with st.sidebar:
    st.title("🚓 交通戰情室選單")
    
    # 這裡列出您所有的功能選項
    choice = st.radio(
        "請選擇功能：",
        [
            "🏠 系統首頁",
            "1️⃣ 🚑 交通事故統計",
            "2️⃣ 🚔 取締重大交通違規統計",
            "3️⃣ 🚛 超載統計",
            "4️⃣ 🚦 加強交通安全執法五項違規",
            "5️⃣ 📈 " + PROJECT_NAME  # 這就是您要新增的功能
        ]
    )
    
    st.markdown("---")
    st.caption(f"最後更新：{datetime.now().strftime('%Y-%m-%d')}")
    st.caption("© 龍潭分局交通組")

# ==========================================
# 3. 主畫面邏輯控制
# ==========================================

# --- 首頁 ---
if choice == "🏠 系統首頁":
    st.title("🚓 桃園市政府警察局龍潭分局 - 交通數據戰情室")
    st.markdown("---")
    st.markdown("""
    ### 👋 歡迎使用自動化統計系統
    請從 **左側選單** 選擇您要使用的功能。
    
    * 目前已整合：事故、違規、超載及專案統計。
    * 資料同步對象：[Google Sheets 雲端試算表]({0})
    """.format(GOOGLE_SHEET_URL))

# --- 原有功能 1~3 (預留位置) ---
elif "1️⃣" in choice or "2️⃣" in choice or "3️⃣" in choice:
    st.header(choice)
    st.info("此功能模組開發中或請嵌入原有程式碼...")

# --- 原有功能 4：五項違規 ---
elif "4️⃣" in choice:
    st.header("🚦 加強交通安全執法取締五項交通違規統計表")
    # 這裡放入您原本『五項違規』的長程式碼邏輯
    st.info("請上傳相關報表進行分析...")

# --- 新增功能 5：專案件數統計 ---
elif "5️⃣" in choice:
    st.header(f"📊 {PROJECT_NAME}")
    
    # 檔案上傳
    uploaded_file = st.file_uploader("請上傳『法條件數統計報表.csv』", type="csv", key="proj_csv_upload")

    if uploaded_file:
        try:
            # 讀取並清洗數據
            df = pd.read_csv(uploaded_file, skiprows=3)
            df = df[df['單位'].notna() & (~df['單位'].isin(['總計', '合計', '列印人員：']))]
            
            # 數值轉換與單位映射
            df['合計'] = pd.to_numeric(df['合計'], errors='coerce').fillna(0)
            df['顯示單位'] = df['單位'].apply(map_unit_name)
            
            # 統計
            summary = df.dropna(subset=['顯示單位']).groupby('顯示單位')['合計'].sum().reset_index()
            summary.columns = ['單位', '取締件數']
            
            # 合計行
            total_df = pd.DataFrame([['合計', summary['取締件數'].sum()]], columns=['單位', '取締件數'])
            
            # 排序
            order = ['合計', '科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '交通分隊']
            summary['排序'] = pd.Categorical(summary['單位'], categories=order, ordered=True)
            final_df = pd.concat([total_df, summary]).sort_values('排序').drop(columns=['排序'])

            # 畫面呈現
            col1, col2 = st.columns([1, 1])
            with col1:
                st.subheader("📋 單位件數清單")
                st.table(final_df)
            with col2:
                st.subheader("📊 圖表分析")
                st.bar_chart(final_df[final_df['單位'] != '合計'].set_index('單位'))

            # 同步功能
            st.markdown("---")
            if st.button("🚀 同步數據至雲端試算表"):
                with st.spinner("同步中..."):
                    try:
                        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                        sh = gc.open_by_url(GOOGLE_SHEET_URL)
                        try:
                            ws = sh.worksheet(PROJECT_NAME)
                        except:
                            ws = sh.add_worksheet(title=PROJECT_NAME, rows=50, cols=5)
                        
                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        data_vals = [[PROJECT_NAME], [f"更新時間：{now_str}"], ["單位", "取締件數"]] + final_df.values.tolist()
                        
                        ws.clear()
                        ws.update(values=data_vals)
                        st.success(f"✅ 已同步至：{PROJECT_NAME}")
                        st.balloons()
                    except Exception as e:
                        st.error(f"雲端同步失敗：{e}")
        except Exception as e:
            st.error(f"檔案讀取失敗：{e}")
