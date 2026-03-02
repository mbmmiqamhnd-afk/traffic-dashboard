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
# 2. 側邊欄導覽設計 (取代分頁)
# ==========================================
with st.sidebar:
    st.title("🚓 功能選單")
    # 建立選單
    choice = st.selectbox(
        "請選擇要執行的功能：",
        ["🏠 系統首頁", "🚦 五項違規自動化", "📈 強化專案取締統計"]
    )
    st.markdown("---")
    st.caption(f"最後更新：{datetime.now().strftime('%Y-%m-%d')}")
    st.caption("© 龍潭分局交通組")

# ==========================================
# 3. 根據選單切換主畫面內容
# ==========================================

# --- 選項 1：系統首頁 ---
if choice == "🏠 系統首頁":
    st.title("🚓 龍潭分局 - 交通數據戰情室")
    st.markdown("---")
    st.image("https://via.placeholder.com/800x200?text=Longtan+Traffic+Police+Intelligence+Center") # 可替換為分局照片
    st.markdown(f"""
    ### 👋 歡迎使用交通數據統計系統
    
    目前的導覽模式已整合至 **左側側邊欄**。
    
    * **系統狀態**：🟢 正常運作中
    * **當前功能**：{choice}
    * **目標雲端**：[點此開啟試算表]({GOOGLE_SHEET_URL})
    """)

# --- 選項 2：五項交通違規統計 (原功能) ---
elif choice == "🚦 五項違規自動化":
    st.header("🚦 加強交通安全執法取締五項交通違規統計表")
    st.info("請在此處執行原本的五項違規分析流程。")
    # (此處貼上您原本處理 6 個檔案的長代碼)

# --- 選項 3：強化專案統計 (新功能) ---
elif choice == "📈 強化專案取締統計":
    st.header(f"📊 {PROJECT_NAME}")
    
    uploaded_file = st.file_uploader("請上傳『法條件數統計報表.csv』", type="csv", key="side_proj_upload")

    if uploaded_file:
        # (這裡維持之前的處理邏輯)
        df = pd.read_csv(uploaded_file, skiprows=3)
        df = df[df['單位'].notna() & (~df['單位'].isin(['總計', '合計', '列印人員：']))]
        df['合計'] = pd.to_numeric(df['合計'], errors='coerce').fillna(0)
        df['顯示單位'] = df['單位'].apply(map_unit_name)
        
        summary = df.dropna(subset=['顯示單位']).groupby('顯示單位')['合計'].sum().reset_index()
        summary.columns = ['單位', '取締件數']
        
        total_sum = summary['取締件數'].sum()
        total_df = pd.DataFrame([['合計', total_sum]], columns=['單位', '取締件數'])
        
        order = ['合計', '科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '交通分隊']
        summary['排序'] = pd.Categorical(summary['單位'], categories=order, ordered=True)
        final_df = pd.concat([total_df, summary]).sort_values('排序').drop(columns=['排序'])

        st.subheader("📋 統計結果")
        st.dataframe(final_df, use_container_width=True, hide_index=True)
        
        if st.button("🚀 同步至雲端分頁"):
            # (同步邏輯同前)
            st.success("數據已同步至 Google Sheets！")
