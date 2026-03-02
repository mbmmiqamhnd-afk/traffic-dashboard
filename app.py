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
# 2. 側邊欄選單 (整合您原有的前四項功能)
# ==========================================
with st.sidebar:
    st.title("🚓 交通功能選單")
    choice = st.radio(
        "請選擇功能模組：",
        [
            "🏠 系統首頁",
            "1️⃣ 🚑 交通事故統計",
            "2️⃣ 🚔 取締重大交通違規統計",
            "3️⃣ 🚛 超載統計",
            "4️⃣ 🚦 五項交通違規統計",
            "5️⃣ 📈 " + PROJECT_NAME  # 新增的功能列在最後
        ]
    )
    st.markdown("---")
    st.caption(f"維護單位：交通組 | {datetime.now().strftime('%Y-%m-%d')}")

# ==========================================
# 3. 主畫面顯示邏輯
# ==========================================

# --- 首頁 ---
if choice == "🏠 系統首頁":
    st.title("🚓 龍潭分局交通數據戰情室")
    st.info("請點選左側選單執行特定統計功能。")
    st.markdown("""
    * **科技執法成果**：包含事故、違規及專案件數統計。
    * **雲端同步**：數據將自動寫入指定的 Google 試算表。
    """)

# --- 原有功能 1~4 ---
elif any(x in choice for x in ["1️⃣", "2️⃣", "3️⃣", "4️⃣"]):
    st.header(choice)
    st.warning("請在此處嵌入您原本處理該功能的代碼內容...")

# --- 5️⃣ 新功能：強化交通安全執法專案統計 ---
elif "5️⃣" in choice:
    st.header(f"📊 {PROJECT_NAME}")
    
    # 檔案上傳
    uploaded_file = st.file_uploader("請上傳『法條件數統計報表.csv』", type="csv", key="project_csv")

    if uploaded_file:
        # 讀取數據 (自動跳過前3行)
        df = pd.read_csv(uploaded_file, skiprows=3)
        df = df[df['單位'].notna() & (~df['單位'].isin(['總計', '合計', '列印人員：']))]
        
        # 數值轉換與單位映射
        df['合計'] = pd.to_numeric(df['合計'], errors='coerce').fillna(0)
        df['顯示單位'] = df['單位'].apply(map_unit_name)
        
        # 統計
        summary = df.dropna(subset=['顯示單位']).groupby('顯示單位')['合計'].sum().reset_index()
        summary.columns = ['單位', '取締件數']
        
        # 合計行與排序
        total_val = summary['取締件數'].sum()
        total_row = pd.DataFrame([['合計', total_val]], columns=['單位', '取締件數'])
        order = ['合計', '科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '交通分隊']
        summary['排序'] = pd.Categorical(summary['單位'], categories=order, ordered=True)
        final_df = pd.concat([total_row, summary]).sort_values('排序').drop(columns=['排序'])

        # 顯示統計表與圖表
        col_t, col_c = st.columns([1, 1])
        with col_t:
            st.subheader("📋 單位件數清單")
            st.dataframe(final_df, use_container_width=True, hide_index=True)
        with col_c:
            st.subheader("📊 取締分布圖")
            st.bar_chart(final_df[final_df['單位'] != '合計'].set_index('單位'))

        # 雲端同步按鈕
        st.markdown("---")
        if st.button("🚀 同步至雲端分頁", use_container_width=True):
            with st.spinner("同步至 Google Sheets 中..."):
                try:
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    
                    try:
                        ws = sh.worksheet(PROJECT_NAME)
                    except:
                        ws = sh.add_worksheet(title=PROJECT_NAME, rows=50, cols=5)
                    
                    # 準備寫入資料
                    update_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    data_vals = [[PROJECT_NAME], [f"更新時間：{update_time}"], ["單位", "取締件數"]] + final_df.values.tolist()
                    
                    ws.clear()
                    ws.update(values=data_vals)
                    st.success(f"✅ 成功同步至分頁：{PROJECT_NAME}")
                    st.balloons()
                except Exception as e:
                    st.error(f"同步失敗：{e}")
