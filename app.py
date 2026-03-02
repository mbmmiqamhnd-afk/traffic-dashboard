import streamlit as st
import pandas as pd
import gspread
import io
from datetime import datetime

# ==========================================
# 1. 頁面基本設定
# ==========================================
st.set_page_config(page_title="龍潭分局交通戰情室", page_icon="🚓", layout="wide")

# 設定區 (Google Sheets)
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
# 2. 標題與簡介
# ==========================================
st.title("🚓 桃園市政府警察局龍潭分局 - 交通數據戰情室")
st.markdown("---")

# ==========================================
# 3. 功能 A：加強交通安全執法取締五項交通違規統計表
# ==========================================
st.header("🚦 (一) 加強交通安全執法取締五項交通違規統計表")
st.info("請上傳相關報表檔案 (本期、本年、去年) 進行週報表分析。")

# --- 這裡請嵌入您原本「五項違規」的長程式碼 (如下載 Excel、發送郵件等) ---
# st.file_uploader("上傳五項違規報表...", accept_multiple_files=True)
# ... [您的原代碼邏輯] ...

st.markdown("<br><hr style='border:2px solid gray'><br>", unsafe_allow_html=True)

# ==========================================
# 4. 功能 B：強化交通安全執法專案勤務取締件數統計表
# ==========================================
st.header(f"📈 (二) {PROJECT_NAME}")

# 檔案上傳
uploaded_file = st.file_uploader("請上傳『法條件數統計報表.csv』進行專案統計", type="csv", key="project_only")

if uploaded_file:
    try:
        # 數據讀取與過濾 (跳過前3行)
        df = pd.read_csv(uploaded_file, skiprows=3)
        df = df[df['單位'].notna() & (~df['單位'].isin(['總計', '合計', '列印人員：']))]
        
        # 轉換數值
        df['合計'] = pd.to_numeric(df['合計'], errors='coerce').fillna(0)
        df['顯示單位'] = df['單位'].apply(map_unit_name)
        
        # 彙整統計
        summary = df.dropna(subset=['顯示單位']).groupby('顯示單位')['合計'].sum().reset_index()
        summary.columns = ['單位', '取締件數']
        
        # 加上合計行
        total_val = summary['取締件數'].sum()
        total_df = pd.DataFrame([['合計', total_val]], columns=['單位', '取締件數'])
        
        # 排序
        order = ['合計', '科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '交通分隊']
        summary['排序'] = pd.Categorical(summary['單位'], categories=order, ordered=True)
        final_df = pd.concat([total_df, summary]).sort_values('排序').drop(columns=['排序'])

        # 呈現結果
        col1, col2 = st.columns([1, 1])
        with col1:
            st.subheader("📋 專案統計清單")
            st.table(final_df)
        with col2:
            st.subheader("📊 取締件數分布")
            # 排除合計畫圖
            chart_df = final_df[final_df['單位'] != '合計']
            st.bar_chart(chart_df.set_index('單位'))

        # 同步按鈕
        if st.button("🚀 同步此專案數據至雲端試算表", use_container_width=True):
            with st.spinner("同步至 Google Sheets..."):
                try:
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    
                    try:
                        ws = sh.worksheet(PROJECT_NAME)
                    except:
                        ws = sh.add_worksheet(title=PROJECT_NAME, rows=50, cols=5)
                    
                    # 準備寫入
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    data_vals = [[PROJECT_NAME], [f"更新時間：{now_str}"], ["單位", "取締件數"]] + final_df.values.tolist()
                    
                    ws.clear()
                    ws.update(values=data_vals)
                    st.success(f"✅ 已成功更新至分頁：{PROJECT_NAME}")
                    st.balloons()
                except Exception as e:
                    st.error(f"同步失敗：{e}")
                    
    except Exception as e:
        st.error(f"檔案處理失敗：{e}")

# ==========================================
# 5. 側邊欄 (僅保留簡單設定與說明)
# ==========================================
with st.sidebar:
    st.title("⚙️ 系統說明")
    st.write("本頁面整合了兩大核心統計功能，請依序向下捲動操作。")
    st.markdown("---")
    st.caption("維護單位：龍潭分局交通組")
