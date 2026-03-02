import streamlit as st
import pandas as pd
import gspread
from datetime import datetime
import io

# 1. 頁面基本設定
st.set_page_config(page_title="龍潭分局交通戰情室", page_icon="🚓", layout="wide")

# 設定區 (確保名稱一致)
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"

u_map = {
    '龍潭交通分隊': '交通分隊', '交通分隊': '交通分隊', '交通組': '科技執法',
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所',
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所'
}

def map_unit_name(raw_name):
    for key, val in u_map.items():
        if key in str(raw_name): return val
    return None

# =========================================================
# 第一部分：原本的五項違規統計 (原功能)
# =========================================================
st.title("🚓 桃園市政府警察局龍潭分局 - 交通數據戰情室")
st.header("🚦 (一) 加強交通安全執法取締五項交通違規統計表")

# ---【注意】請將您原本「五項違規」的所有分析程式碼貼在下方 ---
# --- 請確保這段程式碼裡面沒有 st.stop()，否則會擋住下面的新功能 ---

st.write("這是原本功能的區域...")
# (這裡貼上您原本的上傳器、資料處理、寄信、同步雲端等程式碼)


# =========================================================
# 分隔線
# =========================================================
st.markdown("<br><hr style='border:2px solid #ddd'><br>", unsafe_allow_html=True)


# =========================================================
# 第二部分：新增的專案統計 (新功能)
# =========================================================
st.header(f"📈 (二) {PROJECT_NAME}")

# 使用獨立的 Key，確保與上面的上傳器不衝突
uploaded_file = st.file_uploader("請上傳『法條件數統計報表.csv』進行專案統計", type="csv", key="unique_project_key")

if uploaded_file:
    # 1. 讀取與處理
    df = pd.read_csv(uploaded_file, skiprows=3)
    df = df[df['單位'].notna() & (~df['單位'].isin(['總計', '合計', '列印人員：']))]
    df['合計'] = pd.to_numeric(df['合計'], errors='coerce').fillna(0)
    df['顯示單位'] = df['單位'].apply(map_unit_name)
    
    # 2. 彙整
    summary = df.dropna(subset=['顯示單位']).groupby('顯示單位')['合計'].sum().reset_index()
    summary.columns = ['單位', '取締件數']
    
    # 3. 合計與排序
    total_val = summary['取締件數'].sum()
    total_df = pd.DataFrame([['合計', total_val]], columns=['單位', '取締件數'])
    order = ['合計', '科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '交通分隊']
    summary['排序'] = pd.Categorical(summary['單位'], categories=order, ordered=True)
    final_df = pd.concat([total_df, summary]).sort_values('排序').drop(columns=['排序'])

    # 4. 顯示結果
    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("📋 統計表")
        st.dataframe(final_df, use_container_width=True, hide_index=True)
    with col2:
        st.subheader("📊 分布圖")
        st.bar_chart(final_df[final_df['單位'] != '合計'].set_index('單位'))

    # 5. 同步雲端
    if st.button("🚀 同步此專案數據至雲端", key="sync_btn_proj"):
        try:
            gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
            sh = gc.open_by_url(GOOGLE_SHEET_URL)
            try:
                ws = sh.worksheet(PROJECT_NAME)
            except:
                ws = sh.add_worksheet(title=PROJECT_NAME, rows=50, cols=5)
            
            update_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            data_vals = [[PROJECT_NAME], [f"更新時間：{update_time}"], ["單位", "取締件數"]] + final_df.values.tolist()
            ws.clear()
            ws.update(values=data_vals)
            st.success(f"✅ 成功同步至雲端分頁：{PROJECT_NAME}")
            st.balloons()
        except Exception as e:
            st.error(f"雲端同步失敗：{e}")
