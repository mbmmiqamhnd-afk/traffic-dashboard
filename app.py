import streamlit as st
import pandas as pd
import gspread
from datetime import datetime

# 1. 基本設定
st.set_page_config(page_title="龍潭分局交通戰情室", page_icon="🚓", layout="wide")

# 設定雲端同步的名稱與網址
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"

# 單位映射表 (維持您的習慣)
u_map = {
    '龍潭交通分隊': '交通分隊', '交通分隊': '交通分隊', '交通組': '科技執法',
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所',
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所'
}

def map_unit_name(raw_name):
    for key, val in u_map.items():
        if key in str(raw_name): return val
    return None

# ---------------------------------------------------------
# 2. 建立分頁導覽
# ---------------------------------------------------------
tab_home, tab_five, tab_project = st.tabs([
    "🏠 系統首頁", 
    "🚦 五項違規統計 (原功能)", 
    "📈 強化專案取締件數 (新功能)"
])

# --- 分頁：首頁 ---
with tab_home:
    st.title("🚓 桃園市政府警察局龍潭分局 - 交通數據戰情室")
    st.markdown("---")
    st.markdown("""
    ### 👋 歡迎使用自動化統計系統
    請切換上方分頁執行功能：
    * **🚦 加強交通安全執法取締五項交通違規統計表**：每週例行報表 (酒駕、闖紅燈等 5 項)。
    * **📈 強化交通安全執法專案勤務取締件數統計表**：單一報表總件數統計與雲端同步。
    """)

# --- 分頁：五項違規 (這裡請放入您原本那段長長的代碼) ---
with tab_five:
    st.header("🚦 五項交通違規統計")
    st.info("請在此處繼續執行您原本的自動化流程...")
    # (註：請將您原本 process_data, send_email 等邏輯貼在這裡)

# --- 分頁：新功能專案統計 ---
with tab_project:
    st.header(f"📊 {PROJECT_NAME}")
    
    # 檔案上傳
    stat_file = st.file_uploader("請上傳『法條件數統計報表.csv』", type="csv", key="proj_upload")

    if stat_file:
        # 數據處理邏輯
        df_raw = pd.read_csv(stat_file, skiprows=3)
        df_raw = df_raw[df_raw['單位'].notna()]
        df_raw = df_raw[~df_raw['單位'].isin(['總計', '合計', '列印人員：'])]
        
        # 轉換數值與單位
        df_raw['合計'] = pd.to_numeric(df_raw['合計'], errors='coerce').fillna(0)
        df_raw['顯示單位'] = df_raw['單位'].apply(map_unit_name)
        
        # 統計與排序
        res = df_raw.dropna(subset=['顯示單位']).groupby('顯示單位')['合計'].sum().reset_index()
        res.columns = ['單位', '取締件數']
        
        total_val = res['取締件數'].sum()
        total_row = pd.DataFrame([['合計', total_val]], columns=['單位', '取締件數'])
        
        order = ['合計', '科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '交通分隊']
        res['排序'] = pd.Categorical(res['單位'], categories=order, ordered=True)
        final_df = pd.concat([total_row, res]).sort_values('排序').drop(columns=['排序'])

        # 顯示預覽
        c1, c2 = st.columns([1, 1])
        with c1:
            st.table(final_df)
        with c2:
            st.bar_chart(final_df[final_df['單位'] != '合計'].set_index('單位'))

        # 雲端同步按鈕
        if st.button("🚀 同步至雲端試算表"):
            try:
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL)
                try:
                    ws = sh.worksheet(PROJECT_NAME)
                except:
                    ws = sh.add_worksheet(title=PROJECT_NAME, rows=50, cols=5)
                
                now = datetime.now().strftime("%Y-%m-%d %H:%M")
                data = [[PROJECT_NAME], [f"更新時間：{now}"], ["單位", "取締件數"]] + final_df.values.tolist()
                ws.clear()
                ws.update(values=data)
                st.success("✅ 已成功同步至雲端分頁！")
                st.balloons()
            except Exception as e:
                st.error(f"同步失敗：{e}")
