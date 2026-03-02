import streamlit as st
import pandas as pd
import gspread
import io
import smtplib
import re
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# ==========================================
# 0. 基本設定與頁面配置
# ==========================================
st.set_page_config(page_title="龍潭分局交通戰情室", page_icon="🚓", layout="wide")

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"

# 單位名稱映射表
u_map = {
    '龍潭交通分隊': '交通分隊', '交通分隊': '交通分隊', '交通組': '科技執法', '科技執法': '科技執法',
    '聖亭派出所': '聖亭所', '聖亭所': '聖亭所', '龍潭派出所': '龍潭所', '龍潭所': '龍潭所',
    '中興派出所': '中興所', '中興所': '中興所', '石門派出所': '石門所', '石門所': '石門所',
    '高平派出所': '高平所', '高平所': '高平所', '三和派出所': '三和所', '三和所': '三和所'
}

def map_unit_name(raw_name):
    for key, val in u_map.items():
        if key in str(raw_name): return val
    return None

# ==========================================
# 1. 導覽標題
# ==========================================
st.title("🚓 桃園市政府警察局龍潭分局 - 交通數據戰情室")
st.markdown("---")

# ==========================================
# 2. 功能 (一)：五項交通違規統計 (原功能)
# ==========================================
st.header("🚦 (一) 加強交通安全執法取締五項交通違規統計表")

# --- 此處請保留您原本處理五項違規的程式碼邏輯 ---
# 包含：uploaded_files = st.file_uploader(...), process_data, send_email 等
# 注意：請確保原本的代碼中沒有 st.stop()，否則會看不到下方的專案功能。

st.info("請在此執行原本的五項違規每週報表分析流程...")

# ==========================================
# 3. 分隔線
# ==========================================
st.markdown("<br><hr style='border:2px solid #ddd'><br>", unsafe_allow_html=True)

# ==========================================
# 4. 功能 (二)：專案勤務取締統計 (新功能)
# ==========================================
st.header(f"📈 (二) {PROJECT_NAME}")

# 支援 CSV 與 Excel 上傳
uploaded_proj = st.file_uploader(
    "請上傳『法條件數統計報表』(CSV 或 Excel)", 
    type=["csv", "xlsx"], 
    key="project_stats_uploader"
)

if uploaded_proj:
    try:
        # 根據檔案類型讀取，鎖定第 4 列為標題 (skiprows=3)
        if uploaded_proj.name.endswith('.csv'):
            df_proj = pd.read_csv(uploaded_proj, skiprows=3)
        else:
            df_proj = pd.read_excel(uploaded_proj, skiprows=3)
        
        # 1. 清理數據
        df_proj.columns = [str(c).strip() for c in df_proj.columns]
        df_proj = df_proj[df_proj['單位'].notna()]
        df_proj = df_proj[~df_proj['單位'].isin(['合計', '總計', '小計', '列印人員：'])]
        
        # 2. 轉換數值與單位映射
        df_proj['合計'] = pd.to_numeric(df_proj['合計'], errors='coerce').fillna(0)
        df_proj['顯示單位'] = df_proj['單位'].apply(map_unit_name)
        
        # 3. 彙整統計
        summary_res = df_proj.dropna(subset=['顯示單位']).groupby('顯示單位')['合計'].sum().reset_index()
        summary_res.columns = ['單位', '取締件數']
        
        # 4. 計算合計列
        total_val = summary_res['取締件數'].sum()
        total_row = pd.DataFrame([['合計', total_val]], columns=['單位', '取締件數'])
        
        # 5. 排序邏輯
        unit_order = ['合計', '科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '交通分隊']
        summary_res['排序'] = pd.Categorical(summary_res['單位'], categories=unit_order, ordered=True)
        final_summary = pd.concat([total_row, summary_res]).sort_values('排序').drop(columns=['排序'])

        # 6. 介面呈現
        col_t, col_c = st.columns([1, 1])
        with col_t:
            st.subheader("📋 專案統計明細")
            st.dataframe(final_summary, use_container_width=True, hide_index=True)
        
        with col_c:
            st.subheader("📊 取締件數分布圖")
            # 排除合計行後畫圖
            chart_df = final_summary[final_summary['單位'] != '合計']
            st.bar_chart(chart_df.set_index('單位'))

        # 7. 雲端同步按鈕
        st.markdown("---")
        if st.button("🚀 同步專案數據至雲端試算表", use_container_width=True):
            with st.spinner("同步至 Google Sheets..."):
                try:
                    # 使用原本功能已設定好的 Secrets
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    
                    # 取得或新增分頁
                    try:
                        ws = sh.worksheet(PROJECT_NAME)
                    except:
                        ws = sh.add_worksheet(title=PROJECT_NAME, rows=50, cols=5)
                    
                    # 準備資料
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    update_data = [
                        [PROJECT_NAME],
                        [f"資料更新時間：{now_str}"],
                        ["單位", "總取締件數"]
                    ] + final_summary.values.tolist()
                    
                    ws.clear()
                    ws.update(values=update_data)
                    
                    st.success(f"✅ 同步成功！分頁：{PROJECT_NAME}")
                    st.balloons()
                except Exception as e:
                    st.error(f"雲端同步失敗：{e}")
                    
    except Exception as e:
        st.error(f"檔案處理發生錯誤：{e}")

# ==========================================
# 5. 側邊欄說明
# ==========================================
with st.sidebar:
    st.title("⚙️ 系統說明")
    st.markdown("""
    1. **上方區塊**：執行五項違規週報。
    2. **下方區塊**：執行專案勤務統計。
    ---
    **維護單位**：交通組
    """)
