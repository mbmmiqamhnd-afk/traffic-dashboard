import streamlit as st
import pandas as pd
import gspread
from datetime import datetime

# ==========================================
# 0. 基本設定與專案目標值
# ==========================================
st.set_page_config(page_title="龍潭分局交通戰情室", page_icon="🚓", layout="wide")

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"

# 類別順序：酒駕, 闖紅燈, 嚴重超速, 車不讓人, 行人違規, 大型車違規
TARGET_CONFIG = {
    '聖亭所': [5, 115, 5, 16, 7, 10],
    '龍潭所': [6, 145, 7, 20, 9, 12],
    '中興所': [5, 115, 5, 16, 7, 10],
    '石門所': [3, 80, 4, 11, 5, 7],
    '高平所': [3, 80, 4, 11, 5, 7],
    '三和所': [2, 40, 2, 6, 2, 5],
    '交通分隊': [5, 115, 4, 16, 6, 8]
}

CATS = ["酒後駕車", "闖紅燈", "嚴重超速", "車不讓人", "行人違規", "大型車違規"]

# 法條對應關鍵字
LAW_MAP = {
    "酒後駕車": ["35條", "73條2項", "73條3項"],
    "闖紅燈": ["53條"],
    "嚴重超速": ["43條", "40條"],
    "車不讓人": ["44條", "48條"],
    "行人違規": ["78條"],
    "大型車違規": ["29", "30", "18之1", "33條1項03款"]
}

def map_unit_name(raw_name):
    u_map = {'交通分隊': '交通分隊', '聖亭': '聖亭所', '龍潭': '龍潭所', '中興': '中興所', '石門': '石門所', '高平': '高平所', '三和': '三和所'}
    for key, val in u_map.items():
        if key in str(raw_name): return val
    return None

def get_counts(df, unit, categories_list):
    """從 DataFrame 中提取特定單位與類別的數據"""
    row = df[df['單位'].apply(map_unit_name) == unit]
    counts = {}
    for cat in categories_list:
        keywords = LAW_MAP[cat]
        matched_cols = [c for c in df.columns if any(k in c for k in keywords)]
        counts[cat] = int(row[matched_cols].sum(axis=1).values[0]) if not row.empty else 0
    return counts

# ==========================================
# 1. 介面顯示
# ==========================================
st.title("🚓 龍潭分局交通數據戰情室")
st.header("🚦 (一) 加強交通安全執法取締五項交通違規統計表")
st.info("此處為原本的五項違規報表區塊...")

st.markdown("<br><hr style='border:2px solid #ddd'><br>", unsafe_allow_html=True)

# ==========================================
# 2. 強化專案功能 (跨檔案統計)
# ==========================================
st.header(f"📈 (二) {PROJECT_NAME}")

col1, col2 = st.columns(2)
file1 = col1.file_uploader("1. 上傳第一份報表 (統計前五項)", type=["csv", "xlsx"], key="f1")
file2 = col2.file_uploader("2. 上傳第二份報表 (統計大型車)", type=["csv", "xlsx"], key="f2")

if file1 and file2:
    # 讀取檔案
    df1 = pd.read_csv(file1, skiprows=3) if file1.name.endswith('.csv') else pd.read_excel(file1, skiprows=3)
    df2 = pd.read_csv(file2, skiprows=3) if file2.name.endswith('.csv') else pd.read_excel(file2, skiprows=3)
    
    # 清洗欄位名稱
    df1.columns = [str(c).strip() for c in df1.columns]
    df2.columns = [str(c).strip() for c in df2.columns]

    final_results = []
    for unit, targets in TARGET_CONFIG.items():
        # 從第一個檔案提取前 5 項
        data_1to5 = get_counts(df1, unit, CATS[:5])
        # 從第二個檔案提取第 6 項 (大型車)
        data_6 = get_counts(df2, unit, [CATS[5]])
        
        # 合併數據
        all_counts = {**data_1to5, **data_6}
        
        unit_row = [unit]
        for i, cat in enumerate(CATS):
            count = all_counts[cat]
            target = targets[i]
            ratio = f"{(count/target*100):.0f}%" if target > 0 else "0%"
            unit_row.extend([count, target, ratio])
        final_results.append(unit_row)

    # 建立統計表格
    headers = ["單位"]
    for cat in CATS:
        headers.extend([f"{cat}_取締", f"{cat}_目標", f"{cat}_達成率"])
    
    df_final = pd.DataFrame(final_results, columns=headers)
    
    # 計算合計
    totals = ["合計"]
    for i in range(1, len(headers), 3):
        c_sum = df_final.iloc[:, i].sum()
        t_sum = df_final.iloc[:, i+1].sum()
        r_sum = f"{(c_sum/t_sum*100):.0f}%" if t_sum > 0 else "0%"
        totals.extend([int(c_sum), int(t_sum), r_sum])
    
    df_final = pd.concat([pd.DataFrame([totals], columns=headers), df_final]).reset_index(drop=True)

    st.subheader("📊 雙檔案彙整統計表")
    st.dataframe(df_final, use_container_width=True)

    # 同步雲端按鈕
    if st.button("🚀 同步整合數據至 Google Sheets", use_container_width=True):
        try:
            gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
            sh = gc.open_by_url(GOOGLE_SHEET_URL)
            ws = sh.worksheet(PROJECT_NAME)
            
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            h1 = [f"{PROJECT_NAME} (雙檔同步：{now_str})"] + [""] * 18
            h2 = [""] + [c for c in CATS for _ in range(3)]
            h3 = ["單位"] + ["取締", "目標", "達成率"] * 6
            
            ws.clear()
            ws.update(values=[h1, h2, h3] + df_final.values.tolist())
            st.success("✅ 雙檔案數據已成功彙整並同步至雲端！")
            st.balloons()
        except Exception as e:
            st.error(f"同步失敗：{e}")
