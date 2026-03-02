import streamlit as st
import pandas as pd
import gspread
from datetime import datetime

# ==========================================
# 0. 基本設定與目標值 (依據您上傳的專案格式)
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

# 法條與欄位定義 (對應自選匯出 (7) 系列檔案)
LAW_MAP = {
    "酒後駕車": ["35條", "73條2項", "73條3項"],
    "闖紅燈": ["53條"],
    "嚴重超速": ["43條", "40條"],
    "車不讓人": ["44條", "48條"],
    "大型車違規": ["29", "30", "18之1", "33條1項03款"]
}

def map_unit_name(raw_name):
    u_map = {'交通分隊': '交通分隊', '聖亭': '聖亭所', '龍潭': '龍潭所', '中興': '中興所', '石門': '石門所', '高平': '高平所', '三和': '三和所'}
    for key, val in u_map.items():
        if key in str(raw_name): return val
    return None

# ==========================================
# 1. 介面導覽
# ==========================================
st.title("🚓 龍潭分局交通數據戰情室")
st.header("🚦 (一) 加強交通安全執法取締五項交通違規統計表")
st.info("請在此執行原本的五項違規分析流程...")
# [保留原本代碼...]

st.markdown("<br><hr style='border:2px solid #ddd'><br>", unsafe_allow_html=True)

# ==========================================
# 2. 強化專案功能 (對應新檔案格式)
# ==========================================
st.header(f"📈 (二) {PROJECT_NAME}")

col_f1, col_f2 = st.columns(2)
file_statute = col_f1.file_uploader("1. 上傳『法條件數統計報表』", type="csv", key="f_stat")
file_type = col_f2.file_uploader("2. 上傳『案件類型-件數統計報表』", type="csv", key="f_type")

if file_statute and file_type:
    # 讀取法條數據 (skip 3 rows)
    df_s = pd.read_csv(file_statute, skiprows=3)
    df_s.columns = [str(c).strip() for c in df_s.columns]
    
    # 讀取案件類型數據 (用於抓取行人違規)
    df_t = pd.read_csv(file_type, skiprows=3)
    df_t.columns = [str(c).strip() for c in df_t.columns]

    results = []
    for unit, targets in TARGET_CONFIG.items():
        # 匹配單位
        row_s = df_s[df_s['單位'].apply(map_unit_name) == unit]
        row_t = df_t[df_t['單位'].apply(map_unit_name) == unit]
        
        unit_data = [unit]
        for i, cat in enumerate(CATS):
            count = 0
            if cat == "行人違規":
                # 從『案件類型』中抓取『行人攤販』欄位
                if not row_t.empty and '行人攤販' in df_t.columns:
                    count = row_t['行人攤販'].values[0]
            else:
                # 從『法條報表』中抓取對應法條
                keywords = LAW_MAP[cat]
                matched_cols = [c for c in df_s.columns if any(k in c for k in keywords)]
                count = row_s[matched_cols].sum(axis=1).values[0] if not row_s.empty else 0
            
            target = targets[i]
            ratio = f"{(count/target*100):.0f}%" if target > 0 else "0%"
            unit_data.extend([int(count), target, ratio])
        results.append(unit_data)

    # 建立統計表
    cols_header = ["單位"]
    for cat in CATS:
        cols_header.extend([f"{cat}_取締", f"{cat}_目標", f"{cat}_達成率"])
    df_final = pd.DataFrame(results, columns=cols_header)
    
    # 計算合計
    totals = ["合計"]
    for i in range(1, len(cols_header), 3):
        c_sum = df_final.iloc[:, i].sum()
        t_sum = df_final.iloc[:, i+1].sum()
        r_sum = f"{(c_sum/t_sum*100):.0f}%" if t_sum > 0 else "0%"
        totals.extend([int(c_sum), int(t_sum), r_sum])
    df_final = pd.concat([pd.DataFrame([totals], columns=cols_header), df_final]).reset_index(drop=True)

    st.subheader("📊 專案數據彙整表")
    st.dataframe(df_final, use_container_width=True)

    # 同步按鈕
    if st.button("🚀 同步至 Google Sheets (大型車違規更新版)"):
        try:
            gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
            sh = gc.open_by_url(GOOGLE_SHEET_URL)
            ws = sh.worksheet(PROJECT_NAME)
            
            now_dt = datetime.now().strftime("%Y-%m-%d %H:%M")
            h1 = [f"{PROJECT_NAME} (更新：{now_dt})"] + [""] * 18
            h2 = [""] + [c for c in CATS for _ in range(3)]
            h3 = ["單位"] + ["取締", "目標", "達成率"] * 6
            ws.clear()
            ws.update(values=[h1, h2, h3] + df_final.values.tolist())
            st.success("✅ 大型車違規數據已成功同步！")
            st.balloons()
        except Exception as e:
            st.error(f"同步失敗：{e}")
