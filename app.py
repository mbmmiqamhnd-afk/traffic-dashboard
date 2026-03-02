import streamlit as st
import pandas as pd
import gspread
from datetime import datetime
import io

# ==========================================
# 0. 基本配置與專案目標值
# ==========================================
st.set_page_config(page_title="龍潭分局交通戰情室", page_icon="🚓", layout="wide")

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"

# 專案各單位目標值
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
    row = df[df['單位'].apply(map_unit_name) == unit]
    counts = {}
    for cat in categories_list:
        keywords = LAW_MAP[cat]
        matched_cols = [c for c in df.columns if any(k in c for k in keywords)]
        counts[cat] = int(row[matched_cols].sum(axis=1).values[0]) if not row.empty else 0
    return counts

# ==========================================
# 1. 側邊欄導覽 (依 Pages 順序)
# ==========================================
with st.sidebar:
    st.title("🚓 功能導覽")
    choice = st.radio(
        "功能選單",
        [
            "1️⃣ 📊 交通事故統計",
            "2️⃣ 🚔 取締重大交通違規統計",
            "3️⃣ 🚛 超載統計",
            "4️⃣ 🚦 加強交通安全執法取締五項交通違規統計表",
            "5️⃣ 🛡️ 科技執法成效",
            "6️⃣ 📈 " + PROJECT_NAME,
            "7️⃣ 🏷️ 商標頁碼工具",
            "8️⃣ 📄 PDF轉檔工具"
        ],
        index=5,  # 預設直接顯示第 6 項：強化專案統計
        label_visibility="collapsed"
    )

# ==========================================
# 2. 獨立功能頁面邏輯
# ==========================================

# --- 第 6 頁：強化交通安全執法專案 ---
if "6️⃣" in choice:
    st.title(PROJECT_NAME)
    
    # 兩份檔案上傳區 (均為法條件數統計報表)
    col1, col2 = st.columns(2)
    f1 = col1.file_uploader("📂 上傳第一份報表 (統計前五項)", type=["csv", "xlsx"], key="file_1_5")
    f2 = col2.file_uploader("📂 上傳第二份報表 (統計大型車)", type=["csv", "xlsx"], key="file_6")

    if f1 and f2:
        # 讀取數據
        df1 = pd.read_csv(f1, skiprows=3) if f1.name.endswith('.csv') else pd.read_excel(f1, skiprows=3)
        df2 = pd.read_csv(f2, skiprows=3) if f2.name.endswith('.csv') else pd.read_excel(f2, skiprows=3)
        df1.columns = [str(c).strip() for c in df1.columns]
        df2.columns = [str(c).strip() for c in df2.columns]

        final_rows = []
        for unit, targets in TARGET_CONFIG.items():
            # 分別從兩份檔案提取數據
            data1 = get_counts(df1, unit, CATS[:5])
            data2 = get_counts(df2, unit, [CATS[5]])
            all_counts = {**data1, **data2}
            
            unit_row = [unit]
            for i, cat in enumerate(CATS):
                cnt = all_counts[cat]
                tgt = targets[i]
                ratio = f"{(cnt/tgt*100):.0f}%" if tgt > 0 else "0%"
                unit_row.extend([cnt, tgt, ratio])
            final_rows.append(unit_row)

        headers = ["單位"]
        for cat in CATS: headers.extend([f"{cat}_取締", f"{cat}_目標", f"{cat}_達成率"])
        df_final = pd.DataFrame(final_rows, columns=headers)

        # 合計列
        totals = ["合計"]
        for i in range(1, len(headers), 3):
            c_sum = df_final.iloc[:, i].sum()
            t_sum = df_final.iloc[:, i+1].sum()
            r_sum = f"{(c_sum/t_sum*100):.0f}%" if t_sum > 0 else "0%"
            totals.extend([int(c_sum), int(t_sum), r_sum])
        df_final = pd.concat([pd.DataFrame([totals], columns=headers), df_final]).reset_index(drop=True)

        # 顯示統計結果表格
        st.subheader("📊 取締件數與目標值達成率統計")
        st.dataframe(df_final, use_container_width=True, hide_index=True)

        # 雲端同步按鈕
        if st.button("🚀 將統計結果同步至雲端試算表", use_container_width=True):
            try:
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL)
                ws = sh.worksheet(PROJECT_NAME)
                
                # 準備寫入內容
                now = datetime.now().strftime("%Y-%m-%d %H:%M")
                h1 = [f"{PROJECT_NAME} (同步日期：{now})"] + [""] * 18
                h2 = [""] + [c for c in CATS for _ in range(3)]
                h3 = ["單位"] + ["取締", "目標", "達成率"] * 6
                
                ws.clear()
                ws.update(values=[h1, h2, h3] + df_final.values.tolist())
                st.success("✅ 已完成雲端數據更新！")
                st.balloons()
            except Exception as e:
                st.error(f"同步過程中發生錯誤：{e}")

# --- 其他功能頁面 (預留空間) ---
else:
    st.title(choice)
    st.info("此模組正在等待嵌入相關功能代碼...")
