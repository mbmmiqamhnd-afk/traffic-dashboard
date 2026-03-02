import streamlit as st
import pandas as pd
import gspread
from datetime import datetime
import io

# ==========================================
# 0. 基本設定與專案目標值 (維持之前設定)
# ==========================================
st.set_page_config(page_title="龍潭分局交通戰情室", page_icon="🚓", layout="wide")

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"

# 單位與目標值 (酒駕, 闖紅燈, 嚴重超速, 車不讓人, 行人違規, 大型車違規)
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
# 1. 側邊欄導覽選單 (依序顯示)
# ==========================================
with st.sidebar:
    st.title("🚓 交通戰情室")
    choice = st.radio(
        "功能選單",
        [
            "🏠 系統首頁",
            "📊 交通事故統計",
            "🚔 取締重大交通違規統計",
            "🚛 超載統計",
            "🚦 加強交通安全執法取締五項交通違規統計表",
            "🛡️ 科技執法成效",
            "📈 " + PROJECT_NAME,
            "🏷️ 商標頁碼工具",
            "📄 PDF轉檔工具"
        ]
    )
    st.markdown("---")
    st.caption(f"最後更新：{datetime.now().strftime('%Y-%m-%d')}")

# ==========================================
# 2. 主畫面功能切換
# ==========================================

# --- 系統首頁 ---
if choice == "🏠 系統首頁":
    st.title("🚓 桃園市政府警察局龍潭分局")
    st.subheader("交通數據自動化統計系統")
    st.markdown("""
    請從左側選單選擇特定功能進行數據分析與雲端同步。
    * **系統狀態**：🟢 正常
    * **雲端連結**：[點此開啟試算表]({0})
    """.format(GOOGLE_SHEET_URL))

# --- 強化專案功能 (您的核心需求) ---
elif choice == "📈 " + PROJECT_NAME:
    st.header(f"📊 {PROJECT_NAME}")
    st.write("說明：前五項依據第一份報表，第六項(大型車)依據第二份報表。")
    
    col1, col2 = st.columns(2)
    file1 = col1.file_uploader("1. 上傳第一份報表 (前五項)", type=["csv", "xlsx"], key="p_f1")
    file2 = col2.file_uploader("2. 上傳第二份報表 (大型車)", type=["csv", "xlsx"], key="p_f2")

    if file1 and file2:
        df1 = pd.read_csv(file1, skiprows=3) if file1.name.endswith('.csv') else pd.read_excel(file1, skiprows=3)
        df2 = pd.read_csv(file2, skiprows=3) if file2.name.endswith('.csv') else pd.read_excel(file2, skiprows=3)
        df1.columns = [str(c).strip() for c in df1.columns]
        df2.columns = [str(c).strip() for c in df2.columns]

        final_rows = []
        for unit, targets in TARGET_CONFIG.items():
            data1 = get_counts(df1, unit, CATS[:5])
            data2 = get_counts(df2, unit, [CATS[5]])
            all_data = {**data1, **data2}
            
            unit_row = [unit]
            for i, cat in enumerate(CATS):
                cnt = all_data[cat]
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

        st.dataframe(df_final, use_container_width=True)

        if st.button("🚀 同步整合數據至 Google Sheets"):
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
                st.success("✅ 同步成功！")
                st.balloons()
            except Exception as e:
                st.error(f"同步失敗：{e}")

# --- 其他功能預留位置 (可自行嵌入原有代碼) ---
else:
    st.header(choice)
    st.info(f"此模組 {choice} 正在整合中，請將對應的程式碼放入邏輯區塊。")
