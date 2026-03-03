import streamlit as st
import pandas as pd
import gspread
from datetime import datetime

# ==========================================
# 0. 頁面配置與目標值設定
# ==========================================
st.set_page_config(page_title="強化專案統計 - 龍潭分局", layout="wide")

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"

# 單位目標值 (酒駕, 闖紅燈, 嚴重超速, 車不讓人, 行人違規, 大型車違規)
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

# 法條定義邏輯 (僅保留前五項，大型車改依新報表統計)
LAW_MAP = {
    "酒後駕車": ["35條", "73條2項", "73條3項"],
    "闖紅燈": ["53條"],
    "嚴重超速": ["43條", "40條"],
    "車不讓人": ["44條", "48條"],
    "行人違規": ["78條"]
}

def map_unit_name(raw_name):
    """單位名稱標準化 (排除科技執法與警備隊)"""
    u_map = {
        '交通分隊': '交通分隊', '龍潭交通分隊': '交通分隊',
        '聖亭': '聖亭所', '聖亭派出所': '聖亭所',
        '龍潭': '龍潭所', '龍潭派出所': '龍潭所',
        '中興': '中興所', '中興派出所': '中興所',
        '石門': '石門所', '石門派出所': '石門所',
        '高平': '高平所', '高平派出所': '高平所',
        '三和': '三和所', '三和派出所': '三和所'
    }
    for key, val in u_map.items():
        if key in str(raw_name): return val
    return None

def get_counts(df, unit, categories_list):
    """從第一份資料表(法條統計)中加總特定法條的件數"""
    row = df[df['單位'].apply(map_unit_name) == unit]
    counts = {}
    for cat in categories_list:
        keywords = LAW_MAP.get(cat, [])
        matched_cols = [c for c in df.columns if any(k in str(c) for k in keywords)]
        counts[cat] = int(row[matched_cols].sum(axis=1).values[0]) if not row.empty else 0
    return counts

# ==========================================
# 1. 畫面顯示與檔案上傳 (雙檔案分流邏輯)
# ==========================================
st.title(f"📈 {PROJECT_NAME}")

col1, col2 = st.columns(2)
f1 = col1.file_uploader("📂 1. 上傳『法條件數報表』\n(統計前5項)", type=["csv", "xlsx"], key="f_top5")
f2 = col2.file_uploader("📂 2. 上傳『大型車輛違規績效統計表』\n(統計大型車)", type=["csv", "xlsx"], key="f_heavy")

if f1 and f2:
    try:
        # --- 處理第一份報表 (固定跳過前3行) ---
        df1 = pd.read_csv(f1, skiprows=3) if f1.name.endswith('.csv') else pd.read_excel(f1, skiprows=3)
        df1.columns = [str(c).strip() for c in df1.columns]

        # --- 處理第二份報表 (動態尋找標頭列，因行數不固定) ---
        df2 = pd.read_csv(f2, header=None) if f2.name.endswith('.csv') else pd.read_excel(f2, header=None)
        
        header_idx_2 = None
        for idx, row in df2.head(20).iterrows():
            row_vals = [str(x).strip() for x in row.values]
            # 尋找同時包含「單位」與「舉發總數」的那一行作為欄位標題
            if '單位' in row_vals and '舉發總數' in row_vals:
                header_idx_2 = idx
                break

        if header_idx_2 is None:
            st.error("❌ 在第二份報表中找不到『單位』與『舉發總數』欄位，請確認上傳了正確的檔案！")
            st.stop()
            
        cols = [str(c).strip() for c in df2.iloc[header_idx_2]]
        
        # 處理新報表中常見的「重複欄位名稱」(如兩個 '單位' 欄位)
        seen = {}
        new_cols = []
        for c in cols:
            if c in seen:
                seen[c] += 1
                new_cols.append(f"{c}.{seen[c]}")
            else:
                seen[c] = 0
                new_cols.append(c)
                
        df2_clean = df2.iloc[header_idx_2+1:].reset_index(drop=True)
        df2_clean.columns = new_cols
        
        df2_clean['標準單位'] = df2_clean['單位'].apply(map_unit_name)
        df2_clean['舉發總數'] = pd.to_numeric(df2_clean['舉發總數'], errors='coerce').fillna(0)

        # --- 開始彙整兩份檔案的數據 ---
        final_results = []
        for unit in TARGET_CONFIG.keys():
            # 1. 提取前五項 (來自 df1)
            data_1to5 = get_counts(df1, unit, CATS[:5])
            
            # 2. 提取第六項 (來自 df2_clean 的『舉發總數』)
            unit_rows = df2_clean[df2_clean['標準單位'] == unit]
            heavy_count = int(unit_rows['舉發總數'].sum()) if not unit_rows.empty else 0
            
            all_c = {**data_1to5, CATS[5]: heavy_count}
            
            # 3. 計算達成率
            unit_row = [unit]
            for i, cat in enumerate(CATS):
                cnt = all_c[cat]
                tgt = TARGET_CONFIG[unit][i]
                ratio = f"{(cnt/tgt*100):.1f}%" if tgt > 0 else "0%"
                unit_row.extend([cnt, tgt, ratio])
            final_results.append(unit_row)

        # 建立 DataFrame
        headers = ["單位"]
        for cat in CATS:
            headers.extend([f"{cat}_取締", f"{cat}_目標", f"{cat}_達成率"])
        df_final = pd.DataFrame(final_results, columns=headers)

        # 計算最下方的合計列
        totals = ["合計"]
        for i in range(1, len(headers), 3):
            c_sum = df_final.iloc[:, i].sum()
            t_sum = df_final.iloc[:, i+1].sum()
            r_sum = f"{(c_sum/t_sum*100):.1f}%" if t_sum > 0 else "0%"
            totals.extend([int(c_sum), int(t_sum), r_sum])
        
        df_final = pd.concat([pd.DataFrame([totals], columns=headers), df_final]).reset_index(drop=True)

        # 顯示結果
        st.subheader("📊 雙檔案整合分析結果")
        st.dataframe(df_final, use_container_width=True, hide_index=True)

        # ==========================================
        # 2. 雲端同步功能
        # ==========================================
        if st.button("🚀 同步數據至雲端試算表", use_container_width=True):
            with st.spinner("同步中，請稍候..."):
                try:
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    
                    try:
                        ws = sh.worksheet(PROJECT_NAME)
                    except:
                        ws = sh.add_worksheet(title=PROJECT_NAME, rows=50, cols=20)
                    
                    now = datetime.now().strftime("%Y-%m-%d %H:%M")
                    h1 = [f"{PROJECT_NAME} (同步時間：{now})"] + [""] * 18
                    h2 = [""] + [c for c in CATS for _ in range(3)]
                    h3 = ["單位"] + ["取締件數", "目標值", "達成率"] * 6
                    
                    ws.clear()
                    ws.update(values=[h1, h2, h3] + df_final.values.tolist())
                    
                    # 合併表頭單元格
                    requests = [{"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 19}, "mergeType": "MERGE_ALL"}}]
                    for j in range(1, 19, 3):
                        requests.append({"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": j, "endColumnIndex": j+3}, "mergeType": "MERGE_ALL"}})
                    
                    sh.batch_update({"requests": requests})
                    st.success("✅ 數據已成功同步並自動格式化！")
                    st.balloons()
                except Exception as e:
                    st.error(f"雲端連線失敗：{e}")

    except Exception as e:
        st.error(f"處理檔案時發生錯誤：{e}")
