import streamlit as st
import pandas as pd
import gspread
import re
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
        # --- 動態擷取「統計期間」並格式化 ---
        f1.seek(0)
        try:
            df1_head = pd.read_csv(f1, nrows=10, header=None) if f1.name.endswith('.csv') else pd.read_excel(f1, nrows=10, header=None)
        except UnicodeDecodeError:
            f1.seek(0)
            df1_head = pd.read_csv(f1, nrows=10, header=None, encoding='cp950') if f1.name.endswith('.csv') else pd.DataFrame()
        except Exception:
            df1_head = pd.DataFrame()
            
        date_range_str = "未知期間"
        if not df1_head.empty:
            for _, row in df1_head.iterrows():
                for cell in row.values:
                    cell_str = str(cell)
                    if '統計期間' in cell_str:
                        # 自動擷取日期字串
                        match = re.search(r'統計期間.*?[：:](?:\s*\(入案日\))?\s*([0-9年月日\-至]+)', cell_str)
                        if match:
                            date_raw = match.group(1).strip()
                            parts = date_raw.split('至') if '至' in date_raw else date_raw.split('-')
                            clean_parts = []
                            for p in parts:
                                p = p.strip()
                                if '年' in p:
                                    p = p.split('年')[-1]
                                elif len(p) == 7 and p.isdigit():
                                    p = p[3:]
                                clean_parts.append(p)
                            date_range_str = "-".join(clean_parts)
        f1.seek(0)

        # --- 處理第一份報表 (固定跳過前3行) ---
        try:
            df1 = pd.read_csv(f1, skiprows=3) if f1.name.endswith('.csv') else pd.read_excel(f1, skiprows=3)
        except UnicodeDecodeError:
            f1.seek(0)
            df1 = pd.read_csv(f1, skiprows=3, encoding='cp950') if f1.name.endswith('.csv') else pd.DataFrame()
            
        df1.columns = [str(c).strip() for c in df1.columns]

        # --- 處理第二份報表 (動態尋找標頭列，因行數不固定) ---
        df2 = pd.read_csv(f2, header=None) if f2.name.endswith('.csv') else pd.read_excel(f2, header=None)
        
        header_idx_2 = None
        for idx, row in df2.head(20).iterrows():
            row_vals = [str(x).strip() for x in row.values]
            if '單位' in row_vals and '舉發總數' in row_vals:
                header_idx_2 = idx
                break

        if header_idx_2 is None:
            st.error("❌ 在第二份報表中找不到『單位』與『舉發總數』欄位，請確認上傳了正確的檔案！")
            st.stop()
            
        cols = [str(c).strip() for c in df2.iloc[header_idx_2]]
        
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
            data_1to5 = get_counts(df1, unit, CATS[:5])
            
            unit_rows = df2_clean[df2_clean['標準單位'] == unit]
            heavy_count = int(unit_rows['舉發總數'].sum()) if not unit_rows.empty else 0
            
            all_c = {**data_1to5, CATS[5]: heavy_count}
            
            unit_row = [unit]
            for i, cat in enumerate(CATS):
                cnt = all_c[cat]
                tgt = TARGET_CONFIG[unit][i]
                ratio = f"{(cnt/tgt*100):.1f}%" if tgt > 0 else "0%"
                unit_row.extend([cnt, tgt, ratio])
            final_results.append(unit_row)

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

        # ==========================================
        # 👑 計算總達成率，並標記最後三名
        # ==========================================
        unit_overall_ratio = {}
        for unit in TARGET_CONFIG.keys():
            idx_list = df_final.index[df_final['單位'] == unit].tolist()
            if idx_list:
                idx = idx_list[0]
                total_c = df_final.loc[idx, [f"{cat}_取締" for cat in CATS]].sum()
                total_t = df_final.loc[idx, [f"{cat}_目標" for cat in CATS]].sum()
                unit_overall_ratio[unit] = total_c / total_t if total_t > 0 else 0

        # 取前三名 (達成率墊底的三名)
        bottom_3_units = sorted(unit_overall_ratio, key=unit_overall_ratio.get)[:3]

        def highlight_bottom(val):
            if val in bottom_3_units:
                return 'color: red; font-weight: bold;'
            return ''

        styled_df = df_final.style.map(highlight_bottom, subset=['單位'])

        # 顯示整合結果標題 (帶上日期)
        st.subheader(f"📊 雙檔案整合分析結果 (統計期間：{date_range_str})")
        st.markdown(f"*(提示：總達成率最後三名單位 **{', '.join(bottom_3_units)}** 已標示為紅色)*")
        st.dataframe(styled_df, use_container_width=True, hide_index=True)

        # ==========================================
        # 2. 雲端同步功能 (只覆寫數值，保護格式)
        # ==========================================
        if st.button("🚀 同步數據與顏色至雲端試算表", use_container_width=True):
            with st.spinner("同步數據與格式中，請稍候..."):
                try:
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    
                    try:
                        ws = sh.worksheet(PROJECT_NAME)
                    except:
                        ws = sh.add_worksheet(title=PROJECT_NAME, rows=50, cols=20)
                    
                    # 雲端表格首行文字更新為「統計期間」
                    h1 = [f"{PROJECT_NAME} (統計期間：{date_range_str})"] + [""] * 18
                    h2 = [""] + [c for c in CATS for _ in range(3)]
                    h3 = ["單位"] + ["取締件數", "目標值", "達成率"] * 6
                    
                    # 只更新數值內容，完全不清除既有的排版格式！
                    ws.update(values=[h1, h2, h3] + df_final.values.tolist())
                    
                    # 準備 Google Sheets 的字體變色請求
                    requests = []
                    
                    # 1. 將『單位』欄位全部初始化為黑色字體 (防呆機制，清除前一次的紅色)
                    requests.append({
                        "repeatCell": {
                            "range": {
                                "sheetId": ws.id,
                                "startRowIndex": 3,  # 標題佔了 0,1,2 行
                                "endRowIndex": 3 + len(df_final),
                                "startColumnIndex": 0,
                                "endColumnIndex": 1
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "textFormat": {
                                        "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0},
                                        "bold": False
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.textFormat(foregroundColor,bold)"
                        }
                    })

                    # 2. 將『最後三名』的儲存格標記為紅色粗體
                    for unit in bottom_3_units:
                        idx_list = df_final.index[df_final['單位'] == unit].tolist()
                        if idx_list:
                            # 加上表頭的 3 行偏移量
                            sheet_row_idx
