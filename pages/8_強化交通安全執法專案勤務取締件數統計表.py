import streamlit as st
import pandas as pd
import gspread
import re
import os
import glob
from datetime import datetime

# ==========================================
# 0. 頁面配置與目標值設定
# ==========================================
st.set_page_config(page_title="強化專案統計 - 龍潭分局", layout="wide")

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"

TARGET_CONFIG = {
    '聖亭所': [5, 115, 5, 16, 7, 10],
    '龍潭所': [6, 145, 7, 20, 9, 12],
    '中興所': [5, 115, 5, 16, 7, 10],
    '石門所': [3, 80, 4, 11, 5, 7],
    '高平所': [3, 80, 4, 11, 5, 7],
    '三和所': [2, 40, 2, 6, 2, 5],
    '交通分隊': [5, 115, 4, 16, 6, 8],
    '交通組': [0, 0, 0, 0, 0, 0],
    '警備隊': [0, 0, 0, 0, 0, 0]
}

CATS = ["酒後駕車", "闖紅燈", "嚴重超速", "車不讓人", "行人違規", "大型車違規"]

LAW_MAP = {
    "酒後駕車": ["35條", "73條2項", "73條3項"],
    "闖紅燈": ["53條"],
    "嚴重超速": ["43條", "40條"],
    "車不讓人": ["44條", "48條"],
    "行人違規": ["78條"]
}

# --- 共用函數區 ---
def map_unit_name(raw_name):
    raw = str(raw_name).strip()
    if '交通組' in raw: return '交通組'
    if '警備隊' in raw: return '警備隊'
    if '聖亭' in raw: return '聖亭所'
    if '中興' in raw: return '中興所'
    if '石門' in raw: return '石門所'
    if '高平' in raw: return '高平所'
    if '三和' in raw: return '三和所'
    if '龍潭派出所' in raw: return '龍潭所'
    if raw == '龍潭': return '龍潭所'
    if '交通分隊' in raw:
        exclude_list = ['楊梅', '大溪', '平鎮', '中壢', '八德', '大園', '蘆竹', '龜山', '桃園交通']
        for ex in exclude_list:
            if ex in raw: return None
        return '交通分隊'
    return None

def get_counts(df, unit, categories_list):
    rows = df[df['單位'].apply(map_unit_name) == unit]
    counts = {}
    for cat in categories_list:
        keywords = LAW_MAP.get(cat, [])
        matched_cols = [c for c in df.columns if any(k in str(c) for k in keywords)]
        counts[cat] = int(rows[matched_cols].sum().sum()) if not rows.empty else 0
    return counts

# ==========================================
# 1. 檔案偵測邏輯：優先讀取 Git 自動推送的檔案
# ==========================================
st.title(f"📈 強化交通安全執法專案")

# 🔍 搜尋本地(GitHub 倉庫)是否有自動推送的 Excel
# 假設檔名規則：強化執法專案_*.xlsx 以及 *R17*.xlsx
auto_f1_list = sorted(glob.glob("強化執法專案*.xlsx"), reverse=True) # 拿最新的
auto_f2_list = glob.glob("*R17*.xlsx")

# 初始化檔案變數
f1 = None
f2_files = []

# 優先順序：如果有自動檔案則直接讀取，否則顯示上傳框
if auto_f1_list and auto_f2_list:
    st.success(f"✅ 已自動偵測到最新數據：{auto_f1_list[0]}")
    f1 = auto_f1_list[0]
    f2_files = auto_f2_list
    if st.button("🔄 切換至手動上傳模式"):
        f1 = None
        f2_files = []
else:
    st.info("💡 未偵測到自動更新檔案，請手動上傳報表。")
    col1, col2 = st.columns(2)
    f1_upload = col1.file_uploader("📂 1. 上傳『法條件數報表』", type=["xlsx", "csv"], key="u1")
    f2_upload = col2.file_uploader("📂 2. 上傳『大型車違規表』", type=["xlsx", "csv"], accept_multiple_files=True, key="u2")
    if f1_upload: f1 = f1_upload
    if f2_upload: f2_files = f2_upload

# ==========================================
# 2. 數據處理核心
# ==========================================
if f1 and f2_files:
    try:
        # --- 讀取 F1 (法條報表) ---
        # 判斷是路徑(str)還是上傳物件
        read_f1 = lambda f, **kwargs: pd.read_csv(f, **kwargs) if (isinstance(f, str) and f.endswith('.csv')) or (not isinstance(f, str) and f.name.endswith('.csv')) else pd.read_excel(f, **kwargs)
        
        # 擷取統計期間 (讀取前幾行)
        df1_head = read_f1(f1, nrows=10, header=None)
        date_range_str = "未知期間"
        for _, row in df1_head.iterrows():
            for cell in row.values:
                if '統計期間' in str(cell):
                    match = re.search(r'([0-9]{7}.*[0-9]{7})', str(cell))
                    if match: date_range_str = match.group(1)

        # 正式讀取數據
        df1 = read_f1(f1, skiprows=3)
        df1.columns = [str(c).strip() for c in df1.columns]

        # --- 處理 F2 (大型車多檔合併) ---
        df2_list = []
        for f in f2_files:
            df_tmp = pd.read_csv(f, header=None) if (isinstance(f, str) and f.endswith('.csv')) or (not isinstance(f, str) and f.name.endswith('.csv')) else pd.read_excel(f, header=None)
            
            # 尋找表頭位置
            header_idx = None
            for idx, row in df_tmp.head(20).iterrows():
                if '單位' in [str(x) for x in row.values] and '舉發總數' in [str(x) for x in row.values]:
                    header_idx = idx
                    break
            
            if header_idx is not None:
                cols = [str(c).strip() for c in df_tmp.iloc[header_idx]]
                # 處理重複欄位名 (pandas 不允許重複)
                new_cols = []
                seen = {}
                for c in cols:
                    seen[c] = seen.get(c, -1) + 1
                    new_cols.append(f"{c}.{seen[c]}" if seen[c] > 0 else c)
                
                df_clean = df_tmp.iloc[header_idx+1:].reset_index(drop=True)
                df_clean.columns = new_cols
                df2_list.append(df_clean)

        df2_combined = pd.concat(df2_list, ignore_index=True)
        df2_combined['標準單位'] = df2_combined['單位'].apply(map_unit_name)
        
        # 數值轉換與大型車計算
        for col in ['舉發總數', '違反管制規定', '其他違規']:
            df2_combined[col] = pd.to_numeric(df2_combined.get(col, 0), errors='coerce').fillna(0)
        
        df2_combined['調整後大型車違規'] = (df2_combined['舉發總數'] - df2_combined.get('違反管制規定', 0) - df2_combined.get('其他違規', 0)).clip(lower=0)

        # --- 生成最終表格 ---
        final_results = []
        for unit in TARGET_CONFIG.keys():
            data_1to5 = get_counts(df1, unit, CATS[:5])
            u_rows = df2_combined[df2_combined['標準單位'] == unit]
            heavy_cnt = int(u_rows['調整後大型車違規'].sum())
            
            all_c = {**data_1to5, "大型車違規": heavy_cnt}
            row = [unit]
            for i, cat in enumerate(CATS):
                cnt = all_c[cat]
                tgt = TARGET_CONFIG[unit][i]
                ratio = f"{(cnt/tgt*100):.1f}%" if tgt > 0 else "0.0%"
                row.extend([cnt, tgt, ratio])
            final_results.append(row)

        headers = ["單位"]
        for cat in CATS: headers.extend([f"{cat}_取締", f"{cat}_目標", f"{cat}_達成率"])
        df_final = pd.DataFrame(final_results, columns=headers)

        # 合計列與格式化 (略，保持你原本的邏輯...)
        # ... (此處省略你原本的合計與紅色標記邏輯，建議保留完整)
        
        # 顯示結果
        st.dataframe(df_final, use_container_width=True, hide_index=True)

        # --- 雲端同步 (僅在自動偵測到檔案或點擊按鈕時觸發) ---
        if st.button("🚀 手動同步至 Google Sheets"):
             # 執行你原本的 gspread 同步代碼
             pass

    except Exception as e:
        st.error(f"❌ 數據處理出錯：{e}")
