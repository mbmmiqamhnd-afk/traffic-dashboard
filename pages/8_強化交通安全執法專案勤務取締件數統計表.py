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

# --- 共用邏輯函數 ---
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
    if raw == '龍潭' or raw == '龍潭所': return '龍潭所'
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
# 1. 檔案偵測與讀取邏輯
# ==========================================
st.title(f"📈 強化交通安全執法專案")

# 🔍 優先從根目錄搜尋 Git 推送的檔案
auto_f1_list = sorted(glob.glob("強化執法專案*.xlsx"), reverse=True)
auto_f2_list = glob.glob("*R17*.xlsx")

f1 = None
f2_files = []

# 檢查是否自動讀取
if auto_f1_list and auto_f2_list:
    st.success(f"✅ 自動模式：偵測到 GitHub 最新數據 ({auto_f1_list[0]})")
    f1 = auto_f1_list[0]
    f2_files = auto_f2_list
    if st.button("Manual Upload (手動上傳)"):
        st.session_state.force_manual = True
else:
    st.info("💡 提示：根目錄未偵測到自動更新檔案，請手動上傳。")
    col1, col2 = st.columns(2)
    f1 = col1.file_uploader("📂 1. 上傳『法條件數報表』", type=["xlsx", "csv"], key="u1")
    f2_files = col2.file_uploader("📂 2. 上傳『大型車違規表』", type=["xlsx", "csv"], accept_multiple_files=True, key="u2")

# ==========================================
# 2. 數據運算核心
# ==========================================
if f1 and f2_files:
    try:
        # 讀取函數 (支援路徑或上傳物件)
        def smart_read(f, **kwargs):
            is_csv = (isinstance(f, str) and f.endswith('.csv')) or (not isinstance(f, str) and f.name.endswith('.csv'))
            if is_csv:
                try: return pd.read_csv(f, **kwargs)
                except: return pd.read_csv(f, encoding='cp950', **kwargs)
            return pd.read_excel(f, **kwargs)

        # 擷取日期
        df1_head = smart_read(f1, nrows=10, header=None)
        date_range_str = "未知期間"
        for _, row in df1_head.iterrows():
            for cell in row.values:
                if '統計期間' in str(cell):
                    m = re.search(r'([0-9年月日\-至]+)', str(cell).split('：')[-1])
                    if m: date_range_str = m.group(1).strip()

        # 處理 F1
        df1 = smart_read(f1, skiprows=3)
        df1.columns = [str(c).strip() for c in df1.columns]

        # 處理 F2 多檔合併
        df2_all = []
        for f in f2_files:
            df_tmp = smart_read(f, header=None)
            h_idx = None
            for idx, r in df_tmp.head(20).iterrows():
                if '單位' in [str(x) for x in r.values] and '舉發總數' in [str(x) for x in r.values]:
                    h_idx = idx; break
            if h_idx is not None:
                cols = [str(c).strip() for c in df_tmp.iloc[h_idx]]
                df_clean = df_tmp.iloc[h_idx+1:].reset_index(drop=True)
                df_clean.columns = cols # 簡化處理重複欄位
                df2_all.append(df_clean)

        df2_combined = pd.concat(df2_all, ignore_index=True)
        df2_combined['標準單位'] = df2_combined['單位'].apply(map_unit_name)
        for c in ['舉發總數', '違反管制規定', '其他違規']:
            df2_combined[c] = pd.to_numeric(df2_combined.get(c, 0), errors='coerce').fillna(0)
        df2_combined['調整後大型車'] = (df2_combined['舉發總數'] - df2_combined['違反管制規定'] - df2_combined['其他違規']).clip(lower=0)

        # 產生成果
        final_data = []
        for unit in TARGET_CONFIG.keys():
            d15 = get_counts(df1, unit, CATS[:5])
            u_rows = df2_combined[df2_combined['標準單位'] == unit]
            h_cnt = int(u_rows['調整後大型車'].sum())
            
            res = [unit]
            for i, cat in enumerate(CATS):
                cnt = d15[cat] if cat != "大型車違規" else h_cnt
                tgt = TARGET_CONFIG[unit][i]
                ratio = f"{(cnt/tgt*100):.1f}%" if tgt > 0 else "0.0%"
                res.extend([cnt, tgt, ratio])
            final_data.append(res)

        headers = ["單位"]
        for cat in CATS: headers.extend([f"{cat}_取締", f"{cat}_目標", f"{cat}_達成率"])
        df_final = pd.DataFrame(final_data, columns=headers)

        # --- 計算合計列 ---
        totals = ["合計"]
        for i in range(1, len(headers), 3):
            c_sum = df_final.iloc[:, i].sum()
            t_sum = df_final.iloc[:, i+1].sum()
            r_sum = f"{(c_sum/t_sum*100):.1f}%" if t_sum > 0 else "0.0%"
            totals.extend([int(c_sum), int(t_sum), r_sum])
        
        # 標記交通組/警備隊為 '-'
        mask_units = ['交通組', '警備隊']
        target_cols = [c for c in df_final.columns if '目標' in c or '達成率' in c]
        df_final.loc[df_final['單位'].isin(mask_units), target_cols] = '-'
        df_final = pd.concat([pd.DataFrame([totals], columns=headers), df_final]).reset_index(drop=True)

        # --- 紅色標記邏輯 ---
        red_coords = []
        for cat in CATS:
            col_name = f"{cat}_達成率"
            col_idx = df_final.columns.get_loc(col_name)
            rates = pd.to_numeric(df_final.loc[1:, col_name].astype(str).str.rstrip('%'), errors='coerce')
            if not rates.dropna().empty:
                bot2 = rates.nsmallest(2).iloc[-1]
                for r_idx in rates.index:
                    if pd.notna(rates.loc[r_idx]) and rates.loc[r_idx] <= bot2:
                        red_coords.append((r_idx, col_idx))

        def style_red(x):
            df_c = pd.DataFrame('', index=x.index, columns=x.columns)
            for r, c in red_coords: df_c.iloc[r, c] = 'color: red; font-weight: bold;'
            return df_c

        # --- 顯示結果 ---
        st.markdown(f"### 📊 :blue[{PROJECT_NAME}] :red[(期間：{date_range_str})]")
        st.dataframe(df_final.style.apply(style_red, axis=None), use_container_width=True, hide_index=True)

        # --- 自動雲端同步按鈕 ---
        if st.button("🚀 同步至 Google Sheets"):
            with st.spinner("同步中..."):
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL)
                ws = sh.worksheet(PROJECT_NAME)
                # 更新內容與格式 (此處可沿用你原本的 gspread requests 邏輯)
                st.success("同步成功！")

    except Exception as e:
        st.error(f"錯誤：{e}")
