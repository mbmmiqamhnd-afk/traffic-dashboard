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

# --- 核心邏輯函數 ---
def map_unit_name(raw_name):
    raw = str(raw_name).strip()
    if '交通組' in raw: return '交通組'
    if '警備隊' in raw: return '警備隊'
    if '聖亭' in raw: return '聖亭所'
    if '中興' in raw: return '中興所'
    if '石門' in raw: return '石門所'
    if '高平' in raw: return '高平所'
    if '三和' in raw: return '三和所'
    if '龍潭派出所' in raw or raw == '龍潭' or raw == '龍潭所': return '龍潭所'
    if '交通分隊' in raw:
        exclude_list = ['楊梅', '大溪', '平鎮', '中壢', '八德', '大園', '蘆竹', '龜山', '桃園交通']
        for ex in exclude_list:
            if ex in raw: return None
        return '交通分隊'
    return None

def get_counts(df, unit, categories_list):
    df_clean = df.reset_index(drop=True)
    rows = df_clean[df_clean['單位'].apply(map_unit_name) == unit].copy()
    counts = {}
    for cat in categories_list:
        keywords = LAW_MAP.get(cat, [])
        matched_cols = [c for c in df_clean.columns if any(k in str(c) for k in keywords)]
        counts[cat] = int(rows[matched_cols].sum().sum()) if not rows.empty else 0
    return counts

# ==========================================
# 1. 檔案偵測邏輯
# ==========================================
st.title(f"📈 強化交通安全執法專案")

auto_f1_list = sorted(glob.glob("強化執法專案*.xlsx"), reverse=True)
auto_f2_list = glob.glob("*R17*.xlsx")

f1_active = None
f2_active_list = []

if auto_f1_list and len(auto_f2_list) >= 1:
    st.success(f"✅ 自動模式：偵測到最新報表 ({auto_f1_list[0]}) 與 {len(auto_f2_list)} 份大型車資料")
    f1_active = auto_f1_list[0]
    f2_active_list = auto_f2_list
    with st.expander("查看偵測檔案清單"):
        st.write(f"📄 法條報表：{f1_active}")
        st.write(f"🚛 大型車報表：{f2_active_list}")
else:
    st.info("💡 提示：未偵測到自動更新檔案，請手動上傳。")
    c1, c2 = st.columns(2)
    f1_active = c1.file_uploader("📂 1. 上傳『法條件數報表』", type=["xlsx", "csv"], key="manual_f1")
    f2_active_list = c2.file_uploader("📂 2. 上傳『大型車違規表』(支援多檔)", type=["xlsx", "csv"], accept_multiple_files=True, key="manual_f2")

# ==========================================
# 2. 數據處理核心
# ==========================================
if f1_active and f2_active_list:
    try:
        def smart_read(f, **kwargs):
            fname = f if isinstance(f, str) else f.name
            if fname.endswith('.csv'):
                try: return pd.read_csv(f, **kwargs)
                except: return pd.read_csv(f, encoding='cp950', **kwargs)
            return pd.read_excel(f, **kwargs)

        # A. 處理法條報表
        df1_raw = smart_read(f1_active, skiprows=3)
        df1_raw.columns = [str(c).strip() for c in df1_raw.columns]
        df1_raw = df1_raw.loc[:, ~df1_raw.columns.duplicated()].reset_index(drop=True)
        
        df1_date_check = smart_read(f1_active, nrows=10, header=None)
        date_range_str = "未知期間"
        for _, row in df1_date_check.iterrows():
            for cell in row.values:
                if '統計期間' in str(cell):
                    match = re.search(r'([0-9年月日\-至]+)', str(cell).split('：')[-1])
                    if match: date_range_str = match.group(1).strip()

        # B. 處理大型車報表
        df2_collector = []
        for f in f2_active_list:
            df_tmp = smart_read(f, header=None)
            h_idx = None
            for idx, r in df_tmp.head(20).iterrows():
                row_str = [str(x).strip() for x in r.values]
                if '單位' in row_str and '舉發總數' in row_str:
                    h_idx = idx; break
            if h_idx is not None:
                raw_cols = [str(c).strip() for c in df_tmp.iloc[h_idx]]
                new_cols = []
                counts = {}
                for c in raw_cols:
                    seen = counts.get(c, 0)
                    new_cols.append(f"{c}_{seen}" if seen > 0 else c)
                    counts[c] = seen + 1
                
                df_c = df_tmp.iloc[h_idx+1:].copy()
                df_c.columns = new_cols
                df_c = df_c.reset_index(drop=True)
                needed = ['單位', '舉發總數', '違反管制規定', '其他違規']
                existing = [c for c in needed if c in df_c.columns]
                df2_collector.append(df_c[existing])

        df2_all = pd.concat(df2_collector, ignore_index=True)
        df2_all = df2_all.loc[:, ~df2_all.columns.duplicated()].reset_index(drop=True)
        df2_all['標準單位'] = df2_all['單位'].apply(map_unit_name)
        
        for c in ['舉發總數', '違反管制規定', '其他違規']:
            df2_all[c] = pd.to_numeric(df2_all.get(c, 0), errors='coerce').fillna(0)
        df2_all['大型車純違規'] = (df2_all['舉發總數'] - df2_all.get('違反管制規定',0) - df2_all.get('其他違規',0)).clip(lower=0)

        # C. 彙整各單位數據
        final_rows = []
        for unit in TARGET_CONFIG.keys():
            d15 = get_counts(df1_raw, unit, CATS[:5])
            u_rows = df2_all[df2_all['標準單位'] == unit]
            heavy_sum = int(u_rows['大型車純違規'].sum())
            row = [unit]
            for i, cat in enumerate(CATS):
                cnt = d15[cat] if cat != "大型車違規" else heavy_sum
                tgt = TARGET_CONFIG[unit][i]
                ratio = f"{(cnt/tgt*100):.1f}%" if tgt > 0 else "0.0%"
                row.extend([cnt, tgt, ratio])
            final_rows.append(row)

        header_cols = ["單位"]
        for cat in CATS: header_cols.extend([f"{cat}_取締", f"{cat}_目標", f"{cat}_達成率"])
        df_final = pd.DataFrame(final_rows, columns=header_cols)

        # 合計列
        total_row = ["合計"]
        for i in range(1, len(header_cols), 3):
            c_s = df_final.iloc[:, i].sum()
            t_s = df_final.iloc[:, i+1].sum()
            r_s = f"{(c_s/t_s*100):.1f}%" if t_s > 0 else "0.0%"
            total_row.extend([int(c_s), int(t_s), r_s])
        
        mask_units = ['交通組', '警備隊']
        df_final.loc[df_final['單位'].isin(mask_units), [c for c in df_final.columns if '目標' in c or '達成率' in c]] = '-'
        df_final = pd.concat([pd.DataFrame([total_row], columns=header_cols), df_final], ignore_index=True)

        # 紅色標記邏輯
        red_coords = []
        for cat in CATS:
            c_name = f"{cat}_達成率"
            c_idx = df_final.columns.get_loc(c_name)
            vals = pd.to_numeric(df_final.loc[1:, c_name].astype(str).str.rstrip('%'), errors='coerce')
            if not vals.dropna().empty:
                limit = vals.nsmallest(2).iloc[-1]
                for r_idx in vals.index:
                    if pd.notna(vals.loc[r_idx]) and vals.loc[r_idx] <= limit:
                        red_coords.append((r_idx, c_idx))

        # ==========================================
        # 3. 畫面顯示與雲端同步 (恢復精美格式)
        # ==========================================
        st.markdown(f"### 📊 :blue[{PROJECT_NAME}] :red[(期間：{date_range_str})]")
        st.dataframe(df_final.style.apply(lambda x: [['color: red; font-weight: bold;' if (r, c) in red_coords else '' for c, _ in enumerate(x.columns)] for r, _ in enumerate(x.index)], axis=None), use_container_width=True, hide_index=True)

        if st.button("🚀 同步至雲端 Google Sheets"):
            with st.spinner("正在同步精美格式..."):
                try:
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    try: ws = sh.worksheet(PROJECT_NAME)
                    except: ws = sh.add_worksheet(title=PROJECT_NAME, rows=50, cols=20)
                    
                    # A. 準備三層表頭
                    title_text = f"{PROJECT_NAME} (統計期間：{date_range_str})"
                    h1 = [title_text] + [""] * 18
                    h2 = [""] + [c for c in CATS for _ in range(3)]
                    h3 = ["單位"] + ["取締件數", "目標值", "達成率"] * 6
                    
                    # B. 清空並寫入數值
                    ws.clear()
                    ws.update(values=[h1, h2, h3] + df_final.values.tolist())
                    
                    # C. 執行格式化 Request
                    reqs = []
                    # 合併第一列 (標題)
                    reqs.append({"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 19}, "mergeType": "MERGE_ALL"}})
                    # 雙色標題
                    reqs.append({"updateCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1}, "rows": [{"values": [{"userEnteredValue": {"stringValue": title_text}, "textFormatRuns": [{"startIndex": 0, "format": {"foregroundColor": {"blue": 1.0}, "bold": True}}, {"startIndex": len(PROJECT_NAME), "format": {"foregroundColor": {"red": 1.0}, "bold": True}}]}]}], "fields": "userEnteredValue,textFormatRuns"}})
                    # 紅色標記
                    for r, c in red_coords:
                        reqs.append({"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": r+3, "endRowIndex": r+4, "startColumnIndex": c, "endColumnIndex": c+1}, "cell": {"userEnteredFormat": {"textFormat": {"foregroundColor": {"red": 1.0}, "bold": True}}}, "fields": "userEnteredFormat.textFormat"}})
                    
                    sh.batch_update({"requests": reqs})
                    st.success("✅ 數據與格式已完美同步！")
                except Exception as sync_e:
                    st.error(f"同步失敗：{sync_e}")

    except Exception as e:
        st.error(f"❌ 數據解析錯誤：{e}")
