import streamlit as st
import pandas as pd
import gspread
import re
import os
from datetime import datetime

# ==========================================
# 0. 頁面配置與目標值設定
# ==========================================
st.set_page_config(page_title="強化專案統計 - 龍潭分局", layout="wide")

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"

TARGET_CONFIG = {
    '聖亭所': [5, 115, 5, 16, 7, 10], '龍潭所': [6, 145, 7, 20, 9, 12],
    '中興所': [5, 115, 5, 16, 7, 10], '石門所': [3, 80, 4, 11, 5, 7],
    '高平所': [3, 80, 4, 11, 5, 7], '三和所': [2, 40, 2, 6, 2, 5],
    '交通分隊': [5, 115, 4, 16, 6, 8], '交通組': [0, 0, 0, 0, 0, 0], '警備隊': [0, 0, 0, 0, 0, 0]
}

CATS = ["酒後駕車", "闖紅燈", "嚴重超速", "車不讓人", "行人違規", "大型車違規"]
LAW_MAP = {
    "酒後駕車": ["35條", "73條2項", "73條3項"], "闖紅燈": ["53條"],
    "嚴重超速": ["43條", "40條"], "車不讓人": ["44條", "48條"], "行人違規": ["78條"]
}

# --- 核心輔助函數 ---
def map_unit_name(raw_name):
    raw = str(raw_name).strip()
    if '交通分隊' in raw:
        if '龍潭' in raw: return '交通分隊'
        if not any(ex in raw for ex in ['楊梅', '大溪', '平鎮', '中壢', '八德', '蘆竹', '龜山', '大園', '桃園']):
            return '交通分隊'
    if '交通組' in raw: return '交通組'
    if '警備隊' in raw: return '警備隊'
    for k in ['聖亭', '中興', '石門', '高平', '三和']:
        if k in raw: return k + '所'
    if '龍潭派出所' in raw or raw in ['龍潭', '龍潭所']: return '龍潭所'
    return None

def make_columns_unique(df):
    cols = pd.Series(df.columns.map(str))
    for dup in cols[cols.duplicated()].unique():
        cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols
    return df

def get_counts(df, unit, categories_list):
    df_c = df.reset_index(drop=True)
    if '單位' not in df_c.columns: return {cat: 0 for cat in categories_list}
    rows = df_c[df_c['單位'].apply(map_unit_name) == unit].copy()
    counts = {}
    for cat in categories_list:
        keywords = LAW_MAP.get(cat, [])
        matched = [c for c in df_c.columns if any(k in str(c) for k in keywords)]
        counts[cat] = int(rows[matched].sum().sum()) if not rows.empty else 0
    return counts

# --- 1. 智慧上傳介面 (已移除 GitHub 自動偵測) ---
st.title(f"📈 強化交通安全執法專案")
st.markdown("##### 💡 操作說明：請將今天的「3 份 Excel 報表」全選後一次拖入下方框中即可。")

all_uploads = st.file_uploader("📂 請拖入所有報表（自動辨識法條、大型車大隊/分局報表）", 
                               type=["xlsx", "csv"], 
                               accept_multiple_files=True)

f1_active = None
f2_active = []

if all_uploads:
    for f in all_uploads:
        if "強化" in f.name:
            f1_active = f
            st.success(f"✔️ 已識別法條報表: **{f.name}**")
        elif "R17" in f.name:
            f2_active.append(f)
            st.success(f"✔️ 已識別大型車報表: **{f.name}**")

# --- 2. 數據處理核心 ---
if f1_active and len(f2_active) >= 1:
    try:
        def smart_read(f, **kwargs):
            fname = f.name
            if fname.endswith('.csv'):
                try:
                    return pd.read_csv(f, **kwargs)
                except:
                    return pd.read_csv(f, encoding='cp950', **kwargs)
            return pd.read_excel(f, **kwargs)

        # 抓日期
        date_range_str = "未知期間"
        df1_h = smart_read(f1_active, nrows=10, header=None)
        for _, r in df1_h.iterrows():
            for cell in r.values:
                if '統計期間' in str(cell):
                    raw = str(cell).replace('(入案日)', '').split('：')[-1].split(':')[-1].strip()
                    m = re.search(r'([0-9年月日\-至\s]+)', raw)
                    if m:
                        date_range_str = m.group(1).replace('115', '').strip()

        # 讀取 F1
        df1 = make_columns_unique(smart_read(f1_active, skiprows=3)).reset_index(drop=True)

        # 讀取 F2 (支援多個檔案)
        df2_list = []
        for f in f2_active:
            df_t = smart_read(f, header=None)
            h_idx = None
            for i, row in df_t.head(30).iterrows():
                row_vals = [str(x).strip() for x in row.values]
                if '單位' in row_vals and '舉發總數' in row_vals:
                    h_idx = i
                    break

            if h_idx is not None:
                df_c = df_t.iloc[h_idx+1:].copy()
                df_c.columns = [str(x).strip() for x in df_t.iloc[h_idx].values]
                df_c = make_columns_unique(df_c).reset_index(drop=True)
                needed = ['單位', '舉發總數', '違反管制規定', '其他違規']
                existing = [c for c in needed if c in df_c.columns]
                
                if '單位' in existing:
                    df_c['來源檔名'] = str(f.name)
                    df2_list.append(df_c)

        if not df2_list:
            st.error("❌ 大型車報表欄位抓取失敗")
            st.stop()

        df2_all = pd.concat(df2_list, ignore_index=True).reset_index(drop=True)

        # 數值化處理
        for c in ['舉發總數', '違反管制規定', '其他違規']:
            if c not in df2_all.columns: df2_all[c] = 0
            df2_all[c] = pd.to_numeric(df2_all[c], errors='coerce').fillna(0)

        df2_all['大型車純違規'] = (df2_all['舉發總數'] - df2_all['違反管制規定'] - df2_all['其他違規']).clip(lower=0)

        # 彙整數據
        final_rows = []
        for unit in TARGET_CONFIG.keys():
            d15 = get_counts(df1, unit, CATS[:5])

            if unit == '交通分隊':
                u_rows = df2_all[
                    (df2_all['來源檔名'].str.contains('大隊|交大', na=False)) &
                    (df2_all['單位'].str.contains('龍潭', na=False))
                ]
            else:
                u_rows = df2_all[
                    (df2_all['單位'].apply(map_unit_name) == unit) &
                    (~df2_all['來源檔名'].str.contains('大隊|交大', na=False))
                ]

            h_sum = int(u_rows['大型車純違規'].sum()) if not u_rows.empty else 0

            res = [unit]
            for i, cat in enumerate(CATS):
                cnt = d15.get(cat, 0) if cat != "大型車違規" else h_sum
                tgt = TARGET_CONFIG[unit][i]
                res.extend([cnt, tgt, f"{(cnt/tgt*100):.1f}%" if tgt > 0 else "0.0%"])
            final_rows.append(res)

        headers = ["單位"]
        for cat in CATS:
            headers.extend([f"{cat}_取締件數", f"{cat}_目標值", f"{cat}_達成率"])
        df_f = pd.DataFrame(final_rows, columns=headers)

        # 合計列
        total = ["合計"]
        for i in range(1, len(headers), 3):
            cs = df_f.iloc[:, i].sum()
            ts = df_f.iloc[:, i+1].sum()
            total.extend([int(cs), int(ts), f"{(cs/ts*100):.1f}%" if ts > 0 else "0.0%"])

        df_f.loc[df_f['單位'].isin(['交通組', '警備隊']), [c for c in df_f.columns if '目標值' in c or '達成率' in c]] = '-'
        df_f = pd.concat([pd.DataFrame([total], columns=headers), df_f], ignore_index=True)

        # 顯示結果
        st.divider()
        st.markdown(f"### 📊 :blue[{PROJECT_NAME}] :red[(統計期間：{date_range_str})]")
        st.dataframe(df_f, use_container_width=True, hide_index=True)

        # 雲端同步按鈕
        if st.button("🚀 同步至雲端 Google Sheets"):
            with st.spinner("同步中..."):
                gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                sh = gc.open_by_url(GOOGLE_SHEET_URL)
                ws = sh.worksheet(PROJECT_NAME)
                full_t = f"{PROJECT_NAME} (統計期間：{date_range_str})"
                ws.clear()
                ws.update(values=[
                    [full_t] + [""] * 18,
                    [""] + [c for c in CATS for _ in range(3)],
                    ["單位"] + ["取締件數", "目標值", "達成率"] * 6
                ] + df_f.values.tolist())
                st.success("✅ 同步完成！")

    except Exception as e:
        st.error(f"❌ 解析錯誤：{e}")
else:
    st.warning("⏳ 請上傳報表檔案以開始分析。")
