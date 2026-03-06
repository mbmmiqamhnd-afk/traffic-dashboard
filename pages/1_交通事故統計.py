import streamlit as st
import pandas as pd
import io
import re
import gspread

# ==========================================
# 🔐 1. 安全設定與環境配置
# ==========================================
st.set_page_config(page_title="交通事故統計系統", layout="wide", page_icon="🚑")

try:
    GCP_CREDS = st.secrets["gcp_service_account"]
except Exception as e:
    st.error("❌ 找不到 Secrets 設定！請配置 [gcp_service_account]。")
    st.stop()

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

# ==========================================
# 🛠️ 2. 工具函式
# ==========================================

def parse_raw(f):
    try:
        f.seek(0)
        try: return pd.read_csv(f, header=None)
        except: 
            f.seek(0)
            return pd.read_excel(f, header=None)
    except Exception as e:
        st.error(f"檔案 {f.name} 讀取失敗: {e}")
        return None

def clean_data(df_raw):
    df_raw[0] = df_raw[0].astype(str)
    df_data = df_raw[df_raw[0].str.contains("所|總計|合計", na=False)].copy()
    # 根據您的 CSV 結構：Index 0 單位, Index 5 A1死亡, Index 9 A2受傷
    cols = {0: "Station", 5: "A1_Deaths", 9: "A2_Injuries"}
    df_data = df_data.rename(columns=cols)
    for c in [5, 9]:
        target = cols[c]
        df_data[target] = pd.to_numeric(df_data[target].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
    df_data['Station_Short'] = df_data['Station'].str.replace('派出所', '所').str.replace('總計', '合計').str.strip()
    return df_data

# ==========================================
# 📊 3. 核心計算邏輯 (合併 4 個檔案)
# ==========================================

def build_final_table(df_wk, df_prev, df_cur, df_lst, stations, col_name, is_a2=False):
    # 合併四份資料
    m = pd.merge(df_wk[['Station_Short', col_name]], df_prev[['Station_Short', col_name]], on='Station_Short', suffixes=('_wk', '_prev'))
    m = pd.merge(m, df_cur[['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_cur'})
    m = pd.merge(m, df_lst[['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_lst'})
    
    # 排序單位
    m = m[m['Station_Short'].isin(stations)].copy()
    m['Station_Short'] = pd.Categorical(m['Station_Short'], categories=stations, ordered=True)
    m = m.sort_values('Station_Short')
    
    # 計算合計
    total = m.select_dtypes(include='number').sum().to_dict()
    total['Station_Short'] = '合計'
    m = pd.concat([pd.DataFrame([total]), m], ignore_index=True)
    
    # 計算比較 (本年累計 - 去年累計)
    m['Diff'] = m[col_name+'_cur'] - m[col_name+'_lst']
    if is_a2:
        m['Pct'] = m.apply(lambda x: f"{(x['Diff']/x[col_name+'_lst']):.2%}" if x[col_name+'_lst'] != 0 else "0.00%", axis=1)
    
    return m

# ==========================================
# ☁️ 4. 雲端同步 (將前期數據輸出至 C 欄)
# ==========================================

def sync_to_gsheet(df_a1, df_a2):
    try:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        
        # A1 類更新 (索引 2)
        ws_a1 = sh.get_worksheet(2)
        ws_a1.batch_clear(["A3:Z100"])
        # A1 欄位: [單位, 本期, 本年累計, 去年累計, 比較]
        ws_a1.update('A3', df_a1.values.tolist())
        
        # A2 類更新 (索引 3)
        ws_a2 = sh.get_worksheet(3)
        ws_a2.batch_clear(["A3:Z100"])
        
        a2_final_rows = []
        for _, row in df_a2.iterrows():
            # 強制對齊：A欄:單位, B欄:本期, C欄:前期(來自第4個檔案), D欄:本年累計, E欄:去年累計, F欄:比較, G欄:比例
            line = [
                row['Station_Short'], 
                row['A2_Injuries_wk'], 
                row['A2_Injuries_prev'], # 👈 這就是輸出在 C 欄的數值
                row['A2_Injuries_cur'], 
                row['A2_Injuries_lst'], 
                row['Diff'], 
                row['Pct']
            ]
            # 轉整數確保雲端不顯示小數點
            line = [int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in line]
            a2_final_rows.append(line)
        
        ws_a2.update('A3', a2_final_rows)
        return True, "✅ 同步成功：前期 A2 受傷人數已填入 C 欄"
    except Exception as e:
        return False, f"❌ 同步失敗: {e}"

# ==========================================
# 🚀 5. 主程式
# ==========================================

st.title("🚑 交通事故統計 (4 檔案版本)")

files = st.file_uploader("請上傳 4 個檔案：1.本期 2.前期 3.本年累計 4.去年累計", accept_multiple_files=True)

if files and len(files) == 4:
    with st.status("正在分析資料並分配期別...") as status:
        try:
            meta = []
            for f in files:
                df = parse_raw(f)
                text = str(df.iloc[:5, :5].values)
                dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', text)
                if len(dates) >= 2:
                    meta.append({
                        'df': clean_data(df),
                        'year': int(dates[1][0]),
                        'start_day': int(dates[0][2]), # 用來判斷本期 vs 前期
                        'is_cumu': (int(dates[0][1]) == 1)
                    })

            if len(meta) < 4:
                st.error("日期解析失敗，請確認檔案。")
                st.stop()

            # --- 關鍵：四個檔案的分配邏輯 ---
            # 1. 去年累計 (年份最小)
            df_lst = sorted([f for f in meta if f['year'] < max(m['year'] for m in meta)], key=lambda x: x['year'])[-1]
            # 2. 今年累計 (今年且從 1/1 開始)
            df_cur = [f for f in meta if f['year'] == max(m['year'] for m in meta) and f['is_cumu']][0]
            # 3. 本期 vs 前期 (今年、非累計，且日期較晚的是本期)
            recent_files = sorted([f for f in meta if f['year'] == max(m['year'] for m in meta) and not f['is_cumu']], key=lambda x: x['start_day'])
            df_prev = recent_files[0] # 較早的日期
            df_wk = recent_files[1]   # 較晚的日期

            stations = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
            
            # 執行計算
            a1_res = build_final_table(df_wk['df'], df_prev['df'], df_cur['df'], df_lst['df'], stations, 'A1_Deaths')
            a2_res = build_final_table(df_wk['df'], df_prev['df'], df_cur['df'], df_lst['df'], stations, 'A2_Injuries', is_a2=True)

            # 同步至雲端
            ok, msg = sync_to_gsheet(a1_res, a2_res)
            st.write(msg)
            
            status.update(label="✅ 同步完成", state="complete")
            st.success("統計已更新：C 欄現在顯示「前期」受傷人數。")
            st.write("A2 類最終合併結果：")
            st.dataframe(a2_res)

        except Exception as e:
            st.error(f"分析失敗：{e}")
else:
    st.info("請同時選取 4 個報表檔案上傳。")
