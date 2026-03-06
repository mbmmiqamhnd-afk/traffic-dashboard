import streamlit as st
import pandas as pd
import io
import re
import gspread

# ==========================================
# 🔐 1. 系統配置與安全設定
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
    cols = {0: "Station", 5: "A1_Deaths", 9: "A2_Injuries"}
    df_data = df_data.rename(columns=cols)
    for c in [5, 9]:
        target = cols[c]
        df_data[target] = pd.to_numeric(df_data[target].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
    df_data['Station_Short'] = df_data['Station'].str.replace('派出所', '所').str.replace('總計', '合計').str.strip()
    return df_data

# ==========================================
# 📊 3. 核心合併邏輯 (日期區間顯示在各標題)
# ==========================================

def build_final_table(df_wk, df_prev, df_cur, df_lst, stations, col_name, labels, is_a2=False):
    # 合併四份報表
    m = pd.merge(df_wk[['Station_Short', col_name]], df_prev[['Station_Short', col_name]], on='Station_Short', suffixes=('_wk', '_prev'))
    m = pd.merge(m, df_cur[['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_cur'})
    m = pd.merge(m, df_lst[['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_lst'})
    
    # 單位排序
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
        res = m[['Station_Short', col_name+'_wk', col_name+'_prev', col_name+'_cur', col_name+'_lst', 'Diff', 'Pct']]
        # 🌟 標題顯示各自區間
        res.columns = [
            '統計期間', 
            f'本期({labels["wk"]})', 
            f'前期({labels["prev"]})', 
            f'本年累計({labels["cur"]})', 
            f'去年累計({labels["lst"]})', 
            '本年與去年同期比較', 
            '增減比例'
        ]
    else:
        res = m[['Station_Short', col_name+'_wk', col_name+'_cur', col_name+'_lst', 'Diff']]
        res.columns = [
            '統計期間', 
            f'本期({labels["wk"]})', 
            f'本年累計({labels["cur"]})', 
            f'去年累計({labels["lst"]})', 
            '本年與去年同期比較'
        ]
    
    return res

# ==========================================
# ☁️ 4. 雲端同步
# ==========================================

def sync_to_gsheet(df_a1, df_a2):
    try:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        
        ws1 = sh.get_worksheet(2)
        ws1.batch_clear(["A2:F100"])
        ws1.update('A2', [df_a1.columns.tolist()])
        ws1.update('A3', df_a1.values.tolist())
        
        ws2 = sh.get_worksheet(3)
        ws2.batch_clear(["A2:G100"])
        ws2.update('A2', [df_a2.columns.tolist()])
        
        data_to_save = [[int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in row] for row in df_a2.values.tolist()]
        ws2.update('A3', data_to_save)
        
        return True, "✅ 同步成功：各期別標題已包含日期區間。"
    except Exception as e:
        return False, f"❌ 同步失敗: {e}"

# ==========================================
# 🚀 5. Streamlit 主流程
# ==========================================

files = st.file_uploader("請上傳 4 個報表檔案", accept_multiple_files=True)

if files and len(files) == 4:
    with st.status("正在解析各檔案日期...") as status:
        try:
            meta = []
            for f in files:
                df = parse_raw(f)
                text = str(df.iloc[:5, :5].values)
                # 偵測日期
                dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', text)
                if len(dates) >= 2:
                    # 格式化日期區間為 0220-0226
                    d_range = f"{int(dates[0][1]):02d}{int(dates[0][2]):02d}-{int(dates[1][1]):02d}{int(dates[1][2]):02d}"
                    meta.append({
                        'df': clean_data(df),
                        'year': int(dates[1][0]),
                        'start_day': int(dates[0][1])*100 + int(dates[0][2]),
                        'range': d_range,
                        'is_cumu': (int(dates[0][1]) == 1 and int(dates[0][2]) == 1)
                    })

            if len(meta) < 4:
                st.error("日期解析失敗，請確認檔案格式。")
                st.stop()

            # --- 檔案分配 ---
            this_year = max(m['year'] for m in meta)
            # 1. 去年累計
            f_lst = sorted([f for f in meta if f['year'] < this_year], key=lambda x: x['year'])[-1]
            # 2. 今年累計
            f_cur = [f for f in meta if f['year'] == this_year and f['is_cumu']][0]
            # 3. 本期 vs 前期
            period_files = sorted([f for f in meta if f['year'] == this_year and not f['is_cumu']], key=lambda x: x['start_day'])
            f_prev = period_files[0]
            f_wk = period_files[1]

            # 建立動態標題標籤
            labels = {
                "wk": f_wk['range'],
                "prev": f_prev['range'],
                "cur": f_cur['range'],
                "lst": f_lst['range']
            }

            stations = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
            
            a1_res = build_final_table(f_wk['df'], f_prev['df'], f_cur['df'], f_lst['df'], stations, 'A1_Deaths', labels)
            a2_res = build_final_table(f_wk['df'], f_prev['df'], f_cur['df'], f_lst['df'], stations, 'A2_Injuries', labels, is_a2=True)

            ok, msg = sync_to_gsheet(a1_res, a2_res)
            
            status.update(label="✅ 處理完成", state="complete")
            st.success(msg)
            st.table(a2_res)

        except Exception as e:
            st.error(f"分析失敗：{e}")
else:
    st.info("💡 請同時上傳 4 個報表檔案。")
