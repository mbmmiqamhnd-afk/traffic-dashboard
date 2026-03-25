import streamlit as st
import pandas as pd
import io
import re
import gspread
from datetime import datetime

# ==========================================
# 🔐 1. 系統配置
# ==========================================
st.set_page_config(page_title="交通事故統計", layout="wide", page_icon="🚑")

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

# ==========================================
# 🛠️ 2. 工具函式 (交通事故專用)
# ==========================================

def get_gsheet_rich_text_req(sheet_id, row_idx, col_idx, text):
    """Google Sheets 紅字格式請求：數字與符號轉紅"""
    text = str(text)
    pattern = r'([0-9\(\)\/\-]+)'
    tokens = re.split(pattern, text)
    runs = []
    current_pos = 0
    for token in tokens:
        if not token: continue
        color = {"red": 1, "green": 0, "blue": 0} if re.match(pattern, token) else {"red": 0, "green": 0, "blue": 0}
        runs.append({"startIndex": current_pos, "format": {"foregroundColor": color, "bold": True}})
        current_pos += len(token)
    return {
        "updateCells": {
            "rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}],
            "fields": "userEnteredValue,textFormatRuns",
            "range": {"sheetId": sheet_id, "startRowIndex": row_idx, "endRowIndex": row_idx + 1, "startColumnIndex": col_idx, "endColumnIndex": col_idx + 1}
        }
    }

def clean_traffic_data(df_raw):
    """交通事故原始數據清洗"""
    df_raw[0] = df_raw[0].astype(str)
    df_data = df_raw[df_raw[0].str.contains("所|總計|合計", na=False)].copy()
    cols = {0: "Station", 5: "A1_Deaths", 9: "A2_Injuries"}
    df_data = df_data.rename(columns=cols)
    for c in [5, 9]:
        target = cols[c]
        df_data[target] = pd.to_numeric(df_data[target].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
    df_data['Station_Short'] = df_data['Station'].str.replace('派出所', '所').str.replace('總計', '合計').str.strip()
    return df_data

def build_traffic_table(df_wk, df_prev, df_cur, df_lst, stations, col_name, labels, is_a2=False):
    """建立 A1 或 A2 的統計總表"""
    m = pd.merge(df_wk[['Station_Short', col_name]], df_prev[['Station_Short', col_name]], on='Station_Short', suffixes=('_wk', '_prev'))
    m = pd.merge(m, df_cur[['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_cur'})
    m = pd.merge(m, df_lst[['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_lst'})
    m = m[m['Station_Short'].isin(stations)].copy()
    m['Station_Short'] = pd.Categorical(m['Station_Short'], categories=stations, ordered=True)
    m = m.sort_values('Station_Short')
    total = m.select_dtypes(include='number').sum().to_dict()
    total['Station_Short'] = '合計'
    m = pd.concat([pd.DataFrame([total]), m], ignore_index=True)
    m['Diff'] = m[col_name+'_cur'] - m[col_name+'_lst']
    if is_a2:
        m['Pct'] = m.apply(lambda x: f"{(x['Diff']/x[col_name+'_lst']):.2%}" if x[col_name+'_lst'] != 0 else "0.00%", axis=1)
        res = m[['Station_Short', col_name+'_wk', col_name+'_prev', col_name+'_cur', col_name+'_lst', 'Diff', 'Pct']]
        res.columns = ['統計期間', f'本期({labels["wk"]})', f'前期({labels["prev"]})', f'本年累計({labels["cur"]})', f'去年累計({labels["lst"]})', '本年與去年同期比較', '增減比例']
    else:
        res = m[['Station_Short', col_name+'_wk', col_name+'_cur', col_name+'_lst', 'Diff']]
        res.columns = ['統計期間', f'本期({labels["wk"]})', f'本年累計({labels["cur"]})', f'去年累計({labels["lst"]})', '本年與去年同期比較']
    return res

# ==========================================
# 🚀 3. Streamlit 主入口 (雙通道接收)
# ==========================================

st.title("🚑 龍潭分局 - 交通事故統計系統")
st.markdown("##### 🚀 **操作說明：** 需準備 4 個交通事故報表，系統將自動解析、計算並完成雲端同步。")

# 🌟【關鍵修改區】：雙通道接收檔案 🌟
all_files = None
if "auto_files_accident" in st.session_state and st.session_state["auto_files_accident"]:
    st.info("📥 系統已自動載入從「首頁」分配過來的檔案！")
    all_files = st.session_state["auto_files_accident"]
    
    # 防呆機制：取消載入
    if st.button("❌ 取消自動載入，改為手動上傳"):
        del st.session_state["auto_files_accident"]
        st.rerun()
else:
    all_files = st.file_uploader("📂 請全選並拖入所有 Excel 報表 (需 4 份)", type=["xlsx", "csv"], accept_multiple_files=True)

# 以下為您原本強大的分析邏輯 (完全保留)
if all_files:
    # --- 交通事故處理邏輯 (偵測到 4 個檔案) ---
    if len(all_files) == 4:
        with st.status("📊 偵測到交通事故報表，啟動智慧分析...", expanded=True) as status:
            try:
                meta = []
                for f in all_files:
                    # 讀取檔案
                    f.seek(0)
                    try: df_raw = pd.read_csv(f, header=None)
                    except: 
                        f.seek(0)
                        df_raw = pd.read_excel(f, header=None)
                    
                    # 日期解析
                    sample_text = str(df_raw.iloc[:5, :5].values)
                    dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', sample_text)
                    if len(dates) >= 2:
                        d_range = f"{int(dates[0][1]):02d}{int(dates[0][2]):02d}-{int(dates[1][1]):02d}{int(dates[1][2]):02d}"
                        meta.append({
                            'df': clean_traffic_data(df_raw),
                            'year': int(dates[1][0]),
                            'start_day': int(dates[0][1])*100 + int(dates[0][2]),
                            'range': d_range,
                            'is_cumu': (int(dates[0][1]) == 1 and int(dates[0][2]) == 1)
                        })

                if len(meta) == 4:
                    # 檔案自動分流
                    this_year = max(m['year'] for m in meta)
                    f_lst = sorted([f for f in meta if f['year'] < this_year], key=lambda x: x['year'])[-1]
                    f_cur = [f for f in meta if f['year'] == this_year and f['is_cumu']][0]
                    period_files = sorted([f for f in meta if f['year'] == this_year and not f['is_cumu']], key=lambda x: x['start_day'])
                    f_prev, f_wk = period_files[0], period_files[1]

                    labels = {"wk": f_wk['range'], "prev": f_prev['range'], "cur": f_cur['range'], "lst": f_lst['range']}
                    stations = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
                    
                    # 計算 A1 & A2
                    a1_res = build_traffic_table(f_wk['df'], f_prev['df'], f_cur['df'], f_lst['df'], stations, 'A1_Deaths', labels)
                    a2_res = build_traffic_table(f_wk['df'], f_prev['df'], f_cur['df'], f_lst['df'], stations, 'A2_Injuries', labels, is_a2=True)

                    # 顯示預覽
                    st.subheader(f"📅 分析期間：{labels['wk']}")
                    col1, col2 = st.columns(2)
                    col1.write("A1 死亡人數統計")
                    col1.dataframe(a1_res, hide_index=True)
                    col2.write("A2 受傷人數統計")
                    col2.dataframe(a2_res, hide_index=True)

                    # --- 自動同步雲端 ---
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    
                    # 處理 A1 (WS Index 2) 與 A2 (WS Index 3)
                    for ws_idx, df in zip([2, 3], [a1_res, a2_res]):
                        ws = sh.get_worksheet(ws_idx)
                        ws.batch_clear(["A2:G20"])
                        
                        # 同步標題 (含紅字)
                        reqs = []
                        for c_idx, c_name in enumerate(df.columns):
                            reqs.append(get_gsheet_rich_text_req(ws.id, 1, c_idx, c_name))
                        sh.batch_update({"requests": reqs})
                        
                        # 同步內容
                        data_rows = [[int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in row] for row in df.values.tolist()]
                        ws.update('A3', data_rows)

                    status.update(label="✅ 交通事故統計完成，雲端紅字格式已更新！", state="complete")
                else:
                    st.error("日期解析失敗，請確認檔案內容或數量。")

            except Exception as e:
                st.error(f"分析失敗：{e}")
    
    # --- 如果上傳的是其他數量的檔案 ---
    else:
        st.warning(f"⚠️ 注意：交通事故分析需剛好 4 份檔案，目前收到 {len(all_files)} 份檔案。")

else:
    st.info("⏳ 準備就緒，請透過首頁分配或手動上傳報表檔案...")
