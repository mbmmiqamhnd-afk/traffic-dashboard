import streamlit as st
import pandas as pd
import io
import re
import smtplib
import gspread
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ==========================================
# 🔐 1. 安全設定與環境配置 (Secrets)
# ==========================================
st.set_page_config(page_title="交通事故統計系統", layout="wide", page_icon="🚑")

try:
    MY_EMAIL = st.secrets["email"]["user"]
    MY_PASSWORD = st.secrets["email"]["password"]
    GCP_CREDS = st.secrets["gcp_service_account"]
except Exception as e:
    st.error("❌ 找不到 Secrets 設定！請在 .streamlit/secrets.toml 中配置 [email] 與 [gcp_service_account]。")
    st.stop()

# 指定的雲端試算表網址與設定
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
TO_EMAIL = MY_EMAIL 

# ==========================================
# 🛠️ 2. 工具函式 (Parsing & Formatting)
# ==========================================

def parse_raw(f):
    """解析上傳檔案 (支援 CSV 與 Excel)"""
    try:
        f.seek(0)
        if f.name.endswith('.csv'):
            return pd.read_csv(f, header=None)
        else:
            return pd.read_excel(f, header=None)
    except Exception as e:
        st.error(f"檔案 {f.name} 讀取失敗: {e}")
        return None

def clean_data(df_raw):
    """清洗報表，提取關鍵欄位"""
    df_raw[0] = df_raw[0].astype(str)
    df_data = df_raw[df_raw[0].str.contains("所|總計|合計", na=False)].copy()
    cols = {0: "Station", 5: "A1_Deaths", 9: "A2_Injuries"}
    df_data = df_data.rename(columns=cols)
    
    # 數值清理 (處理千分位逗號)
    for c in [5, 9]:
        target = cols[c]
        df_data[target] = pd.to_numeric(df_data[target].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
    
    df_data['Station_Short'] = df_data['Station'].str.replace('派出所', '所').str.replace('總計', '合計')
    return df_data

def get_gsheet_rich_text_req(sheet_id, row_idx, col_idx, text):
    """產製 Google Sheets 紅字格式請求"""
    text = str(text)
    tokens = re.split(r'([0-9\(\)\/\-\.\%]+)', text)
    runs = []
    current_pos = 0
    for token in tokens:
        if not token: continue
        color = {"red": 1, "green": 0, "blue": 0} if re.match(r'^[0-9\(\)\/\-\.\%]+$', token) else {"red": 0, "green": 0, "blue": 0}
        runs.append({"startIndex": current_pos, "format": {"foregroundColor": color, "bold": True}})
        current_pos += len(token)
    return {
        "updateCells": {
            "rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}],
            "fields": "userEnteredValue,textFormatRuns",
            "range": {"sheetId": sheet_id, "startRowIndex": row_idx, "endRowIndex": row_idx + 1, "startColumnIndex": col_idx, "endColumnIndex": col_idx + 1}
        }
    }

# ==========================================
# 📊 3. 核心計算邏輯 (修正排序與對齊)
# ==========================================

def build_a1_table(wk_df, cur_df, lst_df, stations, date_labels):
    col = 'A1_Deaths'
    m = pd.merge(wk_df[['Station_Short', col]], cur_df[['Station_Short', col]], on='Station_Short', suffixes=('_wk', '_cur'))
    m = pd.merge(m, lst_df[['Station_Short', col]], on='Station_Short').rename(columns={col: col+'_lst'})
    
    # 🌟 僅保留指定單位並自定義排序
    m = m[m['Station_Short'].isin(stations)].copy()
    m['Station_Short'] = pd.Categorical(m['Station_Short'], categories=stations, ordered=True)
    m = m.sort_values('Station_Short')
    
    # 計算合計並置頂
    total = m.select_dtypes(include='number').sum().to_dict()
    total['Station_Short'] = '合計'
    m = pd.concat([pd.DataFrame([total]), m], ignore_index=True)
    
    m['Diff'] = m[col+'_cur'] - m[col+'_lst']
    m = m[['Station_Short', col+'_wk', col+'_cur', col+'_lst', 'Diff']]
    m.columns = ['統計期間', f'本期({date_labels["wk"]})', f'本年累計({date_labels["cur"]})', f'去年累計({date_labels["lst"]})', '比較']
    return m

def build_a2_table(wk_df, cur_df, lst_df, stations, date_labels):
    col = 'A2_Injuries'
    m = pd.merge(wk_df[['Station_Short', col]], cur_df[['Station_Short', col]], on='Station_Short', suffixes=('_wk', '_cur'))
    m = pd.merge(m, lst_df[['Station_Short', col]], on='Station_Short').rename(columns={col: col+'_lst'})
    
    # 🌟 僅保留指定單位並自定義排序
    m = m[m['Station_Short'].isin(stations)].copy()
    m['Station_Short'] = pd.Categorical(m['Station_Short'], categories=stations, ordered=True)
    m = m.sort_values('Station_Short')
    
    # 計算合計並置頂
    total = m.select_dtypes(include='number').sum().to_dict()
    total['Station_Short'] = '合計'
    m = pd.concat([pd.DataFrame([total]), m], ignore_index=True)
    
    m['Diff'] = m[col+'_cur'] - m[col+'_lst']
    m['Pct'] = m.apply(lambda x: f"{(x['Diff']/x[col+'_lst']):.2%}" if x[col+'_lst'] != 0 else "0.00%", axis=1)
    
    # 7 欄位排列: 精準插入 'Prev'
    m.insert(2, 'Prev', '-')
    m = m[['Station_Short', col+'_wk', 'Prev', col+'_cur', col+'_lst', 'Diff', 'Pct']]
    m.columns = ['統計期間', f'本期({date_labels["wk"]})', '前期', f'本年累計({date_labels["cur"]})', f'去年累計({date_labels["lst"]})', '比較', '增減比例']
    return m

# ==========================================
# ☁️ 4. 雲端同步
# ==========================================

def sync_to_gsheet(df_a1, df_a2):
    try:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        
        def update_ws(ws_index, df, title_text):
            ws = sh.get_worksheet(ws_index)
            ws.batch_clear(["A3:Z100"]) 
            ws.update_acell('A1', title_text)
            data = [[int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in row] for row in df.values.tolist()]
            ws.update('A3', data)
            
            # 套用 Rich Text (紅字數字)
            reqs = [get_gsheet_rich_text_req(ws.id, 1, i, col) for i, col in enumerate(df.columns)]
            sh.batch_update({"requests": reqs})
        
        update_ws(2, df_a1, "A1類交通事故死亡人數統計表") # 第3個分頁
        update_ws(3, df_a2, "A2類交通事故受傷人數統計表") # 第4個分頁
        return True, "✅ 雲端試算表同步成功"
    except Exception as e:
        return False, f"❌ 試算表同步失敗: {e}"

# ==========================================
# 🚀 5. 主程式流程
# ==========================================

st.title("🚑 交通事故統計系統 (單位排序修正版)")

uploaded_files = st.file_uploader("請同時上傳 3 個報表檔案", accept_multiple_files=True)

if not uploaded_files:
    st.info("💡 提示：請一次選擇三個報表檔案 (CSV/Excel) 上傳。")
elif len(uploaded_files) != 3:
    st.warning(f"目前已上傳 {len(uploaded_files)} 個檔案，還差 {3-len(uploaded_files)} 個才能啟動。")
else:
    with st.status("正在處理數據與同步...", expanded=True) as status:
        try:
            files_meta = []
            for f in uploaded_files:
                df_raw = parse_raw(f)
                if df_raw is None: continue
                
                # 偵測標題中的日期 (民國年)
                dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', str(df_raw.iloc[:5, :5].values))
                if len(dates) >= 2:
                    d_range = f"{int(dates[0][1]):02d}{int(dates[0][2]):02d}-{int(dates[1][1]):02d}{int(dates[1][2]):02d}"
                    files_meta.append({
                        'df': clean_data(df_raw),
                        'year': int(dates[1][0]),
                        'date_range': d_range,
                        'is_cumu': (int(dates[0][1]) == 1)
                    })
            
            if len(files_meta) < 3:
                st.error("❌ 檔案日期辨識失敗。請確認報表首幾列包含民國年月。")
                st.stop()

            # 排序與分配
            files_meta.sort(key=lambda x: x['year'], reverse=True)
            cur_year = files_meta[0]['year']
            
            df_wk = [f for f in files_meta if f['year'] == cur_year and not f['is_cumu']][0]
            df_cur = [f for f in files_meta if f['year'] == cur_year and f['is_cumu']][0]
            df_lst = [f for f in files_meta if f['year'] < cur_year][0]
            
            date_labels = {"wk": df_wk['date_range'], "cur": df_cur['date_range'], "lst": df_lst['date_range']}
            
            # 🌟 指定順序
            stations_order = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']

            # 計算表格
            a1_res = build_a1_table(df_wk['df'], df_cur['df'], df_lst['df'], stations_order, date_labels)
            a2_res = build_a2_table(df_wk['df'], df_cur['df'], df_lst['df'], stations_order, date_labels)

            # 同步雲端
            gs_ok, gs_msg = sync_to_gsheet(a1_res, a2_res)
            st.write(gs_msg)
            
            status.update(label="✅ 處理完成！", state="complete", expanded=False)
            
            st.success("統計數據已更新，順序：合計 > 聖亭 > 龍潭 > 中興 > 石門 > 高平 > 三和")
            st.dataframe(a1_res, use_container_width=True)
            st.dataframe(a2_res, use_container_width=True)

        except Exception as e:
            st.error(f"分析發生錯誤: {e}")
