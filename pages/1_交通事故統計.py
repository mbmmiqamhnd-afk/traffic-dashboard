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
# 🔐 1. 安全設定與環境配置
# ==========================================
st.set_page_config(page_title="交通事故統計系統", layout="wide", page_icon="🚑")

try:
    MY_EMAIL = st.secrets["email"]["user"]
    MY_PASSWORD = st.secrets["email"]["password"]
    GCP_CREDS = st.secrets["gcp_service_account"]
except Exception as e:
    st.error("❌ 找不到 Secrets 設定！請在 .streamlit/secrets.toml 中配置 [email] 與 [gcp_service_account]。")
    st.stop()

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

# ==========================================
# 🛠️ 2. 工具函式
# ==========================================

def parse_raw(f):
    try:
        f.seek(0)
        # 優先嘗試以 CSV 讀取，若失敗則用 Excel
        try:
            return pd.read_csv(f, header=None)
        except:
            f.seek(0)
            return pd.read_excel(f, header=None)
    except Exception as e:
        st.error(f"檔案讀取失敗: {e}")
        return None

def clean_data(df_raw):
    """清洗報表，提取關鍵欄位：A1死亡(Index 5), A2受傷(Index 9)"""
    df_raw[0] = df_raw[0].astype(str)
    # 篩選包含派出所或合計的列
    df_data = df_raw[df_raw[0].str.contains("所|總計|合計", na=False)].copy()
    
    # 根據您的檔案：Index 0 是單位, Index 5 是 A1死亡, Index 9 是 A2受傷
    cols = {0: "Station", 5: "A1_Deaths", 9: "A2_Injuries"}
    df_data = df_data.rename(columns=cols)
    
    for c in [5, 9]:
        target = cols[c]
        df_data[target] = pd.to_numeric(df_data[target].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
    
    df_data['Station_Short'] = df_data['Station'].str.replace('派出所', '所').str.replace('總計', '合計').str.strip()
    return df_data

# ==========================================
# 📊 3. 核心計算邏輯
# ==========================================

def build_table(wk_df, cur_df, lst_df, stations, col_name, is_a2=False):
    # 合併三份資料
    m = pd.merge(wk_df[['Station_Short', col_name]], cur_df[['Station_Short', col_name]], on='Station_Short', suffixes=('_wk', '_cur'))
    m = pd.merge(m, lst_df[['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_lst'})
    
    # 過濾指定單位並排序
    m = m[m['Station_Short'].isin(stations)].copy()
    m['Station_Short'] = pd.Categorical(m['Station_Short'], categories=stations, ordered=True)
    m = m.sort_values('Station_Short')
    
    # 計算合計列
    total = m.select_dtypes(include='number').sum().to_dict()
    total['Station_Short'] = '合計'
    m = pd.concat([pd.DataFrame([total]), m], ignore_index=True)
    
    # 計算比較值與增減比例
    m['Diff'] = m[col_name+'_cur'] - m[col_name+'_lst']
    if is_a2:
        m['Pct'] = m.apply(lambda x: f"{(x['Diff']/x[col_name+'_lst']):.2%}" if x[col_name+'_lst'] != 0 else "0.00%", axis=1)
    
    return m

# ==========================================
# ☁️ 4. 雲端同步 (A2受傷數值輸出至 C 欄)
# ==========================================

def sync_to_gsheet(df_a1, df_a2, labels):
    try:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        
        # A1 類更新 (第 3 個分頁)
        ws_a1 = sh.get_worksheet(2)
        ws_a1.batch_clear(["A3:Z100"])
        a1_rows = [[int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in row] for row in df_a1.values.tolist()]
        ws_a1.update('A3', a1_rows)
        
        # A2 類更新 (第 4 個分頁，受傷人數置於 C 欄)
        ws_a2 = sh.get_worksheet(3)
        ws_a2.batch_clear(["A3:Z100"])
        
        a2_final_rows = []
        for _, row in df_a2.iterrows():
            # 格式：[單位, 前期佔位, 本期受傷(C欄), 本年累計, 去年累計, 比較, 比例]
            line = [row[0], "-", row[1], row[2], row[3], row[4], row[5]]
            line = [int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in line]
            a2_final_rows.append(line)
        
        ws_a2.update('A3', a2_final_rows)
        return True, "✅ 同步成功：A2 受傷人數已輸出至 C 欄"
    except Exception as e:
        return False, f"❌ 同步失敗: {e}"

# ==========================================
# 🚀 5. 主程式
# ==========================================

st.title("🚑 交通事故統計 (受傷人數 C 欄輸出版)")

files = st.file_uploader("請上傳 3 個報表檔案 (本期、今年累計、去年累計)", accept_multiple_files=True)

if files and len(files) == 3:
    with st.status("分析中...", expanded=True) as status:
        try:
            meta = []
            for f in files:
                df = parse_raw(f)
                # 偵測日期標題區塊
                text_block = str(df.iloc[:5, :5].values)
                dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', text_block)
                if len(dates) >= 2:
                    d_label = f"{int(dates[0][1]):02d}{int(dates[0][2]):02d}-{int(dates[1][1]):02d}{int(dates[1][2]):02d}"
                    meta.append({'df': clean_data(df), 'year': int(dates[1][0]), 'range': d_label, 'is_cumu': (int(dates[0][1]) == 1)})
            
            if len(meta) < 3:
                st.error("無法辨識日期。請檢查報表最上方是否包含民國年日期。")
                st.stop()

            # 分配資料
            meta.sort(key=lambda x: x['year'], reverse=True)
            wk = [f for f in meta if f['year'] == meta[0]['year'] and not f['is_cumu']][0]
            cur = [f for f in meta if f['year'] == meta[0]['year'] and f['is_cumu']][0]
            lst = [f for f in meta if f['year'] < meta[0]['year']][0]
            
            stations = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
            
            a1_res = build_table(wk['df'], cur['df'], lst['df'], stations, 'A1_Deaths')
            a2_res = build_table(wk['df'], cur['df'], lst['df'], stations, 'A2_Injuries', is_a2=True)
            
            # 同步
            ok, msg = sync_to_gsheet(a1_res, a2_res, None)
            st.write(msg)
            
            status.update(label="✅ 完成", state="complete")
            st.success("統計已更新！請檢查雲端試算表 C 欄。")
            st.dataframe(a2_res)

        except Exception as e:
            st.error(f"處理失敗：{e}")
else:
    st.info("請同時選取並上傳 3 個報表檔案（.csv 或 .xlsx）。")
