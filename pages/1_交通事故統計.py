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
try:
    MY_EMAIL = st.secrets["email"]["user"]
    MY_PASSWORD = st.secrets["email"]["password"]
    GCP_CREDS = st.secrets["gcp_service_account"]
except Exception as e:
    st.error("❌ 找不到 Secrets 設定！請在設定區配置 [email] 與 [gcp_service_account]。")
    st.stop()

TO_EMAIL = MY_EMAIL
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

st.set_page_config(page_title="交通事故統計系統", layout="wide", page_icon="🚑")

# ==========================================
# 🛠️ 2. 工具函式 (必須放在最上方)
# ==========================================

def parse_raw(f):
    """解析 CSV 或 Excel 檔案"""
    try:
        f.seek(0)
        return pd.read_csv(f, header=None)
    except:
        f.seek(0)
        return pd.read_excel(f, header=None)

def clean_data(df_raw):
    """清洗報表原始資料，提取關鍵欄位"""
    df_raw[0] = df_raw[0].astype(str)
    # 篩選包含「所」或「合計」的行
    df_data = df_raw[df_raw[0].str.contains("所|總計|合計", na=False)].copy()
    cols = {0: "Station", 5: "A1_Deaths", 9: "A2_Injuries"}
    df_data = df_data.rename(columns=cols)
    
    # 轉換數值並處理千分位逗號
    for c in [5, 9]:
        col_name = cols[c]
        df_data[col_name] = pd.to_numeric(df_data[col_name].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
    
    # 統一單位名稱
    df_data['Station_Short'] = df_data['Station'].str.replace('派出所', '所').str.replace('總計', '合計')
    return df_data

def format_html_header(text):
    """HTML 顯示紅字數字"""
    text = str(text)
    tokens = re.split(r'([0-9\(\)\/\-\.\%]+)', text)
    html_str = "".join([f'<span style="color: red;">{t}</span>' if re.match(r'^[0-9\(\)\/\-\.\%]+$', t) else f'<span>{t}</span>' for t in tokens])
    return html_str

def render_styled_table(df, title):
    """在 Streamlit 渲染美化表格"""
    st.subheader(title)
    style = "<style>table.acc_table {width:100%; border-collapse:collapse;} th, td {border:1px solid black; padding:8px; text-align:center;}</style>"
    html = f"{style}<table class='acc_table'><thead><tr>"
    for col in df.columns: html += f"<th>{format_html_header(col)}</th>"
    html += "</tr></thead><tbody>"
    for _, row in df.iterrows():
        html += "<tr>"
        for col_name, val in row.items():
            color = "red" if ("比較" in col_name or "增減" in col_name) and str(val) != "0.00%" and "-" not in str(val) and str(val) != "0" else "black"
            html += f'<td style="color: {color};">{val}</td>'
        html += "</tr>"
    st.markdown(html + "</tbody></table>", unsafe_allow_html=True)

# ==========================================
# 📊 3. 核心計算函式 (參數傳遞避免 NameError)
# ==========================================

def build_a1_table(wk_df, cur_df, lst_df, stations):
    col = 'A1_Deaths'
    m = pd.merge(wk_df[['Station_Short', col]], cur_df[['Station_Short', col]], on='Station_Short', suffixes=('_wk', '_cur'))
    m = pd.merge(m, lst_df[['Station_Short', col]], on='Station_Short').rename(columns={col: col+'_lst'})
    
    m = m[m['Station_Short'].isin(stations)].copy()
    total = m.select_dtypes(include='number').sum().to_dict()
    total['Station_Short'] = '合計'
    m = pd.concat([pd.DataFrame([total]), m], ignore_index=True)
    
    m['Diff'] = m[col+'_cur'] - m[col+'_lst']
    # 5 欄排列
    m = m[['Station_Short', col+'_wk', col+'_cur', col+'_lst', 'Diff']]
    return m

def build_a2_table(wk_df, cur_df, lst_df, stations):
    col = 'A2_Injuries'
    m = pd.merge(wk_df[['Station_Short', col]], cur_df[['Station_Short', col]], on='Station_Short', suffixes=('_wk', '_cur'))
    m = pd.merge(m, lst_df[['Station_Short', col]], on='Station_Short').rename(columns={col: col+'_lst'})
    
    m = m[m['Station_Short'].isin(stations)].copy()
    total = m.select_dtypes(include='number').sum().to_dict()
    total['Station_Short'] = '合計'
    m = pd.concat([pd.DataFrame([total]), m], ignore_index=True)
    
    m['Diff'] = m[col+'_cur'] - m[col+'_lst']
    m['Pct'] = m.apply(lambda x: f"{(x['Diff']/x[col+'_lst']):.2%}" if x[col+'_lst'] != 0 else "0.00%", axis=1)
    
    # 7 欄排列，精準插入 'Prev' 於索引 2 (第3欄)
    m.insert(2, 'Prev', '-')
    m = m[['Station_Short', col+'_wk', 'Prev', col+'_cur', col+'_lst', 'Diff', 'Pct']]
    return m

# ==========================================
# 🚀 4. Streamlit 主流程
# ==========================================

st.title("🚑 交通事故統計 (格式對齊修正版)")
uploaded_files = st.file_uploader("請上傳 3 個報表檔案 (本期、本年累計、去年同期累計)", accept_multiple_files=True)

if uploaded_files and len(uploaded_files) == 3:
    with st.spinner("⚡ 數據分析中..."):
        try:
            files_meta = []
            for f in uploaded_files:
                df_raw = parse_raw(f)
                # 偵測日期 (民國年格式)
                dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', str(df_raw.iloc[:5, :5].values))
                if len(dates) >= 2:
                    d_str = f"{int(dates[0][1]):02d}{int(dates[0][2]):02d}-{int(dates[1][1]):02d}{int(dates[1][2]):02d}"
                    files_meta.append({
                        'df': clean_data(df_raw), 
                        'year': int(dates[1][0]), 
                        'date_range': d_str, 
                        'is_cumu': (int(dates[0][1]) == 1)
                    })

            if len(files_meta) < 3:
                st.error("❌ 無法辨識日期，請確認檔案標題包含民國年月日區間。")
                st.stop()

            # 分配變數
            files_meta.sort(key=lambda x: x['year'], reverse=True)
            cur_year = files_meta[0]['year']
            
            df_wk = [f for f in files_meta if f['year'] == cur_year and not f['is_cumu']][0]
            df_cur = [f for f in files_meta if f['year'] == cur_year and f['is_cumu']][0]
            df_lst = [f for f in files_meta if f['year'] < cur_year][0]

            # 指定派出所
            stations = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']

            # 產生結果 DataFrame (此處已解決 NameError)
            a1_res = build_a1_table(df_wk['df'], df_cur['df'], df_lst['df'], stations)
            a1_res.columns = ['統計期間', f'本期({df_wk["date_range"]})', f'本年累計({df_cur["date_range"]})', f'去年累計({df_lst["date_range"]})', '比較']

            a2_res = build_a2_table(df_wk['df'], df_cur['df'], df_lst['df'], stations)
            a2_res.columns = ['統計期間', f'本期({df_wk["date_range"]})', '前期', f'本年累計({df_cur["date_range"]})', f'去年累計({df_lst["date_range"]})', '比較', '增減比例']

            # 顯示表格
            col1, col2 = st.columns(2)
            with col1: render_styled_table(a1_res, "📊 A1 死亡人數")
            with col2: render_styled_table(a2_res, "📊 A2 受傷人數")

            st.success("✅ 數據對齊完成！數值已按照「單位 > 本期 > 前期 > 累計」順序排列。")

        except Exception as e:
            st.error(f"分析發生錯誤：{e}")
