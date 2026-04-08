import streamlit as st
import pandas as pd
import io
import re
import gspread
import calendar
import traceback
from datetime import datetime, timedelta, date

# ==========================================
# 0. 系統初始化
# ==========================================
st.set_page_config(page_title="龍潭分局交通智慧戰情室", page_icon="🚓", layout="wide")

try:
    from gspread_formatting import *
    HAS_FORMATTING = True
except ImportError:
    HAS_FORMATTING = False

# ==========================================
# 1. 全局常數與設定
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

try:
    GCP_CREDS = dict(st.secrets.get("gcp_service_account", {}))
except:
    GCP_CREDS = None

MAJOR_UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
MAJOR_TARGETS = {'聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479, '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006}
MAJOR_FOOTNOTE = "重大交通違規指：「酒駕」、「闖紅燈」、「嚴重超速」、「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」及「不暫停讓行人」"

# ==========================================
# 2. 核心組件：重大違規處理
# ==========================================

def get_robust_date(df):
    """
    🛠️ 強化版「入案日」日期抓取：
    精確鎖定 7 位數民國日期，並僅擷取後 4 碼 (MMDD)。
    """
    try:
        # 讀取前 10 列並清理字串
        text_block = df.head(10).astype(str).to_string()
        text_block = text_block.replace("\n", " ").replace(" ", "")
        
        # 模式：尋找 7 位數字 (民國日期 1150401)
        # 我們找尋 1 開頭且後面接 6 個數字的組合
        date_matches = re.findall(r'1\d{6}', text_block)
        
        if len(date_matches) >= 2:
            # 取第一組與最後一組的後 4 碼
            start_mmdd = date_matches[0][-4:]
            end_mmdd = date_matches[-1][-4:]
            return f"{start_mmdd}-{end_mmdd}"
        
        # 備援模式：尋找 115/04/01 這種帶斜線的格式
        slashed_matches = re.findall(r'1\d{2}/(\d{2})/(\d{2})', text_block)
        if len(slashed_matches) >= 2:
            start_mmdd = f"{slashed_matches[0][0]}{slashed_matches[0][1]}"
            end_mmdd = f"{slashed_matches[-1][0]}{slashed_matches[-1][1]}"
            return f"{start_mmdd}-{end_mmdd}"
            
        return "" 
    except:
        return ""

def parse_major_data(f, sheet_kw, col_pair):
    """解析單一檔案的特定數據欄位"""
    f.seek(0)
    xl = pd.ExcelFile(f)
    sn = next((s for s in xl.sheet_names if sheet_kw in s), xl.sheet_names[0])
    df = pd.read_excel(xl, sheet_name=sn, header=None)
    
    dt_str = get_robust_date(df)
    
    def clean_unit(n):
        n = str(n).strip()
        if '分隊' in n: return '交通分隊'
        if any(k in n for k in ['科技', '交通組']): return '科技執法'
        if '警備' in n: return '警備隊'
        for k in ['聖亭', '龍潭', '中興', '石門', '高平', '三和']:
            if k in n: return k + '所'
        return None

    def to_i(v):
        try: return int(float(str(v).replace(',', '').strip()))
        except: return 0

    res = {}
    for _, r in df.iterrows():
        u = clean_unit(r.iloc[0])
        if u and "合計" not in str(r.iloc[0]):
            res[u] = {'stop': to_i(r.iloc[col_pair[0]]), 'cit': to_i(r.iloc[col_pair[1]])}
    return res, dt_str

def process_major_module(files):
    """重大違規處理主邏輯"""
    if len(files) < 2:
        st.error("❌ 錯誤：請確認是否同時上傳『本期(週)』與『(1)累計』報表。")
        return

    # 檔案識別 (大者為累計)
    sorted_files = sorted(files, key=lambda x: x.size)
    f_wk, f_year = sorted_files[0], sorted_files[1]
    
    # 若檔名標註相反則修正
    if "(1)" in f_wk.name: f_wk, f_year = f_year, f_wk

    # 數據抓取
    d_wk, date_wk = parse_major_data(f_wk, "重點違規", [15, 16])
    d_year, date_yr = parse_major_data(f_year, "(1)", [15, 16])
    d_last, _ = parse_major_data(f_year, "(1)", [18, 19])

    # 建立表格數據
    table_rows = []
    summary = {k: 0 for k in ['ws', 'wc', 'ys', 'yc', 'ls', 'lc', 'diff', 'tgt']}
    
    for u in MAJOR_UNIT_ORDER:
        w_data = d_wk.get(u, {'stop':0, 'cit':0})
        y_data = d_year.get(u, {'stop':0, 'cit':0})
        l_data = d_last.get(u, {'stop':0, 'cit':0})
        
        y_total = y_data['stop'] + y_data['cit']
        l_total = l_data['stop'] + l_data['cit']
        tgt = MAJOR_TARGETS.get(u, 0)
        diff = int(y_total - l_total)
        
        rate = f"{(y_total/tgt):.1%}" if tgt > 0 else "0%"
        
        if u != '警備隊':
            summary['diff'] += diff; summary['tgt'] += tgt
            
        table_rows.append([u, w_data['stop'], w_data['cit'], y_data['stop'], y_data['cit'], l_data['stop'], l_data['cit'], diff if u != '警備隊' else "—", tgt, rate if u != '警備隊' else "—"])
        
        summary['ws'] += w_data['stop']; summary['wc'] += w_data['cit']
        summary['ys'] += y_data['stop']; summary['yc'] += y_data['cit']
        summary['ls'] += l_data['stop']; summary['lc'] += l_data['cit']

    # 合計列
    total_rate = f"{((summary['ys']+summary['yc'])/summary['tgt']):.1%}" if summary['tgt'] > 0 else "0%"
    table_rows.insert(0, ['合計', summary['ws'], summary['wc'], summary['ys'], summary['yc'], summary['ls'], summary['lc'], summary['diff'], summary['tgt'], total_rate])
    table_rows.append([MAJOR_FOOTNOTE] + [""] * 9)
    
    # 組合標題標籤
    h_wk = f"本期({date_wk})" if date_wk else "本期"
    h_yr = f"本年累計({date_yr})" if date_yr else "本年累計"
    h_ls = f"去年累計({date_yr})" if date_yr else "去年累計"
    
    header_1 = ['統計期間', h_wk, h_wk, h_yr, h_yr, h_ls, h_ls, '本年與去年同期比較', '目標值', '達成率']
    header_2 = ['取締方式', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '', '', '']
    
    df_result = pd.DataFrame(table_rows, columns=pd.MultiIndex.from_arrays([header_1, header_2]))
    
    st.subheader("📊 重大交通違規自動化統計")
    st.dataframe(df_result, use_container_width=True)

    if GCP_CREDS:
        if push_to_gsheet(df_result):
            st.success(f"✅ 雲端同步成功！(識別日期：{date_wk})")

def push_to_gsheet(df):
    """推送到 Google Sheets，從 A2 開始寫入"""
    try:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        ws = sh.get_worksheet(0)
        
        titles = df.columns.tolist()
        data = [[t[0] for t in titles], [t[1] for t in titles]] + df.values.tolist()
        ws.update(range_name='A2', values=data)
        return True
    except Exception as e:
        st.error(f"雲端同步失敗：{e}")
        return False

# ==========================================
# 3. Streamlit 介面
# ==========================================
st.title("🚓 龍潭分局交通智慧戰情室")
st.markdown("---")

uploaded_files = st.file_uploader("📂 請上傳重大違規相關報表", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 啟動數據批次處理"):
        target_files = [f for f in uploaded_files if any(k in f.name for k in ["重大", "重點"])]
        process_major_module(target_files)
