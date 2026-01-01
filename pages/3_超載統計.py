import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import smtplib
import gspread
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# ==========================================
# 🛑 核心偵測區：如果你沒看到這行字，代表檔案沒存成功
# ==========================================
st.set_page_config(page_title="超載統計", layout="wide", page_icon="🚛")

st.markdown("# 🛑 這是 v27 強制偵測版")
st.error("如果您看到這行紅字，代表程式碼已成功更新。")

# --- 核心重置按鈕 ---
if st.button("🧹 點我徹底重置環境 (解決標題與合計問題)", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    for key in st.session_state.keys():
        del st.session_state[key]
    st.success("✅ 已清空快取！請現在重新整理頁面 (F5)。")
    st.stop()

# ==========================================
# 0. 設定與參數區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

TARGETS = {
    '科技執法': 0, '聖亭所': 24, '龍潭所': 32, '中興所': 24, 
    '石門所': 19, '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30
}

UNIT_MAP = {
    '交通組': '科技執法', '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', 
    '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', 
    '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'
}

UNIT_DATA_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. 核心寫入函數 (A2 起始，含標題)
# ==========================================
def update_sheet_from_a2(df, sheet_url):
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Secrets 未設定！")
            return False

        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(1) # 分頁 2 (Index 1)
        
        # 標題列 + 數據列
        header = df.columns.tolist()
        values = df.values.tolist()
        payload = [header] + values
        
        # 從 A2 開始覆蓋寫入
        try:
            ws.update(range_name='A2', values=payload)
        except:
            ws.update('A2', payload)
            
        return True
    except Exception as e:
        st.error(f"❌ 寫入失敗: {e}")
        return False

# ==========================================
# 2. 解析函數
# ==========================================
def parse_stone_report(f):
    if not f: return {}, None
    unit_counts = {}
    report_date = None
    try:
        f.seek(0)
        df_top = pd.read_excel(f, header=None, nrows=15)
        text = df_top.to_string()
        date_match = re.search(r'(?:至|~|迄)\s*(\d{3})(\d{2})(\d{2})', text)
        if date_match:
            y, m, d = map(int, date_match.groups())
            report_date = date(y + 1911, m, d)
        
        f.seek(0)
        xls = pd.ExcelFile(f)
        for s_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=s_name, header=None)
            active_unit = None
            for _, row in df.iterrows():
                row_str = " ".join(row.astype(str))
                if "舉發單位：" in row_str:
                    m = re.search(r"舉發單位：(\S+)", row_str)
                    if m: active_unit = m.group(1).strip()
                if "總計" in row_str and active_unit:
                    nums = [float(str(x).replace(',','')) for x in row if str(x).replace('.','',1).isdigit()]
                    if nums:
                        short = UNIT_MAP.get(active_unit, active_unit)
                        unit_counts[short] = unit_counts.get(short, 0) + int(nums[-1])
                        active_unit = None
        return unit_counts, report_date
    except: return {}, None

# ==========================================
# 3. 主程式流程
# ==========================================
files = st.file_uploader("上傳 3 個 stoneCnt 報表", accept_multiple_files=True, type=['xlsx', 'xls'], key="uploader_v27")

if files and len(files) >= 3:
    try:
        f_week, f_ytd, f_lytd = None, None, None
        for f in files:
            if "(1)" in f.name: f_ytd = f
            elif "(2)" in f.name: f_lytd = f
            else: f_week = f
        
        d_wk, _ = parse_stone_report(f_week)
        d_yt, end_dt = parse_stone_report(f_ytd)
        d_ly, _ = parse_stone_report(f_lytd)

        # 1. 各單位數據列
        body_rows = []
        for u in UNIT_DATA_ORDER:
            yt = d_yt.get(u, 0)
            target = TARGETS.get(u, 0)
            rate = f"{yt/target:.0%}" if target > 0 else "—"
            body_rows.append({
                '統計期間': u, '本期': d_wk.get(u, 0), '本年累計': yt, '去年累計': d_ly.get(u, 0),
                '本年與去年同期比較': yt - d_ly.get(u, 0), '目標值': target, '達成率': rate
            })
        
        # 2. 合計列
        df_temp = pd.DataFrame(body_rows)
        sum_data = df_temp[df_temp['統計期間'] != '警備隊'][['本期', '本年累計', '去年累計', '目標值']].sum()
        total_rate = f"{sum_data['本年累計']/sum_data['目標值']:.0%}" if sum_data['目標值'] > 0 else "0%"
        total_row = pd.DataFrame([{
            '統計期間': '合計', '本期': sum_data['本期'], '本年累計': sum_data['本年累計'], '去年累計': sum_data['去年累計'],
            '本年與去年同期比較': sum_data['本年累計'] - sum_data['去年累計'], '目標值': sum_data['目標值'], '達成率': total_rate
        }])

        # 3. 最終組合 (合計置頂)
        df_final = pd.concat([total_row, df_temp], ignore_index=True)

        st.success("✅ 解析成功")
        st.dataframe(df_final, use_container_width=True, hide_index=True)

        # 自動化
        if "v27_executed" not in st.session_state:
            with st.status("🚀 執行寫入...") as s:
                if update_sheet_from_a2(df_final, GOOGLE_SHEET_URL):
                    st.write("✅ A2 標題與 A3 合計寫入成功")
                s.update(label="完成", state="complete")
            st.session_state.v27_executed = True

    except Exception as e:
        st.error(f"錯誤：{e}")
