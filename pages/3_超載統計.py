import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import smtplib
import gspread
from datetime import date
import calendar
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# 強制清除快取
try:
    st.cache_data.clear()
    st.cache_resource.clear()
except: pass

st.set_page_config(page_title="超載統計", layout="wide", page_icon="🚛")
st.title("🚛 超載自動統計 (v30 欄位動態日期版)")

# --- 核心重置按鈕 ---
if st.button("🧹 徹底重置環境 (若日期抓取不對請按我)", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    for key in st.session_state.keys():
        del st.session_state[key]
    st.success("✅ 已清空快取！請重新整理頁面 (F5) 並重新上傳檔案。")
    st.stop()

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

TARGETS = {
    '聖亭所': 24, '龍潭所': 32, '中興所': 24, '石門所': 19, 
    '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30
}

UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', 
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', 
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'
}

UNIT_DATA_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. Google Sheets 核心寫入函數
# ==========================================
def update_sheet_with_footer(df, footer_text, sheet_url):
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(1) 
        
        header = df.columns.tolist()
        values = df.values.tolist()
        footer_row = [footer_text] + [""] * (len(header) - 1)
        
        payload = [header] + values + [footer_row]
        
        try:
            ws.update(range_name='A2', values=payload)
        except:
            ws.update('A2', payload)
        return True
    except Exception as e:
        st.error(f"❌ 寫入失敗: {e}")
        return False

# ==========================================
# 2. 解析函數 (升級：抓取起始與結束日期)
# ==========================================
def parse_stone_report(f):
    if not f: return {}, None, None
    unit_counts = {}
    start_str, end_str = "0000000", "0000000"
    
    try:
        f.seek(0)
        df_top = pd.read_excel(f, header=None, nrows=15)
        text = df_top.to_string()
        
        # 抓取「入案日期：XXX 至 YYY」
        date_match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', text)
        if date_match:
            start_str, end_str = date_match.group(1), date_match.group(2)
        
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
                        if short in UNIT_DATA_ORDER:
                            unit_counts[short] = unit_counts.get(short, 0) + int(nums[-1])
                        active_unit = None
        return unit_counts, start_str, end_str
    except: return {}, None, None

# ==========================================
# 3. 主程式流程
# ==========================================
files = st.file_uploader("請上傳 3 個 stoneCnt 報表檔案", accept_multiple_files=True, type=['xlsx', 'xls'])

if files and len(files) >= 3:
    try:
        f_week, f_ytd, f_lytd = None, None, None
        for f in files:
            if "(1)" in f.name: f_ytd = f
            elif "(2)" in f.name: f_lytd = f
            else: f_week = f
        
        # 解析三個檔案的數據與日期區間
        d_wk, wk_s, wk_e = parse_stone_report(f_week)
        d_yt, yt_s, yt_e = parse_stone_report(f_ytd)
        d_ly, ly_s, ly_e = parse_stone_report(f_lytd)

        # 定義動態欄位名稱
        # 本期顯示：(月日~月日) -> 取 7 位數 ROC 日期的後四碼
        col_name_wk = f"本期 ({wk_s[-4:]}~{wk_e[-4:]})"
        # 本年與去年顯示：(年月日~年月日)
        col_name_yt = f"本年累計 ({yt_s}~{yt_e})"
        col_name_ly = f"去年累計 ({ly_s}~{ly_e})"

        # 計算年度進度說明文字
        footer_text = ""
        try:
            y, m, d = int(yt_e[:3])+1911, int(yt_e[3:5]), int(yt_e[5:])
            end_dt = date(y, m, d)
            days_passed = (end_dt - date(y, 1, 1)).days + 1
            total_days = 366 if calendar.isleap(y) else 365
            progress_rate = days_passed / total_days
            footer_text = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {yt_e[:3]}年{yt_e[3:5]}月{yt_e[5:]}日 (入案日期)應達成率為{progress_rate:.1%}"
        except:
            footer_text = "日期格式解析錯誤"

        # 1. 建立數據列
        body_rows = []
        for u in UNIT_DATA_ORDER:
            yt_val = d_yt.get(u, 0)
            target_val = TARGETS.get(u, 0)
            rate_str = f"{yt_val/target_val:.0%}" if target_val > 0 else "—"
            
            body_rows.append({
                '統計期間': u, 
                col_name_wk: d_wk.get(u, 0), 
                col_name_yt: yt_val, 
                col_name_ly: d_ly.get(u, 0),
                '本年與去年同期比較': yt_val - d_ly.get(u, 0), 
                '目標值': target_val, 
                '達成率': rate_str
            })
        
        # 2. 合計列
        df_temp = pd.DataFrame(body_rows)
        sum_data = df_temp[df_temp['統計期間'] != '警備隊'][[col_name_wk, col_name_yt, col_name_ly, '目標值']].sum()
        total_rate = f"{sum_data[col_name_yt]/sum_data['目標值']:.0%}" if sum_data['目標值'] > 0 else "0%"
        
        total_row = pd.DataFrame([{
            '統計期間': '合計', 
            col_name_wk: sum_data[col_name_wk], 
            col_name_yt: sum_data[col_name_yt], 
            col_name_ly: sum_data[col_name_ly],
            '本年與去年同期比較': sum_data[col_name_yt] - sum_data[col_name_ly],
            '目標值': sum_data['目標值'],
            '達成率': total_rate
        }])

        # 3. 最終組合
        df_final = pd.concat([total_row, df_temp], ignore_index=True)

        st.success("✅ 數據解析成功")
        st.dataframe(df_final, use_container_width=True, hide_index=True)
        
        # 執行動作
        if "executed_v30" not in st.session_state: st.session_state.executed_v30 = ""
        current_hash = "".join(sorted([f.name for f in files]))

        if st.button("🚀 執行寫入與自動化程序"):
            with st.status("正在處理中...") as s:
                if update_sheet_with_footer(df_final, footer_text, GOOGLE_SHEET_URL):
                    st.write(f"✅ 試算表寫入成功！欄位日期已更新。")
                s.update(label="全部完成！", state="complete")
                st.balloons()

    except Exception as e:
        st.error(f"錯誤：{e}")
