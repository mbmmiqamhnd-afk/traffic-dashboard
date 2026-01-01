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
st.title("🚛 超載自動統計 (v32 網址鎖定版)")

# --- 核心重置按鈕 ---
if st.button("🧹 徹底重置環境 (若網頁沒反應請點我)", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    for key in st.session_state.keys():
        del st.session_state[key]
    st.success("✅ 已清空快取！請現在重新整理頁面 (F5)。")
    st.stop()

# ==========================================
# 0. 設定區 (網址已直接寫入)
# ==========================================
# 鎖定您的試算表網址
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
# 1. 核心寫入函數 (A2 起始)
# ==========================================
def update_sheet_final(df, footer_text, sheet_url):
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(1) # 鎖定分頁 2 (Index 1)
        
        # 標題
        header = df.columns.tolist()
        # 數據
        values = df.values.tolist()
        # 備註列
        footer_row = [footer_text] + [""] * (len(header) - 1)
        
        # 打包所有資料
        payload = [header] + values + [footer_row]
        
        # 執行寫入 (從 A2 開始，包含標題)
        ws.update(range_name='A2', values=payload)
        return True
    except Exception as e:
        st.error(f"❌ 試算表連動失敗: {e}")
        return False

# ==========================================
# 2. 解析函數
# ==========================================
def parse_stone_report(f):
    if not f: return {}, "0000000", "0000000"
    unit_counts = {}
    s_str, e_str = "0000000", "0000000"
    try:
        f.seek(0)
        df_top = pd.read_excel(f, header=None, nrows=15)
        text = df_top.to_string()
        
        # 抓取入案日期區間
        date_match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', text)
        if date_match:
            s_str, e_str = date_match.group(1), date_match.group(2)
        
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
        return unit_counts, s_str, e_str
    except: return {}, "0000000", "0000000"

# ==========================================
# 3. 主程式流程
# ==========================================
files = st.file_uploader("請上傳 3 個 stoneCnt 報表檔案 (週報、本年、去年)", accept_multiple_files=True, type=['xlsx', 'xls'])

if files and len(files) >= 3:
    try:
        f_week, f_ytd, f_lytd = None, None, None
        for f in files:
            if "(1)" in f.name: f_ytd = f
            elif "(2)" in f.name: f_lytd = f
            else: f_week = f
        
        # 解析數據
        d_wk, wk_s, wk_e = parse_stone_report(f_week)
        d_yt, yt_s, yt_e = parse_stone_report(f_ytd)
        d_ly, ly_s, ly_e = parse_stone_report(f_lytd)

        # 動態標題命名
        col_wk = f"本期 ({wk_s[-4:]}~{wk_e[-4:]})"
        col_yt = f"本年累計 ({yt_s}~{yt_e})"
        col_ly = f"去年累計 ({ly_s}~{ly_e})"

        # 說明文字計算
        try:
            y, m, d = int(yt_e[:3])+1911, int(yt_e[3:5]), int(yt_e[5:])
            end_dt = date(y, m, d)
            progress = ((end_dt - date(y, 1, 1)).days + 1) / (366 if calendar.isleap(y) else 365)
            footer_text = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {yt_e[:3]}年{yt_e[3:5]}月{yt_e[5:]}日 (入案日期)應達成率為{progress:.1%}"
        except: footer_text = "日期格式有誤，請檢查報表標頭。"

        # 建立數據
        body = []
        for u in UNIT_DATA_ORDER:
            yt_v = d_yt.get(u, 0)
            target = TARGETS.get(u, 0)
            # 達成率整數化
            rate_str = f"{yt_v/target:.0%}" if target > 0 else "—"
            body.append({
                '統計期間': u, 
                col_wk: d_wk.get(u, 0), 
                col_yt: yt_v, 
                col_ly: d_ly.get(u, 0),
                '本年與去年同期比較': yt_v - d_ly.get(u, 0), 
                '目標值': target, 
                '達成率': rate_str
            })
        
        # 合計列
        df_temp = pd.DataFrame(body)
        sum_data = df_temp[df_temp['統計期間'] != '警備隊'][[col_wk, col_yt, col_ly, '目標值']].sum()
        total_rate = f"{sum_data[col_yt]/sum_data['目標值']:.0%}" if sum_data['目標值'] > 0 else "0%"
        total_row = pd.DataFrame([{
            '統計期間': '合計', 
            col_wk: sum_data[col_wk], 
            col_yt: sum_data[col_yt], 
            col_ly: sum_data[col_ly],
            '本年與去年同期比較': sum_data[col_yt] - sum_data[col_ly],
            '目標值': sum_data['目標值'],
            '達成率': total_rate
        }])

        df_final = pd.concat([total_row, df_temp], ignore_index=True)

        st.success("✅ 數據解析完成")
        st.dataframe(df_final, use_container_width=True, hide_index=True)
        st.info(f"💡 備註內容：\n{footer_text}")

        # 寫入按鈕
        st.markdown("---")
        if st.button("🚀 執行寫入：同步至 Google 試算表", type="primary"):
            with st.status("正在同步數據...") as s:
                if update_sheet_final(df_final, footer_text, GOOGLE_SHEET_URL):
                    st.write(f"✅ 成功寫入試算表：{GOOGLE_SHEET_URL}")
                    st.balloons()
                s.update(label="同步結束", state="complete")

        # 下載
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
            df_final.to_excel(wr, index=False)
        st.download_button("📥 下載 Excel 報表", out.getvalue(), f"超載統計_{yt_e}.xlsx")

    except Exception as e:
        st.error(f"執行出錯：{e}")
