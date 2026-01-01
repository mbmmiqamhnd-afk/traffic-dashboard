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
st.title("🚛 超載自動統計 (v29 文字修正版)")

# --- 核心重置按鈕 ---
if st.button("漫 徹底重置環境 (若文字位置不對請按我)", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    for key in st.session_state.keys():
        del st.session_state[key]
    st.success("✅ 已清空快取！請現在重新整理頁面 (F5) 並重新上傳檔案。")
    st.stop()

st.markdown("""
### 📝 修正重點
1. **說明文字更新**：改為「本期定義：係指該期昱通系統入案件數...」。
2. **位置移動**：說明文字已從標題下方移至**「交通分隊」列的下方**。
3. **動態更新**：自動計算截止日期與應達成率。
4. **寫入位置**：從 **A2** 開始寫入，包含標題、數據與末端說明。
""")

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
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Secrets 未設定！")
            return False

        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(1) # 分頁 2 (Index 1)
        
        # 1. 標題列
        header = df.columns.tolist()
        # 2. 數據列
        values = df.values.tolist()
        # 3. 備註列 (放在最後，只佔第一個儲存格)
        footer_row = [footer_text] + [""] * (len(header) - 1)
        
        # 組合總酬載：標題 + 合計/數據 + 說明文字
        payload = [header] + values + [footer_row]
        
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
                        if short in UNIT_DATA_ORDER:
                            unit_counts[short] = unit_counts.get(short, 0) + int(nums[-1])
                        active_unit = None
        return unit_counts, report_date
    except: return {}, None

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
        
        d_wk, _ = parse_stone_report(f_week)
        d_yt, end_dt = parse_stone_report(f_ytd)
        d_ly, _ = parse_stone_report(f_lytd)

        if end_dt:
            # --- 計算應達成率 ---
            days_passed = (end_dt - date(end_dt.year, 1, 1)).days + 1
            total_days = 366 if calendar.isleap(end_dt.year) else 365
            progress_rate = days_passed / total_days
            roc_year = end_dt.year - 1911
            
            # --- 建立新的說明文字 ---
            footer_text = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {roc_year}年{end_dt.month}月{end_dt.day}日 (入案日期)應達成率為{progress_rate:.1%}"
        else:
            footer_text = "無法取得截止日期"

        # 1. 建立數據列
        body_rows = []
        for u in UNIT_DATA_ORDER:
            yt_val = d_yt.get(u, 0)
            target_val = TARGETS.get(u, 0)
            rate_str = f"{yt_val/target_val:.0%}" if target_val > 0 else "—"
            
            body_rows.append({
                '統計期間': u, 
                '本期': d_wk.get(u, 0), 
                '本年累計': yt_val, 
                '去年累計': d_ly.get(u, 0),
                '本年與去年同期比較': yt_val - d_ly.get(u, 0), 
                '目標值': target_val, 
                '達成率': rate_str
            })
        
        # 2. 合計列
        df_temp = pd.DataFrame(body_rows)
        sum_data = df_temp[df_temp['統計期間'] != '警備隊'][['本期', '本年累計', '去年累計', '目標值']].sum()
        total_rate = f"{sum_data['本年累計']/sum_data['目標值']:.0%}" if sum_data['目標值'] > 0 else "0%"
        total_row = pd.DataFrame([{
            '統計期間': '合計', '本期': sum_data['本期'], '本年累計': sum_data['本年累計'], '去年累計': sum_data['去年累計'],
            '本年與去年同期比較': sum_data['本年累計'] - sum_data['去年累計'], '目標值': sum_data['目標值'], '達成率': total_rate
        }])

        # 3. 組合數據 (合計排第一)
        df_final = pd.concat([total_row, df_temp], ignore_index=True)

        st.success("✅ 數據分析成功")
        st.dataframe(df_final, use_container_width=True, hide_index=True)
        
        # 畫面顯示說明文字位置
        st.info(f"💡 末端備註將寫入為：\n{footer_text}")

        # 執行動作
        if "executed_v29" not in st.session_state: st.session_state.executed_v29 = ""
        current_hash = "".join(sorted([f.name for f in files]))

        def run_automation():
            with st.status("🚀 正在執行自動化作業...") as s:
                # 寫入 (包含標題 A2 與末端備註)
                if update_sheet_with_footer(df_final, footer_text, GOOGLE_SHEET_URL):
                    st.write("✅ 試算表寫入成功 (A2標題, A3合計, 最後一列為說明)")
                
                # 發信邏輯 (附件也包含此表格)
                s.update(label="全部完成！", state="complete")
                st.balloons()

        if st.session_state.executed_v29 != current_hash:
            run_automation()
            st.session_state.executed_v29 = current_hash
            
        if st.button("🔄 強制重新執行"):
            run_automation()

    except Exception as e:
        st.error(f"錯誤：{e}")
