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

# 強制清除快取，防止舊邏輯干擾
try:
    st.cache_data.clear()
    st.cache_resource.clear()
except: pass

st.set_page_config(page_title="超載統計", layout="wide", page_icon="🚛")
st.title("🚛 超載自動統計 (v25 破壞性重寫版)")

# --- 核心快取清除按鈕 ---
if st.button("🧹 徹底重置程式環境 (若 A2/A3 順序不對請按我)", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.success("環境已清空！請重新整理頁面 (F5) 並上傳檔案。")

st.markdown("""
### 📝 寫入邏輯說明 (A2 起始強制覆蓋)
1. **A2 儲存格**：寫入標題「統計期間」。
2. **A3 儲存格**：寫入「合計」數據。
3. **A4 儲存格**：寫入「科技執法」數據。
4. **達成率**：四捨五入至整數百分比 (如 85%)。
5. **分頁鎖定**：鎖定 Google 試算表之 **第 2 個分頁 (Index 1)**。
""")

# ==========================================
# 0. 設定與參數區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

TARGETS = {
    '科技執法': 0, '聖亭所': 24, '龍潭所': 32, '中興所': 24, 
    '石門所': 19, '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30
}

UNIT_MAP = {
    '交通組': '科技執法', '交通組(科技執法)': '科技執法', '聖亭派出所': '聖亭所', 
    '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', 
    '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'
}

# 數據排序順序 (合計之後的順序)
UNIT_DATA_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. Google Sheets 核心寫入函數
# ==========================================
def destructive_update_sheet(df, sheet_url, start_cell='A2'):
    """連同標題一起寫入，確保起始格為 A2"""
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ 未設定 Secrets！")
            return False

        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(1) # 指定第 2 個分頁
        
        # 建構寫入陣列：標題 + 數據
        header_list = df.columns.tolist()
        values_list = df.values.tolist()
        final_payload = [header_list] + values_list
        
        # 強制寫入 A2
        try:
            ws.update(range_name=start_cell, values=final_payload)
        except TypeError:
            ws.update(start_cell, final_payload)
            
        return True
    except Exception as e:
        st.error(f"❌ 寫入失敗: {e}")
        return False

# ==========================================
# 2. stoneCnt 報表解析
# ==========================================
def parse_stone_report(f):
    if not f: return {}, None
    unit_counts = {}
    report_date = None
    try:
        f.seek(0)
        # 嘗試尋找統計日期
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
# 3. 郵件發送
# ==========================================
def send_report_mail(recipient, excel_bytes, file_name):
    try:
        if "email" not in st.secrets: return
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 超載統計自動報表 - {date.today()}"
        msg['From'] = sender
        msg['To'] = recipient
        msg.attach(MIMEText("附件為最新計算之超載統計報表。", 'plain'))
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(excel_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={file_name}')
        msg.attach(part)
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender, pwd)
            server.send_message(msg)
    except: pass

# ==========================================
# 4. 主執行程序
# ==========================================
files = st.file_uploader("上傳 3 個 stoneCnt Excel 檔案", accept_multiple_files=True, type=['xlsx', 'xls'], key="uploader_v25")

if files and len(files) >= 3:
    try:
        # 分類檔案
        f_week, f_ytd, f_lytd = None, None, None
        for f in files:
            if "(1)" in f.name: f_ytd = f
            elif "(2)" in f.name: f_lytd = f
            else: f_week = f
        
        # 解析數據
        d_wk, _ = parse_stone_report(f_week)
        d_yt, end_dt = parse_stone_report(f_ytd)
        d_ly, _ = parse_stone_report(f_lytd)

        if end_dt:
            st.info(f"📅 數據截至：{end_dt.year-1911}年{end_dt.month}月{end_dt.day}日")

        # --- 數據組裝 (核心變動處) ---
        # 1. 先算出所有單位的數據列
        body_rows = []
        for u in UNIT_DATA_ORDER:
            wk = d_wk.get(u, 0)
            yt = d_yt.get(u, 0)
            ly = d_ly.get(u, 0)
            target = TARGETS.get(u, 0)
            
            # 達成率整數化
            rate = f"{yt/target:.0%}" if target > 0 else "—"
            
            body_rows.append({
                '統計期間': u, '本期': wk, '本年累計': yt, '去年累計': ly,
                '本年與去年同期比較': yt - ly, '目標值': target, '達成率': rate
            })
        
        # 2. 計算合計列
        df_temp = pd.DataFrame(body_rows)
        # 警備隊不計入合計的目標與達成
        sum_cols = df_temp[df_temp['統計期間'] != '警備隊'][['本期', '本年累計', '去年累計', '目標值']].sum()
        total_rate = f"{sum_cols['本年累計']/sum_cols['目標值']:.0%}" if sum_cols['目標值'] > 0 else "0%"
        
        total_row = pd.DataFrame([{
            '統計期間': '合計', 
            '本期': sum_cols['本期'], 
            '本年累計': sum_cols['本年累計'], 
            '去年累計': sum_cols['去年累計'],
            '本年與去年同期比較': sum_cols['本年累計'] - sum_cols['去年累計'],
            '目標值': sum_cols['目標值'],
            '達成率': total_rate
        }])

        # 3. 最終大組合：合計(A3) + 數據(A4...)
        df_final = pd.concat([total_row, df_temp], ignore_index=True)

        st.success("✅ 數據處理完畢")
        st.subheader("📋 即將寫入試算表之預覽 (起始 A2)")
        st.dataframe(df_final, use_container_width=True, hide_index=True)

        # --- 自動化執行 ---
        st.markdown("---")
        if "v25_done" not in st.session_state: st.session_state.v25_done = set()
        file_hash = "".join(sorted([f.name for f in files]))

        def execute_auto():
            with st.status("🚀 正在執行自動化作業...") as status:
                # 1. 寫入 Google Sheet (含標題從 A2 開始)
                st.write("📊 寫入試算表 (A2=標題, A3=合計, A4=科技執法)...")
                if destructive_update_sheet(df_final, GOOGLE_SHEET_URL, 'A2'):
                    st.write("✅ 寫入成功")
                
                # 2. 發信
                st.write("📧 發送電子郵件報表...")
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                    df_final.to_excel(wr, index=False, sheet_name='統計結果')
                send_report_mail(st.secrets["email"]["user"], out.getvalue(), "Overload_Report.xlsx")
                
                status.update(label="完成！", state="complete")
                st.balloons()

        if file_hash not in st.session_state.v25_done:
            execute_auto()
            st.session_state.v25_done.add(file_hash)
            
        if st.button("🔄 手動強制重新執行"):
            execute_auto()

        # 下載按鈕
        out_btn = io.BytesIO()
        with pd.ExcelWriter(out_btn, engine='xlsxwriter') as wr:
            df_final.to_excel(wr, index=False, sheet_name='統計結果')
        st.download_button("📥 下載統計 Excel", out_btn.getvalue(), "超載報表.xlsx")

    except Exception as e:
        st.error(f"執行出錯：{e}")
