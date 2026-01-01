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

# 強制清除快取
try:
    st.cache_data.clear()
    st.cache_resource.clear()
except: pass

st.set_page_config(page_title="超載統計", layout="wide", page_icon="🚛")
st.title("🚛 超載 (stoneCnt) 自動統計 - 修正版")

# --- 核心重置按鈕 ---
if st.button("🧹 徹底重置環境 (若欄位或順序不對請點我)", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    for key in st.session_state.keys():
        del st.session_state[key]
    st.success("✅ 已清空快取！請現在重新整理頁面 (F5) 並重新上傳檔案。")
    st.stop()

st.markdown("""
### 📝 修正說明
1. **刪除科技執法**：已移除交通組資料，不再進行統計。
2. **統計期間**：第一欄標題鎖定。
3. **合計置頂**：數據從 A3 開始為合計。
4. **寫入位置**：從 **A2** 開始寫入 (含標題)。
""")

# ==========================================
# 0. 設定區 (已移除科技執法)
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

# 單位目標值 (已移除科技執法)
TARGETS = {
    '聖亭所': 24, '龍潭所': 32, '中興所': 24, '石門所': 19, 
    '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30
}

# 單位名稱轉換 (已移除交通組)
UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', 
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', 
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'
}

# 報表顯示順序 (合計之後的順序，已移除科技執法)
UNIT_DATA_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. Google Sheets 核心寫入函數
# ==========================================
def update_sheet_from_a2(df, sheet_url):
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Secrets 未設定！")
            return False

        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        # 鎖定第 2 個分頁 (Index 1)
        ws = sh.get_worksheet(1) 
        
        # 建構寫入陣列：標題 + 數據
        header = df.columns.tolist()
        values = df.values.tolist()
        payload = [header] + values
        
        # 從 A2 開始寫入
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
                        # 如果單位不在我們的清單中(例如交通組)，就不統計
                        if short in UNIT_DATA_ORDER:
                            unit_counts[short] = unit_counts.get(short, 0) + int(nums[-1])
                        active_unit = None
        return unit_counts, report_date
    except: return {}, None

# ==========================================
# 3. 郵件發送
# ==========================================
def send_email_report(excel_bytes):
    try:
        if "email" not in st.secrets: return
        sender = st.secrets["email"]["user"]
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 超載統計報表(不含科技執法) - {date.today()}"
        msg['From'] = sender
        msg['To'] = sender
        msg.attach(MIMEText("自動報表發送。", 'plain'))
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(excel_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename=Overload_Report.xlsx')
        msg.attach(part)
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender, st.secrets["email"]["password"])
            server.send_message(msg)
    except: pass

# ==========================================
# 4. 主程式流程
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

        # 1. 建立各單位數據
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
        
        # 2. 計算合計列 (置頂用)
        df_temp = pd.DataFrame(body_rows)
        # 排除警備隊來算合計目標值與達成率
        sum_data = df_temp[df_temp['統計期間'] != '警備隊'][['本期', '本年累計', '去年累計', '目標值']].sum()
        total_rate = f"{sum_data['本年累計']/sum_data['目標值']:.0%}" if sum_data['目標值'] > 0 else "0%"
        
        total_row = pd.DataFrame([{
            '統計期間': '合計', 
            '本期': sum_data['本期'], 
            '本年累計': sum_data['本年累計'], 
            '去年累計': sum_data['去年累計'],
            '本年與去年同期比較': sum_data['本年累計'] - sum_data['去年累計'],
            '目標值': sum_data['目標值'],
            '達成率': total_rate
        }])

        # 3. 最終組合 (合計排在第 1 列)
        df_final = pd.concat([total_row, df_temp], ignore_index=True)

        st.success("✅ 數據分析成功")
        st.subheader("📋 報表預覽 (已移除科技執法)")
        st.dataframe(df_final, use_container_width=True, hide_index=True)

        # 執行寫入
        if "executed_files" not in st.session_state: st.session_state.executed_files = ""
        current_hash = "".join(sorted([f.name for f in files]))

        def run_automation():
            with st.status("🚀 正在執行自動化作業...") as s:
                # 寫入 (從 A2 開始，包含標題)
                if update_sheet_from_a2(df_final, GOOGLE_SHEET_URL):
                    st.write("✅ 試算表 A2 寫入成功 (含標題，合計置頂)")
                
                # 發送郵件
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                    df_final.to_excel(wr, index=False)
                send_email_report(out.getvalue())
                
                s.update(label="全部完成！", state="complete")
                st.balloons()

        if st.session_state.executed_files != current_hash:
            run_automation()
            st.session_state.executed_files = current_hash
            
        if st.button("🔄 強制重新執行"):
            run_automation()

        # 下載按鈕
        out_excel = io.BytesIO()
        with pd.ExcelWriter(out_excel, engine='xlsxwriter') as wr:
            df_final.to_excel(wr, index=False)
        st.download_button("📥 下載超載統計 Excel", out_excel.getvalue(), "Overload_Report.xlsx")

    except Exception as e:
        st.error(f"錯誤：{e}")
