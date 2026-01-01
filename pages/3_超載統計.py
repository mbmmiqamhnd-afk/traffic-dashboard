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
st.title("🚛 超載自動統計 (v24 終極修正版)")

# --- 強制清除快取按鈕 ---
if st.button("🧹 徹底清除快取 (若欄位名稱或位置不對請按我)", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.success("快取已清除！請重新整理頁面 (F5) 並重新上傳檔案。")

st.markdown("""
### 📝 v24 更新說明
1. **B3 儲存格**：寫入標題「統計期間」。
2. **B4 儲存格**：寫入「合計」的數據。
3. **B5 儲存格**：寫入「科技執法」的數據。
4. **達成率**：四捨五入至整數 (0%)。
""")

# ==========================================
# 0. 設定區
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

UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. Google Sheets 寫入函數 (連同標題寫入)
# ==========================================
def update_google_sheet(df_with_header, sheet_url, start_cell='B3'):
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ 未設定 Secrets！")
            return False

        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(1) # 分頁 2
        
        # 準備資料：標題列 + 所有數據列
        # 我們要把 DataFrame 轉換成 List of Lists
        header = df_with_header.columns.tolist()
        values = df_with_header.values.tolist()
        data_to_write = [header] + values
        
        st.write(f"正在將資料寫入 **{ws.title}** 的 **{start_cell}** 位置...")
        
        # 寫入 (gspread 自動處理範圍)
        try:
            ws.update(range_name=start_cell, values=data_to_write)
        except TypeError:
            ws.update(start_cell, data_to_write)
            
        return True
    except Exception as e:
        st.error(f"❌ 寫入失敗: {e}")
        return False

# ==========================================
# 2. 寄信函數
# ==========================================
def send_email(recipient, subject, body, file_bytes, filename):
    try:
        if "email" not in st.secrets: return False
        sender = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(file_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=Header(filename, 'utf-8').encode())
        msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, recipient, msg.as_string())
        server.quit()
        return True
    except: return False

# ==========================================
# 3. 解析函數
# ==========================================
def parse_stone(f):
    if not f: return {}, None
    counts = {}
    found_date = None
    try:
        f.seek(0)
        df_head = pd.read_excel(f, header=None, nrows=20)
        text_content = df_head.to_string()
        match = re.search(r'(?:至|~|迄)\s*(\d{3})(\d{2})(\d{2})', text_content)
        if match:
            y, m, d = map(int, match.groups())
            found_date = date(y + 1911, m, d)
        
        f.seek(0)
        xls = pd.ExcelFile(f)
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            curr_unit = None
            for _, row in df.iterrows():
                row_str = row.astype(str).str.cat(sep=' ')
                if "舉發單位：" in row_str:
                    m = re.search(r"舉發單位：(\S+)", row_str)
                    if m: curr_unit = m.group(1).strip()
                if "總計" in row_str and curr_unit:
                    nums = [float(str(x).replace(',','')) for x in row if str(x).replace('.','',1).isdigit()]
                    if nums:
                        short_name = UNIT_MAP.get(curr_unit, curr_unit)
                        counts[short_name] = counts.get(short_name, 0) + int(nums[-1])
                        curr_unit = None
        return counts, found_date
    except: return {}, None

# ==========================================
# 4. 主程式執行
# ==========================================
uploaded_files = st.file_uploader("上傳 3 個 stoneCnt 報表", accept_multiple_files=True, type=['xlsx', 'xls'], key="stone_v24_key")

if uploaded_files and len(uploaded_files) >= 3:
    try:
        files_map = {"Week": None, "YTD": None, "Last_YTD": None}
        for f in uploaded_files:
            if "(1)" in f.name: files_map["YTD"] = f
            elif "(2)" in f.name: files_map["Last_YTD"] = f
            else: files_map["Week"] = f
        
        d_wk, _ = parse_stone(files_map["Week"])
        d_yt, end_date = parse_stone(files_map["YTD"])
        d_ly, _ = parse_stone(files_map["Last_YTD"])

        prog_text = ""
        if end_date:
            days = (end_date - date(end_date.year, 1, 1)).days + 1
            total = 366 if end_date.year % 4 == 0 else 365
            prog_text = f"統計截至 {end_date.year-1911}年{end_date.month}月{end_date.day}日，進度 {days/total:.1%}"
            st.info(f"📅 {prog_text}")

        # 1. 建立各單位數據 (第一欄直接叫 統計期間)
        unit_rows = []
        for u in UNIT_ORDER:
            w, y, l = d_wk.get(u,0), d_yt.get(u,0), d_ly.get(u,0)
            tgt = TARGETS.get(u,0)
            if u == '警備隊': w, y, l, tgt = 0, 0, 0, 0
            
            # 四捨五入到整數
            rate_str = f"{y/tgt:.0%}" if tgt > 0 else "—"
            unit_rows.append({
                '統計期間': u, '本期': w, '本年累計': y, '去年累計': l, 
                '本年與去年同期比較': y - l, '目標值': tgt, '達成率': rate_str
            })
        
        # 2. 計算合計
        df_temp = pd.DataFrame(unit_rows)
        s = df_temp[['本期', '本年累計', '去年累計', '目標值']].sum()
        total_rate = f"{s['本年累計']/s['目標值']:.0%}" if s['目標值'] > 0 else "0%"
        total_row = {
            '統計期間': '合計', '本期': s['本期'], '本年累計': s['本年累計'], '去年累計': s['去年累計'], 
            '本年與去年同期比較': s['本年累計'] - s['去年累計'], '目標值': s['目標值'], '達成率': total_rate
        }
        
        # 3. 組合：合計置頂 (Row 0 是 合計)
        df_final = pd.concat([pd.DataFrame([total_row]), df_temp], ignore_index=True)
        
        st.success("✅ 分析完成！")
        
        # 顯示預覽
        st.subheader("📋 報表結構預覽")
        st.write("寫入 B3：標題「統計期間」等")
        st.write("寫入 B4：合計數據")
        st.dataframe(df_final, use_container_width=True, hide_index=True)

        # Excel 產生
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Sheet1', startrow=3)
            ws = writer.sheets['Sheet1']
            ws.write('A1', '超載取締統計表')
            ws.write('A2', prog_text)
        excel_data = output.getvalue()

        st.markdown("---")
        if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
        f_ids = ",".join(sorted([f.name for f in uploaded_files]))
        
        def run_auto():
            with st.status("🚀 執行自動化程序...") as status:
                # 寄信
                email = st.secrets["email"]["user"] if "email" in st.secrets else None
                if email: send_email(email, "📊 超載統計自動報表", "附件為超載報表。", excel_data, "超載統計.xlsx")
                
                # 寫入 (連標題一起寫)
                if update_google_sheet(df_final, GOOGLE_SHEET_URL, 'B3'):
                    st.write("✅ 試算表更新成功！B3=統計期間, B4=合計")
                else:
                    st.write("❌ 試算表更新失敗")
                status.update(label="全部執行完畢", state="complete")
                st.balloons()
        
        if f_ids not in st.session_state["sent_cache"]:
            run_auto()
            st.session_state["sent_cache"].add(f_ids)
            
        if st.button("🔄 強制重新執行"):
            run_auto()
            
        st.download_button("📥 下載 Excel", excel_data, "超載統計.xlsx")

    except Exception as e: st.error(f"錯誤：{e}")
