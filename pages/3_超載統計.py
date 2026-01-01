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
st.title("🚛 超載 (stoneCnt) 自動統計 (v20 絕對修正版)")

# --- 強制清除快取按鈕 ---
if st.button("🧹 清除快取 (若更新無效請按此)", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.success("快取已清除！請重新整理頁面 (F5) 並重新上傳檔案。")

st.markdown("""
### 📝 使用說明
1. 請上傳 **3 個** `stoneCnt` 系列的 Excel 檔案。
2. **第一欄名稱已修正為「統計期間」**。
3. **「合計」列強制排在第一位**。
4. 寫入位置：**第 2 個分頁 (Index 1) 的 B3** (純數據)。
""")

st.warning("""
⚠️ **寫入位置注意：**
程式將從 **B3** 開始寫入數據 (合計的數據)。
請確認您的 Google 試算表 **A3** 儲存格是 **「合計」**。
若 A3 是其他單位，請手動修改試算表 A 欄順序，否則數據會錯位。
""")

# ==========================================
# 0. 設定區
# ==========================================
# 請將您的 Google 試算表網址貼在這裡
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

TARGETS = {'聖亭所': 24, '龍潭所': 32, '中興所': 24, '石門所': 19, '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30}
UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'}
UNIT_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. Google Sheets 寫入函數
# ==========================================
def update_google_sheet(df, sheet_url, start_cell='B3'): # <--- 確認為 B3
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ 未設定 GCP Service Account Secrets！")
            return False

        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        
        # 抓取第 2 個工作表 (Index 1)
        try:
            ws = sh.get_worksheet(1) 
            if ws is None: raise Exception("找不到 Index 1 的工作表")
        except Exception as e:
            st.error(f"❌ 無法取得第 2 個工作表 (Index 1)，請確認試算表是否有至少 2 個分頁。錯誤: {e}")
            return False
        
        st.info(f"📂 寫入目標工作表：**「{ws.title}」**")
        
        # 準備寫入資料 (純數據，無標題)
        df_clean = df.fillna("").replace([np.inf, -np.inf], 0)
        
        # 轉成 List (不含 Header)
        data = df_clean.values.tolist()
        
        # 寫入
        try:
            ws.update(range_name=start_cell, values=data)
        except TypeError:
            ws.update(start_cell, data)
        except Exception as e:
            st.error(f"❌ Google 試算表寫入失敗: {e}")
            return False
        return True
        
    except Exception as e:
        st.error(f"❌ 未知錯誤: {e}")
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
# 3. 資料解析函數
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
        if not match:
            match = re.search(r'(?:至|~|迄)\s*(\d{3})[./\-年](\d{1,2})[./\-月](\d{1,2})', text_content)
        
        if match:
            y, m, d = map(int, match.groups())
            if 100 <= y <= 200 and 1 <= m <= 12 and 1 <= d <= 31:
                found_date = date(y + 1911, m, d)
        
        f.seek(0)
        xls = pd.ExcelFile(f)
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            curr = None
            for _, row in df.iterrows():
                s = row.astype(str).str.cat(sep=' ')
                if "舉發單位：" in s:
                    m = re.search(r"舉發單位：(\S+)", s)
                    if m: curr = m.group(1).strip()
                if "總計" in s and curr:
                    nums = [float(x) for x in row if str(x).replace('.','',1).isdigit()]
                    if nums:
                        short = UNIT_MAP.get(curr, curr)
                        counts[short] = counts.get(short, 0) + int(nums[-1])
                        curr = None
        return counts, found_date
    except Exception as e:
        st.error(f"解析檔案 {f.name} 錯誤: {e}")
        return {}, None

# ==========================================
# 4. 主程式執行
# ==========================================
# ★★★ v20 Key ★★★
uploaded_files = st.file_uploader("請拖曳 3 個 stoneCnt 檔案至此", accept_multiple_files=True, type=['xlsx', 'xls'], key="stone_uploader_v20_final")

if uploaded_files:
    if len(uploaded_files) < 3:
        st.warning("⏳ 檔案不足 3 個，請繼續上傳...")
    else:
        try:
            files_config = {"Week": None, "YTD": None, "Last_YTD": None}
            for f in uploaded_files:
                if "(1)" in f.name: files_config["YTD"] = f
                elif "(2)" in f.name: files_config["Last_YTD"] = f
                else: files_config["Week"] = f
            
            d_wk, _ = parse_stone(files_config["Week"])
            d_yt, end_date = parse_stone(files_config["YTD"])
            d_ly, _ = parse_stone(files_config["Last_YTD"])

            prog_text = ""
            if end_date:
                start_of_year = date(end_date.year, 1, 1)
                days_passed = (end_date - start_of_year).days + 1
                total_days = 366 if (end_date.year % 4 == 0 and end_date.year % 100 != 0) or (end_date.year % 400 == 0) else 365
                progress_rate = days_passed / total_days
                prog_text = f"統計截至 {end_date.year-1911}年{end_date.month}月{end_date.day}日 (入案日期)，年度時間進度為 {progress_rate:.1%}"
                st.info(f"📅 {prog_text}")
            else:
                st.warning("⚠️ 無法從「本年累計」檔案中找到截止日期。")

            unit_rows = []
            for u in UNIT_ORDER:
                w = d_wk.get(u,0)
                y = d_yt.get(u,0)
                l = d_ly.get(u,0)
                tgt = TARGETS.get(u,0)
                
                # 警備隊數值歸零
                if u == '警備隊': w, y, l, tgt = 0, 0, 0, 0
                
                # 計算數值
                diff = y - l
                rate_str = f"{y/tgt:.0%}" if tgt > 0 else "0%"
                if u == '警備隊': rate_str = "—"

                # ★★★ Key 確保為 統計期間 ★★★
                unit_rows.append({
                    '統計期間': u,  
                    '本期': w, 
                    '本年累計': y, 
                    '去年累計': l, 
                    '本年與去年同期比較': diff,
                    '目標值': tgt,
                    '達成率': rate_str
                })
            
            # 建立單位 DataFrame
            df_units = pd.DataFrame(unit_rows)
            
            # 計算合計
            total_s = df_units[['本期', '本年累計', '去年累計', '目標值']].sum()
            total_diff = total_s['本年累計'] - total_s['去年累計']
            total_rate_str = f"{total_s['本年累計']/total_s['目標值']:.0%}" if total_s['目標值']>0 else "0%"
            
            # ★★★ 欄位名稱確保為 統計期間 ★★★
            total_row = {
                '統計期間': '合計',
                '本期': total_s['本期'],
                '本年累計': total_s['本年累計'],
                '去年累計': total_s['去年累計'],
                '本年與去年同期比較': total_diff,
                '目標值': total_s['目標值'],
                '達成率': total_rate_str
            }
            
            # ★★★ 合計置頂 ★★★
            final_rows = [total_row] + unit_rows

            cols = ['統計期間', '本期', '本年累計', '去年累計', '本年與去年同期比較', '目標值', '達成率']
            df_final = pd.DataFrame(final_rows, columns=cols)
            
            # 準備寫入的 DataFrame (移除第一欄)
            df_write = df_final.drop(columns=['統計期間'])
            
            st.success("✅ 分析完成！")
            st.dataframe(df_final, use_container_width=True, hide_index=True)
            
            # 預覽寫入內容
            st.caption("▼ 即將寫入分頁2 (B3) 的數據預覽 (無標題，第一列為合計數據)：")
            st.dataframe(df_write, use_container_width=True)

            # --- 產生 Excel ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='超載統計', startrow=3)
                workbook = writer.book
                worksheet = writer.sheets['超載統計']
                fmt_title = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center'})
                fmt_subtitle = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': 'blue', 'align': 'left'})
                worksheet.merge_range('A1:G1', '超載取締統計表', fmt_title)
                if prog_text:
                    worksheet.merge_range('A2:G2', f"說明：{prog_text}", fmt_subtitle)
                worksheet.set_column(0, 0, 15)
                worksheet.set_column(1, 6, 12)
            excel_data = output.getvalue()
            file_name_out = '超載統計表.xlsx'

            # --- 自動化與手動執行區塊 ---
            st.markdown("---")
            st.subheader("🚀 執行動作")
            
            if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
            file_ids = ",".join(sorted([f.name for f in uploaded_files]))
            
            def run_automation():
                with st.status("正在處理中...", expanded=True) as status:
                    st.write("📧 準備寄送 Email...")
                    mail_body = "附件為超載統計報表。"
                    if prog_text: mail_body += f"\n\n{prog_text}"
                    email_receiver = st.secrets["email"]["user"] if "email" in st.secrets else None
                    
                    if email_receiver:
                        if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", mail_body, excel_data, file_name_out):
                            st.write(f"✅ Email 已發送")
                    
                    st.write("📊 正在寫入 Google 試算表 (B3)...")
                    if update_google_sheet(df_write, GOOGLE_SHEET_URL, start_cell='B3'):
                        st.write("✅ Google 試算表已更新！")
                    else:
                        st.write("❌ Google 試算表更新失敗")
                    
                    status.update(label="執行完畢", state="complete", expanded=False)
                    st.balloons()
            
            if file_ids not in st.session_state["sent_cache"]:
                run_automation()
                st.session_state["sent_cache"].add(file_ids)
            else:
                st.info("✅ 已自動執行過。")

            if st.button("🔄 強制重新執行 (寫入 + 寄信)", type="primary"):
                run_automation()

            st.download_button(label="📥 下載 Excel", data=excel_data, file_name=file_name_out, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e: st.error(f"發生錯誤：{e}")
