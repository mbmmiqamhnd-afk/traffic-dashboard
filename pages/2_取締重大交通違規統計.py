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

st.set_page_config(page_title="重大交通違規統計", layout="wide", page_icon="🚨")
st.title("🚨 重大交通違規自動統計")

st.markdown("""
### 📝 使用說明
1. 請上傳 **3 個** 相關統計 Excel 檔案。
2. 系統將自動計算數據與年度時間進度。
3. 自動寄信並寫入 Google 試算表 **(寫入同一個檔案的第 1 個分頁)**。
4. **若沒反應，請點擊下方的「🔄 強制手動執行」按鈕。**
""")

# ==========================================
# 0. 設定區
# ==========================================
# ★★★ 重要：請確認這裡填入的是與「超載統計」完全一樣的 Google 試算表網址 ★★★
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

TARGETS = {'聖亭所': 24, '龍潭所': 32, '中興所': 24, '石門所': 19, '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30}
UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'}
UNIT_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. Google Sheets 寫入函數 (強力相容版)
# ==========================================
def update_google_sheet(df, sheet_url, start_cell='A3'):
    try:
        # 1. 檢查 Secrets
        if "gcp_service_account" not in st.secrets:
            st.error("❌ 錯誤：未設定 Secrets！")
            return False

        # 2. 連線測試
        try:
            gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
            sh = gc.open_by_url(sheet_url)
        except Exception as e:
            st.error(f"❌ 連線失敗 (請檢查網址或機器人權限): {e}")
            return False
        
        # 3. 抓取工作表 (鎖定第 1 個，索引為 0)
        try:
            ws = sh.get_worksheet(0) # <--- 0 代表第 1 個分頁
            if ws is None: raise Exception("找不到 Index 0 的工作表")
        except Exception as e:
            st.error(f"❌ 找不到第 1 個工作表: {e}")
            return False
        
        # 4. 準備資料
        df_clean = df.fillna("").replace([np.inf, -np.inf], 0)
        data = [df_clean.columns.values.tolist()] + df_clean.values.tolist()
        
        # 5. 執行寫入 (雙重寫法相容)
        try:
            # 新版 gspread (v6.0+)
            ws.update(range_name=start_cell, values=data)
        except TypeError:
            try:
                # 舊版 gspread
                ws.update(start_cell, data)
            except Exception as e_inner:
                st.error(f"❌ 寫入數據失敗 (舊版寫法): {e_inner}")
                return False
        except Exception as e:
            st.error(f"❌ 寫入數據失敗: {e}")
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
        if "email" not in st.secrets:
            st.error("❌ 未設定 Email Secrets！")
            return False
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
    except Exception as e:
        st.error(f"❌ 寄信失敗: {e}")
        return False

# ==========================================
# 3. 資料解析函數
# ==========================================
def parse_report(f):
    if not f: return {}, None
    counts = {}
    found_date = None
    try:
        f.seek(0)
        df_head = pd.read_excel(f, header=None, nrows=20)
        text_content = df_head.to_string()
        
        # 日期抓取 (支援分隔符號與連續數字)
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
uploaded_files = st.file_uploader("請拖曳 3 個統計檔案至此", accept_multiple_files=True, type=['xlsx', 'xls'])

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
            
            d_wk, _ = parse_report(files_config["Week"])
            d_yt, end_date = parse_report(files_config["YTD"])
            d_ly, _ = parse_report(files_config["Last_YTD"])

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

            rows = []
            for u in UNIT_ORDER:
                rows.append({
                    '單位': u, 
                    '本期': d_wk.get(u,0), 
                    '本年累計': d_yt.get(u,0), 
                    '去年累計': d_ly.get(u,0), 
                    '目標值': TARGETS.get(u,0)
                })
            
            df = pd.DataFrame(rows)
            df_calc = df.copy()
            mask_guard = df_calc['單位'] == '警備隊'
            df_calc.loc[mask_guard, ['本期', '本年累計', '去年累計', '目標值']] = 0
            
            total = df_calc[['本期', '本年累計', '去年累計', '目標值']].sum().to_dict()
            total['單位'] = '合計'
            
            df_final = pd.concat([pd.DataFrame([total]), df], ignore_index=True)
            df_final['本年與去年同期比較'] = df_final['本年累計'] - df_final['去年累計']
            df_final['達成率'] = df_final.apply(lambda x: f"{x['本年累計']/x['目標值']:.2%}" if x['目標值']>0 else "—", axis=1)
            df_final.loc[df_final['單位']=='警備隊', ['本年與去年同期比較', '目標值', '達成率']] = "—"
            
            cols = ['單位', '本期', '本年累計', '去年累計', '本年與去年同期比較', '目標值', '達成率']
            df_final = df_final[cols]
            
            st.success("✅ 分析完成！")
            st.dataframe(df_final, use_container_width=True, hide_index=True)
            
            # --- Excel 下載用 ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='交通違規統計', startrow=3)
                workbook = writer.book
                worksheet = writer.sheets['交通違規統計']
                fmt_title = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center'})
                fmt_subtitle = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': 'blue', 'align': 'left'})
                worksheet.merge_range('A1:G1', '重大交通違規統計表', fmt_title)
                if prog_text:
                    worksheet.merge_range('A2:G2', f"說明：{prog_text}", fmt_subtitle)
                worksheet.set_column(0, 0, 15)
                worksheet.set_column(1, 6, 12)
            excel_data = output.getvalue()
            file_name_out = '交通違規統計表.xlsx'

            # --- 自動化/手動執行區 ---
            st.markdown("---")
            st.subheader("🚀 執行動作")
            
            if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
            file_ids = ",".join(sorted([f.name for f in uploaded_files]))
            
            def run_automation():
                with st.status("正在執行...", expanded=True) as status:
                    # 1. 寄信
                    st.write("📧 正在寄信...")
                    mail_body = "附件為重大交通違規統計報表。"
                    if prog_text: mail_body += f"\n\n{prog_text}"
                    email_receiver = st.secrets["email"]["user"] if "email" in st.secrets else None
                    if email_receiver:
                        if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", mail_body, excel_data, file_name_out):
                            st.write(f"✅ Email 已發送")
                        else:
                            st.write("❌ Email 發送失敗")
                    else:
                        st.write("⚠️ 未設定 Email")

                    # 2. 寫入
                    st.write("📊 正在寫入 Google 試算表 (第 1 分頁)...")
                    if update_google_sheet(df_final, GOOGLE_SHEET_URL, start_cell='A3'):
                        st.write("✅ Google 試算表寫入成功！")
                    else:
                        st.write("❌ Google 試算表寫入失敗")
                    
                    status.update(label="執行結束", state="complete", expanded=False)
                    st.balloons()
            
            # 自動執行
            if file_ids not in st.session_state["sent_cache"]:
                run_automation()
                st.session_state["sent_cache"].add(file_ids)
            else:
                st.info("✅ 已自動執行過。")

            # ★★★ 手動按鈕在這裡 ★★★
            if st.button("🔄 強制重新執行 (寫入 + 寄信)", type="primary"):
                run_automation()

            st.download_button(label="📥 下載 Excel", data=excel_data, file_name=file_name_out, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e: st.error(f"發生錯誤：{e}")
