import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import smtplib
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

st.set_page_config(page_title="超載統計", layout="wide", page_icon="🚛")
st.title("🚛 超載 (stoneCnt) 自動統計")

st.markdown("""
### 📝 使用說明
1. 請上傳 **3 個** `stoneCnt` 系列的 Excel 檔案。
2. 系統將依據 **「至」** 或 **「~」** 後的日期作為 **入案截止日**。
3. 年度時間達成率會直接顯示於 Excel 表頭。
""")

# ==========================================
# 1. 參數設定
# ==========================================
TARGETS = {'聖亭所': 24, '龍潭所': 32, '中興所': 24, '石門所': 19, '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30}
UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'}
UNIT_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 2. 寄信函數
# ==========================================
def send_email(recipient, subject, body, file_bytes, filename):
    try:
        if "email" not in st.secrets:
            st.error("❌ 未設定 Secrets！")
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
# 3. 資料解析函數 (修正日期抓取邏輯)
# ==========================================
def parse_stone(f):
    if not f: return {}, None
    counts = {}
    found_date = None
    try:
        # 1. 抓取日期：只抓取「統計區間」的結束日
        f.seek(0)
        df_head = pd.read_excel(f, header=None, nrows=20)
        text_content = df_head.to_string()
        
        # 關鍵修正：只尋找「至」或「~」後面的日期
        # 格式範例： 113/01/01 至 113/05/31  或  113.01.01 ~ 113.05.31
        # regex 解釋：尋找 "至" 或 "~" 接著任意空白，接著 年/月/日
        match = re.search(r'(?:至|~)\s*(\d{3})[./\-年](\d{1,2})[./\-月](\d{1,2})', text_content)
        
        if match:
            y, m, d = map(int, match.groups())
            # 簡單檢核日期合理性
            if 100 <= y <= 200 and 1 <= m <= 12 and 1 <= d <= 31:
                found_date = date(y + 1911, m, d)
        
        # 如果上方沒抓到，嘗試抓取同一行有多個日期的情況 (取後者)
        if not found_date:
            # 找出所有類似日期的字串
            all_dates_raw = re.findall(r'(\d{3})[./-](\d{1,2})[./-](\d{1,2})', text_content)
            # 如果發現很多日期，通常統計表頭的格式是 [開始日] [結束日] [列印日]
            # 我們要避免抓到列印日。通常統計結束日會跟在開始日後面。
            # 這裡做一個保守估計：如果有找到日期，但沒找到「至」，則暫時不回傳日期，以免誤用列印日。
            pass 

        # 2. 讀取數據
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
uploaded_files = st.file_uploader("請拖曳 3 個 stoneCnt 檔案至此", accept_multiple_files=True, type=['xlsx', 'xls'])

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
            
            # 開始解析
            d_wk, _ = parse_stone(files_config["Week"])
            d_yt, end_date = parse_stone(files_config["YTD"]) # 關鍵：從本年累計抓日期
            d_ly, _ = parse_stone(files_config["Last_YTD"])

            # 計算年度時間進度
            prog_text = ""
            if end_date:
                start_of_year = date(end_date.year, 1, 1)
                days_passed = (end_date - start_of_year).days + 1
                total_days = 366 if (end_date.year % 4 == 0 and end_date.year % 100 != 0) or (end_date.year % 400 == 0) else 365
                progress_rate = days_passed / total_days
                
                prog_text = f"統計截至 {end_date.year-1911}年{end_date.month}月{end_date.day}日 (入案日期)，年度時間進度為 {progress_rate:.1%}"
                st.info(f"📅 {prog_text}")
            else:
                st.warning("⚠️ 無法從「本年累計」檔案中找到「至 11x/xx/xx」格式的日期，無法計算時間進度。")

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
            
            # --- 產生 Excel (包含標題與時間進度) ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # 從第 4 列開始寫入表格 (保留上方給標題)
                df_final.to_excel(writer, index=False, sheet_name='超載統計', startrow=3)
                
                workbook = writer.book
                worksheet = writer.sheets['超載統計']
                
                # 設定格式
                fmt_title = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center'})
                fmt_subtitle = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': 'blue', 'align': 'left'})
                
                # 寫入標題 (合併儲存格)
                worksheet.merge_range('A1:G1', '超載取締統計表', fmt_title)
                
                # 寫入時間進度 (在第 2 列)
                if prog_text:
                    worksheet.merge_range('A2:G2', f"說明：{prog_text}", fmt_subtitle)
                
                # 自動調整欄寬
                worksheet.set_column(0, 0, 15) # 單位欄寬一點
                worksheet.set_column(1, 6, 12) # 數據欄

            excel_data = output.getvalue()
            file_name_out = '超載統計表.xlsx'

            # 自動寄信邏輯
            if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
            file_ids = ",".join(sorted([f.name for f in uploaded_files]))
            email_receiver = st.secrets["email"]["user"]
            
            if file_ids not in st.session_state["sent_cache"]:
                with st.spinner(f"正在自動寄送報表至 {email_receiver}..."):
                    mail_body = "附件為超載統計報表。"
                    if prog_text: mail_body += f"\n\n{prog_text}"
                    
                    if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", mail_body, excel_data, file_name_out):
                        st.balloons()
                        st.success(f"✅ 郵件已發送至 {email_receiver}")
                        st.session_state["sent_cache"].add(file_ids)
            else:
                st.info(f"✅ 報表已於剛才發送至 {email_receiver}")

            st.download_button(label="📥 下載 Excel", data=excel_data, file_name=file_name_out, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e: st.error(f"錯誤：{e}")
