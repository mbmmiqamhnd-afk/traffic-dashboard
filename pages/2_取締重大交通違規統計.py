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
st.title("🚨 重大交通違規自動統計 (超級通用版)")

st.markdown("""
### 📝 使用說明
1. 上傳檔案後，系統會嘗試 **多種策略** 抓取數據。
2. **請觀察下方的「📄 檔案內容預覽」**，確認程式讀到的內容是否正確。
3. 自動寫入 Google 試算表 **(第 1 個分頁，從 A4 開始)**。
""")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

UNIT_MAP = {
    '聖亭': '聖亭所', '龍潭': '龍潭所', '中興': '中興所', '石門': '石門所', 
    '高平': '高平所', '三和': '三和所', '警備': '警備隊', '交通分隊': '交通分隊'
}
UNIT_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
TARGETS = {'聖亭所': 24, '龍潭所': 32, '中興所': 24, '石門所': 19, '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30}

# ==========================================
# 1. Google Sheets 寫入函數
# ==========================================
def update_google_sheet(df, sheet_url, start_cell='A4'):
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ 錯誤：未設定 Secrets！")
            return False
        try:
            gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
            sh = gc.open_by_url(sheet_url)
        except Exception as e:
            st.error(f"❌ 連線失敗: {e}")
            return False
        
        try:
            ws = sh.get_worksheet(0) # Index 0 = 第1個分頁
            if ws is None: raise Exception("找不到 Index 0 的工作表")
        except Exception as e:
            st.error(f"❌ 找不到第 1 個工作表: {e}")
            return False
        
        df_clean = df.fillna("").replace([np.inf, -np.inf], 0)
        data = [df_clean.columns.values.tolist()] + df_clean.values.tolist()
        
        try:
            ws.update(range_name=start_cell, values=data)
        except TypeError:
            ws.update(start_cell, data)
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
# 3. 超級通用解析函數
# ==========================================
def parse_report(f, file_label=""):
    if not f: return {}, None
    counts = {}
    found_date = None
    
    # 顯示檔案預覽 (讓使用者知道程式看到了什麼)
    st.markdown(f"#### 📄 正在分析：{file_label}")
    try:
        f.seek(0)
        
        # 1. 抓日期
        df_head = pd.read_excel(f, header=None, nrows=20)
        text_content = df_head.to_string()
        match = re.search(r'(?:至|~|迄)\s*(\d{3})(\d{2})(\d{2})', text_content)
        if not match:
            match = re.search(r'(?:至|~|迄)\s*(\d{3})[./\-年](\d{1,2})[./\-月](\d{1,2})', text_content)
        if match:
            y, m, d = map(int, match.groups())
            if 100 <= y <= 200 and 1 <= m <= 12 and 1 <= d <= 31:
                found_date = date(y + 1911, m, d)
        
        # 2. 讀取所有內容並顯示給使用者看
        f.seek(0)
        xls = pd.ExcelFile(f)
        
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            
            # --- 在網頁上顯示前 5 行，方便除錯 ---
            with st.expander(f"預覽工作表內容: {sheet} (點擊展開)", expanded=False):
                st.dataframe(df.head(10)) 
            # -----------------------------------

            # 開始掃描每一行
            for idx, row in df.iterrows():
                row_str = row.astype(str).str.cat(sep=' ')
                
                # 策略 A: 找 "舉發單位：" (StoneCnt 格式)
                stone_match = re.search(r"舉發單位：(\S+)", row_str)
                if stone_match:
                    raw_unit = stone_match.group(1).strip()
                    # 映射標準名稱
                    final_unit = UNIT_MAP.get(raw_unit[:2], raw_unit) # 嘗試只取前兩個字對應
                    # 找同一行的數字
                    nums = [float(x) for x in row if str(x).replace('.','',1).isdigit()]
                    if nums:
                        counts[final_unit] = counts.get(final_unit, 0) + int(nums[-1])
                    continue # 這一行處理完了，換下一行

                # 策略 B: 直接掃描單位名稱 (Focus 格式)
                matched_unit = None
                for keyword, official_name in UNIT_MAP.items():
                    if keyword in row_str:
                        # 避免 "龍潭分局" 被誤判為 "龍潭所"
                        if keyword == '龍潭' and '分隊' in row_str: continue
                        matched_unit = official_name
                        break
                
                if matched_unit:
                    # 找到單位，找這行最後一個「合理的」數字
                    nums = []
                    for x in row:
                        try:
                            val = float(str(x).replace(',', ''))
                            # 排除 NaN, Inf, 還有可能是年份的數字 (例如 113)
                            # 這裡做一個假設：統計數字通常不會剛好是 113 (除非剛好)
                            # 但為了保險，我們只排除 NaN/Inf
                            if not pd.isna(val) and val != float('inf'):
                                nums.append(val)
                        except: pass
                    
                    if nums:
                        val = int(nums[-1]) # 取最後一個
                        # 簡單防呆：如果最後一個數字是年份 (如 113)，且倒數第二個才是數據？
                        # Focus 報表通常最後一欄是合計，所以取 -1 應該是對的
                        counts[matched_unit] = counts.get(matched_unit, 0) + val

        return counts, found_date
    except Exception as e:
        st.error(f"解析錯誤 {file_label}: {e}")
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
            
            # 解析
            d_wk, _ = parse_report(files_config["Week"], "本期")
            d_yt, end_date = parse_report(files_config["YTD"], "本年累計")
            d_ly, _ = parse_report(files_config["Last_YTD"], "去年累計")

            # 計算進度
            prog_text = ""
            if end_date:
                start_of_year = date(end_date.year, 1, 1)
                days_passed = (end_date - start_of_year).days + 1
                total_days = 366 if (end_date.year % 4 == 0 and end_date.year % 100 != 0) or (end_date.year % 400 == 0) else 365
                progress_rate = days_passed / total_days
                prog_text = f"統計截至 {end_date.year-1911}年{end_date.month}月{end_date.day}日 (入案日期)，年度時間進度為 {progress_rate:.1%}"
                st.info(f"📅 {prog_text}")
            else:
                st.warning("⚠️ 無法找到日期，將不顯示年度進度。")

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
            
            # --- Excel ---
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

            # --- 自動執行區 ---
            st.markdown("---")
            st.subheader("🚀 執行動作")
            
            if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
            file_ids = ",".join(sorted([f.name for f in uploaded_files]))
            
            def run_automation():
                with st.status("正在執行...", expanded=True) as status:
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

                    st.write("📊 正在寫入 Google 試算表 (第 1 分頁, A4)...")
                    if update_google_sheet(df_final, GOOGLE_SHEET_URL, start_cell='A4'):
                        st.write("✅ Google 試算表寫入成功！")
                    else:
                        st.write("❌ Google 試算表寫入失敗")
                    
                    status.update(label="執行結束", state="complete", expanded=False)
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
