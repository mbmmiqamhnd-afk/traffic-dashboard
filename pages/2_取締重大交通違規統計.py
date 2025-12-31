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
st.title("🚨 重大交通違規自動統計 (Focus 專用版)")

st.markdown("""
### 📝 使用說明
1. 請上傳 **3 個** `Focus` 系列 Excel 檔案。
2. 系統會自動搜尋含有 **單位名稱** (如：聖亭、龍潭...) 的列並加總數據。
3. 自動寄信並寫入 Google 試算表 **(第 1 個分頁，從 A4 開始)**。
4. **數值若有誤，請展開下方的「🕵️‍♀️ 詳細抓取過程」檢查。**
""")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

# 這裡設定要搜尋的單位關鍵字 (左邊是報表上可能出現的字，右邊是統一名稱)
# 程式會掃描 Excel 裡是否包含左邊的字眼
UNIT_MAP = {
    '聖亭': '聖亭所', 
    '龍潭': '龍潭所', 
    '中興': '中興所', 
    '石門': '石門所', 
    '高平': '高平所', 
    '三和': '三和所', 
    '警備': '警備隊', 
    '交通分隊': '交通分隊'
}

# 這是最後要顯示的順序
UNIT_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# 目標值 (請確認 Focus 的目標值是否正確)
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
            # 鎖定第 1 個分頁 (Index 0)
            ws = sh.get_worksheet(0) 
            if ws is None: raise Exception("找不到 Index 0 的工作表")
        except Exception as e:
            st.error(f"❌ 找不到第 1 個工作表: {e}")
            return False
        
        df_clean = df.fillna("").replace([np.inf, -np.inf], 0)
        data = [df_clean.columns.values.tolist()] + df_clean.values.tolist()
        
        try:
            # 嘗試寫入 (A4)
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
# 3. 資料解析函數 (Focus 專用邏輯)
# ==========================================
def parse_report(f, file_label=""):
    if not f: return {}, None, []
    counts = {}
    found_date = None
    debug_logs = []
    
    try:
        debug_logs.append(f"🔵 開始解析: {file_label}")
        f.seek(0)
        
        # --- 1. 抓取日期 ---
        # 讀取前 20 行找日期
        df_head = pd.read_excel(f, header=None, nrows=20)
        text_content = df_head.to_string()
        
        # 搜尋類似 113/05/20 或 1130520 的日期
        match = re.search(r'(?:至|~|迄)\s*(\d{3})(\d{2})(\d{2})', text_content)
        if not match:
            match = re.search(r'(?:至|~|迄)\s*(\d{3})[./\-年](\d{1,2})[./\-月](\d{1,2})', text_content)
        
        if match:
            y, m, d = map(int, match.groups())
            if 100 <= y <= 200 and 1 <= m <= 12 and 1 <= d <= 31:
                found_date = date(y + 1911, m, d)
                debug_logs.append(f"📅 抓到日期: {found_date}")
        else:
            debug_logs.append("⚠️ 未抓到日期 (可能影響進度計算)")
        
        # --- 2. 抓取數據 (核心修改) ---
        f.seek(0)
        xls = pd.ExcelFile(f)
        for sheet in xls.sheet_names:
            debug_logs.append(f"📄 掃描工作表: {sheet}")
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            
            for idx, row in df.iterrows():
                # 將整列轉為字串方便搜尋
                row_str = row.astype(str).str.cat(sep=' ')
                
                # 掃描是否包含我們定義的單位名稱 (如 '聖亭', '龍潭')
                matched_unit = None
                for keyword, official_name in UNIT_MAP.items():
                    # 這裡排除掉 "龍潭交通分隊" 裡的 "龍潭" 被誤判為 "龍潭所" 的情況
                    # 邏輯：如果找到關鍵字，且沒有被更長的關鍵字涵蓋 (簡單版暫不處理複雜邏輯，通常 Focus 表格分隊與派出所是分開的)
                    if keyword in row_str:
                        # 特殊處理：避免「龍潭分局」被當作「龍潭所」
                        if keyword == '龍潭' and '分隊' in row_str: continue 
                        
                        matched_unit = official_name
                        break # 找到一個就停，避免重複匹配
                
                if matched_unit:
                    # 找到單位了，現在找這行裡面的數字
                    # 排除掉小數點、文字，只抓純數字
                    nums = []
                    for x in row:
                        try:
                            # 嘗試轉為浮點數
                            val = float(str(x).replace(',', '')) # 去除千分位逗號
                            # 排除 NaN 和無限大
                            if not pd.isna(val) and val != float('inf'):
                                nums.append(val)
                        except:
                            continue
                    
                    if nums:
                        # 策略：Focus 報表通常最後一欄是合計，或者是數值最大的是合計
                        # 這裡我們取「最後一個數字」作為該單位的統計值 (通常是總計)
                        # 如果您的報表總計在第一欄，請告訴我，我再改
                        val = int(nums[-1])
                        
                        # 累加 (防止同一個單位出現在多行)
                        counts[matched_unit] = counts.get(matched_unit, 0) + val
                        debug_logs.append(f"   ✅ Row {idx}: 發現 [{matched_unit}] -> 抓到數字序列 {nums} -> 取用: {val}")
                    else:
                        debug_logs.append(f"   ⚠️ Row {idx}: 發現 [{matched_unit}] 但該行沒有數字")

        return counts, found_date, debug_logs
    except Exception as e:
        st.error(f"解析檔案 {f.name} 錯誤: {e}")
        return {}, None, [f"❌ 發生錯誤: {e}"]

# ==========================================
# 4. 主程式執行
# ==========================================
uploaded_files = st.file_uploader("請拖曳 3 個 Focus 統計檔案至此", accept_multiple_files=True, type=['xlsx', 'xls'])

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
            d_wk, _, logs_wk = parse_report(files_config["Week"], "本期")
            d_yt, end_date, logs_yt = parse_report(files_config["YTD"], "本年累計")
            d_ly, _, logs_ly = parse_report(files_config["Last_YTD"], "去年累計")

            # --- 除錯區 (預設收合) ---
            with st.expander("🕵️‍♀️ 查看詳細抓取過程 (若數值有誤請點此檢查)", expanded=False):
                c1, c2, c3 = st.columns(3)
                with c1: 
                    st.caption("本期日誌")
                    for l in logs_wk: st.text(l)
                with c2: 
                    st.caption("本年累計日誌")
                    for l in logs_yt: st.text(l)
                with c3: 
                    st.caption("去年累計日誌")
                    for l in logs_ly: st.text(l)

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

            # 建立表格
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
            # 警備隊數值歸零 (依需求)
            mask_guard = df_calc['單位'] == '警備隊'
            df_calc.loc[mask_guard, ['本期', '本年累計', '去年累計', '目標值']] = 0
            
            total = df_calc[['本期', '本年累計', '去年累計', '目標值']].sum().to_dict()
            total['單位'] = '合計'
            
            df_final = pd.concat([pd.DataFrame([total]), df], ignore_index=True)
            df_final['本年與去年同期比較'] = df_final['本年累計'] - df_final['去年累計']
            df_final['達成率'] = df_final.apply(lambda x: f"{x['本年累計']/x['目標值']:.2%}" if x['目標值']>0 else "—", axis=1)
            # 警備隊顯示 —
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
