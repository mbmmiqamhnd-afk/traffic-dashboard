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

st.set_page_config(page_title="取締重大交通違規統計", layout="wide", page_icon="🚔")
st.title("🚔 取締重大交通違規統計 (v22 新增跨欄標題版)")

# --- 強制清除快取按鈕 ---
if st.button("🧹 清除快取 (若更新無效請按此)", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.success("快取已清除！請重新整理頁面 (F5) 並重新上傳檔案。")

st.markdown("""
### 📝 使用說明 (v22)
1. **Excel 排版更新**：
   - 刪除原本的「一、統計期間...」文字列。
   - 新增 **跨欄標題** (本期、本年累計、去年累計) 並自動帶入日期。
2. **目標值與達成率維持空白**。
3. **自動寄信** 與 **Google Sheet 寫入** 功能保留。
""")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所',
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所',
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '交通組': '科技執法'
}
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. Google Sheets 寫入函數
# ==========================================
def update_google_sheet(df, sheet_url, start_cell='B4'):
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ 錯誤：未設定 Secrets！")
            return False
        
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(0)
        if ws is None: raise Exception("找不到 Index 0 的工作表")
        
        st.info(f"📂 寫入目標工作表：**「{ws.title}」**")

        df_clean = df.fillna("").replace([np.inf, -np.inf], 0)
        data = df_clean.values.tolist()
        
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
def parse_focus_report(uploaded_file):
    if not uploaded_file: return None
    file_name = uploaded_file.name
    try:
        content = uploaded_file.getvalue()
        start_date, end_date = "", ""
        df = None; header_idx = -1
        
        df_raw = pd.read_excel(io.BytesIO(content), header=None, nrows=25)
        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
            if not start_date:
                match = re.search(r'入案日期[：:]?\s*(\d{3,7}).*至\s*(\d{3,7})', row_str)
                if match: start_date, end_date = match.group(1), match.group(2)
            if "單位" in row_str:
                header_idx = i
                if start_date: break
        
        if header_idx == -1:
            st.warning(f"⚠️ 檔案 {file_name} 解析警告：找不到標題列。")
            return None

        df = pd.read_excel(io.BytesIO(content), header=header_idx)
        keywords = ["酒後", "闖紅燈", "嚴重超速", "逆向", "轉彎", "蛇行", "不暫停讓行人", "機車"]
        stop_cols = []; cit_cols = []
        
        for i in range(len(df.columns)):
            col_str = str(df.columns[i])
            if any(k in col_str for k in keywords) and "路肩" not in col_str and "大型車" not in col_str:
                stop_cols.append(i); cit_cols.append(i+1)
        
        unit_data = {}
        for _, row in df.iterrows():
            raw_unit = str(row['單位']).strip()
            if raw_unit == 'nan' or not raw_unit or "合計" in raw_unit: continue
            
            unit_name = UNIT_MAP.get(raw_unit, raw_unit)
            s, c = 0, 0
            
            for col in stop_cols:
                try:
                    val = row.iloc[col]
                    if pd.isna(val) or str(val).strip() == "": val = 0
                    s += float(str(val).replace(',', ''))
                except: pass
            
            for col in cit_cols:
                try:
                    val = row.iloc[col]
                    if pd.isna(val) or str(val).strip() == "": val = 0
                    c += float(str(val).replace(',', ''))
                except: pass

            unit_data[unit_name] = {'stop': s, 'cit': c}

        duration = 0
        try:
            if start_date and end_date:
                s_d = re.sub(r'[^\d]', '', start_date); e_d = re.sub(r'[^\d]', '', end_date)
                d1 = date(int(s_d[:3])+1911, int(s_d[3:5]), int(s_d[5:]))
                d2 = date(int(e_d[:3])+1911, int(e_d[3:5]), int(e_d[5:]))
                duration = (d2 - d1).days
        except: duration = 0
        if not start_date: start_date = "0000000"
        if not end_date: end_date = "0000000"
        return {'data': unit_data, 'start': start_date, 'end': end_date, 'duration': duration, 'filename': file_name}
    except Exception as e:
        st.warning(f"⚠️ 檔案 {file_name} 錯誤: {e}")
        return None

# 小工具：取出日期字串的後4碼 (月日)
def get_mmdd(date_str):
    clean = re.sub(r'[^\d]', '', str(date_str))
    return clean[-4:] if len(clean) >= 4 else clean

# ==========================================
# 4. 主程式
# ==========================================
# ★★★ v22 Key ★★★
uploaded_files = st.file_uploader("請拖曳 3 個 Focus 統計檔案至此", accept_multiple_files=True, type=['xlsx', 'xls'], key="focus_uploader_v22_merge_header")

if uploaded_files:
    if len(uploaded_files) < 3: st.warning("⏳ 檔案不足 (需 3 個)...")
    else:
        try:
            parsed_files = []
            for f in uploaded_files:
                res = parse_focus_report(f)
                if res: parsed_files.append(res)
            
            if len(parsed_files) < 3: 
                st.error("❌ 解析失敗。")
                st.stop()

            parsed_files.sort(key=lambda x: x['start'])
            file_last_year = parsed_files[0]
            others = parsed_files[1:]
            others.sort(key=lambda x: x['duration'], reverse=True)
            file_year = others[0]
            file_week = others[1]

            unit_rows = []
            accum = {'ws':0, 'wc':0, 'ys':0, 'yc':0, 'ls':0, 'lc':0}
            
            for u in UNIT_ORDER:
                w = file_week['data'].get(u, {'stop':0, 'cit':0})
                y = file_year['data'].get(u, {'stop':0, 'cit':0})
                l = file_last_year['data'].get(u, {'stop':0, 'cit':0})
                
                if u == '科技執法': w['stop'], y['stop'], l['stop'] = 0, 0, 0
                y_total = y['stop'] + y['cit']; l_total = l['stop'] + l['cit']
                
                w_s, w_c = int(w['stop']), int(w['cit'])
                y_s, y_c = int(y['stop']), int(y['cit'])
                l_s, l_c = int(l['stop']), int(l['cit'])

                row_data = [u, w_s, w_c, y_s, y_c, l_s, l_c]
                
                if u == '警備隊': 
                    row_data.extend(['—', '', '']) 
                else:
                    diff = int(y_total - l_total)
                    row_data.append(diff)
                    if u == '科技執法':
                        row_data.extend(['', ''])
                    else:
                        row_data.extend(['', '']) 
                
                accum['ws']+=w_s; accum['wc']+=w_c
                accum['ys']+=y_s; accum['yc']+=y_c
                accum['ls']+=l_s; accum['lc']+=l_c
                unit_rows.append(row_data)

            t_diff = (accum['ys']+accum['yc']) - (accum['ls']+accum['lc'])
            total_row = ['合計', accum['ws'], accum['wc'], accum['ys'], accum['yc'], accum['ls'], accum['lc'], t_diff, '', '']
            final_rows = [total_row] + unit_rows

            cols = ['取締方式', '本期_攔停', '本期_逕舉', '本年_攔停', '本年_逕舉', '去年_攔停', '去年_逕舉', '本年與去年比較', '目標值', '達成率']
            df_final = pd.DataFrame(final_rows, columns=cols)
            df_write = df_final.drop(columns=['取締方式'])

            st.success("✅ 分析完成！Excel 標題已更新")
            st.dataframe(df_final, use_container_width=True, hide_index=True)

            # ==========================================
            # Excel 產生邏輯 (大幅修改：新增跨欄標題列)
            # ==========================================
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # 1. 將 DataFrame 寫在第 3 列 (Index 2)，預留上方給標題
                df_final.to_excel(writer, index=False, sheet_name='Sheet1', startrow=2)
                workbook = writer.book
                ws = writer.sheets['Sheet1']
                
                # 樣式
                fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
                fmt_header_center = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFEB9C'}) # 淺黃色底強調
                fmt_header_plain = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})

                # 2. 寫入大標題 (A1)
                ws.merge_range('A1:J1', '取締重大交通違規件數統計表', fmt_title)

                # 3. 準備跨欄標題的日期文字
                # 格式: 本期 (0101~0107)
                str_week = f"本期 ({get_mmdd(file_week['start'])}~{get_mmdd(file_week['end'])})"
                str_year = f"本年累計 ({get_mmdd(file_year['start'])}~{get_mmdd(file_year['end'])})"
                str_last = f"去年累計 ({get_mmdd(file_last_year['start'])}~{get_mmdd(file_last_year['end'])})"

                # 4. 寫入第二列 (A2 ~ G2) 的跨欄標題
                # A2: 統計期間 (單格)
                ws.write('A2', '統計期間', fmt_header_center)
                
                # B2-C2: 本期 (跨欄)
                ws.merge_range('B2:C2', str_week, fmt_header_center)
                
                # D2-E2: 本年累計 (跨欄)
                ws.merge_range('D2:E2', str_year, fmt_header_center)
                
                # F2-G2: 去年累計 (跨欄)
                ws.merge_range('F2:G2', str_last, fmt_header_center)
                
                # H2, I2, J2: 這裡可以留白，或是給簡單的框線，為了美觀我們加上框線
                ws.write('H2', '', fmt_header_plain)
                ws.write('I2', '', fmt_header_plain)
                ws.write('J2', '', fmt_header_plain)

                # 5. 調整欄寬
                ws.set_column(0, 0, 15) # 取締方式
                ws.set_column(1, 6, 10) # 數據欄
                ws.set_column(7, 9, 12) # 後面幾欄
            
            excel_data = output.getvalue()
            file_name_out = f'重點違規統計_{file_year["end"]}.xlsx'

            st.markdown("---")
            if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
            file_ids = ",".join(sorted([f.name for f in uploaded_files]))
            
            def run_automation():
                with st.status("🚀 執行自動化任務...", expanded=True) as status:
                    st.write("📧 正在寄送 Email...")
                    email_receiver = st.secrets["email"]["user"] if "email" in st.secrets else None
                    if email_receiver:
                        if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", "附件為重點違規統計報表(新版格式)。", excel_data, file_name_out):
                            st.write(f"✅ Email 已發送")
                    else: st.warning("⚠️ 未設定 Email Secrets")
                    
                    st.write("📊 正在寫入 Google 試算表 (B4)...")
                    if update_google_sheet(df_write, GOOGLE_SHEET_URL, start_cell='B4'):
                        st.write("✅ 寫入成功！")
                    else: st.write("❌ 寫入失敗")
                    
                    status.update(label="執行完畢", state="complete", expanded=False)
                    st.balloons()
            
            if file_ids not in st.session_state["sent_cache"]:
                run_automation()
                st.session_state["sent_cache"].add(file_ids)
            else: st.info("✅ 已自動執行過。")

            if st.button("🔄 強制執行", type="primary"): run_automation()

            st.download_button(label="📥 下載 Excel", data=excel_data, file_name=file_name_out, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e: 
            st.error(f"❌ 發生嚴重錯誤：{e}")
