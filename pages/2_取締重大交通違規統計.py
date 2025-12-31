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

st.set_page_config(page_title="取締重大交通違規統計", layout="wide", page_icon="🚔")
st.title("🚔 取締重大交通違規統計 (含攔停/逕舉)")

st.markdown("""
### 📝 使用說明
1. 請上傳 **3 個** 重點違規報表 (focus系列)。
2. 系統會自動區分 **攔停** 與 **逕舉** 件數。
3. 自動寄信並寫入 Google 試算表 **(只寫入數據，從 B4 開始)**。
4. **若沒反應，請點擊下方的「🔄 強制手動執行」按鈕。**
""")

# ==========================================
# 0. 設定區
# ==========================================
# ★★★ 請確認這裡填入的是正確的 Google 試算表網址 ★★★
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

# 單位名稱對應
UNIT_MAP = {
    '聖亭派出所': '聖亭所', 
    '龍潭派出所': '龍潭所', 
    '中興派出所': '中興所', 
    '石門派出所': '石門所', 
    '高平派出所': '高平所', 
    '三和派出所': '三和所', 
    '警備隊': '警備隊', 
    '龍潭交通分隊': '交通分隊',
    '交通組': '科技執法' 
}

# 顯示順序
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# 目標值
TARGETS = {
    '聖亭所': 1838, '龍潭所': 2451, '中興所': 1838, '石門所': 1488, 
    '高平所': 1226, '三和所': 400, '交通分隊': 2576, '警備隊': 263, '科技執法': 0
}

# ==========================================
# 1. Google Sheets 寫入函數
# ==========================================
def update_google_sheet(df, sheet_url, start_cell='B4'): # <--- 預設改為 B4 (數據區)
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
        
        # 處理資料格式
        df_clean = df.fillna("").replace([np.inf, -np.inf], 0)
        
        # ★★★ 關鍵修改：不含標題列 ★★★
        # 只取 values，不取 columns
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
# 3. Focus 專用解析函數
# ==========================================
def parse_focus_report(uploaded_file):
    if not uploaded_file: return None
    
    content = uploaded_file.getvalue()
    start_date, end_date = "", ""
    df = None
    header_idx = -1
    
    try:
        # 1. 嘗試讀取前20行找日期與標題列
        df_raw = pd.read_excel(io.BytesIO(content), header=None, nrows=20)
        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
            
            # 抓取日期
            if not start_date:
                match = re.search(r'入案日期[：:]?\s*(\d{3,7}).*至\s*(\d{3,7})', row_str)
                if match: start_date, end_date = match.group(1), match.group(2)
            
            # 抓取標題列
            if "單位" in row_str and "酒後" in row_str: 
                header_idx = i
                
        if header_idx != -1: 
            df = pd.read_excel(io.BytesIO(content), header=header_idx)
        else:
            return None 

        if df is None: return None

        # 2. 定義重點欄位
        keywords = ["酒後", "闖紅燈", "嚴重超速", "逆向", "轉彎", "蛇行", "不暫停讓行人", "機車"]
        stop_cols = [] 
        cit_cols = []  
        
        for i in range(len(df.columns)):
            col_str = str(df.columns[i])
            if any(k in col_str for k in keywords) and "路肩" not in col_str and "大型車" not in col_str:
                stop_cols.append(i)
                cit_cols.append(i+1)
        
        # 3. 統計數據
        unit_data = {}
        for _, row in df.iterrows():
            raw_unit = str(row['單位']).strip()
            if raw_unit == 'nan' or not raw_unit: continue
            
            unit_name = UNIT_MAP.get(raw_unit, raw_unit)
            
            s, c = 0, 0
            for col in stop_cols:
                try: s += float(str(row.iloc[col]).replace(',', ''))
                except: pass
            for col in cit_cols:
                try: c += float(str(row.iloc[col]).replace(',', ''))
                except: pass
            
            unit_data[unit_name] = {'stop': s, 'cit': c}

        # 4. 計算天數
        duration = 0
        try:
            s_d = re.sub(r'[^\d]', '', start_date)
            e_d = re.sub(r'[^\d]', '', end_date)
            if len(s_d) < 7: s_d = s_d.zfill(7)
            if len(e_d) < 7: e_d = e_d.zfill(7)
            d1 = date(int(s_d[:3])+1911, int(s_d[3:5]), int(s_d[5:]))
            d2 = date(int(e_d[:3])+1911, int(e_d[3:5]), int(e_d[5:]))
            duration = (d2 - d1).days
        except: duration = 0
        
        return {'data': unit_data, 'start': start_date, 'end': end_date, 'duration': duration}

    except Exception as e:
        st.error(f"解析錯誤: {e}")
        return None

# ==========================================
# 4. 主程式執行
# ==========================================
# 使用新 key 強制重置上傳狀態
uploaded_files = st.file_uploader("請拖曳 3 個 Focus 統計檔案至此", accept_multiple_files=True, type=['xlsx', 'xls'], key="focus_uploader_B4_no_header")

if uploaded_files:
    if len(uploaded_files) < 3:
        st.warning("⏳ 檔案不足 3 個，請繼續上傳...")
    else:
        try:
            # 1. 解析所有檔案
            parsed_files = []
            for f in uploaded_files:
                res = parse_focus_report(f)
                if res: parsed_files.append(res)
            
            if len(parsed_files) < 3:
                st.error("❌ 有檔案解析失敗，請確認檔案格式")
                st.stop()

            # 2. 自動判斷
            parsed_files.sort(key=lambda x: x['start']) 
            file_last_year = parsed_files[0] 
            
            others = parsed_files[1:]
            others.sort(key=lambda x: x['duration'], reverse=True)
            file_year = others[0] 
            file_week = others[1] 

            # 3. 進度
            prog_text = ""
            try:
                end_str = re.sub(r'[^\d]', '', file_year['end'])
                if len(end_str) < 7: end_str = end_str.zfill(7)
                curr_y = int(end_str[:3]) + 1911
                curr_m = int(end_str[3:5])
                curr_d = int(end_str[5:])
                target_date = date(curr_y, curr_m, curr_d)
                start_of_year = date(curr_y, 1, 1)
                days_passed = (target_date - start_of_year).days + 1
                total_days = 366 if (curr_y % 4 == 0 and curr_y % 100 != 0) or (curr_y % 400 == 0) else 365
                progress_rate = days_passed / total_days
                prog_text = f"統計截至 {curr_y-1911}年{curr_m}月{curr_d}日 (入案日期)，年度時間進度為 {progress_rate:.1%}"
                st.info(f"📅 {prog_text}")
            except: pass

            # 4. 組合最終表格
            rows = []
            accum = {'ws':0, 'wc':0, 'ys':0, 'yc':0, 'ls':0, 'lc':0}
            
            for u in UNIT_ORDER:
                w = file_week['data'].get(u, {'stop':0, 'cit':0})
                y = file_year['data'].get(u, {'stop':0, 'cit':0})
                l = file_last_year['data'].get(u, {'stop':0, 'cit':0})
                
                # 科技執法歸零
                if u == '科技執法': w['stop'], y['stop'], l['stop'] = 0, 0, 0

                y_total = y['stop'] + y['cit']
                l_total = l['stop'] + l['cit']
                
                row_data = [u, w['stop'], w['cit'], y['stop'], y['cit'], l['stop'], l['cit']]
                
                if u == '警備隊': 
                    row_data.extend(['—', '—', '—'])
                else:
                    diff = int(y_total - l_total)
                    tgt = TARGETS.get(u, 0)
                    row_data.append(diff)
                    if u == '科技執法':
                        row_data.extend(['—', '—'])
                    else:
                        rate_str = f"{y_total/tgt:.2%}" if tgt > 0 else "0.00%"
                        row_data.extend([tgt, rate_str])
                
                accum['ws']+=w['stop']; accum['wc']+=w['cit']
                accum['ys']+=y['stop']; accum['yc']+=y['cit']
                accum['ls']+=l['stop']; accum['lc']+=l['cit']
                
                rows.append(row_data)

            # 合計列
            total_target = sum([v for k,v in TARGETS.items() if k not in ['警備隊', '科技執法']])
            t_diff = (accum['ys']+accum['yc']) - (accum['ls']+accum['lc'])
            t_rate = (accum['ys']+accum['yc'])/total_target if total_target > 0 else 0
            
            total_row = ['合計', accum['ws'], accum['wc'], accum['ys'], accum['yc'], accum['ls'], accum['lc'], t_diff, total_target, f"{t_rate:.2%}"]
            
            cols = ['單位', '本期_攔停', '本期_逕舉', '本年_攔停', '本年_逕舉', '去年_攔停', '去年_逕舉', '本年與去年比較', '目標值', '達成率']
            df_final = pd.DataFrame([total_row] + rows, columns=cols)

            st.success("✅ 分析完成！")
            st.dataframe(df_final, use_container_width=True, hide_index=True)

            # --- 產生 Excel ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Sheet1', startrow=3)
                workbook = writer.book
                ws = writer.sheets['Sheet1']
                fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
                ws.merge_range('A1:J1', '取締重大交通違規件數統計表', fmt_title)
                ws.write('A2', f"一、統計期間：{file_year['start']}~{file_year['end']}")
                if prog_text:
                    ws.write('A3', f"二、{prog_text}")
                ws.set_column(0, 0, 15) 
            
            excel_data = output.getvalue()
            file_name_out = f'重點違規統計_{file_year["end"]}.xlsx'

            # --- 自動化流程 ---
            st.markdown("---")
            st.subheader("🚀 執行動作")
            
            if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
            file_ids = ",".join(sorted([f.name for f in uploaded_files]))
            
            def run_automation():
                with st.status("正在執行自動化作業...", expanded=True) as status:
                    # 1. 寄信
                    st.write("📧 正在寄送 Email...")
                    email_receiver = st.secrets["email"]["user"] if "email" in st.secrets else None
                    if email_receiver:
                        if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", "附件為重點違規統計報表。", excel_data, file_name_out):
                            st.write(f"✅ Email 已發送至 {email_receiver}")
                        else:
                            st.write("❌ Email 發送失敗")
                    else:
                        st.write("⚠️ 未設定 Email")

                    # 2. 寫入 Google Sheet (移除單位欄，寫入 B4)
                    st.write("📊 正在寫入 Google 試算表 (第 1 分頁, B4, 不含標題/單位)...")
                    
                    # 建立不含「單位」欄的 DataFrame
                    df_write = df_final.drop(columns=['單位'])
                    
                    if update_google_sheet(df_write, GOOGLE_SHEET_URL, start_cell='B4'): 
                        st.write("✅ Google 試算表寫入成功！")
                    else:
                        st.write("❌ Google 試算表寫入失敗")
                    
                    status.update(label="執行完畢", state="complete", expanded=False)
                    st.balloons()
            
            if file_ids not in st.session_state["sent_cache"]:
                run_automation()
                st.session_state["sent_cache"].add(file_ids)
            else:
                st.info("✅ 此組檔案已自動執行過。")

            if st.button("🔄 強制重新執行 (寫入 + 寄信)", type="primary"):
                run_automation()

            st.download_button(label="📥 下載 Excel", data=excel_data, file_name=file_name_out, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e: st.error(f"發生錯誤：{e}")
