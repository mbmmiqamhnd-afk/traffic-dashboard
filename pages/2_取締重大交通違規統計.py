import streamlit as st
import pandas as pd
import io
import re
import smtplib
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
2. **上傳後自動分析** 8 大重點項目。
3. 自動計算 **年度時間進度** 供績效參考。
""")

# --- 寄信函數 ---
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

# --- 主程式 ---
uploaded_files = st.file_uploader("請上傳 3 個檔案", accept_multiple_files=True, key="focus_uploader")

if uploaded_files:
    if len(uploaded_files) < 3:
        st.warning("⏳ 檔案不足 3 個，請繼續上傳...")
    else:
        try:
            def parse_file_content(uploaded_file):
                content = uploaded_file.getvalue()
                df = None; start_date = ""; end_date = ""; header_idx = -1
                is_excel = uploaded_file.name.endswith(('.xlsx', '.xls'))
                try:
                    if is_excel:
                        df_raw = pd.read_excel(io.BytesIO(content), header=None, nrows=20)
                        for i, row in df_raw.iterrows():
                            row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
                            if not start_date:
                                # 抓取日期格式 1130101 或 113/01/01
                                match = re.search(r'入案日期[：:]?\s*(\d{3,7}).*至\s*(\d{3,7})', row_str)
                                if match: 
                                    start_date, end_date = match.group(1), match.group(2)
                            if "單位" in row_str and "酒後" in row_str: header_idx = i
                        if header_idx != -1: df = pd.read_excel(io.BytesIO(content), header=header_idx)
                    else:
                        try: text = content.decode('utf-8')
                        except: text = content.decode('cp950', errors='ignore')
                        lines = text.splitlines()
                        for i, line in enumerate(lines):
                            match = re.search(r'入案日期[：:]?\s*(\d{3,7}).*至\s*(\d{3,7})', line)
                            if match: start_date, end_date = match.group(1), match.group(2)
                            if "單位" in line and "酒後" in line: header_idx = i
                        if header_idx != -1: df = pd.read_csv(io.StringIO(text), header=header_idx)
                except: return None

                if df is None: return None
                keywords = ["酒後", "闖紅燈", "嚴重超速", "逆向", "轉彎", "蛇行", "不暫停讓行人", "機車"]
                stop_cols = []; cit_cols = []
                for i in range(len(df.columns)):
                    col_str = str(df.columns[i])
                    if any(k in col_str for k in keywords) and "路肩" not in col_str and "大型車" not in col_str:
                        stop_cols.append(i); cit_cols.append(i+1)
                
                unit_data = {}
                for _, row in df.iterrows():
                    unit = str(row['單位']).strip()
                    if unit == 'nan' or not unit: continue
                    s, c = 0, 0
                    for col in stop_cols:
                        try: s += float(str(row.iloc[col]).replace(',', ''))
                        except: pass
                    for col in cit_cols:
                        try: c += float(str(row.iloc[col]).replace(',', ''))
                        except: pass
                    unit_data[unit] = {'stop': s, 'cit': c}
                
                # 計算日期差距
                duration = 0
                try:
                    # 清理日期字串 (移除 / . 等符號，統一變成 7 位數)
                    s_d = re.sub(r'[^\d]', '', start_date)
                    e_d = re.sub(r'[^\d]', '', end_date)
                    if len(s_d) < 7: s_d = s_d.zfill(7) # 補0
                    if len(e_d) < 7: e_d = e_d.zfill(7)

                    d1 = date(int(s_d[:3])+1911, int(s_d[3:5]), int(s_d[5:]))
                    d2 = date(int(e_d[:3])+1911, int(e_d[3:5]), int(e_d[5:]))
                    duration = (d2 - d1).days
                except: duration = 0
                
                return {'data': unit_data, 'start': start_date, 'end': end_date, 'duration': duration}

            parsed_files = []
            for f in uploaded_files:
                res = parse_file_content(f)
                if res: parsed_files.append(res)
            
            if len(parsed_files) < 3: st.error("有效檔案不足！"); st.stop()

            # 排序邏輯
            # 1. 找開始日期最早的 -> 去年
            # 2. 剩下兩個，天數長的 -> 本年累計
            parsed_files.sort(key=lambda x: x['start']) 
            file_last_year = parsed_files[0]
            
            others = parsed_files[1:]
            others.sort(key=lambda x: x['duration'], reverse=True)
            file_year = others[0] # 本年累計
            file_week = others[1] # 本期

            # --- 計算年度達成率基準 ---
            prog_text = ""
            try:
                # 解析本年累計的「結束日期」
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
                
                prog_text = f"📅 統計截至 **{curr_y-1911}年{curr_m}月{curr_d}日**，年度時間進度為 **{progress_rate:.1%}**"
                st.info(prog_text)
            except:
                pass # 日期解析失敗則不顯示

            st.success(f"✅ 檔案識別成功：本年({file_year['start']}~{file_year['end']})")

            unit_mapping = {'交通組': '科技執法', '龍潭交通分隊': '交通分隊', '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊'}
            display_order = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
            targets = {'聖亭所': 1838, '龍潭所': 2451, '中興所': 1838, '石門所': 1488, '高平所': 1226, '三和所': 400, '交通分隊': 2576, '警備隊': 263, '科技執法': 0}

            rows = []
            accum = {'ws':0, 'wc':0, 'ys':0, 'yc':0, 'ls':0, 'lc':0}
            rev_map = {v: k for k, v in unit_mapping.items()}

            for disp_name in display_order:
                src_name = rev_map.get(disp_name, disp_name)
                w = file_week['data'].get(src_name, {'stop':0, 'cit':0})
                y = file_year['data'].get(src_name, {'stop':0, 'cit':0})
                l = file_last_year['data'].get(src_name, {'stop':0, 'cit':0})
                if disp_name == '科技執法': w['stop'], y['stop'], l['stop'] = 0, 0, 0
                
                y_total = y['stop'] + y['cit']; l_total = l['stop'] + l['cit']
                row_data = [disp_name, w['stop'], w['cit'], y['stop'], y['cit']]
                if disp_name == '警備隊': row_data.extend(['—']*5)
                else:
                    diff = int(y_total - l_total); tgt = targets.get(disp_name, 0)
                    row_data.extend([l['stop'], l['cit'], diff])
                    if disp_name == '科技執法': row_data.extend(['—', '—'])
                    else: row_data.extend([tgt, f"{y_total/tgt:.2%}" if tgt>0 else 0])
                
                accum['ws']+=w['stop']; accum['wc']+=w['cit']; accum['ys']+=y['stop']; accum['yc']+=y['cit']; accum['ls']+=l['stop']; accum['lc']+=l['cit']
                rows.append(row_data)

            total_target = sum([v for k,v in targets.items() if k not in ['警備隊', '科技執法']])
            t_diff = (accum['ys']+accum['yc']) - (accum['ls']+accum['lc'])
            t_rate = (accum['ys']+accum['yc'])/total_target if total_target>0 else 0
            total_row = ['合計', accum['ws'], accum['wc'], accum['ys'], accum['yc'], accum['ls'], accum['lc'], t_diff, total_target, f"{t_rate:.2%}"]

            cols_header = ['單位', '本期_攔停', '本期_逕舉', '本年_攔停', '本年_逕舉', '去年_攔停', '去年_逕舉', '本年與去年比較', '目標值', '達成率']
            df_final = pd.DataFrame([total_row] + rows, columns=cols_header)

            st.subheader("📊 統計結果"); st.dataframe(df_final, use_container_width=True)

            # 產生 Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, sheet_name='Sheet1', startrow=3, index=False)
                ws = writer.sheets['Sheet1']
                fmt = writer.book.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
                ws.merge_range('A1:J1', '取締重大交通違規件數統計表', fmt)
                ws.write('A2', f"一、統計期間：{file_year['start']}~{file_year['end']}")
                if prog_text:
                    # 去除 markdown 符號寫入 Excel
                    clean_prog = prog_text.replace('*', '').replace('📅 ', '')
                    ws.write('A3', f"二、{clean_prog}")
            
            excel_data = output.getvalue()
            file_name_out = f'重點違規統計_{file_year["end"]}.xlsx'

            # 自動寄信邏輯
            if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
            file_ids = ",".join(sorted([f.name for f in uploaded_files]))
            email_receiver = st.secrets["email"]["user"]
            
            if file_ids not in st.session_state["sent_cache"]:
                with st.spinner(f"正在自動寄送報表至 {email_receiver}..."):
                    if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", "附件為重點違規統計報表(Excel)。", excel_data, file_name_out):
                        st.balloons()
                        st.success(f"✅ 郵件已發送至 {email_receiver}")
                        st.session_state["sent_cache"].add(file_ids)
            else:
                st.info(f"✅ 報表已於剛才發送至 {email_receiver}")

            st.download_button(label="📥 下載 Excel", data=excel_data, file_name=file_name_out, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e: st.error(f"發生錯誤：{e}")
