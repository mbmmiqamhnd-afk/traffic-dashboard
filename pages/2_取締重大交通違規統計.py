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
2. **自動分析** 並 **自動寄出**。
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
            # --- 1. 資料解析區 ---
            def parse_file_content(uploaded_file):
                content = uploaded_file.getvalue()
                df = None; start_date = ""; header_idx = -1
                is_excel = uploaded_file.name.endswith(('.xlsx', '.xls'))
                try:
                    if is_excel:
                        df_raw = pd.read_excel(io.BytesIO(content), header=None, nrows=20)
                        for i, row in df_raw.iterrows():
                            row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
                            if not start_date:
                                match = re.search(r'入案日期[：:]?\s*(\d{3,7})\s*至\s*(\d{3,7})', row_str)
                                if match: start_date, end_date = match.group(1), match.group(2)
                            if "單位" in row_str and "酒後" in row_str: header_idx = i
                        if header_idx != -1: df = pd.read_excel(io.BytesIO(content), header=header_idx)
                    else:
                        try: text = content.decode('utf-8')
                        except: text = content.decode('cp950', errors='ignore')
                        lines = text.splitlines()
                        for i, line in enumerate(lines):
                            match = re.search(r'入案日期[：:]?\s*(\d{3,7})\s*至\s*(\d{3,7})', line)
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
                
                try:
                    d1 = date(int(start_date[:3])+1911, int(start_date[3:5]), int(start_date[5:]))
                    d2 = date(int(end_date[:3])+1911, int(end_date[3:5]), int(end_date[5:]))
                    duration = (d2 - d1).days
                except: duration = 0
                return {'data': unit_data, 'start': start_date, 'end': end_date, 'duration': duration}

            parsed_files = []
            for f in uploaded_files:
                res = parse_file_content(f)
                if res: parsed_files.append(res)
            
            if len(parsed_files) < 3: st.error("有效檔案不足！"); st.stop()

            parsed_files.sort(key=lambda x: int(x['start'].replace('/','').replace('.','')))
            file_last_year = parsed_files[0]
            others = parsed_files[1:]
            others.sort(key=lambda x: x['duration'], reverse=True)
            file_year = others[0]; file_week = others[1]

            st.success(f"✅ 檔案識別成功：本年({file_year['start']})、去年({file_last_year['start']})、本期({file_week['start']})")

            # --- 2. 統計運算區 ---
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

            # --- 3. 檔案產生與自動寄信區 ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, sheet_name='Sheet1', startrow=3, index=False)
                ws = writer.sheets['Sheet1']
                fmt = writer.book.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
                ws.merge_range('A1:J1', '取締重大交通違規件數統計表', fmt)
                ws.write('A2', f"一、統計期間：{file_year['start']}~{file_year['end']}")
            
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
