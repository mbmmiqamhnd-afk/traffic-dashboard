import streamlit as st
import pandas as pd
import io
import re
import smtplib
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

st.set_page_config(page_title="交通事故統計", layout="wide", page_icon="🚑")
st.title("🚑 交通事故統計 (A1/A2)")

st.markdown("""
### 📝 使用說明
1. 請上傳 3 個原始報表檔案 (本週、今年累計、去年累計)。
2. 系統會**全自動掃描日期**並進行邏輯判斷。
3. **自動計算合計數值**，確保數據準確。
4. 上傳後自動分析，完成後可寄信。
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
uploaded_files = st.file_uploader("請上傳 3 個事故報表檔案", accept_multiple_files=True, key="acc_uploader")

if uploaded_files:
    if len(uploaded_files) < 3:
        st.warning("⏳ 請上傳滿 3 個檔案以開始計算...")
    else:
        try:
            # ==========================================
            # 1. 讀取與清理函數
            # ==========================================
            def parse_raw(file_obj):
                try: 
                    file_obj.seek(0)
                    return pd.read_csv(file_obj, header=None)
                except: 
                    file_obj.seek(0)
                    return pd.read_excel(file_obj, header=None)

            def extract_date_info(df):
                head_str = df.head(20).to_string()
                matches = re.findall(r'(\d{3})[./-](\d{1,2})[./-](\d{1,2})', head_str)
                if len(matches) >= 2:
                    y1, m1, d1 = map(int, matches[0])
                    y2, m2, d2 = map(int, matches[1])
                    start_dt = datetime(y1 + 1911, m1, d1)
                    end_dt = datetime(y2 + 1911, m2, d2)
                    return start_dt, end_dt, f"{y1}/{m1:02d}/{d1:02d}~{y2}/{m2:02d}/{d2:02d}"
                return None, None, None

            def clean_data(df_raw):
                # 1. 只抓取「派出所」資料，排除原始的總計(以免格式跑掉)
                df_data = df_raw[df_raw[0].astype(str).str.contains("派出所", na=False)].copy()
                df_data = df_data.reset_index(drop=True)
                
                # 2. 定義欄位
                for i in range(11):
                    if i not in df_data.columns: df_data[i] = 0
                
                target_cols = {
                    0: "Station", 1: "Total_Cases", 2: "Total_Deaths", 3: "Total_Injuries",
                    4: "A1_Cases", 5: "A1_Deaths", 6: "A1_Injuries",
                    7: "A2_Cases", 8: "A2_Deaths", 9: "A2_Injuries", 10: "A3_Cases"
                }
                df_data = df_data.rename(columns=target_cols)
                df_data = df_data[list(target_cols.values())] 
                
                # 3. 轉數值
                for col in list(target_cols.values())[1:]:
                    df_data[col] = pd.to_numeric(df_data[col].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
                
                # 4. 簡化單位名稱
                df_data['Station_Short'] = df_data['Station'].astype(str).str.replace('派出所', '所').str.strip()

                # 5. 🔥 關鍵修改：自動重新計算「合計」
                # 將所有數值欄位加總
                sum_row = df_data.sum(numeric_only=True)
                # 將 Series 轉為 DataFrame (單列)
                sum_df = pd.DataFrame([sum_row])
                # 補上單位名稱
                sum_df['Station_Short'] = '合計'
                
                # 合併：合計放第一列，後面接各派出所
                df_final = pd.concat([sum_df, df_data], ignore_index=True)
                
                return df_final

            # ==========================================
            # 2. 智慧辨識檔案邏輯
            # ==========================================
            file_info_list = []
            
            for f in uploaded_files:
                df = parse_raw(f)
                start_dt, end_dt, raw_date_str = extract_date_info(df)
                
                if start_dt:
                    duration = (end_dt - start_dt).days
                    file_info_list.append({
                        'file_obj': f, 'df': df, 'start_dt': start_dt, 'end_dt': end_dt,
                        'duration': duration, 'raw_date': raw_date_str, 'name': f.name
                    })

            if len(file_info_list) < 3:
                st.error("❌ 無法識別出 3 個有效檔案的日期，請檢查檔案內容。")
            else:
                file_info_list.sort(key=lambda x: x['start_dt'])
                data_lst = file_info_list[0] # 去年
                
                remaining = file_info_list[1:]
                remaining.sort(key=lambda x: x['duration'], reverse=True)
                data_cur = remaining[0] # 本年
                data_wk = remaining[1]  # 本期
                
                df_wk = clean_data(data_wk['df'])
                df_cur = clean_data(data_cur['df'])
                df_lst = clean_data(data_lst['df'])
                
                h_wk = data_wk['raw_date']
                h_cur = data_cur['raw_date']
                h_lst = data_lst['raw_date']

                st.info(f"✅ 成功辨識：本期({h_wk})、本年({h_cur})、去年({h_lst})")

                # ==========================================
                # 3. 統計運算區
                # ==========================================
                # --- A1 ---
                a1_wk = df_wk[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'wk'})
                a1_cur = df_cur[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'cur'})
                a1_lst = df_lst[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'last'})
                
                m_a1 = pd.merge(a1_wk, a1_cur, on='Station_Short', how='outer')
                m_a1 = pd.merge(m_a1, a1_lst, on='Station_Short', how='outer').fillna(0)
                m_a1['Diff'] = m_a1['cur'] - m_a1['last']

                # --- A2 ---
                a2_wk = df_wk[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'wk'})
                a2_cur = df_cur[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'cur'})
                a2_lst = df_lst[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'last'})
                
                m_a2 = pd.merge(a2_wk, a2_cur, on='Station_Short', how='outer')
                m_a2 = pd.merge(m_a2, a2_lst, on='Station_Short', how='outer').fillna(0)
                m_a2['Diff'] = m_a2['cur'] - m_a2['last']
                m_a2['Pct_Str'] = m_a2.apply(lambda x: f"{(x['Diff']/x['last']):.2%}" if x['last']!=0 else "-", axis=1)
                m_a2['Prev'] = "-"

                # 排序
                target_order = ['合計', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
                for m in [m_a1, m_a2]:
                    m['Station_Short'] = m['Station_Short'].astype(str)
                    m['Station_Short'] = pd.Categorical(m['Station_Short'], categories=target_order, ordered=True)
                    m.sort_values('Station_Short', inplace=True)

                a1_final = m_a1[['Station_Short', 'wk', 'cur', 'last', 'Diff']].copy()
                a1_final.columns = ['單位', f'本期({h_wk})', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較']
                
                a2_final = m_a2[['Station_Short', 'wk', 'Prev', 'cur', 'last', 'Diff', 'Pct_Str']].copy()
                a2_final.columns = ['單位', f'本期({h_wk})', '前期', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較', '本年較去年增減比例']

                st.subheader("📊 A1 死亡人數統計")
                st.dataframe(a1_final, use_container_width=True, hide_index=True)
                
                st.subheader("📊 A2 受傷人數統計")
                st.dataframe(a2_final, use_container_width=True, hide_index=True)

                # ==========================================
                # 4. 檔案產生與寄信
                # ==========================================
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    a1_final.to_excel(writer, index=False, sheet_name='A1死亡人數')
                    a2_final.to_excel(writer, index=False, sheet_name='A2受傷人數')
                
                excel_data = output.getvalue()
                file_name_out = f'交通事故統計表_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx'

                # 自動寄信
                if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
                file_ids = ",".join(sorted([f.name for f in uploaded_files]))
                email_receiver = st.secrets["email"]["user"]
                
                if file_ids not in st.session_state["sent_cache"]:
                    with st.spinner(f"正在自動寄送報表至 {email_receiver}..."):
                        if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", "附件為本期事故統計報表(Excel)。", excel_data, file_name_out):
                            st.balloons()
                            st.success(f"✅ 郵件已發送至 {email_receiver}")
                            st.session_state["sent_cache"].add(file_ids)
                else:
                    st.info(f"✅ 報表已於剛才發送至 {email_receiver}")

                st.download_button(label="📥 下載 Excel", data=excel_data, file_name=file_name_out, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e:
            st.error(f"發生未預期的錯誤：{e}")
