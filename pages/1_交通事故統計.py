import streamlit as st
import pandas as pd
import io
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

st.set_page_config(page_title="交通事故統計", layout="wide", page_icon="🚑")
st.title("🚑 交通事故統計 (A1/A2)")

st.markdown("""
### 📝 使用說明
1. 請上傳 3 個原始報表檔案。
2. 系統自動分辨日期並計算。
3. **完成後自動寄送報表至您的信箱**。
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
            # --- 1. 資料處理區 ---
            def parse_raw(file_obj):
                try: return pd.read_csv(file_obj, header=None)
                except: file_obj.seek(0); return pd.read_excel(file_obj, header=None)

            def clean_data(df_raw):
                df_data = df_raw[df_raw[0].notna()].copy()
                df_data = df_data[df_data[0].astype(str).str.contains("總計|派出所")].copy()
                df_data = df_data.reset_index(drop=True)
                columns_map = {
                    0: "Station", 1: "Total_Cases", 2: "Total_Deaths", 3: "Total_Injuries",
                    4: "A1_Cases", 5: "A1_Deaths", 6: "A1_Injuries",
                    7: "A2_Cases", 8: "A2_Deaths", 9: "A2_Injuries", 10: "A3_Cases"
                }
                for i in range(11):
                    if i not in df_data.columns: df_data[i] = 0
                df_data = df_data.rename(columns=columns_map)
                df_data = df_data[list(columns_map.values())]
                for col in list(columns_map.values())[1:]:
                    df_data[col] = pd.to_numeric(df_data[col].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
                df_data['Station_Short'] = df_data['Station'].astype(str).str.replace('派出所', '所').str.replace('總計', '合計')
                return df_data

            file_data_map = {}
            for uploaded_file in uploaded_files:
                df = parse_raw(uploaded_file)
                try:
                    raw_str = str(df.iloc[1, 0])
                    date_str = raw_str.replace("統計日期：", "").strip()
                    dates = re.findall(r'(\d{3})/(\d{2})/(\d{2})', date_str)
                    if dates:
                        start_y, start_m, start_d = map(int, dates[0])
                        end_y, end_m, end_d = map(int, dates[1])
                        month_diff = (end_y - start_y) * 12 + (end_m - start_m)
                        category = 'weekly' if (month_diff == 0 and (end_d - start_d) < 20) else 'cumulative'
                        file_data_map[uploaded_file.name] = {'df': df, 'category': category, 'year': start_y, 'raw_date': date_str}
                except: pass

            df_wk = None; df_cur = None; df_lst = None
            h_wk = ""; h_cur = ""; h_lst = ""

            for fname, data in file_data_map.items():
                if data['category'] == 'weekly':
                    df_wk = clean_data(data['df']); h_wk = data['raw_date']; break
            
            cumu_files = [d for d in file_data_map.values() if d['category'] == 'cumulative']
            if len(cumu_files) >= 2:
                cumu_files.sort(key=lambda x: x['year'], reverse=True)
                df_cur = clean_data(cumu_files[0]['df']); h_cur = cumu_files[0]['raw_date']
                df_lst = clean_data(cumu_files[1]['df']); h_lst = cumu_files[1]['raw_date']

            if df_wk is None or df_cur is None or df_lst is None:
                st.error("❌ 無法識別完整的 3 份檔案，請檢查檔案內容日期。")
            else:
                # --- 2. 統計運算區 ---
                a1_wk = df_wk[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'wk'})
                a1_cur = df_cur[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'cur'})
                a1_lst = df_lst[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'last'})
                m_a1 = pd.merge(a1_wk, a1_cur, on='Station_Short', how='outer')
                m_a1 = pd.merge(m_a1, a1_lst, on='Station_Short', how='outer').fillna(0)
                m_a1['Diff'] = m_a1['cur'] - m_a1['last']

                a2_wk = df_wk[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'wk'})
                a2_cur = df_cur[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'cur'})
                a2_lst = df_lst[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'last'})
                m_a2 = pd.merge(a2_wk, a2_cur, on='Station_Short', how='outer')
                m_a2 = pd.merge(m_a2, a2_lst, on='Station_Short', how='outer').fillna(0)
                m_a2['Diff'] = m_a2['cur'] - m_a2['last']
                m_a2['Pct_Str'] = m_a2.apply(lambda x: f"{(x['Diff']/x['last']):.2%}" if x['last']!=0 else "-", axis=1)
                m_a2['Prev'] = "-"

                target_order = ['合計', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
                for m in [m_a1, m_a2]:
                    m['Station_Short'] = pd.Categorical(m['Station_Short'], categories=target_order, ordered=True)
                    m.sort_values('Station_Short', inplace=True)

                a1_final = m_a1[['Station_Short', 'wk', 'cur', 'last', 'Diff']].copy()
                a1_final.columns = ['單位', f'本期({h_wk})', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較']
                
                a2_final = m_a2[['Station_Short', 'wk', 'Prev', 'cur', 'last', 'Diff', 'Pct_Str']].copy()
                a2_final.columns = ['單位', f'本期({h_wk})', '前期', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較', '本年較去年增減比例']

                st.success("✅ 分析完成！")
                st.subheader("📊 A1 死亡人數統計"); st.dataframe(a1_final, use_container_width=True, hide_index=True)
                st.subheader("📊 A2 受傷人數統計"); st.dataframe(a2_final, use_container_width=True, hide_index=True)

                # --- 3. 檔案產生與自動寄信區 ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    a1_final.to_excel(writer, index=False, sheet_name='A1死亡人數')
                    a2_final.to_excel(writer, index=False, sheet_name='A2受傷人數')
                
                excel_data = output.getvalue()
                file_name_out = f'交通事故統計表_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx'

                # 自動寄信邏輯 (防重複發送)
                if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
                file_ids = ",".join(sorted([f.name for f in uploaded_files])) # 產生本次上傳的唯一碼

                email_receiver = st.secrets["email"]["user"]
                
                if file_ids not in st.session_state["sent_cache"]:
                    with st.spinner(f"正在自動寄送報表至 {email_receiver}..."):
                        if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", "附件為本期事故統計報表(Excel)。", excel_data, file_name_out):
                            st.balloons()
                            st.success(f"✅ 郵件已發送至 {email_receiver}")
                            st.session_state["sent_cache"].add(file_ids) # 標記為已發送
                else:
                    st.info(f"✅ 報表已於剛才發送至 {email_receiver} (若需重寄請重新整理頁面)")

                st.download_button(label="📥 下載 Excel", data=excel_data, file_name=file_name_out, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e: st.error(f"發生錯誤：{e}")
