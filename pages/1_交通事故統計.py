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
from openpyxl.styles import Font, Alignment, Border, Side

# ==========================================
# 👇👇👇 【使用者設定區】 已設定完成 👇👇👇
# ==========================================

# 1. 您的 Gmail (寄件者)
MY_EMAIL = "mbmmiqamhnd@gmail.com" 

# 2. 您的應用程式密碼
MY_PASSWORD = "kvpw ymgn xawe qxnl" 

# 3. 收件者 (寄給自己)
TO_EMAIL = "mbmmiqamhnd@gmail.com"

# 4. SMTP 伺服器 (Gmail 預設不用改)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
# ==========================================


st.set_page_config(page_title="交通事故統計 (自動寄信版)", layout="wide", page_icon="🚑")
st.title("🚑 交通事故統計自動化系統")
st.markdown("### 📝 說明：系統將自動計算並將報表寄送至您的信箱。")

uploaded_files = st.file_uploader("請上傳 3 個事故報表檔案", accept_multiple_files=True, key="acc_uploader")

# --- 寄信函數 ---
def send_email_auto(attachment_data, filename):
    try:
        msg = MIMEMultipart()
        msg['From'] = MY_EMAIL
        msg['To'] = TO_EMAIL
        msg['Subject'] = f"交通事故統計報表 ({pd.Timestamp.now().strftime('%Y/%m/%d')})"
        
        body = "長官好，\n\n檢送本期交通事故統計報表如附件，請查照。\n\n(此郵件由系統自動發送)"
        msg.attach(MIMEText(body, 'plain'))

        # 附件處理
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment_data.getvalue())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={filename}')
        msg.attach(part)

        # 發送
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
            s.starttls()
            s.login(MY_EMAIL, MY_PASSWORD)
            s.send_message(msg)
        return True, f"✅ 報表已自動寄送至：{TO_EMAIL}"
    except smtplib.SMTPAuthenticationError:
        return False, "❌ 登入失敗：請確認應用程式密碼是否正確，或 Google 帳號設定有誤。"
    except Exception as e:
        return False, f"❌ 寄送失敗：{e}"

# --- 主程式 ---
if uploaded_files and st.button("🚀 開始分析並寄送", key="btn_acc"):
    with st.spinner("正在處理資料、生成報表並寄送郵件中..."):
        try:
            # === (A) 資料讀取與清理 ===
            def parse_raw(file_obj):
                try: return pd.read_csv(file_obj, header=None)
                except: file_obj.seek(0); return pd.read_excel(file_obj, header=None)

            def clean_data(df_raw):
                df_raw[0] = df_raw[0].astype(str)
                df_data = df_raw[df_raw[0].str.contains("總計|派出所|合計", na=False)].copy()
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

            # === (B) 智慧辨識 (跨年邏輯) ===
            files_meta = []
            for uploaded_file in uploaded_files:
                uploaded_file.seek(0)
                df = parse_raw(uploaded_file)
                found_dates = []
                for r in range(min(5, len(df))):
                    for c in range(min(3, len(df.columns))):
                        val = str(df.iloc[r, c])
                        dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', val)
                        if len(dates) >= 2:
                            found_dates = dates
                            break
                    if found_dates: break

                if found_dates:
                    start_y, start_m, start_d = map(int, found_dates[0])
                    end_y, end_m, end_d = map(int, found_dates[1])
                    d_start = date(start_y + 1911, start_m, start_d)
                    d_end = date(end_y + 1911, end_m, end_d)
                    duration_days = (d_end - d_start).days
                    raw_date_str = f"{start_y}/{start_m:02d}/{start_d:02d}-{end_y}/{end_m:02d}/{end_d:02d}"
                    files_meta.append({'file': uploaded_file, 'df': df, 'start_tuple': (start_y, start_m, start_d),
                                       'end_year': end_y, 'duration': duration_days, 'raw_date': raw_date_str})
                else:
                    files_meta.append({'file': uploaded_file, 'end_year': 0}) 

            # === (C) 檔案分配 ===
            files_meta.sort(key=lambda x: x.get('end_year', 0), reverse=True)
            df_wk = None; df_cur = None; df_lst = None
            h_wk = ""; h_cur = ""; h_lst = ""

            valid_files = [f for f in files_meta if f.get('end_year', 0) > 0]
            
            if len(valid_files) >= 3:
                current_year_end = valid_files[0]['end_year']
                current_files = [f for f in valid_files if f['end_year'] == current_year_end]
                past_files = [f for f in valid_files if f['end_year'] < current_year_end]

                if past_files:
                    past_files.sort(key=lambda x: x['end_year'], reverse=True)
                    t = past_files[0]
                    df_lst = clean_data(t['df']); h_lst = t['raw_date']

                if len(current_files) >= 2:
                    starts_on_jan1 = [f for f in current_files if f['start_tuple'][1] == 1 and f['start_tuple'][2] == 1]
                    cumu, wk = None, None
                    if len(starts_on_jan1) == 1:
                        cumu = starts_on_jan1[0]
                        wk = [f for f in current_files if f != cumu][0]
                    else:
                        current_files.sort(key=lambda x: x['duration'])
                        wk = current_files[0]; cumu = current_files[-1]
                    if cumu: df_cur = clean_data(cumu['df']); h_cur = cumu['raw_date']
                    if wk: df_wk = clean_data(wk['df']); h_wk = wk['raw_date']

            if df_wk is None or df_cur is None or df_lst is None:
                st.error("❌ 檔案辨識失敗，請確認檔案內容包含：去年累計、今年累計、今年週報。"); st.stop()

            # === (D) 計算 ===
            # A1
            a1_wk = df_wk[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'wk'})
            a1_cur = df_cur[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'cur'})
            a1_lst = df_lst[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'last'})
            m_a1 = pd.merge(a1_wk, a1_cur, on='Station_Short', how='outer')
            m_a1 = pd.merge(m_a1, a1_lst, on='Station_Short', how='outer').fillna(0)
            m_a1['Diff'] = m_a1['cur'] - m_a1['last']
            # A2
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

            # === (E) 產生 Excel 與 自動寄信 ===
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                a1_final.to_excel(writer, index=False, sheet_name='A1死亡人數')
                a2_final.to_excel(writer, index=False, sheet_name='A2受傷人數')
                # 樣式設定
                font_normal = Font(name='標楷體', size=12)
                font_red_bold = Font(name='標楷體', size=12, bold=True, color="FF0000")
                font_bold = Font(name='標楷體', size=12, bold=True)
                align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
                border_style = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                for sheet_name in ['A1死亡人數', 'A2受傷人數']:
                    ws = writer.book[sheet_name]
                    for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 20
                    for cell in ws[1]:
                        cell.alignment = align_center
                        cell.border = border_style
                        if any(x in str(cell.value) for x in ["本期", "累計", "/"]): cell.font = font_red_bold
                        else: cell.font = font_bold
                    for row in ws.iter_rows(min_row=2):
                        for cell in row:
                            cell.alignment = align_center
                            cell.border = border_style
                            cell.font = font_normal
            
            # --- 🚀 自動寄信觸發點 ---
            filename_excel = f'交通事故統計表_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx'
            
            # 呼叫寄信
            success, msg = send_email_auto(output, filename_excel)
            
            # 顯示結果
            if success:
                st.success(msg)
            else:
                st.warning(msg)

            # 顯示表格與下載按鈕
            col1, col2 = st.columns(2)
            with col1: st.subheader("📊 A1 死亡人數"); st.dataframe(a1_final, hide_index=True)
            with col2: st.subheader("📊 A2 受傷人數"); st.dataframe(a2_final, hide_index=True)

            st.download_button(label="📥 下載 Excel 報表", data=output.getvalue(), file_name=filename_excel, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e:
            st.error(f"系統錯誤：{e}")
