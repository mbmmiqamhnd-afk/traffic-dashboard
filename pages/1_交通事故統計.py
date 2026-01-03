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
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

# ==========================================
# 👇👇👇 【使用者設定區】 👇👇👇
# ==========================================
MY_EMAIL = "mbmmiqamhnd@gmail.com" 
MY_PASSWORD = "kvpw ymgn xawe qxnl" 
TO_EMAIL = "mbmmiqamhnd@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
# ==========================================

st.set_page_config(page_title="交通事故統計 (通用字體版)", layout="wide", page_icon="🚑")
st.title("🚑 交通事故統計 (上傳即寄出)")
st.markdown("### 📝 狀態：Excel 報表改用通用字體 (Calibri)，不再強制標楷體。")

# 1. 檔案上傳區
uploaded_files = st.file_uploader("請一次選取或拖曳 3 個報表檔案", accept_multiple_files=True, key="acc_uploader")

# --- 核心工具：產生紅黑雙色的 HTML 標題 ---
def format_html_header(text):
    """將文字轉換為 HTML，數字符號轉紅，其餘轉黑"""
    text = str(text)
    tokens = re.split(r'([0-9\(\)\/\-\.\%]+)', text)
    html_str = ""
    for token in tokens:
        if not token: continue
        if re.match(r'^[0-9\(\)\/\-\.\%]+$', token):
            html_str += f'<span style="color: red;">{token}</span>'
        else:
            html_str += f'<span style="color: black;">{token}</span>'
    return html_str

def render_styled_table(df, title):
    """在 Streamlit 渲染漂亮的 HTML 表格 (通用字體)"""
    st.subheader(title)
    
    df_display = df.copy()
    
    style = """
    <style>
        table.acc_table {
            font-family: sans-serif; /* 網頁也使用通用字體 */
            border-collapse: collapse;
            width: 100%;
            font-size: 16px;
        }
        table.acc_table th {
            border: 1px solid #000;
            padding: 8px;
            text-align: center !important;
            font-weight: bold;
            background-color: #f0f2f6;
        }
        table.acc_table td {
            border: 1px solid #000;
            padding: 8px;
            text-align: center !important;
            color: black;
        }
    </style>
    """
    
    html = f"{style}<table class='acc_table'><thead><tr>"
    
    for col in df_display.columns:
        styled_header = format_html_header(col)
        html += f"<th>{styled_header}</th>"
    html += "</tr></thead><tbody>"
    
    for _, row in df_display.iterrows():
        html += "<tr>"
        for val in row:
            html += f"<td>{val}</td>"
        html += "</tr>"
    html += "</tbody></table>"
    
    st.markdown(html, unsafe_allow_html=True)

# 2. 寄信函數
def send_email_auto(attachment_data, filename):
    try:
        msg = MIMEMultipart()
        msg['From'] = MY_EMAIL
        msg['To'] = TO_EMAIL
        msg['Subject'] = f"交通事故統計報表 ({pd.Timestamp.now().strftime('%Y/%m/%d')})"
        
        body = "長官好，\n\n檢送本期交通事故統計報表如附件 (Excel 已改為通用字體)，請查照。\n\n(此郵件由系統自動發送)"
        msg.attach(MIMEText(body, 'plain'))

        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(attachment_data.getvalue())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={filename}')
        msg.attach(part)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
            s.starttls()
            s.login(MY_EMAIL, MY_PASSWORD)
            s.send_message(msg)
        return True, f"✅ 報表已自動寄送至：{TO_EMAIL}"
    except Exception as e:
        return False, f"❌ 寄送失敗：{e}"

# 3. 自動處理邏輯
if uploaded_files:
    if len(uploaded_files) != 3:
        st.warning(f"⚠️ 目前已上傳 {len(uploaded_files)} 個檔案，請補齊至 3 個檔案。")
        st.stop()
    
    with st.spinner("⚡ 正在分析、格式化並寄送中..."):
        try:
            # === (A) 資料讀取與清理 ===
            def parse_raw(file_obj):
                try: return pd.read_csv(file_obj, header=None)
                except: file_obj.seek(0); return pd.read_excel(file_obj, header=None)

            def clean_data(df_raw):
                df_raw[0] = df_raw[0].astype(str)
                df_data = df_raw[df_raw[0].str.contains("所|總計|合計", na=False)].copy()
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

            # === (B) 智慧辨識 ===
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
                    raw_date_str = f"{start_m:02d}/{start_d:02d}-{end_m:02d}/{end_d:02d}"
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
                st.error("❌ 檔案辨識失敗。"); st.stop()

            # === (D) 合併與計算 ===
            target_stations = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']

            def process_and_sum(df_main, value_cols):
                df_sub = df_main[df_main['Station_Short'].isin(target_stations)].copy()
                df_sub['Station_Short'] = pd.Categorical(df_sub['Station_Short'], categories=target_stations, ordered=True)
                df_sub.sort_values('Station_Short', inplace=True)
                sum_values = df_sub[value_cols].sum()
                row_total = pd.DataFrame([{'Station_Short': '合計', **sum_values.to_dict()}])
                return pd.concat([row_total, df_sub], ignore_index=True)

            # A1
            a1_wk = df_wk[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'wk'})
            a1_cur = df_cur[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'cur'})
            a1_lst = df_lst[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'last'})
            m_a1 = pd.merge(a1_wk, a1_cur, on='Station_Short', how='outer')
            m_a1 = pd.merge(m_a1, a1_lst, on='Station_Short', how='outer').fillna(0)
            m_a1 = process_and_sum(m_a1, ['wk', 'cur', 'last'])
            m_a1['Diff'] = m_a1['cur'] - m_a1['last']

            # A2
            a2_wk = df_wk[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'wk'})
            a2_cur = df_cur[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'cur'})
            a2_lst = df_lst[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'last'})
            m_a2 = pd.merge(a2_wk, a2_cur, on='Station_Short', how='outer')
            m_a2 = pd.merge(m_a2, a2_lst, on='Station_Short', how='outer').fillna(0)
            m_a2 = process_and_sum(m_a2, ['wk', 'cur', 'last'])
            m_a2['Diff'] = m_a2['cur'] - m_a2['last']
            m_a2['Pct_Str'] = m_a2.apply(lambda x: f"{(x['Diff']/x['last']):.2%}" if x['last']!=0 else "-", axis=1)
            m_a2['Prev'] = "-"

            # 整理欄位
            a1_final = m_a1[['Station_Short', 'wk', 'cur', 'last', 'Diff']].copy()
            a1_final.columns = ['統計期間', f'本期({h_wk})', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較']
            
            a2_final = m_a2[['Station_Short', 'wk', 'Prev', 'cur', 'last', 'Diff', 'Pct_Str']].copy()
            a2_final.columns = ['統計期間', f'本期({h_wk})', '前期', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較', '本年較去年增減比例']

            # === (E) 產生 Excel (改為 Calibri 通用字體) ===
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                a1_final.to_excel(writer, index=False, sheet_name='A1死亡人數')
                a2_final.to_excel(writer, index=False, sheet_name='A2受傷人數')
                
                # ⬇️ 這裡改用 Calibri (Excel 預設字體)
                font_black = InlineFont(rFont='Calibri', sz=12, b=True, color='000000')
                font_red = InlineFont(rFont='Calibri', sz=12, b=True, color='FF0000')
                font_normal_cell = Font(name='Calibri', size=12)
                
                align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
                border_style = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                
                def make_rich_text(text):
                    text = str(text)
                    rich_text = CellRichText()
                    tokens = re.split(r'([0-9\(\)\/\-\.\%]+)', text)
                    for token in tokens:
                        if not token: continue
                        if re.match(r'^[0-9\(\)\/\-\.\%]+$', token):
                            rich_text.append(TextBlock(font_red, token))
                        else:
                            rich_text.append(TextBlock(font_black, token))
                    return rich_text

                for sheet_name in ['A1死亡人數', 'A2受傷人數']:
                    ws = writer.book[sheet_name]
                    for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 20
                    
                    for i, cell in enumerate(ws[1]):
                        original_value = cell.value
                        cell.value = make_rich_text(original_value)
                        cell.alignment = align_center
                        cell.border = border_style

                    for row in ws.iter_rows(min_row=2):
                        for cell in row:
                            cell.alignment = align_center
                            cell.border = border_style
                            cell.font = font_normal_cell
            
            # 🔥 自動寄信
            filename_excel = f'交通事故統計表_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx'
            success, msg = send_email_auto(output, filename_excel)
            
            if success:
                st.balloons()
                st.success(msg)
            else:
                st.error(msg)

            # === 🔥 網頁顯示 ===
            col1, col2 = st.columns(2)
            with col1: 
                render_styled_table(a1_final, "📊 A1 死亡人數")
            with col2: 
                render_styled_table(a2_final, "📊 A2 受傷人數")

        except Exception as e:
            st.error(f"系統錯誤：{e}")
