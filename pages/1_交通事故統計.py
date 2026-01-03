import streamlit as st
import pandas as pd
import io
import re
import smtplib
import gspread
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

# Google Sheet 設定
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
# ==========================================

st.set_page_config(page_title="交通事故統計 (純寫入數據版)", layout="wide", page_icon="🚑")
st.title("🚑 交通事故統計 (上傳即寄出 + 純數據寫入)")
st.markdown("### 📝 狀態：同步時**完全保留**試算表原本的格式 (合併、底色、邊框)，僅更新數值與紅黑字。")

# 1. 檔案上傳區
uploaded_files = st.file_uploader("請一次選取或拖曳 3 個報表檔案", accept_multiple_files=True, key="acc_uploader")

# --- 工具函數 1: HTML 標題專用 ---
def format_html_header(text):
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

# --- 工具函數 2: Google Sheets API Rich Text 專用 ---
def get_gsheet_rich_text_req(sheet_id, row_idx, col_idx, text):
    """
    產生 textFormatRuns 請求。
    🔥 關鍵：fields 僅指定 userEnteredValue 和 textFormatRuns。
    這保證了背景色、邊框、水平對齊等格式【絕對不會】被更動。
    """
    text = str(text)
    tokens = re.split(r'([0-9\(\)\/\-\.\%]+)', text)
    runs = []
    current_pos = 0
    
    for token in tokens:
        if not token: continue
        
        # 只設定顏色與粗體，不設定 fontSize 或 fontFamily，沿用原本試算表的設定
        color = {"red": 0, "green": 0, "blue": 0}
        
        if re.match(r'^[0-9\(\)\/\-\.\%]+$', token):
            color = {"red": 1, "green": 0, "blue": 0} # 紅色
            
        runs.append({
            "startIndex": current_pos,
            "format": {
                "foregroundColor": color,
                "bold": True
            }
        })
        current_pos += len(token)
    
    return {
        "updateCells": {
            "rows": [{
                "values": [{
                    "userEnteredValue": {"stringValue": text},
                    "textFormatRuns": runs
                }]
            }],
            "fields": "userEnteredValue,textFormatRuns", # 🔥 鎖定更新範圍，保護格式
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row_idx,
                "endRowIndex": row_idx + 1,
                "startColumnIndex": col_idx,
                "endColumnIndex": col_idx + 1
            }
        }
    }

def render_styled_table(df, title):
    st.subheader(title)
    df_display = df.copy()
    style = """
    <style>
        table.acc_table {font-family: sans-serif; border-collapse: collapse; width: 100%; font-size: 16px; background-color: #ffffff;}
        table.acc_table th {border: 1px solid #000; padding: 8px; text-align: center !important; font-weight: bold; background-color: #f0f2f6; color: #000000;}
        table.acc_table td {border: 1px solid #000; padding: 8px; text-align: center !important; background-color: #ffffff !important;}
    </style>
    """
    html = f"{style}<table class='acc_table'><thead><tr>"
    for col in df_display.columns:
        html += f"<th>{format_html_header(col)}</th>"
    html += "</tr></thead><tbody>"
    for _, row in df_display.iterrows():
        html += "<tr>"
        for col_name, val in row.items():
            color = "#000000"
            display_val = f"{int(val)}" if isinstance(val, (int, float)) else str(val)
            if "比較" in col_name and isinstance(val, (int, float)) and val > 0: color = "red"
            elif "增減" in col_name and "-" not in display_val and display_val != "0.00%" and display_val != "-": color = "red"
            html += f'<td style="color: {color};">{display_val}</td>'
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
        body = "長官好，\n\n檢送本期交通事故統計報表如附件 (數據已同步至 Google 試算表，原始格式已保留)，請查照。\n\n(此郵件由系統自動發送)"
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

# 3. Google Sheets 同步函數 (🔥 核心修改：純寫入模式)
def sync_to_gsheet(df_a1, df_a2):
    try:
        if "gcp_service_account" not in st.secrets:
            return False, "❌ Secrets 中找不到 [gcp_service_account] 設定，無法同步。"

        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        
        def update_sheet_values_only(ws_index, df, title_text):
            try:
                ws = sh.get_worksheet(ws_index)
                
                # 1. 清除舊數據 (僅清除值 Values，保留格式)
                # "batch_clear" 預設行為就是清除 values，不會動 formatting
                ws.batch_clear(["A3:Z100"]) 
                
                # 2. 更新 Row 1 (大標題) - 僅更新文字，不動合併/字體
                ws.update_acell('A1', title_text)
                
                # 3. 更新 Row 3+ (數據內容) - 僅更新值，不動邊框/對齊
                data_rows = []
                for row in df.values.tolist():
                    data_rows.append([int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in row])
                
                if data_rows:
                    ws.update(range_name='A3', values=data_rows)

                # 4. 更新 Row 2 (欄位標題) - 更新文字並套用紅黑字，但不依賴 userEnteredFormat
                reqs = []
                for col_idx, col_name in enumerate(df.columns):
                    reqs.append(get_gsheet_rich_text_req(ws.id, 1, col_idx, col_name))
                
                if reqs:
                    sh.batch_update({"requests": reqs})
                    
                return True
            except Exception as e:
                raise e

        # 執行 A1 同步 (第3分頁)
        try:
            update_sheet_values_only(2, df_a1, "A1類交通事故死亡人數統計表")
        except Exception as e:
            return False, f"❌ A1 同步失敗: {e}"

        # 執行 A2 同步 (第4分頁)
        try:
            update_sheet_values_only(3, df_a2, "A2類交通事故受傷人數統計表")
        except Exception as e:
            return False, f"❌ A2 同步失敗: {e}"
        
        return True, "✅ Google 試算表同步成功 (格式完美保留)"
    except Exception as e:
        return False, f"❌ Google 試算表連線失敗: {e}"

# 4. 主流程
if uploaded_files:
    if len(uploaded_files) != 3:
        st.warning(f"⚠️ 目前已上傳 {len(uploaded_files)} 個檔案，請補齊至 3 個檔案。")
        st.stop()
    
    with st.spinner("⚡ 正在分析、同步雲端並寄送中..."):
        try:
            # (A) 資料讀取與清理
            def parse_raw(file_obj):
                try: return pd.read_csv(file_obj, header=None)
                except: file_obj.seek(0); return pd.read_excel(file_obj, header=None)

            def clean_data(df_raw):
                df_raw[0] = df_raw[0].astype(str)
                df_data = df_raw[df_raw[0].str.contains("所|總計|合計", na=False)].copy()
                df_data = df_data.reset_index(drop=True)
                columns_map = {0: "Station", 1: "Total_Cases", 2: "Total_Deaths", 3: "Total_Injuries", 4: "A1_Cases", 5: "A1_Deaths", 6: "A1_Injuries", 7: "A2_Cases", 8: "A2_Deaths", 9: "A2_Injuries", 10: "A3_Cases"}
                for i in range(11):
                    if i not in df_data.columns: df_data[i] = 0
                df_data = df_data.rename(columns=columns_map)
                df_data = df_data[list(columns_map.values())]
                for col in list(columns_map.values())[1:]:
                    df_data[col] = pd.to_numeric(df_data[col].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
                df_data['Station_Short'] = df_data['Station'].astype(str).str.replace('派出所', '所').str.replace('總計', '合計')
                return df_data

            # (B) 智慧辨識
            files_meta = []
            for uploaded_file in uploaded_files:
                uploaded_file.seek(0)
                df = parse_raw(uploaded_file)
                found_dates = []
                for r in range(min(5, len(df))):
                    for c in range(min(3, len(df.columns))):
                        val = str(df.iloc[r, c])
                        dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', val)
                        if len(dates) >= 2: found_dates = dates; break
                    if found_dates: break
                if found_dates:
                    start_y, start_m, start_d = map(int, found_dates[0])
                    end_y, end_m, end_d = map(int, found_dates[1])
                    d_start = date(start_y + 1911, start_m, start_d)
                    d_end = date(end_y + 1911, end_m, end_d)
                    duration_days = (d_end - d_start).days
                    raw_date_str = f"{start_m:02d}/{start_d:02d}-{end_m:02d}/{end_d:02d}"
                    files_meta.append({'file': uploaded_file, 'df': df, 'end_year': end_y, 'duration': duration_days, 'raw_date': raw_date_str, 'start_tuple': (start_y, start_m, start_d)})
                else: files_meta.append({'file': uploaded_file, 'end_year': 0})

            # (C) 檔案分配
            files_meta.sort(key=lambda x: x.get('end_year', 0), reverse=True)
            df_wk, df_cur, df_lst, h_wk, h_cur, h_lst = None, None, None, "", "", ""
            valid_files = [f for f in files_meta if f.get('end_year', 0) > 0]
            if len(valid_files) >= 3:
                current_year_end = valid_files[0]['end_year']
                current_files = [f for f in valid_files if f['end_year'] == current_year_end]
                past_files = [f for f in valid_files if f['end_year'] < current_year_end]
                if past_files:
                    past_files.sort(key=lambda x: x['end_year'], reverse=True)
                    t = past_files[0]; df_lst = clean_data(t['df']); h_lst = t['raw_date']
                if len(current_files) >= 2:
                    starts_on_jan1 = [f for f in current_files if f['start_tuple'][1] == 1 and f['start_tuple'][2] == 1]
                    cumu, wk = None, None
                    if len(starts_on_jan1) == 1: cumu = starts_on_jan1[0]; wk = [f for f in current_files if f != cumu][0]
                    else: current_files.sort(key=lambda x: x['duration']); wk = current_files[0]; cumu = current_files[-1]
                    if cumu: df_cur = clean_data(cumu['df']); h_cur = cumu['raw_date']
                    if wk: df_wk = clean_data(wk['df']); h_wk = wk['raw_date']
            if df_wk is None or df_cur is None or df_lst is None: st.error("❌ 檔案辨識失敗。"); st.stop()

            # (D) 合併與計算
            target_stations = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
            def process_and_sum(df_main, value_cols):
                df_sub = df_main[df_main['Station_Short'].isin(target_stations)].copy()
                df_sub['Station_Short'] = pd.Categorical(df_sub['Station_Short'], categories=target_stations, ordered=True)
                df_sub.sort_values('Station_Short', inplace=True)
                sum_values = df_sub[value_cols].sum()
                row_total = pd.DataFrame([{'Station_Short': '合計', **sum_values.to_dict()}])
                return pd.concat([row_total, df_sub], ignore_index=True)

            a1_wk = df_wk[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'wk'})
            a1_cur = df_cur[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'cur'})
            a1_lst = df_lst[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'last'})
            m_a1 = pd.merge(a1_wk, a1_cur, on='Station_Short', how='outer')
            m_a1 = pd.merge(m_a1, a1_lst, on='Station_Short', how='outer').fillna(0)
            m_a1 = process_and_sum(m_a1, ['wk', 'cur', 'last'])
            m_a1['Diff'] = m_a1['cur'] - m_a1['last']

            a2_wk = df_wk[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'wk'})
            a2_cur = df_cur[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'cur'})
            a2_lst = df_lst[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'last'})
            m_a2 = pd.merge(a2_wk, a2_cur, on='Station_Short', how='outer')
            m_a2 = pd.merge(m_a2, a2_lst, on='Station_Short', how='outer').fillna(0)
            m_a2 = process_and_sum(m_a2, ['wk', 'cur', 'last'])
            m_a2['Diff'] = m_a2['cur'] - m_a2['last']
            m_a2['Pct_Str'] = m_a2.apply(lambda x: f"{(x['Diff']/x['last']):.2%}" if x['last']!=0 else "-", axis=1)
            m_a2['Prev'] = "-"

            a1_final = m_a1[['Station_Short', 'wk', 'cur', 'last', 'Diff']].copy()
            a1_final.columns = ['統計期間', f'本期({h_wk})', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較']
            a2_final = m_a2[['Station_Short', 'wk', 'Prev', 'cur', 'last', 'Diff', 'Pct_Str']].copy()
            a2_final.columns = ['統計期間', f'本期({h_wk})', '前期', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較', '本年較去年增減比例']

            # (E) 產生 Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                a1_final.to_excel(writer, index=False, sheet_name='A1死亡人數')
                a2_final.to_excel(writer, index=False, sheet_name='A2受傷人數')
                font_black = InlineFont(rFont='Calibri', sz=12, b=True, color='000000')
                font_red = InlineFont(rFont='Calibri', sz=12, b=True, color='FF0000')
                font_content_black = Font(name='Calibri', size=12, color='000000')
                font_content_red = Font(name='Calibri', size=12, color='FF0000')
                align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
                border_style = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                
                def make_header_rich_text(text):
                    text = str(text)
                    rich = CellRichText()
                    for t in re.split(r'([0-9\(\)\/\-\.\%]+)', text):
                        if t: rich.append(TextBlock(font_red if re.match(r'^[0-9\(\)\/\-\.\%]+$', t) else font_black, t))
                    return rich

                for sn in ['A1死亡人數', 'A2受傷人數']:
                    ws = writer.book[sn]
                    header_names = [c.value for c in ws[1]]
                    for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 20
                    for cell in ws[1]:
                        cell.value = make_header_rich_text(cell.value)
                        cell.alignment = align_center
                        cell.border = border_style
                    for row in ws.iter_rows(min_row=2):
                        for idx, cell in enumerate(row):
                            if isinstance(cell.value, (int, float)): cell.value = int(cell.value)
                            target_font = font_content_black
                            col_n = header_names[idx]
                            if "比較" in str(col_n) and isinstance(cell.value, (int, float)) and cell.value > 0: target_font = font_content_red
                            elif "增減" in str(col_n) and "-" not in str(cell.value) and str(cell.value) not in ["0.00%", "-"]: target_font = font_content_red
                            cell.font = target_font
                            cell.alignment = align_center
                            cell.border = border_style

            # (F) 同步到 Google Sheet (🔥 標題 + Rich Text)
            gs_success, gs_msg = sync_to_gsheet(a1_final, a2_final)
            if gs_success: st.write(gs_msg)
            else: st.error(gs_msg)

            # (G) 自動寄信
            filename_excel = f'交通事故統計表_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx'
            success, msg = send_email_auto(output, filename_excel)
            if success:
                st.balloons()
                st.success(msg)
            else:
                st.error(msg)

            # (H) 網頁顯示
            col1, col2 = st.columns(2)
            with col1: render_styled_table(a1_final, "📊 A1 死亡人數")
            with col2: render_styled_table(a2_final, "📊 A2 受傷人數")

        except Exception as e:
            st.error(f"系統錯誤：{e}")
