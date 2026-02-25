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
# 🔐 【安全設定區 - 已改為 Secrets 模式】 
# ==========================================
# 說明：程式會從 .streamlit/secrets.toml 或 Streamlit Cloud 後台抓取密碼
try:
    MY_EMAIL = st.secrets["email"]["user"]
    MY_PASSWORD = st.secrets["email"]["password"]
    GCP_CREDS = st.secrets["gcp_service_account"]
except Exception as e:
    st.error("❌ 找不到 Secrets 設定！請在設定區配置 [email] 與 [gcp_service_account]。")
    st.stop()

TO_EMAIL = MY_EMAIL # 預設寄給自己，可視需求更改
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

# ==========================================

st.set_page_config(page_title="交通事故統計系統", layout="wide", page_icon="🚑")
st.title("🚑 交通事故統計 (自動寄信 + 格式保留版)")
st.markdown("### 📝 狀態：使用安全 Secrets 機制，支援日期無斜線格式 (如 0101-0107)。")

# --- 工具函數 1: HTML 顯示紅字數字 ---
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

# --- 工具函數 2: Google Sheets API Rich Text (紅黑字) ---
def get_gsheet_rich_text_req(sheet_id, row_idx, col_idx, text):
    text = str(text)
    tokens = re.split(r'([0-9\(\)\/\-\.\%]+)', text)
    runs = []
    current_pos = 0
    for token in tokens:
        if not token: continue
        color = {"red": 1, "green": 0, "blue": 0} if re.match(r'^[0-9\(\)\/\-\.\%]+$', token) else {"red": 0, "green": 0, "blue": 0}
        runs.append({
            "startIndex": current_pos,
            "format": {"foregroundColor": color, "bold": True}
        })
        current_pos += len(token)
    
    return {
        "updateCells": {
            "rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}],
            "fields": "userEnteredValue,textFormatRuns",
            "range": {"sheetId": sheet_id, "startRowIndex": row_idx, "endRowIndex": row_idx + 1, "startColumnIndex": col_idx, "endColumnIndex": col_idx + 1}
        }
    }

# --- 網頁表格顯示 ---
def render_styled_table(df, title):
    st.subheader(title)
    style = """
    <style>
        table.acc_table {font-family: sans-serif; border-collapse: collapse; width: 100%; font-size: 14px;}
        table.acc_table th {border: 1px solid #000; padding: 8px; text-align: center; background-color: #f0f2f6;}
        table.acc_table td {border: 1px solid #000; padding: 8px; text-align: center; background-color: #ffffff;}
    </style>
    """
    html = f"{style}<table class='acc_table'><thead><tr>"
    for col in df.columns: html += f"<th>{format_html_header(col)}</th>"
    html += "</tr></thead><tbody>"
    for _, row in df.iterrows():
        html += "<tr>"
        for col_name, val in row.items():
            color = "red" if ("比較" in col_name or "增減" in col_name) and str(val) != "0.00%" and "-" not in str(val) and str(val) != "0" else "black"
            html += f'<td style="color: {color};">{val}</td>'
        html += "</tr>"
    html += "</tbody></table>"
    st.markdown(html, unsafe_allow_html=True)

# --- 寄信函數 ---
def send_email_auto(attachment_data, filename):
    try:
        msg = MIMEMultipart()
        msg['From'], msg['To'] = MY_EMAIL, TO_EMAIL
        msg['Subject'] = f"交通事故統計報表 ({pd.Timestamp.now().strftime('%Y/%m/%d')})"
        msg.attach(MIMEText("長官好，數據已同步至 Google 試算表，附件為本次統計 Excel。", 'plain'))
        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(attachment_data.getvalue()); encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={filename}')
        msg.attach(part)
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
            s.starttls(); s.login(MY_EMAIL, MY_PASSWORD); s.send_message(msg)
        return True, f"✅ 報表已自動寄送至：{TO_EMAIL}"
    except Exception as e: return False, f"❌ 寄送失敗：{e}"

# --- Google Sheets 同步 ---
def sync_to_gsheet(df_a1, df_a2):
    try:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        
        def update_sheet_values_only(ws_index, df, title_text):
            ws = sh.get_worksheet(ws_index)
            ws.batch_clear(["A3:Z100"])
            ws.update_acell('A1', title_text)
            data_rows = [[int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in row] for row in df.values.tolist()]
            if data_rows: ws.update('A3', data_rows)
            reqs = [get_gsheet_rich_text_req(ws.id, 1, col_idx, col_name) for col_idx, col_name in enumerate(df.columns)]
            if reqs: sh.batch_update({"requests": reqs})
            return True

        update_sheet_values_only(2, df_a1, "A1類交通事故死亡人數統計表")
        update_sheet_values_only(3, df_a2, "A2類交通事故受傷人數統計表")
        return True, "✅ Google 試算表同步成功 (格式保留)"
    except Exception as e: return False, f"❌ Google 試算表失敗: {e}"

# --- 主流程 ---
uploaded_files = st.file_uploader("請上傳 3 個報表檔案", accept_multiple_files=True)

if uploaded_files and len(uploaded_files) == 3:
    with st.spinner("⚡ 處理中..."):
        try:
            # A. 讀取與分配檔案 (邏輯簡化版)
            def parse_raw(f):
                try: return pd.read_csv(f, header=None)
                except: f.seek(0); return pd.read_excel(f, header=None)
            
            def clean_data(df_raw):
                df_raw[0] = df_raw[0].astype(str)
                df_data = df_raw[df_raw[0].str.contains("所|總計|合計", na=False)].copy()
                cols = {0: "Station", 5: "A1_Deaths", 9: "A2_Injuries"}
                df_data = df_data.rename(columns=cols)
                for c in [5, 9]: df_data[cols[c]] = pd.to_numeric(df_data[cols[c]].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
                df_data['Station_Short'] = df_data['Station'].str.replace('派出所', '所').str.replace('總計', '合計')
                return df_data

            files_meta = []
            for f in uploaded_files:
                f.seek(0); df = parse_raw(f)
                dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', str(df.iloc[:5, :3].values))
                if len(dates) >= 2:
                    d_str = f"{int(dates[0][1]):02d}{int(dates[0][2]):02d}-{int(dates[1][1]):02d}{int(dates[1][2]):02d}"
                    files_meta.append({'df': clean_data(df), 'year': int(dates[1][0]), 'date_range': d_str, 'is_cumu': (int(dates[0][1]) == 1)})
            
            # 分配本期、本年累計、去年累計
            files_meta.sort(key=lambda x: x['year'], reverse=True)
            cur_year = files_meta[0]['year']
            df_wk = [f for f in files_meta if f['year'] == cur_year and not f['is_cumu']][0]
            df_cur = [f for f in files_meta if f['year'] == cur_year and f['is_cumu']][0]
            df_lst = [f for f in files_meta if f['year'] < cur_year][0]

            # B. 計算合併 (A1 & A2)
            stations = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
            def build_final(col_name):
                m = pd.merge(df_wk['df'][['Station_Short', col_name]], df_cur['df'][['Station_Short', col_name]], on='Station_Short', suffixes=('_wk', '_cur'))
                m = pd.merge(m, df_lst['df'][['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_lst'})
                m = m[m['Station_Short'].isin(stations)].copy()
                total_row = m.select_dtypes(include='number').sum().to_dict()
                total_row['Station_Short'] = '合計'
                m = pd.concat([pd.DataFrame([total_row]), m], ignore_index=True)
                m['Diff'] = m[col_name+'_cur'] - m[col_name+'_lst']
                return m

            a1_res = build_final('A1_Deaths')
            a1_res.columns = ['統計期間', f'本期({df_wk["date_range"]})', f'本年累計({df_cur["date_range"]})', f'去年累計({df_lst["date_range"]})', '比較']
            
            a2_res = build_final('A2_Injuries')
            a2_res['Pct'] = a2_res.apply(lambda x: f"{(x['Diff']/x['A2_Injuries_lst']):.2%}" if x['A2_Injuries_lst']!=0 else "-", axis=1)
            a2_res.insert(2, 'Prev', '-')
            a2_res.columns = ['統計期間', f'本期({df_wk["date_range"]})', '前期', f'本年累計({df_cur["date_range"]})', f'去年累計({df_lst["date_range"]})', '比較', '增減比例']

            # C. 產製 Excel & 同步 & 寄信
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                a1_res.to_excel(writer, index=False, sheet_name='A1死亡人數')
                a2_res.to_excel(writer, index=False, sheet_name='A2受傷人數')
            
            gs_s, gs_m = sync_to_gsheet(a1_res, a2_res)
            em_s, em_m = send_email_auto(output, "Traffic_Stats.xlsx")
            
            if gs_s and em_s: 
                st.success(f"{gs_m} / {em_m}"); st.balloons()
                col1, col2 = st.columns(2)
                with col1: render_styled_table(a1_res, "📊 A1 死亡人數")
                with col2: render_styled_table(a2_res, "📊 A2 受傷人數")

        except Exception as e: st.error(f"分析失敗：{e}")
