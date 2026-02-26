import streamlit as st
import pandas as pd
import re
import io
import smtplib
import gspread
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
TARGETS = {
    '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479, 
    '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006
}
FOOTNOTE_TEXT = "重大交通違規指：「酒駕」、「闖紅燈」、「嚴重超速」、「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」及「不暫停讓行人」"

def get_standard_unit(raw_name):
    name = str(raw_name).strip()
    if '分隊' in name: return '交通分隊'
    if '科技' in name or '交通組' in name: return '科技執法'
    if '警備' in name: return '警備隊'
    if '聖亭' in name: return '聖亭所'
    if '龍潭' in name: return '龍潭所'
    if '中興' in name: return '中興所'
    if '石門' in name: return '石門所'
    if '高平' in name: return '高平所'
    if '三和' in name: return '三和所'
    return None

# --- 2. 雲端同步功能 (僅更新數據並設定負值紅字) ---
def sync_to_specified_sheet(df):
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        ws = sh.get_worksheet(0)
        
        # 準備數據
        col_tuples = df.columns.tolist()
        top_row = [t[0] for t in col_tuples]
        bottom_row = [t[1] for t in col_tuples]
        data_list = [top_row, bottom_row] + df.values.tolist()
        
        # 僅更新數據，保留原有格式
        ws.update(range_name='A1', values=data_list)
        
        # 設定「本年與去年同期比較」欄位（H 欄，Index 為 7）的負值紅字規則
        # 範圍從第 3 列（Index 2）開始到備註列之前
        data_rows_count = len(data_list) - 1 # 排除最後一列備註
        
        requests = [
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [
                            {
                                "sheetId": ws.id,
                                "startRowIndex": 2,      # 從合計列開始
                                "endRowIndex": data_rows_count, 
                                "startColumnIndex": 7,   # H 欄
                                "endColumnIndex": 8
                            }
                        ],
                        "booleanRule": {
                            "condition": {
                                "type": "NUMBER_LESS",
                                "values": [{"userEnteredValue": "0"}]
                            },
                            "format": {
                                "textFormat": {"foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}}
                            }
                        }
                    },
                    "index": 0
                }
            }
        ]
        
        # 執行條件格式更新（這不會影響你原本手動設定的合併或框線）
        sh.batch_update({"requests": requests})
        
        return True
    except Exception as e:
        st.error(f"雲端同步失敗: {e}")
        return False

# --- 3. 寄信功能 ---
def send_stats_email(df):
    try:
        mail_user = st.secrets["email"]["user"]
        mail_pass = st.secrets["email"]["password"]
        receiver = "mbmmiqamhnd@gmail.com"
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [自動通知] 交通違規統計報表 - {pd.Timestamp.now().strftime('%Y-%m-%d')}"
        msg['From'] = f"交通統計系統 <{mail_user}>"
        msg['To'] = receiver
        
        # 網頁與郵件表格也加入負值紅字樣式
        def color_negative_red(val):
            try:
                color = 'red' if float(val) < 0 else 'black'
            except:
                color = 'black'
            return f'color: {color}'
        
        # 僅針對比較欄位套用樣式
        html_table = df.style.applymap(color_negative_red, subset=df.columns[7:8]).to_html(border=1)
        
        body = f"<h3>您好，以下為本次交通違規統計數據：</h3>{html_table}"
        msg.attach(MIMEText(body, 'html'))
        
        excel_buffer = io.BytesIO()
        df.to_excel(excel_buffer)
        part = MIMEApplication(excel_buffer.getvalue(), Name="Traffic_Stats.xlsx")
        part['Content-Disposition'] = 'attachment; filename="Traffic_Stats.xlsx"'
        msg.attach(part)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(mail_user, mail_pass)
            server.send_message(msg)
        return True
    except: return False

# --- 4. 解析邏輯 ---
def parse_excel_with_cols(uploaded_file, sheet_keyword, col_indices):
    try:
        content = uploaded_file.getvalue()
        xl = pd.ExcelFile(io.BytesIO(content))
        target_sheet = next((s for s in xl.sheet_names if sheet_keyword in s), xl.sheet_names[0])
        df = pd.read_excel(xl, sheet_name=target_sheet, header=None)
        unit_data = {}
        for _, row in df.iterrows():
            u = get_standard_unit(row.iloc[0])
            if u and "合計" not in str(row.iloc[0]):
                def clean(v):
                    try: return int(float(str(v).replace(',', '').strip())) if str(v).strip() not in ['', 'nan', 'None', '-'] else 0
                    except: return 0
                stop_val = 0 if u == '科技執法' else clean(row.iloc[col_indices[0]])
                cit_val = clean(row.iloc[col_indices[1]])
                if u not in unit_data: unit_data[u] = {'stop': stop_val, 'cit': cit_val}
                else: 
                    unit_data[u]['stop'] += stop_val
                    unit_data[u]['cit'] += cit_val
        return unit_data
    except: return None

# --- 5. 主介面 ---
st.title("🚔 交通統計自動化系統")

col_up1, col_up2 = st.columns(2)
with col_up1:
    file_period = st.file_uploader("📂 1. 上傳「本期」檔案", type=['xlsx'])
with col_up2:
    file_year = st.file_uploader("📂 2. 上傳「累計」檔案", type=['xlsx'])

if file_period and file_year:
    d_week = parse_excel_with_cols(file_period, "重點違規統計表", [15, 16])
    d_year = parse_excel_with_cols(file_year, "(1)", [15, 16])
    d_last = parse_excel_with_cols(file_year, "(1)", [18, 19])
    
    if d_week and d_year and d_last:
        rows = []
        t = {k: 0 for k in ['ws', 'wc', 'ys', 'yc', 'ls', 'lc', 'diff', 'tgt']}
        for u in UNIT_ORDER:
            w, y, l = d_week.get(u, {'stop':0, 'cit':0}), d_year.get(u, {'stop':0, 'cit':0}), d_last.get(u, {'stop':0, 'cit':0})
            ys_sum, ls_sum = y['stop'] + y['cit'], l['stop'] + l['cit']
            tgt = TARGETS.get(u, 0)
            
            if u == '警備隊':
                diff_display, rate_display = "—", "—"
            else:
                diff_val = ys_sum - ls_sum
                diff_display = int(diff_val)
                rate_display = f"{(ys_sum/tgt):.1%}" if tgt > 0 else "0%"
                t['diff'] += diff_val
                t['tgt'] += tgt
            
            rows.append([u, w['stop'], w['cit'], y['stop'], y['cit'], l['stop'], l['cit'], diff_display, tgt, rate_display])
            t['ws']+=w['stop']; t['wc']+=w['cit']; t['ys']+=y['stop']; t['yc']+=y['cit']; t['ls']+=l['stop']; t['lc']+=l['cit']
        
        # 合計列置頂
        total_rate = f"{((t['ys']+t['yc'])/t['tgt']):.1%}" if t['tgt']>0 else "0%"
        total_row = ['合計', t['ws'], t['wc'], t['ys'], t['yc'], t['ls'], t['lc'], t['diff'], t['tgt'], total_rate]
        rows.insert(0, total_row)
        
        # 新增註解列
        rows.append([FOOTNOTE_TEXT] + [""] * 9)
        
        # 多層標題
        header_top = ['統計期間', '本期', '本期', '本年累計', '本年累計', '去年累計', '去年累計', '本年與去年同期比較', '目標值', '達成率']
        header_bottom = ['取締方式', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '', '', '']
        
        multi_col = pd.MultiIndex.from_arrays([header_top, header_bottom])
        df_final = pd.DataFrame(rows, columns=multi_col)
        
        st.success("✅ 解析成功！")
        
        # 網頁預覽也加上紅字
        def style_negative(v):
            return 'color: red' if isinstance(v, (int, float)) and v < 0 else None
        
        st.dataframe(df_final.style.applymap(style_negative, subset=df_final.columns[7:8]), use_container_width=True)

        st.divider()
        if st.button("🚀 同步數據並寄出報表", type="primary"):
            if sync_to_specified_sheet(df_final): 
                st.info(f"☁️ 數據已同步！比較欄位的負數已自動設為紅字。")
            
            if send_stats_email(df_final):
                st.balloons()
                st.info("📧 報表已寄送。")
