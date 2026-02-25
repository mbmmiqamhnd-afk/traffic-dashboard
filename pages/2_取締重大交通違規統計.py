import streamlit as st
import pandas as pd
import re
import io
import smtplib
import gspread
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# --- 1. 定義識別與目標 ---
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
TARGETS = {'聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479, '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006}

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

# --- 2. 寄信與同步功能 (維持 v85 穩定版) ---
def send_real_email(df):
    try:
        mail_user = st.secrets["email"]["user"]
        mail_pass = st.secrets["email"]["password"]
        receiver = "mbmmiqamhnd@gmail.com"
        msg = MIMEMultipart()
        msg['Subject'] = f"📊 [自動通知] 交通違規統計報表 - {pd.Timestamp.now().strftime('%Y-%m-%d')}"
        msg['From'] = f"交通統計系統 <{mail_user}>"
        msg['To'] = receiver
        html_table = df.to_html(index=False, border=1)
        body = f"<h3>您好，以下為本次交通違規統計數據：</h3>{html_table}"
        msg.attach(MIMEText(body, 'html'))
        excel_buffer = io.BytesIO()
        df.to_excel(excel_buffer, index=False)
        part = MIMEApplication(excel_buffer.getvalue(), Name="Traffic_Stats.xlsx")
        part['Content-Disposition'] = 'attachment; filename="Traffic_Stats.xlsx"'
        msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(mail_user, mail_pass)
            server.send_message(msg)
        return True
    except: return False

def sync_to_sheets(df):
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open("交通違規統計表")
        ws = sh.get_worksheet(0)
        ws.clear()
        ws.update([df.columns.values.tolist()] + df.values.tolist())
        return True
    except: return False

# --- 3. 解析邏輯 ---
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

# --- 4. 主介面 ---
st.title("🚔 交通統計自動化系統 (v86)")

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
            
            # --- 警備隊特殊處理邏輯 ---
            if u == '警備隊':
                diff_display = "—"
                rate_display = "—"
            else:
                diff_val = ys_sum - ls_sum
                diff_display = int(diff_val)
                rate_display = f"{(ys_sum/tgt):.1%}" if tgt > 0 else "0%"
                # 僅非警備隊的數據才計入合計的比較值與目標值 (若警備隊目標為0則不影響)
                t['diff'] += (ys_sum - ls_sum)
                t['tgt'] += tgt
            
            rows.append([u, w['stop'], w['cit'], y['stop'], y['cit'], l['stop'], l['cit'], diff_display, tgt, rate_display])
            
            # 基礎數值不論是否為警備隊都計入合計
            t['ws']+=w['stop']; t['wc']+=w['cit']; t['ys']+=y['stop']; t['yc']+=y['cit']; t['ls']+=l['stop']; t['lc']+=l['cit']
        
        # 合計列
        total_rate = f"{((t['ys']+t['yc'])/t['tgt']):.1%}" if t['tgt']>0 else "0%"
        total_row = ['合計', t['ws'], t['wc'], t['ys'], t['yc'], t['ls'], t['lc'], t['diff'], t['tgt'], total_rate]
        rows.insert(0, total_row)
        
        df_final = pd.DataFrame(rows, columns=['單位', '本期攔停', '本期逕行', '本年攔停', '本年逕行', '去年攔停', '去年逕行', '增減比較', '目標值', '達成率'])
        st.success("✅ 解析成功！(警備隊已排除比較)")
        st.dataframe(df_final, use_container_width=True)

        st.divider()
        if st.button("🚀 同步並寄出報表", type="primary"):
            if sync_to_sheets(df_final): st.info("☁️ 雲端試算表更新成功！")
            if send_real_email(df_final):
                st.balloons()
                st.info("📧 報表已寄送至 mbmmiqamhnd@gmail.com")
