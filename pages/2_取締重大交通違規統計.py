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

# --- 2. 寄信函式 (修正版) ---
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
    except Exception as e:
        st.error(f"郵件寄送失敗: {e}")
        return False

# --- 3. 雲端同步函式 (移除 oauth2client，使用現代化方式) ---
def sync_to_sheets(df):
    try:
        # 直接從 secrets 讀取字典，不需額外匯入 Credentials 套件
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        # 開啟您的試算表名稱 (請確保 Google Sheet 有共用給 Service Account Email)
        sh = gc.open("交通違規統計表")
        ws = sh.get_worksheet(0) # 開啟第一個工作表
        ws.clear()
        ws.update([df.columns.values.tolist()] + df.values.tolist())
        return True
    except Exception as e:
        st.error(f"雲端同步失敗: {e}")
        return False

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

# --- 5. 介面 ---
st.title("🚔 交通統計自動化系統 (v85)")

col_up1, col_up2 = st.columns(2)
with col_up1:
    file_period = st.file_uploader("📂 1. 上傳「本期」檔案 (週報/月報)", type=['xlsx'])
with col_up2:
    file_year = st.file_uploader("📂 2. 上傳「累計」檔案 (含本年、去年數據)", type=['xlsx'])

if file_period and file_year:
    # 執行數據解析
    d_week = parse_excel_with_cols(file_period, "重點違規統計表", [15, 16])
    d_year = parse_excel_with_cols(file_year, "(1)", [15, 16])
    d_last = parse_excel_with_cols(file_year, "(1)", [18, 19])
    
    if d_week and d_year and d_last:
        rows = []
        t = {k: 0 for k in ['ws', 'wc', 'ys', 'yc', 'ls', 'lc', 'diff', 'tgt']}
        for u in UNIT_ORDER:
            w, y, l = d_week.get(u, {'stop':0, 'cit':0}), d_year.get(u, {'stop':0, 'cit':0}), d_last.get(u, {'stop':0, 'cit':0})
            ys, ls = y['stop'] + y['cit'], l['stop'] + l['cit']
            tgt, diff = TARGETS.get(u, 0), ys - ls
            rows.append([u, w['stop'], w['cit'], y['stop'], y['cit'], l['stop'], l['cit'], diff, tgt, f"{(ys/tgt):.1%}" if tgt > 0 else "0%"])
            t['ws']+=w['stop']; t['wc']+=w['cit']; t['ys']+=y['stop']; t['yc']+=y['cit']; t['ls']+=l['stop']; t['lc']+=l['cit']; t['diff']+=diff; t['tgt']+=tgt
        
        total_row = ['合計', t['ws'], t['wc'], t['ys'], t['yc'], t['ls'], t['lc'], t['diff'], t['tgt'], f"{((t['ys']+t['yc'])/t['tgt']):.1%}" if t['tgt']>0 else "0%"]
        rows.insert(0, total_row)
        df_final = pd.DataFrame(rows, columns=['單位', '本期攔停', '本期逕行', '本年攔停', '本年逕行', '去年攔停', '去年逕行', '增減比較', '目標值', '達成率'])
        
        st.success("✅ 解析成功！")
        st.dataframe(df_final, use_container_width=True)

        st.divider()
        if st.button("🚀 同步並寄出報表", type="primary"):
            # 1. 同步雲端
            success_cloud = sync_to_sheets(df_final)
            if success_cloud:
                st.info("☁️ 雲端試算表更新成功！")
            
            # 2. 寄送郵件
            success_mail = send_real_email(df_final)
            if success_mail:
                st.balloons()
                st.info("📧 報表已寄送至 mbmmiqamhnd@gmail.com")
