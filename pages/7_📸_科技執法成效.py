import streamlit as st
import pandas as pd
import io
import smtplib
import gspread
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ==========================================
# 1. 頁面配置 (必須放在最前面，否則網頁會空白)
# ==========================================
st.set_page_config(page_title="科技執法統計", layout="wide", page_icon="📸")

# ==========================================
# 2. 使用者設定區 (密碼已埋入)
# ==========================================
MY_EMAIL = "mbmmiqamhnd@gmail.com" 
MY_PASSWORD = "kvpw ymgn xawe qxnl"  
TO_EMAIL = "mbmmiqamhnd@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

st.title("📸 科技執法成效分析系統")

# --- 工具函數 ---
def parse_hour(val):
    try: return int(str(int(val)).zfill(4)[:2])
    except: return 0

def create_excel_with_charts(df_loc, df_hour):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_loc.to_excel(writer, sheet_name='路口統計', index=False)
        workbook = writer.book
        ws_loc = writer.sheets['路口統計']
        chart_loc = workbook.add_chart({'type': 'bar'})
        chart_loc.add_series({
            'name': '舉發件數',
            'categories': ['路口統計', 1, 0, len(df_loc), 0],
            'values': ['路口統計', 1, 1, len(df_loc), 1],
            'data_labels': {'value': True},
        })
        chart_loc.set_title({'name': '違規路段排行'})
        ws_loc.insert_chart('D2', chart_loc, {'x_scale': 1.5, 'y_scale': 1.5})

        df_hour.to_excel(writer, sheet_name='時段統計', index=False)
        ws_hour = writer.sheets['時段統計']
        chart_hour = workbook.add_chart({'type': 'column'})
        chart_hour.add_series({
            'name': '舉發件數',
            'categories': ['時段統計', 1, 0, 24, 0],
            'values': ['時段統計', 1, 1, 24, 1],
        })
        chart_hour.set_title({'name': '24小時違規時段分析'})
        ws_hour.insert_chart('D2', chart_hour, {'x_scale': 1.5, 'y_scale': 1.5})
    return output

# --- 同步函數 (修正 index 5 問題) ---
def sync_to_gsheet_tech(df_loc, df_hour):
    try:
        if "gcp_service_account" not in st.secrets:
            return False, "❌ Secrets 遺失 GCP 設定"
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        worksheets = sh.worksheets()
        
        # 同步第 5 個分頁
        if len(worksheets) >= 5:
            ws_loc = sh.get_worksheet(4)
            ws_loc.clear()
            ws_loc.update([df_loc.columns.values.tolist()] + df_loc.values.tolist())
        else:
            return False, "❌ 同步失敗：試算表分頁不足（缺少第 5 個分頁）"

        # 同步第 6 個分頁
        if len(worksheets) >= 6:
            ws_hour = sh.get_worksheet(5)
            ws_hour.clear()
            ws_hour.update([df_hour.columns.values.tolist()] + df_hour.values.tolist())
        else:
            st.warning("⚠️ 提醒：試算表沒有第 6 個分頁，時段數據未同步。請在 Google 試算表按「+」新增分頁。")
        
        return True, "✅ Google 試算表同步成功"
    except Exception as e:
        return False, f"❌ 同步失敗: {e}"

# --- 主程式 ---
uploaded_file = st.file_uploader("請上傳 list2.csv", type=['csv', 'xlsx'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            try: df = pd.read_csv(uploaded_file)
            except: uploaded_file.seek(0); df = pd.read_csv(uploaded_file, encoding='cp950')
        else: df = pd.read_excel(uploaded_file)
        
        df.columns = [str(c).strip() for c in df.columns]
        
        # 精簡地名
        if '違規地點' in df.columns:
            df['違規地點'] = df['違規地點'].astype(str).str.replace('桃園市龍潭區', '', regex=False).str.replace('桃園市', '', regex=False)
        
        df['小時'] = df['違規時間'].apply(parse_hour)
        loc_summary = df['違規地點'].value_counts().head(10).reset_index()
        loc_summary.columns = ['路段名稱', '舉發件數']
        
        hour_all = pd.DataFrame({'小時': range(24)})
        hour_counts = df['小時'].value_counts().reset_index()
        hour_counts.columns = ['小時', '舉發件數']
        hour_summary = pd.merge(hour_all, hour_counts, on='小時', how='left').fillna(0)
        hour_summary['舉發件數'] = hour_summary['舉發件數'].astype(int)

        st.divider()
        c1, c2 = st.columns(2)
        with c1: 
            st.subheader("📍 違規路段排行")
            st.dataframe(loc_summary, use_container_width=True)
        with c2: 
            st.subheader("📊 24H 時段分佈")
            st.bar_chart(hour_summary.set_index('小時'))

        if st.button("🚀 執行自動同步並寄送 Excel 圖表", type="primary"):
            with st.spinner("⚡ 系統處理中..."):
                excel_data = create_excel_with_charts(loc_summary, hour_summary)
                gs_success, gs_msg = sync_to_gsheet_tech(loc_summary, hour_summary)
                st.write(gs_msg)
                
                try:
                    msg = MIMEMultipart()
                    msg['From'] = MY_EMAIL
                    msg['To'] = TO_EMAIL
                    msg['Subject'] = f"科技執法統計報告 - {datetime.now().strftime('%m/%d')}"
                    msg.attach(MIMEText(f"長官好，附件為「違規路段排行」統計報表，請查照。\n\n舉發總數：{len(df)} 件", 'plain'))
                    part = MIMEApplication(excel_data.getvalue(), Name="Report.xlsx")
                    part['Content-Disposition'] = 'attachment; filename="Report.xlsx"'
                    msg.attach(part)
                    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
                        s.starttls()
                        s.login(MY_EMAIL, MY_PASSWORD)
                        s.send_message(msg)
                    st.success(f"✅ 報表已寄送至：{TO_EMAIL}")
                    st.balloons()
                except Exception as e:
                    st.error(f"❌ 寄送失敗：{e}")
    except Exception as e:
        st.error(f"程式出錯：{e}")
else:
    st.info("👋 請上傳 list2.csv 檔案。")
