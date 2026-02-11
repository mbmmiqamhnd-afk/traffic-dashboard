import streamlit as st
import pandas as pd
import io
import smtplib
import re
import gspread
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ==========================================
# 👇👇👇 【使用者設定區：密碼與參數埋入】 👇👇👇
# ==========================================
MY_EMAIL = "mbmmiqamhnd@gmail.com" 
MY_PASSWORD = "kvpw ymgn xawe qxnl"  # 您的 Gmail 應用程式密碼
TO_EMAIL = "mbmmiqamhnd@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
# Google Sheet URL (沿用您提供的試算表)
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
# ==========================================

st.set_page_config(page_title="科技執法成效 (試算表圖表版)", layout="wide", page_icon="📸")

st.title("📸 科技執法成效 (Excel 圖表報表版)")
st.markdown("### 📝 狀態：支援 Excel 內嵌圖表、Google 試算表同步、全自動寄信。")

# --- 工具函數 1: 日期與時間清理 ---
def parse_roc_date(val):
    try:
        s = str(int(val)).zfill(7)
        return f"{int(s[:-4])}/{s[-4:-2]}/{s[-2:]}"
    except: return "N/A"

def parse_hour(val):
    try: return int(str(int(val)).zfill(4)[:2])
    except: return 0

# --- 工具函數 2: 建立含圖表的 Excel ---
def create_excel_with_charts(df_loc, df_hour):
    output = io.BytesIO()
    # 使用 xlsxwriter 作為引擎來產生成報表與圖表
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 1. 寫入路口數據
        df_loc.to_excel(writer, sheet_name='路口統計', index=False)
        workbook = writer.book
        ws_loc = writer.sheets['路口統計']
        
        # 建立路口長條圖 (橫向，確保名字看得見)
        chart_loc = workbook.add_chart({'type': 'bar'})
        chart_loc.add_series({
            'name':       '舉發件數',
            'categories': ['路口統計', 1, 0, len(df_loc), 0],
            'values':     ['路口統計', 1, 1, len(df_loc), 1],
        })
        chart_loc.set_title({'name': '十大違規路段排行'})
        chart_loc.set_x_axis({'name': '件數'})
        ws_loc.insert_chart('D2', chart_loc, {'x_scale': 1.5, 'y_scale': 1.5})

        # 2. 寫入時段數據
        df_hour.to_excel(writer, sheet_name='時段統計', index=False)
        ws_hour = writer.sheets['時段統計']
        
        # 建立時段直條圖
        chart_hour = workbook.add_chart({'type': 'column'})
        chart_hour.add_series({
            'name':       '舉發件數',
            'categories': ['時段統計', 1, 0, 24, 0],
            'values':     ['時段統計', 1, 1, 24, 1],
        })
        chart_hour.set_title({'name': '24小時違規時段分析'})
        ws_hour.insert_chart('D2', chart_hour, {'x_scale': 1.5, 'y_scale': 1.5})
        
    return output

# --- 工具函數 3: 同步至 Google Sheets (比照交通事故統計) ---
def sync_to_gsheet_tech(df_loc, df_hour):
    try:
        if "gcp_service_account" not in st.secrets:
            return False, "❌ Secrets 中找不到 [gcp_service_account]"
        
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        
        # 假設將科技執法數據同步至第 5, 6 分頁 (請視需要調整 index)
        ws_loc = sh.get_worksheet(4) # 第 5 張工作表
        ws_loc.clear()
        ws_loc.update([df_loc.columns.values.tolist()] + df_loc.values.tolist())
        
        ws_hour = sh.get_worksheet(5) # 第 6 張工作表
        ws_hour.clear()
        ws_hour.update([df_hour.columns.values.tolist()] + df_hour.values.tolist())
        
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
        else:
            df = pd.read_excel(uploaded_file)
        
        df.columns = [str(c).strip() for c in df.columns]
        
        # 數據摘要
        df['小時'] = df['違規時間'].apply(parse_hour)
        
        # 產生統計表
        loc_summary = df['違規地點'].value_counts().head(10).reset_index()
        loc_summary.columns = ['違規地點', '舉發件數']
        
        hour_all = pd.DataFrame({'小時': range(24)})
        hour_counts = df['小時'].value_counts().reset_index()
        hour_counts.columns = ['小時', '舉發件數']
        hour_summary = pd.merge(hour_all, hour_counts, on='小時', how='left').fillna(0)
        hour_summary['舉發件數'] = hour_summary['舉發件數'].astype(int)

        # 網頁顯示
        st.divider()
        st.subheader("📊 執法成效數據摘要")
        c1, c2 = st.columns(2)
        with c1: st.dataframe(loc_summary, use_container_width=True)
        with c2: st.bar_chart(df['違規地點'].value_counts().head(10))

        # --- 一鍵寄信與同步按鈕 ---
        st.divider()
        if st.button("🚀 自動同步雲端並寄送試算表圖表報表", type="primary"):
            with st.spinner("⚡ 正在處理中..."):
                # 1. 產生 Excel (內含圖表)
                excel_data = create_excel_with_charts(loc_summary, hour_summary)
                
                # 2. 同步 Google Sheet
                gs_success, gs_msg = sync_to_gsheet_tech(loc_summary, hour_summary)
                st.write(gs_msg)
                
                # 3. 寄送 Email (埋入帳密)
                try:
                    msg = MIMEMultipart()
                    msg['From'] = MY_EMAIL
                    msg['To'] = TO_EMAIL
                    msg['Subject'] = f"科技執法成效報表 (含內嵌圖表) - {datetime.now().strftime('%m/%d')}"
                    msg.attach(MIMEText(f"長官好，\n\n檢送科技執法成效報表，Excel 內已附帶統計圖表，路口名稱已優化顯示，請查照。\n\n舉發總數：{len(df)} 件", 'plain'))
                    
                    part = MIMEApplication(excel_data.getvalue(), Name="Tech_Enforcement_Report.xlsx")
                    part['Content-Disposition'] = 'attachment; filename="Tech_Enforcement_Report.xlsx"'
                    msg.attach(part)
                    
                    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
                        s.starttls()
                        s.login(MY_EMAIL, MY_PASSWORD)
                        s.send_message(msg)
                    st.success(f"✅ 報表已成功寄送至：{TO_EMAIL}")
                    st.balloons()
                except Exception as e:
                    st.error(f"❌ 寄送失敗：{e}")

    except Exception as e:
        st.error(f"系統錯誤：{e}")
