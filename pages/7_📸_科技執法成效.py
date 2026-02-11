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
# 1. 頁面配置 (確保網頁不空白)
# ==========================================
st.set_page_config(page_title="科技執法成效統計", layout="wide", page_icon="📸")

# ==========================================
# 2. 自動化設定區 (密碼已埋入)
# ==========================================
MY_EMAIL = "mbmmiqamhnd@gmail.com" 
MY_PASSWORD = "kvpw ymgn xawe qxnl"  
TO_EMAIL = "mbmmiqamhnd@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

st.title("📸 科技執法成效分析系統")
st.markdown("### 📝 狀態：支援 2 個圖表同步、自動建立分頁、路名精簡化。")

# --- 工具函數 1: 數據處理 ---
def parse_hour(val):
    try: return int(str(int(val)).zfill(4)[:2])
    except: return 0

# --- 工具函數 2: 建立 Excel (含圖表) ---
def create_excel_with_charts(df_loc, df_hour):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 路口排行頁面
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

        # 時段分析頁面
        df_hour.to_excel(writer, sheet_name='時段統計', index=False)
        ws_hour = writer.sheets['時段統計']
        chart_hour = workbook.add_chart({'type': 'column'})
        chart_hour.add_series({
            'name': '舉發件數',
            'categories': ['時段統計', 1, 0, 23, 0],
            'values': ['時段統計', 1, 1, 23, 1],
        })
        chart_hour.set_title({'name': '24小時違規時段分析'})
        ws_hour.insert_chart('D2', chart_hour, {'x_scale': 1.5, 'y_scale': 1.5})
    return output

# --- 工具函數 3: 同步 2 個圖表至 Google Sheets (含自動建立分頁功能) ---
def sync_to_gsheet_tech(df_loc, df_hour):
    try:
        if "gcp_service_account" not in st.secrets:
            return False, "❌ Secrets 遺失 GCP 設定"
        
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        
        # 定義要同步的工作表名稱
        sheet_names = ["科技執法-路段排行", "科技執法-時段分析"]
        data_frames = [df_loc, df_hour]
        
        for name, df in zip(sheet_names, data_frames):
            try:
                # 嘗試開啟工作表
                ws = sh.worksheet(name)
            except gspread.exceptions.WorksheetNotFound:
                # 如果找不到，就自動新增一個
                ws = sh.add_worksheet(title=name, rows="100", cols="20")
                st.info(f"ℹ️ 已自動為您建立新分頁：{name}")
            
            # 清除舊數據並寫入新數據
            ws.clear()
            ws.update([df.columns.values.tolist()] + df.values.tolist())
            
        return True, "✅ 2 個圖表數據已成功同步至 Google 試算表"
    except Exception as e:
        return False, f"❌ 同步失敗: {e}"

# --- 主程式流程 ---
uploaded_file = st.file_uploader("請上傳 list2.csv", type=['csv', 'xlsx'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            try: df = pd.read_csv(uploaded_file)
            except: uploaded_file.seek(0); df = pd.read_csv(uploaded_file, encoding='cp950')
        else: df = pd.read_excel(uploaded_file)
        
        df.columns = [str(c).strip() for c in df.columns]
        
        # 1. 精簡路名：刪除 桃園市 與 龍潭區
        if '違規地點' in df.columns:
            df['違規地點'] = df['違規地點'].astype(str).str.replace('桃園市', '', regex=False).str.replace('龍潭區', '', regex=False).str.strip()
        
        # 2. 統計數據
        df['小時'] = df['違規時間'].apply(parse_hour)
        
        # 圖表 1: 路段排行
        loc_summary = df['違規地點'].value_counts().head(10).reset_index()
        loc_summary.columns = ['路段名稱', '舉發件數']
        
        # 圖表 2: 時段分析
        hour_all = pd.DataFrame({'小時': range(24)})
        hour_counts = df['小時'].value_counts().reset_index()
        hour_counts.columns = ['小時', '舉發件數']
        hour_summary = pd.merge(hour_all, hour_counts, on='小時', how='left').fillna(0)
        hour_summary['舉發件數'] = hour_summary['舉發件數'].astype(int)

        # 3. 網頁顯示
        st.divider()
        c1, c2 = st.columns(2)
        with c1: 
            st.subheader("📍 違規路段排行")
            st.dataframe(loc_summary, use_container_width=True)
        with c2: 
            st.subheader("📊 24H 時段分佈")
            st.bar_chart(hour_summary.set_index('小時'))

        # 4. 執行同步與寄信
        if st.button("🚀 執行 2 個圖表同步並寄送 Excel 報表", type="primary"):
            with st.spinner("⚡ 正在處理中，請稍候..."):
                # A. 產生 Excel
                excel_data = create_excel_with_charts(loc_summary, hour_summary)
                
                # B. 同步至 Google Sheets
                gs_success, gs_msg = sync_to_gsheet_tech(loc_summary, hour_summary)
                if gs_success: st.success(gs_msg)
                else: st.error(gs_msg)
                
                # C. 寄送 Email
                try:
                    msg = MIMEMultipart()
                    msg['From'] = MY_EMAIL
                    msg['To'] = TO_EMAIL
                    msg['Subject'] = f"科技執法成效統計 - {datetime.now().strftime('%m/%d')}"
                    
                    body = f"長官好，科技執法 2 項圖表數據已同步至雲端。\n附件 Excel 內含「違規路段排行」與「時段分析」圖表，請查照。\n\n總舉發件數：{len(df)} 件"
                    msg.attach(MIMEText(body, 'plain'))
                    
                    part = MIMEApplication(excel_data.getvalue(), Name="Tech_Enforcement_Report.xlsx")
                    part['Content-Disposition'] = 'attachment; filename="Tech_Enforcement_Report.xlsx"'
                    msg.attach(part)
                    
                    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
                        s.starttls()
                        s.login(MY_EMAIL, MY_PASSWORD)
                        s.send_message(msg)
                    st.success(f"✅ 報表已寄送至：{TO_EMAIL}")
                    st.balloons()
                except Exception as e:
                    st.error(f"❌ 郵件寄送失敗：{e}")

    except Exception as e:
        st.error(f"程式運行出錯：{e}")
else:
    st.info("👋 請上傳 list2.csv 以開始分析。")
