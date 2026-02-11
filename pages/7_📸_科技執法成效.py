import streamlit as st
import pandas as pd
import io
import smtplib
import gspread
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# 1. 頁面配置
st.set_page_config(page_title="科技執法統計 - 入案時間基準", layout="wide", page_icon="📸")

# 2. 自動化設定
MY_EMAIL = "mbmmiqamhnd@gmail.com" 
MY_PASSWORD = "kvpw ymgn xawe qxnl"  
TO_EMAIL = "mbmmiqamhnd@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

st.title("📸 科技執法成效分析系統 (入案時間版)")
st.markdown("### 📝 狀態：統計期間已修正為以「入案日期」為基準，並維持範本格式產製。")

# --- 工具函數 ---
def parse_hour(val):
    try: return int(str(int(val)).zfill(4)[:2])
    except: return 0

def format_roc_date_range_by_entry(df):
    """從數據中擷取『入案日期』之最小與最大日期並轉換為民國格式"""
    # 優先找『入案日期』，若無則找『入案時間』，再無則退回『違規日期』
    target_col = None
    for col in ['入案日期', '入案時間', '違規日期']:
        if col in df.columns:
            target_col = col
            break
    
    if not target_col:
        return "期間未定"

    try:
        # 確保為數字並過濾掉空值
        valid_dates = pd.to_numeric(df[target_col], errors='coerce').dropna().astype(int)
        if valid_dates.empty:
            return "無有效日期"
            
        start_val = valid_dates.min()
        end_val = valid_dates.max()
        
        def to_roc_str(val):
            s = str(val).zfill(7)
            # 處理 1140101 格式
            year = int(s[:-4])
            month = int(s[-4:-2])
            day = int(s[-2:])
            return f"{year}年{month}月{day}日"
            
        return f"{to_roc_str(start_val)}至{to_roc_str(end_val)}"
    except:
        return "日期格式解析錯誤"

# --- 核心：建立依照範本格式的 Excel ---
def create_formatted_excel(df_loc, df_hour, date_range_text, total_count):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        ws = workbook.add_worksheet('科技執法成效統計')
        
        # 格式設定
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14})
        header_fmt = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1, 'bold': True, 'align': 'center'})
        data_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        total_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFFFCC'})

        # 寫入抬頭
        ws.write('A1', '科技執法成效', title_fmt)
        ws.write('A2', '統計期間')
        ws.write('B2', date_range_text) # 這裡是入案時間的範圍
        
        ws.write('A3', '路口名稱', header_fmt)
        ws.write('B3', '舉發件數', header_fmt)
        
        # 寫入數據
        row_idx = 3
        for _, row in df_loc.iterrows():
            ws.write(row_idx, 0, row['路段名稱'], data_fmt)
            ws.write(row_idx, 1, row['舉發件數'], data_fmt)
            row_idx += 1
        
        # 寫入總計
        ws.write(row_idx, 0, '舉發總數', total_fmt)
        ws.write(row_idx, 1, total_count, total_fmt)
        
        # 插入圖表
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            'name': '舉發件數',
            'categories': ['科技執法成效統計', 3, 0, row_idx - 1, 0],
            'values':     ['科技執法成效統計', 3, 1, row_idx - 1, 1],
            'data_labels': {'value': True},
        })
        chart.set_title({'name': '違規路段排行'})
        ws.insert_chart('D2', chart, {'x_scale': 1.5, 'y_scale': 1.5})

        # 時段分析頁
        df_hour.to_excel(writer, sheet_name='時段分析', index=False)
        
    return output

# --- 同步 Google Sheets ---
def sync_to_gsheet_tech(df_loc, df_hour):
    try:
        if "gcp_service_account" not in st.secrets: return False, "❌ Secrets 遺失"
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        for name, df in zip(["科技執法-路段排行", "科技執法-時段分析"], [df_loc, df_hour]):
            try: ws = sh.worksheet(name)
            except: ws = sh.add_worksheet(title=name, rows="100", cols="20")
            ws.clear()
            ws.update([df.columns.values.tolist()] + df.values.tolist())
        return True, "✅ 數據同步成功"
    except Exception as e: return False, f"❌ 同步失敗: {e}"

# --- 主流程 ---
uploaded_file = st.file_uploader("請上傳清冊檔案 (list2.csv)", type=['csv', 'xlsx'])

if uploaded_file:
    try:
        # 讀檔
        if uploaded_file.name.endswith('.csv'):
            try: df = pd.read_csv(uploaded_file)
            except: uploaded_file.seek(0); df = pd.read_csv(uploaded_file, encoding='cp950')
        else: df = pd.read_excel(uploaded_file)
        
        df.columns = [str(c).strip() for c in df.columns]
        
        # 1. 取得統計期間 (改為入案時間基準)
        date_range_str = format_roc_date_range_by_entry(df)
        total_sum = len(df)
        
        # 2. 地名優化 (刪除 桃園市、龍潭區)
        if '違規地點' in df.columns:
            df['違規地點'] = df['違規地點'].astype(str).str.replace('桃園市', '', regex=False).str.replace('龍潭區', '', regex=False).str.strip()
        
        # 3. 數據統計
        df['小時'] = df['違規時間'].apply(parse_hour)
        loc_summary = df['違規地點'].value_counts().head(10).reset_index()
        loc_summary.columns = ['路段名稱', '舉發件數']
        
        hour_all = pd.DataFrame({'小時': range(24)})
        hour_counts = df['小時'].value_counts().reset_index()
        hour_counts.columns = ['小時', '舉發件數']
        hour_summary = pd.merge(hour_all, hour_counts, on='小時', how='left').fillna(0)
        hour_summary['舉發件數'] = hour_summary['舉發件數'].astype(int)

        # 4. 網頁呈現
        st.divider()
        st.subheader(f"📅 統計期間 (入案時間)：{date_range_str}")
        c1, c2 = st.columns(2)
        with c1: st.dataframe(loc_summary, use_container_width=True)
        with c2: st.bar_chart(hour_summary.set_index('小時'))

        # 5. 執行按鈕
        if st.button("🚀 產製入案基準報表並同步", type="primary"):
            with st.spinner("⚡ 系統處理中..."):
                excel_data = create_formatted_excel(loc_summary, hour_summary, date_range_str, total_sum)
                gs_success, gs_msg = sync_to_gsheet_tech(loc_summary, hour_summary)
                st.write(gs_msg)
                
                try:
                    msg = MIMEMultipart()
                    msg['From'] = MY_EMAIL
                    msg['To'] = TO_EMAIL
                    msg['Subject'] = f"科技執法入案統計({date_range_str})"
                    body = f"長官好，科技執法統計已產製。\n統計基準：入案日期\n統計期間：{date_range_str}\n舉發件數：{total_sum} 件"
                    msg.attach(MIMEText(body, 'plain'))
                    part = MIMEApplication(excel_data.getvalue(), Name="Tech_Report_EntryDate.xlsx")
                    part.add_header('Content-Disposition', 'attachment', filename="Tech_Report_EntryDate.xlsx")
                    msg.attach(part)
                    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
                        s.starttls()
                        s.login(MY_EMAIL, MY_PASSWORD)
                        s.send_message(msg)
                    st.success(f"✅ 報表已寄送至：{TO_EMAIL}")
                    st.balloons()
                except Exception as e: st.error(f"❌ 寄送失敗：{e}")
    except Exception as e: st.error(f"系統錯誤：{e}")
