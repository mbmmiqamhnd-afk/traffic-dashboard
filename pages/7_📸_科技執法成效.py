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
st.set_page_config(page_title="科技執法統計 - 1/1起算版", layout="wide", page_icon="📸")

# 2. 自動化設定
MY_EMAIL = "mbmmiqamhnd@gmail.com" 
MY_PASSWORD = "kvpw ymgn xawe qxnl"  
TO_EMAIL = "mbmmiqamhnd@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

st.title("📸 科技執法成效分析系統")
st.markdown("### 📝 狀態：統計期間已設定為「1月1日」起至「最新入案日」。")

# --- 工具函數 ---
def parse_hour(val):
    try:
        # 處理數字或字串格式的時間 (如 143005 或 14)
        s = str(int(float(val))).zfill(4)
        return int(s[:2])
    except: return 0

def get_col_name(df, possible_names):
    """靈活尋找欄位名稱，忽略空格"""
    for name in possible_names:
        # 檢查是否存在於原始欄位或去除空格後的欄位
        clean_cols = [str(c).strip() for c in df.columns]
        if name in clean_cols:
            return df.columns[clean_cols.index(name)]
    return None

def format_roc_date_range_fixed_start(df):
    """統計期間：1月1日起至入案日最後一天"""
    target_col = get_col_name(df, ['入案日期', '入案時間', '日期', '違規日期'])
    if not target_col: return "期間未定"
    try:
        # 轉換為數字並過濾
        valid_dates = pd.to_numeric(df[target_col], errors='coerce').dropna().astype(int)
        if valid_dates.empty: return "無有效日期"
        
        # 取得資料中最後一天的日期
        max_date_val = valid_dates.max()
        s_max = str(max_date_val).zfill(7)
        
        year = int(s_max[:-4]) # 取得民國年
        month = int(s_max[-4:-2])
        day = int(s_max[-2:])
        
        # 固定從 1 月 1 日開始
        start_text = f"{year}年1月1日"
        end_text = f"{year}年{month}月{day}日"
        
        return f"{start_text}至{end_text}"
    except: return "日期解析錯誤"

# --- 核心：建立 Excel (範本格式) ---
def create_formatted_excel(df_loc, df_hour, date_range_text, total_count):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        ws = workbook.add_worksheet('科技執法成效統計')
        
        # 格式定義
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'left'})
        header_fmt = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1, 'bold': True, 'align': 'center'})
        data_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        total_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFFFCC', 'align': 'center'})

        # 1. 寫入抬頭與期間
        ws.write('A1', '科技執法成效', title_fmt)
        ws.write('A2', '統計期間', workbook.add_format({'align': 'center', 'border': 1}))
        ws.write('B2', date_range_text, workbook.add_format({'border': 1}))
        
        # 2. 寫入欄位標題
        ws.write('A3', '路口名稱', header_fmt)
        ws.write('B3', '舉發件數', header_fmt)
        
        # 3. 寫入十大路段數據
        curr_row = 3
        for _, row in df_loc.iterrows():
            ws.write(curr_row, 0, row['路段名稱'], data_fmt)
            ws.write(curr_row, 1, row['舉發件數'], data_fmt)
            curr_row += 1
        
        # 4. 寫入總計
        ws.write(curr_row, 0, '舉發總數', total_fmt)
        ws.write(curr_row, 1, total_count, total_fmt)
        
        # 5. 插入 Excel 圖表
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            'name': '舉發件數',
            'categories': ['科技執法成效統計', 3, 0, curr_row - 1, 0],
            'values':     ['科技執法成效統計', 3, 1, curr_row - 1, 1],
            'data_labels': {'value': True},
        })
        chart.set_title({'name': '違規路段排行'})
        ws.insert_chart('D2', chart, {'x_scale': 1.5, 'y_scale': 1.5})

        # 6. 時段分析頁
        df_hour.to_excel(writer, sheet_name='時段分析', index=False)
        
    return output

# --- 主流程 ---
uploaded_file = st.file_uploader("請上傳 list2.csv", type=['csv', 'xlsx'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            try: df = pd.read_csv(uploaded_file)
            except: uploaded_file.seek(0); df = pd.read_csv(uploaded_file, encoding='cp950')
        else: df = pd.read_excel(uploaded_file)
        
        # 清理欄位空格
        df.columns = [str(c).strip() for c in df.columns]

        # 1. 自動偵測必要欄位
        loc_col = get_col_name(df, ['違規地點', '路口名稱', '地點'])
        time_col = get_col_name(df, ['入案時間', '違規時間', '時間'])
        
        if not loc_col or not time_col:
            st.error(f"❌ 找不到必要欄位！檔案欄位有：{list(df.columns)}")
            st.stop()

        # 2. 地名精簡 (刪除 桃園市、龍潭區)
        df[loc_col] = df[loc_col].astype(str).str.replace('桃園市', '', regex=False).str.replace('龍潭區', '', regex=False).str.strip()
        
        # 3. 統計處理
        date_range_str = format_roc_date_range_fixed_start(df)
        df['小時'] = df[time_col].apply(parse_hour)
        
        # 十大路段
        loc_summary = df[loc_col].value_counts().head(10).reset_index()
        loc_summary.columns = ['路段名稱', '舉發件數']
        
        # 時段分佈
        hour_counts = df['小時'].value_counts().reindex(range(24), fill_value=0).reset_index()
        hour_counts.columns = ['小時', '舉發件數']

        # 4. 網頁顯示
        st.divider()
        st.subheader(f"📅 統計期間：{date_range_str}")
        c1, c2 = st.columns(2)
        with c1: st.dataframe(loc_summary, use_container_width=True)
        with c2: st.bar_chart(hour_counts.set_index('小時'))

        # 5. 按鈕執行
        if st.button("🚀 產製 1/1 起算之報表並同步寄送", type="primary"):
            with st.spinner("⚡ 報表產製中..."):
                # A. 產製 Excel
                excel_data = create_formatted_excel(loc_summary, hour_counts, date_range_str, len(df))
                
                # B. 同步 Google Sheets (數據同步)
                try:
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    for name, d in zip(["科技執法-路段排行", "科技執法-時段分析"], [loc_summary, hour_counts]):
                        try: ws = sh.worksheet(name)
                        except: ws = sh.add_worksheet(title=name, rows="100", cols="20")
                        ws.clear(); ws.update([d.columns.values.tolist()] + d.values.tolist())
                    st.success("✅ Google 試算表數據同步成功")
                except Exception as e: st.warning(f"⚠️ 雲端同步失敗: {e}")

                # C. 寄送 Email
                try:
                    msg = MIMEMultipart()
                    msg['From'], msg['To'] = MY_EMAIL, TO_EMAIL
                    msg['Subject'] = f"科技執法統計報告({date_range_str})"
                    msg.attach(MIMEText(f"長官好，科技執法成效統計報表（1/1起算）已產製完成。\n\n統計期間：{date_range_str}\n舉發總件數：{len(df)} 件", 'plain'))
                    
                    part = MIMEApplication(excel_data.getvalue(), Name="Enforcement_Report.xlsx")
                    part.add_header('Content-Disposition', 'attachment', filename="Enforcement_Report.xlsx")
                    msg.attach(part)
                    
                    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
                        s.starttls(); s.login(MY_EMAIL, MY_PASSWORD); s.send_message(msg)
                    st.success(f"✅ 報表已寄送至：{TO_EMAIL}")
                    st.balloons()
                except Exception as e: st.error(f"❌ 郵件寄送失敗：{e}")

    except Exception as e:
        st.error(f"系統錯誤：{e}")
