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
st.set_page_config(page_title="科技執法統計 - 格式修正版", layout="wide", page_icon="📸")

# 2. 自動化設定
MY_EMAIL = "mbmmiqamhnd@gmail.com" 
MY_PASSWORD = "kvpw ymgn xawe qxnl"  
TO_EMAIL = "mbmmiqamhnd@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

st.title("📸 科技執法成效分析系統")
st.markdown("### 📝 狀態：已修正欄位偵測邏輯，解決『違規時間』找不到的問題。")

# --- 工具函數 ---
def parse_hour(val):
    try:
        s = str(int(val)).zfill(4)
        return int(s[:2])
    except: return 0

def get_col_name(df, possible_names):
    """從 DataFrame 中尋找可能的欄位名稱"""
    for name in possible_names:
        if name in df.columns:
            return name
    return None

def format_roc_date_range(df):
    """擷取日期範圍並轉為民國格式"""
    target_col = get_col_name(df, ['入案日期', '入案時間', '違規日期', '日期'])
    if not target_col: return "期間未定"
    try:
        valid_dates = pd.to_numeric(df[target_col], errors='coerce').dropna().astype(int)
        if valid_dates.empty: return "無有效日期"
        def to_roc_str(val):
            s = str(val).zfill(7)
            return f"{int(s[:-4])}年{int(s[-4:-2])}月{int(s[-2:])}日"
        return f"{to_roc_str(valid_dates.min())}至{to_roc_str(valid_dates.max())}"
    except: return "日期解析錯誤"

# --- 核心：建立 Excel ---
def create_formatted_excel(df_loc, df_hour, date_range_text, total_count):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        ws = workbook.add_worksheet('科技執法成效統計')
        
        # 格式
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14})
        header_fmt = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1, 'bold': True, 'align': 'center'})
        data_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        total_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFFFCC'})

        ws.write('A1', '科技執法成效', title_fmt)
        ws.write('A2', '統計期間')
        ws.write('B2', date_range_text)
        ws.write('A3', '路口名稱', header_fmt)
        ws.write('B3', '舉發件數', header_fmt)
        
        for i, (_, row) in enumerate(df_loc.iterrows(), 4):
            ws.write(f'A{i}', row['路段名稱'], data_fmt)
            ws.write(f'B{i}', row['舉發件數'], data_fmt)
        
        last_row = 3 + len(df_loc)
        ws.write(last_row, 0, '舉發總數', total_fmt)
        ws.write(last_row, 1, total_count, total_fmt)
        
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            'name': '舉發件數',
            'categories': ['科技執法成效統計', 3, 0, last_row - 1, 0],
            'values':     ['科技執法成效統計', 3, 1, last_row - 1, 1],
            'data_labels': {'value': True},
        })
        chart.set_title({'name': '違規路段排行'})
        ws.insert_chart('D2', chart, {'x_scale': 1.5, 'y_scale': 1.5})
        df_hour.to_excel(writer, sheet_name='時段分析', index=False)
    return output

# --- 主流程 ---
uploaded_file = st.file_uploader("請上傳清冊檔案", type=['csv', 'xlsx'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            try: df = pd.read_csv(uploaded_file)
            except: uploaded_file.seek(0); df = pd.read_csv(uploaded_file, encoding='cp950')
        else: df = pd.read_excel(uploaded_file)
        
        df.columns = [str(c).strip() for c in df.columns]

        # 1. 偵測正確的欄位名稱
        loc_col = get_col_name(df, ['違規地點', '路口名稱', '地點'])
        time_col = get_col_name(df, ['違規時間', '入案時間', '時間'])
        
        if not loc_col or not time_col:
            st.error(f"❌ 找不到必要欄位！目前欄位包含：{list(df.columns)}")
            st.info("請確認檔案包含：違規地點、違規時間(或入案時間)")
            st.stop()

        # 2. 地名精簡 (刪除 桃園市、龍潭區)
        df[loc_col] = df[loc_col].astype(str).str.replace('桃園市', '', regex=False).str.replace('龍潭區', '', regex=False).str.strip()
        
        # 3. 數據統計
        date_range_str = format_roc_date_range(df)
        df['小時'] = df[time_col].apply(parse_hour)
        
        loc_summary = df[loc_col].value_counts().head(10).reset_index()
        loc_summary.columns = ['路段名稱', '舉發件數']
        
        hour_counts = df['小時'].value_counts().reindex(range(24), fill_value=0).reset_index()
        hour_counts.columns = ['小時', '舉發件數']

        # 4. 網頁呈現
        st.divider()
        st.subheader(f"📅 統計期間：{date_range_str}")
        c1, c2 = st.columns(2)
        with c1: st.dataframe(loc_summary, use_container_width=True)
        with c2: st.bar_chart(hour_counts.set_index('小時'))

        # 5. 執行按鈕
        if st.button("🚀 產製報表並同步雲端與寄送", type="primary"):
            with st.spinner("處理中..."):
                excel_data = create_formatted_excel(loc_summary, hour_counts, date_range_str, len(df))
                
                # 同步 Google Sheets
                try:
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    for name, d in zip(["科技執法-路段排行", "科技執法-時段分析"], [loc_summary, hour_counts]):
                        try: ws = sh.worksheet(name)
                        except: ws = sh.add_worksheet(title=name, rows="100", cols="20")
                        ws.clear(); ws.update([d.columns.values.tolist()] + d.values.tolist())
                    st.success("✅ Google 試算表同步成功")
                except Exception as e: st.warning(f"⚠️ 同步提示: {e}")

                # 寄送郵件
                try:
                    msg = MIMEMultipart()
                    msg['From'], msg['To'] = MY_EMAIL, TO_EMAIL
                    msg['Subject'] = f"科技執法統計報告({date_range_str})"
                    msg.attach(MIMEText(f"統計期間：{date_range_str}\n總件數：{len(df)} 件", 'plain'))
                    part = MIMEApplication(excel_data.getvalue(), Name="Report.xlsx")
                    part.add_header('Content-Disposition', 'attachment', filename="Report.xlsx")
                    msg.attach(part)
                    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
                        s.starttls(); s.login(MY_EMAIL, MY_PASSWORD); s.send_message(msg)
                    st.success(f"✅ 報表已寄送至：{TO_EMAIL}")
                    st.balloons()
                except Exception as e: st.error(f"❌ 寄送失敗：{e}")

    except Exception as e:
        st.error(f"系統錯誤：{e}")
