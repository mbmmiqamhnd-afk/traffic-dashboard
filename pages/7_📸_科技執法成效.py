import streamlit as st
import pandas as pd
import io
import smtplib
import gspread
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# 1. 頁面配置
st.set_page_config(page_title="科技執法統計 - 累計至昨日版", layout="wide", page_icon="📸")

# 2. 自動化設定 (改由 Secrets 讀取)
try:
    MY_EMAIL = st.secrets["email"]["user"]
    MY_PASSWORD = st.secrets["email"]["password"]
    GCP_CREDS = st.secrets["gcp_service_account"]
except Exception as e:
    st.error("❌ 找不到 Secrets 設定，請確認設定區包含 [email] 與 [gcp_service_account]")
    st.stop()

TO_EMAIL = "mbmmiqamhnd@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

st.title("📸 科技執法成效分析系統")
st.markdown("### 📝 狀態：統計期間設定為「1月1日」起至「上傳前一日（昨日）」。")

# --- 工具函數 ---
def get_col_name(df, possible_names):
    clean_cols = [str(c).strip() for c in df.columns]
    for name in possible_names:
        if name in clean_cols:
            return df.columns[clean_cols.index(name)]
    return None

def format_roc_date_range_to_yesterday():
    yesterday = datetime.now() - timedelta(days=1)
    roc_year = yesterday.year - 1911
    month = yesterday.month
    day = yesterday.day
    return f"{roc_year}年1月1日至{roc_year}年{month}月{day}日"

# --- 核心：建立 Excel (包含混合顏色標題) ---
def create_formatted_excel(df_loc, date_range_text, total_count):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        ws = workbook.add_worksheet('科技執法成效統計')
        
        # 💡 更新：定義格式 (粗體、24號字、指定顏色)
        blue_title_fmt = workbook.add_format({'bold': True, 'font_size': 24, 'color': 'blue'})
        red_title_fmt = workbook.add_format({'bold': True, 'font_size': 24, 'color': 'red'}) 
        
        header_fmt = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1, 'bold': True, 'align': 'center'})
        data_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        total_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFFFCC', 'align': 'center'})

        # 💡 更新：寫入混合格式：藍色字串 + 紅色字串 (括號及期間)
        ws.write_rich_string('A1', blue_title_fmt, '科技執法成效 ', red_title_fmt, f'({date_range_text})')
        
        ws.write('A2', '統計期間', workbook.add_format({'align': 'center', 'border': 1}))
        ws.write('B2', date_range_text, workbook.add_format({'border': 1, 'color': 'red', 'align': 'center'}))
        ws.write('A3', '路口名稱', header_fmt)
        ws.write('B3', '舉發件數', header_fmt)
        
        curr_row = 3
        for _, row in df_loc.iterrows():
            ws.write(curr_row, 0, row['路段名稱'], data_fmt)
            ws.write(curr_row, 1, row['舉發件數'], data_fmt)
            curr_row += 1
        
        ws.write(curr_row, 0, '舉發總數', total_fmt)
        ws.write(curr_row, 1, total_count, total_fmt)
        
        # 繪製長條圖
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            'name': '舉發件數',
            'categories': ['科技執法成效統計', 3, 0, curr_row - 1, 0],
            'values':      ['科技執法成效統計', 3, 1, curr_row - 1, 1],
            'data_labels': {'value': True},
        })
        chart.set_title({'name': '違規路段排行'})
        ws.insert_chart('D2', chart, {'x_scale': 1.5, 'y_scale': 1.5})
        
    return output

# --- 主流程 ---
uploaded_file = st.file_uploader("請上傳清冊檔案 (如 list2.csv 或 Excel)", type=['csv', 'xlsx'])

if uploaded_file:
    try:
        # 讀取檔案
        if uploaded_file.name.endswith('.csv'):
            try: df = pd.read_csv(uploaded_file)
            except: 
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding='cp950')
        else: 
            df = pd.read_excel(uploaded_file)
        
        df.columns = [str(c).strip() for c in df.columns]
        loc_col = get_col_name(df, ['違規地點', '路口名稱', '地點'])
        
        if not loc_col:
            st.error("❌ 找不到『地點』相關欄位！請確認檔案格式。")
            st.stop()

        # 整理路段名稱與計算排行
        df[loc_col] = df[loc_col].astype(str).str.replace('桃園市', '', regex=False).str.replace('龍潭區', '', regex=False).str.strip()
        date_range_str = format_roc_date_range_to_yesterday()
        
        loc_summary = df[loc_col].value_counts().head(10).reset_index()
        loc_summary.columns = ['路段名稱', '舉發件數']

        st.divider()
        # Streamlit 網頁介面也同步變成前藍後紅
        st.markdown(f"### 📅 統計期間 (累計至昨日)：:blue[科技執法成效 ]:red[({date_range_str})]")
        
        # 顯示路段排行 (置中顯示)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.dataframe(loc_summary, use_container_width=True, hide_index=True)

        if st.button("🚀 產製累計至昨日報表並同步寄送", type="primary", use_container_width=True):
            with st.spinner("⚡ 系統作業中..."):
                excel_data = create_formatted_excel(loc_summary, date_range_str, len(df))
                
                # 同步 Google Sheets
                try:
                    gc = gspread.service_account_from_dict(GCP_CREDS)
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    sheet_name = "科技執法-路段排行"
                    try: 
                        ws = sh.worksheet(sheet_name)
                    except: 
                        ws = sh.add_worksheet(title=sheet_name, rows="100", cols="20")
                    ws.clear()
                    
                    # 1. 先寫入所有資料 (純文字)
                    title_text = f"科技執法成效 ({date_range_str})"
                    update_data = [
                        [title_text, ""],
                        ["路段名稱", "舉發件數"]
                    ] + loc_summary.values.tolist() + [["舉發總數", len(df)]]
                    
                    ws.update(values=update_data)

                    # 2. 💡 更新：使用 batch_update 調整字體大小為 24，前半藍色、後半紅色
                    start_index_of_red = len("科技執法成效 ") 
                    requests = {
                        "requests": [
                            {
                                "updateCells": {
                                    "range": {
                                        "sheetId": ws.id,
                                        "startRowIndex": 0, "endRowIndex": 1,
                                        "startColumnIndex": 0, "endColumnIndex": 1
                                    },
                                    "rows": [
                                        {
                                            "values": [
                                                {
                                                    "userEnteredValue": {"stringValue": title_text},
                                                    "textFormatRuns": [
                                                        {
                                                            "startIndex": 0,
                                                            "format": {
                                                                "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 1.0}, # 藍色
                                                                "bold": True,
                                                                "fontSize": 24  # 24號字
                                                            }
                                                        },
                                                        {
                                                            "startIndex": start_index_of_red,
                                                            "format": {
                                                                "foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}, # 紅色
                                                                "bold": True,
                                                                "fontSize": 24  # 24號字
                                                            }
                                                        }
                                                    ]
                                                }
                                            ]
                                        }
                                    ],
                                    "fields": "userEnteredValue,textFormatRuns"
                                }
                            }
                        ]
                    }
                    sh.batch_update(requests)
                    st.success("✅ Google 試算表『路段排行』同步成功 (標題已更新為 24pt 藍紅雙色)！")
                except Exception as e: 
                    st.warning(f"⚠️ 雲端同步失敗: {e}")

                # 寄送 Email
                try:
                    msg = MIMEMultipart()
                    msg['From'], msg['To'] = MY_EMAIL, TO_EMAIL
                    msg['Subject'] = f"科技執法統計報告({date_range_str})"
                    msg.attach(MIMEText(f"長官好，科技執法路段排行報表已完成。\n\n統計期間：{date_range_str}\n舉發總件數：{len(df)} 件", 'plain'))
                    
                    part = MIMEApplication(excel_data.getvalue(), Name="Tech_Enforcement.xlsx")
                    part.add_header('Content-Disposition', 'attachment', filename="Tech_Enforcement.xlsx")
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
        st.error(f"系統處理錯誤：{e}")
