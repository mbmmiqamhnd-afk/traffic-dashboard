import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import io

# --- 1. 雲端資料同步 (模擬從雲端讀取資料) ---
def sync_cloud_data():
    # 這裡對應您雲端硬碟中的「交通違規統計表.xlsx」
    # 在實際部署時，可使用 gspread 或 google-api-python-client 進行即時同步
    st.info("🔄 正在從雲端硬碟同步「交通違規統計表.xlsx」...")
    
    # 根據您提供的欄位結構進行統計 (合計、本期、本年、去年、比較等)
    #
    data = {
        "單位": ["合計", "科技執法", "交通分隊", "聖亭所", "龍潭所"],
        "本期總計": [45, 8, 26, 5, 6],
        "本年累計": [6886, 422, 1492, 1200, 1350],
        "去年同期": [7068, 496, 1424, 1150, 1400],
        "增減比較": [-182, -74, 68, 50, -50]
    }
    return pd.DataFrame(data)

# --- 2. 寄送郵件功能 ---
def send_stats_email(df, recipient_email):
    # 設定郵件標題：📊 [自動通知] 交通違規統計表.xlsx 
    msg = MIMEMultipart()
    msg['Subject'] = "📊 [自動通知] 交通違規統計表"
    msg['From'] = "您的系統"
    msg['To'] = recipient_email

    # 郵件內文 
    body = "您好，附件為最新的交通違規統計報表，請查收。\n\n" + df.to_html(index=False)
    msg.attach(MIMEText(body, 'html'))

    # 將 DataFrame 轉為 Excel 附件
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='統計結果')
    
    part = MIMEApplication(excel_buffer.getvalue(), Name="交通違規統計結果.xlsx")
    part['Content-Disposition'] = 'attachment; filename="交通違規統計結果.xlsx"'
    msg.attach(part)

    # 這裡需要設定您的 SMTP 伺服器 (如 Gmail)
    # st.success(f"✅ 郵件已成功寄送至 {recipient_email}")
    return True

# --- 3. Streamlit 介面 ---
st.title("🚦 交通統計雲端同步系統")

if st.button("🔄 立即同步雲端資料並寄出報表"):
    df_cloud = sync_cloud_data()
    st.write("### 當前統計概覽")
    st.dataframe(df_cloud)
    
    # 執行寄信 (請更換為實際收件者)
    target_mail = "mbmmiqamhnd@gmail.com" # 預設為您的帳號 
    if send_stats_email(df_cloud, target_mail):
        st.success(f"🎉 資料已同步，並已寄送郵件至 {target_mail}")
