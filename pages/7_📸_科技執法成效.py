import streamlit as st
import pandas as pd
import io
import smtplib
import matplotlib.pyplot as plt
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage

# ==========================================
# 👇👇👇 【自動化寄信設定：密碼已埋入】 👇👇👇
# ==========================================
MY_EMAIL = "mbmmiqamhnd@gmail.com" 
MY_PASSWORD = "kvpw ymgn xawe qxnl" 
TO_EMAIL = "mbmmiqamhnd@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
# ==========================================

st.set_page_config(page_title="科技執法成效統計", layout="wide", page_icon="📸")

st.title("📸 科技執法成效分析 (路口名稱優化版)")
st.markdown("### 📝 狀態：已優化圖表標籤顯示，確保寄送的圖表能看見路口名稱。")

uploaded_file = st.file_uploader("請上傳科技執法清冊 (list2.csv)", type=['csv', 'xlsx'])

if uploaded_file:
    try:
        # 讀取資料
        if uploaded_file.name.endswith('.csv'):
            try: df = pd.read_csv(uploaded_file)
            except: uploaded_file.seek(0); df = pd.read_csv(uploaded_file, encoding='cp950')
        else:
            df = pd.read_excel(uploaded_file)
        
        df.columns = [str(c).strip() for c in df.columns]

        # 日期轉換
        def parse_roc_date(val):
            try:
                s = str(int(val)).zfill(7)
                return datetime(int(s[:-4]) + 1911, int(s[-4:-2]), int(s[-2:]))
            except: return None
        
        df['日期_dt'] = df['違規日期'].apply(parse_roc_date)
        df['小時'] = df['違規時間'].apply(lambda x: int(str(int(x)).zfill(4)[:2]) if pd.notna(x) else 0)

        # 網頁即時預覽 (使用 Streamlit 內建圖表，中文顯示沒問題)
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            loc_counts = df['違規地點'].value_counts().head(10)
            st.subheader("📍 十大違規路段排行")
            st.bar_chart(loc_counts)
        with c2:
            hour_counts = df['小時'].value_counts().sort_index()
            st.subheader("⏰ 違規時段分佈")
            st.bar_chart(hour_counts.combine_first(pd.Series(0, index=range(24))))

        # ==========================================
        # 3. 寄送圖表功能
        # ==========================================
        st.divider()
        if st.button(f"🚀 寄送含路口名稱之圖表至 {TO_EMAIL}", type="primary"):
            try:
                with st.spinner("⚡ 正在產生報表圖片..."):
                    
                    # --- A. 產生路口排行圖片 (改為橫向以顯示長名稱) ---
                    def create_loc_plot(data):
                        # 建立畫布，寬度增加
                        plt.figure(figsize=(12, 8))
                        # 改用橫向長條圖 barh
                        data.sort_values(ascending=True).plot(kind='barh', color='skyblue')
                        plt.title("Top 10 Violation Locations", fontsize=16)
                        plt.xlabel("Count", fontsize=12)
                        # 自動調整佈局，給左側文字留更多空間
                        plt.tight_layout()
                        
                        img_buf = io.BytesIO()
                        plt.savefig(img_buf, format='png', dpi=150)
                        img_buf.seek(0)
                        plt.close()
                        return img_buf

                    # --- B. 產生時段分析圖片 ---
                    def create_hour_plot(data):
                        plt.figure(figsize=(10, 6))
                        data.plot(kind='bar', color='orange')
                        plt.title("Violation Hourly Distribution", fontsize=16)
                        plt.tight_layout()
                        img_buf = io.BytesIO()
                        plt.savefig(img_buf, format='png', dpi=150)
                        img_buf.seek(0)
                        plt.close()
                        return img_buf

                    img_loc = create_loc_plot(loc_counts)
                    img_hour = create_hour_plot(hour_counts.combine_first(pd.Series(0, index=range(24))))

                    # --- C. 建立郵件 ---
                    msg = MIMEMultipart()
                    msg['From'] = MY_EMAIL
                    msg['To'] = TO_EMAIL
                    msg['Subject'] = f"科技執法統計報告 - {datetime.now().strftime('%Y/%m/%d')}"
                    
                    body = f"長官好，檢送本次科技執法統計結果。附件包含路口排行榜圖片與原始數據清冊。"
                    msg.attach(MIMEText(body, 'plain'))

                    # 附加圖表圖片
                    for img_data, name in [(img_loc, "Locations_Chart.png"), (img_hour, "Hours_Chart.png")]:
                        img_part = MIMEImage(img_data.read(), name=name)
                        img_part.add_header('Content-Disposition', f'attachment; filename="{name}"')
                        msg.attach(img_part)

                    # 附加數據
                    csv_buf = io.BytesIO()
                    df.to_csv(csv_buf, index=False, encoding='utf-8-sig')
                    csv_part = MIMEApplication(csv_buf.getvalue(), Name="Data.csv")
                    csv_part.add_header('Content-Disposition', 'attachment; filename="Full_Data.csv"')
                    msg.attach(csv_part)

                    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                        server.starttls()
                        server.login(MY_EMAIL, MY_PASSWORD)
                        server.send_message(msg)
                
                st.balloons()
                st.success(f"✅ 報表已送達：{TO_EMAIL}")
            except Exception as e:
                st.error(f"❌ 寄送失敗：{e}")

    except Exception as e:
        st.error(f"處理失敗：{e}")
