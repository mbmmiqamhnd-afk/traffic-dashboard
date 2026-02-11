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
# 👇👇👇 【使用者自動化設定區】 👇👇👇
# ==========================================
MY_EMAIL = "mbmmiqamhnd@gmail.com" 
MY_PASSWORD = "kvpw ymgn xawe qxnl"  # 您的 Gmail 應用程式密碼
TO_EMAIL = "mbmmiqamhnd@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
# ==========================================

st.set_page_config(page_title="科技執法成效統計", layout="wide", page_icon="📸")

st.title("📸 科技執法成效分析 (一鍵寄送圖表版)")
st.markdown("### 📝 狀態：密碼已內建，支援自動產生圖表並寄送至信箱。")

# 1. 檔案上傳
uploaded_file = st.file_uploader("請上傳科技執法清冊 (如: list2.csv)", type=['csv', 'xlsx'], key="tech_v7")

if uploaded_file:
    try:
        # 讀取資料
        if uploaded_file.name.endswith('.csv'):
            try: df = pd.read_csv(uploaded_file)
            except: uploaded_file.seek(0); df = pd.read_csv(uploaded_file, encoding='cp950')
        else:
            df = pd.read_excel(uploaded_file)
        
        df.columns = [str(c).strip() for c in df.columns]

        # 資料清理與日期轉換
        def parse_roc_date(val):
            try:
                s = str(int(val)).zfill(7)
                return datetime(int(s[:-4]) + 1911, int(s[-4:-2]), int(s[-2:]))
            except: return None
        
        df['日期_dt'] = df['違規日期'].apply(parse_roc_date)
        df['小時'] = df['違規時間'].apply(lambda x: int(str(int(x)).zfill(4)[:2]) if pd.notna(x) else 0)

        # 2. 網頁圖表顯示
        st.divider()
        st.subheader("📊 即時統計預覽")
        c1, c2 = st.columns(2)
        with c1:
            loc_counts = df['違規地點'].value_counts().head(10)
            st.write("📍 十大違規路段")
            st.bar_chart(loc_counts)
        with c2:
            hour_counts = df['小時'].value_counts().sort_index()
            st.write("⏰ 違規時段分佈")
            st.bar_chart(hour_counts.combine_first(pd.Series(0, index=range(24))))

        # ==========================================
        # 3. 寄送圖表與報表功能
        # ==========================================
        st.divider()
        if st.button(f"🚀 寄送統計圖表與報表至 {TO_EMAIL}", type="primary"):
            try:
                with st.spinner("⚡ 系統正在繪製圖表並寄送信件..."):
                    # --- A. 產生圖片 (Matplotlib) ---
                    # 解決 Matplotlib 中文顯示問題 (標題改用英文或不使用特殊字體)
                    def get_chart_img(data, title, is_hour=False):
                        plt.figure(figsize=(8, 5))
                        data.plot(kind='bar', color='skyblue')
                        plt.title(title)
                        plt.tight_layout()
                        img_buf = io.BytesIO()
                        plt.savefig(img_buf, format='png')
                        img_buf.seek(0)
                        plt.close()
                        return img_buf

                    img_loc = get_chart_img(loc_counts, "Top 10 Locations")
                    img_hour = get_chart_img(hour_counts.combine_first(pd.Series(0, index=range(24))), "Hourly Distribution")

                    # --- B. 建立郵件 ---
                    msg = MIMEMultipart()
                    msg['From'] = MY_EMAIL
                    msg['To'] = TO_EMAIL
                    msg['Subject'] = f"科技執法成效報告 - {datetime.now().strftime('%Y/%m/%d')}"
                    
                    body = f"""長官好，

檢送本次科技執法統計結果如下：
- 上傳檔案：{uploaded_file.name}
- 舉發總件數：{len(df)} 件
- 違規最高路段：{df['違規地點'].mode()[0]}

郵件已附加統計圖片(PNG)與完整清冊(CSV)，請查照。
(此郵件由系統自動發送)"""
                    msg.attach(MIMEText(body, 'plain'))

                    # 附加圖表圖片
                    for img_data, name in [(img_loc, "Locations.png"), (img_hour, "Hours.png")]:
                        img_part = MIMEImage(img_data.read(), name=name)
                        img_part.add_header('Content-Disposition', f'attachment; filename="{name}"')
                        msg.attach(img_part)

                    # 附加 CSV 數據
                    csv_buf = io.BytesIO()
                    df.to_csv(csv_buf, index=False, encoding='utf-8-sig')
                    csv_part = MIMEApplication(csv_buf.getvalue(), Name="Data_Report.csv")
                    csv_part.add_header('Content-Disposition', 'attachment; filename="Data_Report.csv"')
                    msg.attach(csv_part)

                    # --- C. SMTP 寄送 ---
                    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                        server.starttls()
                        server.login(MY_EMAIL, MY_PASSWORD)
                        server.send_message(msg)
                
                st.balloons()
                st.success(f"✅ 圖表與報表已送達：{TO_EMAIL}")
            except Exception as e:
                st.error(f"❌ 寄送失敗：{e}")

        with st.expander("🔍 查看詳細清冊"):
            st.dataframe(df)

    except Exception as e:
        st.error(f"解析失敗：{e}")
