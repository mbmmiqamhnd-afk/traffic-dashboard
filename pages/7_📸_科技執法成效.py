import streamlit as st
import pandas as pd
import io
import smtplib
import re
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ==========================================
# 👇👇👇 【自動化寄信設定】 👇👇👇
# ==========================================
# 比照您提供的「交通事故統計」項目設定
MY_EMAIL = "mbmmiqamhnd@gmail.com" 
MY_PASSWORD = "kvpw ymgn xawe qxnl"  # 您的 Gmail 應用程式密碼
TO_EMAIL = "mbmmiqamhnd@gmail.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
# ==========================================

st.set_page_config(page_title="科技執法成效統計", layout="wide", page_icon="📸")

st.title("📸 科技執法成效分析系統 (自動寄送版)")
st.markdown("### 📝 功能：上傳清冊後自動產製圖表，並支援一鍵寄送電子郵件。")

# 1. 檔案上傳區
uploaded_file = st.file_uploader("請上傳科技執法清冊 (如: list2.csv)", type=['csv', 'xlsx'], key="tech_uploader_final")

if uploaded_file:
    try:
        # --- (A) 資料讀取 ---
        if uploaded_file.name.endswith('.csv'):
            try:
                df = pd.read_csv(uploaded_file)
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding='cp950')
        else:
            df = pd.read_excel(uploaded_file)
        
        df.columns = [str(c).strip() for c in df.columns]

        # --- (B) 資料處理邏輯 ---
        # 民國轉西元日期
        def parse_roc_date(val):
            try:
                s = str(int(val)).zfill(7)
                year = int(s[:-4]) + 1911
                month = int(s[-4:-2])
                day = int(s[-2:])
                return datetime(year, month, day)
            except: return None
        
        # 轉換時間
        def parse_hour(val):
            try:
                s = str(int(val)).zfill(4)
                return int(s[:2])
            except: return 0

        df['日期_dt'] = df['違規日期'].apply(parse_roc_date)
        df['小時'] = df['違規時間'].apply(parse_hour)

        # --- (C) 視覺化圖表 ---
        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("📸 舉發總件數", f"{len(df):,} 件")
        m2.metric("📍 違規熱點", df['違規地點'].mode()[0] if not df.empty else "N/A")
        m3.metric("🚙 主要車種", df['車種'].mode()[0] if not df.empty else "N/A")

        col_left, col_right = st.columns(2)
        with col_left:
            st.subheader("📍 十大違規路段排行")
            st.bar_chart(df['違規地點'].value_counts().head(10))
            
        with col_right:
            st.subheader("⏰ 違規時段分佈 (24H)")
            hour_counts = df['小時'].value_counts().sort_index()
            full_hours = pd.Series(0, index=range(24))
            st.bar_chart(hour_counts.combine_first(full_hours))

        # --- (D) 一鍵寄信按鈕 ---
        st.divider()
        st.subheader("📧 報表自動寄送")
        
        if st.button(f"🚀 立即寄送統計報表至 {TO_EMAIL}", type="primary"):
            try:
                with st.spinner("⚡ 正在建立報表並發送郵件..."):
                    # 建立郵件物件
                    msg = MIMEMultipart()
                    msg['From'] = MY_EMAIL
                    msg['To'] = TO_EMAIL
                    msg['Subject'] = f"科技執法成效統計報表 ({datetime.now().strftime('%Y/%m/%d')})"
                    
                    body = f"""長官好，

檢送本次科技執法成效統計報表（檔案：{uploaded_file.name}），摘要如下：
- 舉發總件數：{len(df)} 件
- 違規最高路段：{df['違規地點'].mode()[0]}

詳細數據請參閱附檔。
(此郵件由系統自動發送)"""
                    msg.attach(MIMEText(body, 'plain'))
                    
                    # 製作 CSV 附件
                    csv_buffer = io.BytesIO()
                    df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                    attachment = MIMEApplication(csv_buffer.getvalue(), Name="科技執法統計結果.csv")
                    attachment['Content-Disposition'] = 'attachment; filename="Tech_Enforcement_Report.csv"'
                    msg.attach(attachment)
                    
                    # SMTP 寄送
                    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                        server.starttls()
                        server.login(MY_EMAIL, MY_PASSWORD)
                        server.send_message(msg)
                
                st.success(f"✅ 報表已成功寄送至：{TO_EMAIL}")
                st.balloons()
            except Exception as e:
                st.error(f"❌ 寄送失敗：{e}")

        # 原始資料顯示
        with st.expander("🔍 查看原始資料表"):
            st.dataframe(df, use_container_width=True)

    except Exception as e:
        st.error(f"系統錯誤：{e}")
else:
    st.info("💡 請上傳科技執法清冊 (list2.csv) 以開啟統計功能。")
