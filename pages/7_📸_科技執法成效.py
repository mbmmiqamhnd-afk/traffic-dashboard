import streamlit as st
import pandas as pd
from datetime import datetime
import io
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# 1. 頁面基本配置
st.set_page_config(page_title="科技執法成效統計", layout="wide", page_icon="📸")

st.title("📸 科技執法成效自動化分析系統")
st.info("💡 上傳清冊後，系統將自動分析數據。點擊下方按鈕可一鍵寄送報表至管理信箱。")

# ==========================================
# 2. 數據處理核心邏輯 (針對 list2.csv 格式)
# ==========================================
uploaded_file = st.file_uploader("請上傳科技執法清冊 (如: list2.csv 或 list2.xlsx)", type=['csv', 'xlsx'])

if uploaded_file:
    try:
        # 讀取檔案內容
        if uploaded_file.name.endswith('.csv'):
            try:
                # 嘗試讀取 UTF-8
                df = pd.read_csv(uploaded_file)
            except:
                # 失敗則嘗試 CP950 (Excel CSV 常見編碼)
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding='cp950')
        else:
            df = pd.read_excel(uploaded_file)
        
        # 清理欄位多餘空白
        df.columns = [str(c).strip() for c in df.columns]

        # 日期轉換函數 (民國 1141231 -> 西元 2025-12-31)
        def parse_roc_date(val):
            try:
                s = str(int(val)).zfill(7)
                year = int(s[:-4]) + 1911
                month = int(s[-4:-2])
                day = int(s[-2:])
                return datetime(year, month, day)
            except: return None
        
        # 時間轉換函數 (1240 -> 12時)
        def parse_hour(val):
            try:
                s = str(int(val)).zfill(4)
                return int(s[:2])
            except: return 0

        # 套用轉換
        df['日期_dt'] = df['違規日期'].apply(parse_roc_date)
        df['小時'] = df['違規時間'].apply(parse_hour)

        # ==========================================
        # 3. 視覺化統計圖表 (穩定版)
        # ==========================================
        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("📸 總舉發件數", f"{len(df):,} 件")
        m2.metric("📍 違規熱點", df['違規地點'].mode()[0] if not df.empty else "N/A")
        m3.metric("🚙 主要違規車種", df['車種'].mode()[0] if not df.empty else "N/A")

        col_l, col_r = st.columns(2)
        with col_l:
            st.subheader("📍 十大違規路段排行")
            # 統計地點出現次數並取前十名
            loc_data = df['違規地點'].value_counts().head(10)
            st.bar_chart(loc_data)
            
        with col_r:
            st.subheader("⏰ 24小時違規時段分佈")
            hour_counts = df['小時'].value_counts().sort_index()
            # 確保 0-23 小時都有顯示
            full_hours = pd.Series(0, index=range(24))
            st.bar_chart(hour_counts.combine_first(full_hours))

        st.divider()
        st.subheader("📅 執法成效每日趨勢")
        if not df['日期_dt'].isnull().all():
            trend_df = df.groupby('日期_dt').size()
            st.line_chart(trend_df)

        # ==========================================
        # 4. 全自動寄信功能 (參考現有專案模式)
        # ==========================================
        st.divider()
        st.subheader("📧 報表自動化發送")
        
        # 固定收件人
        target_email = "mbmmiqamhnd@gmail.com"

        # 從 Secrets 抓取帳密，不需要用戶輸入
        if "GMAIL_USER" in st.secrets and "GMAIL_PASS" in st.secrets:
            if st.button(f"🚀 點擊自動寄送報表至 {target_email}", type="primary"):
                try:
                    with st.spinner("正在產生分析附件並發送郵件..."):
                        sender_user = st.secrets["GMAIL_USER"]
                        sender_pass = st.secrets["GMAIL_PASS"]
                        
                        # 準備郵件物件
                        msg = MIMEMultipart()
                        msg['From'] = sender_user
                        msg['To'] = target_email
                        msg['Subject'] = f"【自動報表】科技執法成效統計 - {datetime.now().strftime('%Y/%m/%d')}"
                        
                        # 郵件本文
                        body = f"""您好：
                        
                        附件為科技執法成效統計報表，分析摘要如下：
                        - 上傳檔案：{uploaded_file.name}
                        - 舉發總件數：{len(df)} 件
                        - 統計生成時間：{datetime.now().strftime('%Y/%m/%d %H:%M')}
                        
                        詳細違規清冊請參閱附件 CSV 檔案。
                        """
                        msg.attach(MIMEText(body, 'plain'))
                        
                        # 建立附件
                        csv_buffer = io.BytesIO()
                        df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                        attachment = MIMEApplication(csv_buffer.getvalue(), Name="Tech_Enforcement_Report.csv")
                        attachment['Content-Disposition'] = 'attachment; filename="Tech_Enforcement_Report.csv"'
                        msg.attach(attachment)
                        
                        # 透過 SMTP 發信
                        with smtplib.SMTP('smtp.gmail.com', 587) as server:
                            server.starttls()
                            server.login(sender_user, sender_pass)
                            server.send_message(msg)
                            
                    st.success(f"✅ 郵件已成功自動送達：{target_email}")
                except Exception as e:
                    st.error(f"❌ 寄送失敗。請檢查您的 Secrets 設定或 Gmail 權限。錯誤訊息：{e}")
        else:
            st.error("⚠️ 未偵測到 Secrets 設定！請在 Streamlit Cloud 後台設定 GMAIL_USER 與 GMAIL_PASS。")

        # 原始資料檢視
        with st.expander("🔍 查看清冊原始資料"):
            st.dataframe(df, use_container_width=True)

    except Exception as e:
        st.error(f"系統處理發生異常：{e}")
else:
    st.info("👋 您好，請上傳科技執法清冊 (list2.csv) 以開啟自動化統計分析功能。")
