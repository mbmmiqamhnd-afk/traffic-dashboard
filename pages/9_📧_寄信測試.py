import streamlit as st
import smtplib
from email.mime.text import MIMEText

st.title("🕵️‍♂️ 寄信功能診斷室")

# 1. 檢查 Secrets 是否讀取成功
st.write("### 1. 檢查 Secrets 設定")
if "email" in st.secrets:
    user = st.secrets["email"]["user"]
    # 只顯示前幾個字，確保有讀到
    masked_user = user[:3] + "***" + user.split('@')[-1]
    st.success(f"✅ 成功讀取設定檔！使用者: {masked_user}")
else:
    st.error("❌ 讀取失敗！找不到 [email] 區塊，請檢查 Secrets 格式。")
    st.stop()

# 2. 測試發送
st.write("### 2. 連線測試")
receiver = st.text_input("請輸入您的收件信箱 (建議與寄件者相同)", value=st.secrets["email"]["user"])

if st.button("🚀 發射測試信"):
    status_area = st.empty()
    
    try:
        sender = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        
        msg = MIMEText("恭喜！如果您看到這封信，代表 Streamlit 寄信功能完全正常。")
        msg['Subject'] = "Streamlit 連線測試成功 (自動發送)"
        msg['From'] = sender
        msg['To'] = receiver

        # 詳細連線步驟 (讓您看到卡在哪一步)
        status_area.info("⏳ 1/4 正在連線至 smtp.gmail.com:587 ...")
        server = smtplib.SMTP('smtp.gmail.com', 587)
        
        status_area.info("⏳ 2/4 正在啟動 TLS 加密 ...")
        server.starttls()
        
        status_area.info("⏳ 3/4 正在登入 Google 帳號 ...")
        server.login(sender, password)
        
        status_area.info("⏳ 4/4 正在發送郵件 ...")
        server.sendmail(sender, receiver, msg.as_string())
        server.quit()
        
        status_area.success("🎉 發送成功！請檢查收件匣 (或是垃圾郵件)。")
        st.balloons()

    except Exception as e:
        # 這是最重要的部分：顯示錯誤代碼
        status_area.error("❌ 發送失敗！請截圖下方的錯誤訊息：")
        st.code(str(e))
        
        # 常見錯誤翻譯
        err_msg = str(e)
        if "Username and Password not accepted" in err_msg:
            st.warning("💡 原因分析：應用程式密碼錯誤，或是 Secrets 裡的帳號打錯字。")
        elif "Please log in via your web browser" in err_msg:
            st.warning("💡 原因分析：Google 擋住了連線，請確認兩步驟驗證是否開啟。")
        elif "not define" in err_msg:
            st.warning("💡 原因分析：程式碼變數寫錯了。")
