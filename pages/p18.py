import streamlit as st
import pandas as pd
import io
import sys
import os
import re
import smtplib
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# 自動將上層目錄加入路徑
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
try:
    from app import show_sidebar
except ImportError:
    def show_sidebar():
        pass 

def send_report_email_auto(file_data, filename, year, month):
    """
    自動讀取 Secrets 設定並發送郵件
    """
    try:
        # 從 st.secrets 自動讀取帳號密碼
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = sender  # 自動寄給自己
        msg['Subject'] = f"【系統自動發送】龍潭分局 {year}年{month}月 獎勵金點數統計表"
        
        body = f"郭同仁您好：\n\n系統已自動完成 {year}年{month}月份的獎勵金點數彙整。\n附件為最新產出的統計報表，請查收。"
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # 加入附件 (處理中文檔名)
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename*=UTF-8''{_ul.quote(filename)}")
        msg.attach(part)
        
        # 使用 SSL 465 端口寄送 (與聯合稽查系統一致)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

def p18_page():
    show_sidebar()

    st.title("💰 龍潭分局 - 獎勵金點數統計表產生器")
    st.info("系統將自動修正小計與權重，並在完成後自動同步備份至您的電子信箱。")

    # 1. 點數權重設定
    with st.expander("⚙️ 點數權重設定", expanded=False):
        col1, col2, col3 = st.columns(3)
        p_a2 = col1.number_input("A2 點數/件", value=10.0, step=1.0)
        p_a3 = col2.number_input("A3 點數/件", value=5.0, step=1.0)
        p_traf = col3.number_input("交整點數/小時", value=5.0, step=1.0)

    # 2. 檔案上傳
    st.subheader("📂 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)", type=['xlsx'], accept_multiple_files=True)

    # --- 移除了「請輸入接收郵件地址」的選項 ---

    if st.button("🚀 執行彙整與自動同步信箱", type="primary", use_container_width=True):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種資料皆已完成上傳！")
            return

        with st.spinner("正在重新計算數據並發送電子郵件..."):
            try:
                # --- [數據處理邏輯：包含修正小計與強制重算] ---
                dfs_raw = pd.read_excel(file_template, sheet_name=None, header=None)
                
                # (中間處理 A2/A3、交整時數、小計加總的邏輯維持不變...)
                # 這裡為了簡化顯示省略重複的 Pandas 處理代碼
                # ...
                
                # 假設處理完成後得到的變數為 excel_data 與 final_filename
                # 以及偵測到的年份 ext_year, ext_month
                
                # --- 執行自動寄信 ---
                ok, err = send_report_email_auto(excel_data, final_filename, ext_year, ext_month)
                
                if ok:
                    st.success(f"✅ 報表彙整成功！已自動發送至您的信箱備存。")
                else:
                    st.warning(f"⚠️ 報表已產出，但郵件發送失敗: {err}")

                st.download_button(label="📥 下載統計表到電腦", data=excel_data, file_name=final_filename, use_container_width=True)

            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")

if __name__ == "__main__":
    p18_page()
