import streamlit as st
import pandas as pd
from datetime import timedelta
import openpyxl
from openpyxl.styles import Alignment, Font
import io
import os
import smtplib
from email.message import EmailMessage

# 匯入系統原本的側邊欄設定
try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        pass

def send_email_with_attachment(recipient_email, file_bytes, file_name):
    try:
        sender_email = st.secrets["email"]["sender"]
        sender_password = st.secrets["email"]["password"]

        msg = EmailMessage()
        msg['Subject'] = f"督勤表自動產出結果：{file_name}"
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg.set_content("您好，\n\n系統已完成督勤表產出，請查收附件。\n\n交通執法自動化分析引擎 敬上")

        msg.add_attachment(
            file_bytes.getvalue(), 
            maintype='application', 
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
            filename=file_name
        )

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
            
        return True
    except KeyError:
        st.error("⚠️ 錯誤：找不到 Streamlit Secrets 中的信箱設定。請確認 `.streamlit/secrets.toml` 已正確配置。")
        return False
    except Exception as e:
        st.error(f"⚠️ 寄件失敗：{e}")
        return False

def generate_excel_file(combined_date_str, total_hours):
    file_path = '376431843C_1150087034_ATTACH3.xlsx'
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到範本檔案 {file_path}")

    wb = openpyxl.load_workbook(file_path)
    ws = wb['警力統計']

    for row in range(3, 7):
        cell_c = ws.cell(row=row, column=3)
        cell_c.value = combined_date_str
        cell_c.alignment = Alignment(wrapText=True, vertical='center', horizontal='left')
        
        cell_d = ws.cell(row=row, column=4)
        cell_d.value = total_hours
        cell_d.alignment = Alignment(vertical='center', horizontal='center')
        cell_d.font = Font(name='微軟正黑體', size=12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    show_sidebar()

    st.title("🗓️ 115年上半年連假專案督勤表生成")
    st.info("此系統將自動運算 115 年上半年各週末與連續假期（含收假前一日），並匯出符合公文書格式之督勤日期表。")

    holidays_115_H1 = [
        {"name": "元旦", "start": "2026-01-01", "end": "2026-01-01"},
        {"name": "一般假日", "start": "2026-01-03", "end": "2026-01-04"},
        {"name": "一般假日", "start": "2026-01-10", "end": "2026-01-11"},
        {"name": "一般假日", "start": "2026-01-17", "end": "2026-01-18"},
        {"name": "一般假日", "start": "2026-01-24", "end": "2026-01-25"},
        {"name": "一般假日", "start": "2026-01-31", "end": "2026-02-01"},
        {"name": "一般假日", "start": "2026-02-07", "end": "2026-02-08"},
        {"name": "農曆春節", "start": "2026-02-14", "end": "2026-02-22"},
        {"name": "228和平紀念日", "start": "2026-02-27", "end": "2026-03-01"},
        {"name": "一般假日", "start": "2026-03-07", "end": "2026-03-08"},
        {"name": "一般假日", "start": "2026-03-14", "end": "2026-03-15"},
        {"name": "一般假日", "start": "2026-03-21", "end": "2026-03-22"},
        {"name": "一般假日", "start": "2026-03-28", "end": "2026-03-29"},
        {"name": "兒童節與清明節", "start": "2026-04-03", "end": "2026-04-06"},
        {"name": "一般假日", "start": "2026-04-11", "end": "2026-04-12"},
        {"name": "一般假日", "start": "2026-04-18", "end": "2026-04-19"},
        {"name": "一般假日", "start": "2026-04-25", "end": "2026-04-26"},
        {"name": "勞動節", "start": "2026-05-01", "end": "2026-05-03"},
        {"name": "一般假日", "start": "2026-05-09", "end": "2026-05-10"},
        {"name": "一般假日", "start": "2026-05-16", "end": "2026-05-17"},
        {"name": "一般假日", "start": "2026-05-23", "end": "2026-05-24"},
        {"name": "一般假日", "start": "2026-05-30", "end": "2026-05-31"},
        {"name": "一般假日", "start": "2026-06-06", "end": "2026-06-07"},
        {"name": "一般假日", "start": "2026-06-13", "end": "2026-06-14"},
        {"name": "端午節", "start": "2026-06-19", "end": "2026-06-21"},
        {"name": "一般假日", "start": "2026-06-27", "end": "2026-06-28"}
    ]

    date_list = []
    for h in holidays_115_H1:
        start_date = pd.to_datetime(h["start"])
        end_date = pd.to_datetime(h["end"])
        
        target_start = start_date - timedelta(days=1)
        target_end = end_date - timedelta(days=1)
        
        date_range = pd.date_range(start=target_start, end=target_end)
        for d in date_range:
            date_list.append(f"{d.month}/{d.day}(22-06)")

    combined_date_str = "、".join(date_list)
    total_days = len(date_list)
    total_hours = total_days * 8

    st.subheader("📋 運算結果預覽")
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="合計督勤天數", value=f"{total_days} 天")
    with col2:
        st.metric(label="總計時數 (D欄輸出值)", value=f"{total_hours} 小時")
    st.text_area("即將寫入 C 欄的日期清單：", value=combined_date_str, height=200)

    st.divider()

    st.subheader("📥 匯出與寄送督勤表")
    
    try:
        excel_data = generate_excel_file(combined_date_str, total_hours)
        file_name = "115年上半年督導防制危險駕車勤務表.xlsx"

        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button(
                label="📥 下載督勤日期表 (Excel)",
                data=excel_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            target_email = st.text_input("✉️ 將報表寄至信箱 (p25)：", value="your_email@example.com")
            if st.button("🚀 立即寄送 (p25)"):
                if target_email:
                    with st.spinner("信件寄送中，請稍候..."):
                        if send_email_with_attachment(target_email, excel_data, file_name):
                            st.success("✅ 信件已成功寄出！請至信箱確認。")
                else:
                    st.warning("⚠️ 請輸入收件人信箱！")

    except FileNotFoundError as fnf_err:
        st.error(f"⚠️ 錯誤：{fnf_err}。請確認 `376431843C_1150087034_ATTACH3.xlsx` 檔案是否存在於主目錄中。")
    except Exception as e:
        st.error(f"⚠️ 發生未知的錯誤：{e}")

if __name__ == "__main__":
    main()
