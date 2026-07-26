import streamlit as st
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
    file_path = '376431843C_1150087037_ATTACH4.xlsx'
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

    st.title("👵 行人及護老專案督勤表生成")
    st.info("此系統專為「行人及護老交通安全專案」設計。預設為每月排定 4 日，每日涵蓋 06-10 及 16-20 兩個時段。")

    st.subheader("📝 勤務日期與時段設定")
    
    default_dates = []
    for m in range(1, 7):
        for d in [5, 12, 19, 26]:
            default_dates.append(f"{m}/{d}(06-10)")
            default_dates.append(f"{m}/{d}(16-20)")
            
    default_str = "、".join(default_dates)

    combined_date_str = st.text_area(
        "請確認或修改督勤日期清單 (各時段請以頓號「、」分隔)：", 
        value=default_str, 
        height=250
    )

    hours_per_shift = st.number_input("單一時段勤務時數 (小時)：", min_value=1, max_value=12, value=4, step=1)

    if combined_date_str.strip() == "":
        total_shifts = 0
    else:
        total_shifts = len(combined_date_str.split("、"))
        
    total_hours = total_shifts * hours_per_shift
    total_days = total_shifts // 2

    st.subheader("📋 運算結果預覽")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="合計排定天數", value=f"{total_days} 天")
    with col2:
        st.metric(label="合計督勤班次", value=f"{total_shifts} 班")
    with col3:
        st.metric(label="總計時數 (D欄輸出值)", value=f"{total_hours} 小時")

    st.divider()

    st.subheader("📥 匯出與寄送督勤表")
    
    if total_shifts > 0:
        try:
            excel_data = generate_excel_file(combined_date_str, total_hours)
            file_name = "115年上半年督導行人及護老交通安全勤務表.xlsx"

            col1, col2 = st.columns(2)
            
            with col1:
                st.download_button(
                    label="📥 下載督勤日期表 (Excel)",
                    data=excel_data,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                target_email = st.text_input("✉️ 將報表寄至信箱 (p26)：", value="your_email@example.com")
                if st.button("🚀 立即寄送 (p26)"):
                    if target_email:
                        with st.spinner("信件寄送中，請稍候..."):
                            if send_email_with_attachment(target_email, excel_data, file_name):
                                st.success("✅ 信件已成功寄出！請至信箱確認。")
                    else:
                        st.warning("⚠️ 請輸入收件人信箱！")

        except FileNotFoundError as fnf_err:
            st.error(f"⚠️ 錯誤：{fnf_err}。請確認 `376431843C_1150087037_ATTACH4.xlsx` 檔案是否存在於主目錄中。")
        except Exception as e:
            st.error(f"⚠️ 發生未知的錯誤：{e}")
    else:
        st.warning("請先於上方輸入督勤日期，方可匯出表單。")

if __name__ == "__main__":
    main()
