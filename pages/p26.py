import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Font
import io
import os
import smtplib
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# 匯入系統原本的側邊欄設定
try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        pass

def send_csv_email(file_bytes, file_name):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        
        msg["Subject"] = f"龍潭分局_行人及護老專案督勤表自動產出結果_{file_name}"
        
        body_text = (
            f"您好，\n\n"
            f"系統已自動產出「115年上半年行人及護老專案督勤表」，附件為對應的 Excel 統計檔案。\n\n"
            f"本信件由交通執法自動化分析引擎發送。"
        )
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(file_bytes.getvalue())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(file_name)}")
        msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
            
        return True, None
    except Exception as e:
        return False, str(e)

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
    st.info("此系統專為「行人及護老交通安全專案」設計。預設為每月排定 4 日，每日涵蓋 06-10 及 16-20 兩個時段（每時段固定 4 小時）。")

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

    # 固定每個時段為 4 小時
    hours_per_shift = 4

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
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                if st.button("📧 將此報表一鍵寄至我的信箱", use_container_width=True):
                    with st.spinner("信件發送中，請稍候…"):
                        ok, mail_err = send_csv_email(excel_data, file_name)
                        if ok:
                            st.success("✅ 信件發送成功！報表已隨信夾帶至您的信箱。")
                        else:
                            st.error(f"❌ 發信失敗: {mail_err}")

        except FileNotFoundError as fnf_err:
            st.error(f"⚠️ 錯誤：{fnf_err}。請確認 `376431843C_1150087037_ATTACH4.xlsx` 檔案是否存在於主目錄中。")
        except Exception as e:
            st.error(f"⚠️ 發生未知的錯誤：{e}")
    else:
        st.warning("請先於上方輸入督勤日期，方可匯出表單。")

if __name__ == "__main__":
    main()
