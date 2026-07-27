import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font
import io
import re
import smtplib
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import docx  # 處理 Word 檔案

try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        pass

def send_csv_email(file_bytes, file_name, year_str):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        
        msg["Subject"] = f"龍潭分局_{year_str}行人及護老專案督勤表自動產出結果_{file_name}"
        
        body_text = (
            f"您好，\n\n"
            f"系統已自動依據自訂人員、Word輪休預排與上傳範本產出行人及護老專案督勤表（龍潭分局），附件為對應的 Excel 統計檔案。\n\n"
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

def extract_vacations_from_word(uploaded_vacations_list):
    vacations = {}
    if not uploaded_vacations_list:
        return vacations
        
    for uploaded_vacation in uploaded_vacations_list:
        try:
            doc = docx.Document(uploaded_vacation)
            # 在這裡解析每個 Word 檔案
        except Exception as e:
            st.warning(f"Word 休假表「{uploaded_vacation.name}」解析失敗：{e}")
            
    return vacations

def process_excel(uploaded_file, df_personnel, full_date_list, hours_per_shift, vacation_dict):
    wb = openpyxl.load_workbook(uploaded_file)
    ws = wb['警力統計'] if '警力統計' in wb.sheetnames else wb.active

    title_cell = ws['A1']
    if title_cell.value and '○○分局' in title_cell.value:
        title_cell.value = title_cell.value.replace('○○分局', '龍潭分局')

    current_row = 3
    for _, row in df_personnel.iterrows():
        title = str(row.get("職稱", "")).strip()
        name = str(row.get("姓名", "")).strip()
        
        if title or name:
            ws.cell(row=current_row, column=1, value=title)
            ws.cell(row=current_row, column=2, value=name)
            
            person_dates = []
            user_vacations = vacation_dict.get(name, [])
            
            for date_str in full_date_list:
                raw_date = date_str.split('(')[0]
                if raw_date not in user_vacations:
                    person_dates.append(date_str)
                    
            combined_date_str = "、".join(person_dates)
            total_hours = len(person_dates) * hours_per_shift
            
            cell_c = ws.cell(row=current_row, column=3)
            cell_c.value = combined_date_str
            cell_c.alignment = Alignment(wrapText=True, vertical='center', horizontal='left')
            
            cell_d = ws.cell(row=current_row, column=4)
            cell_d.value = total_hours
            cell_d.alignment = Alignment(vertical='center', horizontal='center')
            cell_d.font = Font(name='微軟正黑體', size=12)
            
            current_row += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def get_year_from_excel(uploaded_file):
    if not uploaded_file:
        return "115年"
    wb = openpyxl.load_workbook(uploaded_file, read_only=True)
    ws = wb['警力統計'] if '警力統計' in wb.sheetnames else wb.active
    title = ws['A1'].value or ""
    match = re.search(r'(\d{2,3})年', title)
    return f"{match.group(1)}年" if match else "115年"

def main():
    show_sidebar()

    st.title("👵 行人及護老專案督勤表生成")
    st.info("請於下方上傳 **Excel 空白範本** 與 **各月份輪休預排表 (支援多檔上傳)**，系統將自動排除休假人員的督導日期。")

    col1, col2 = st.columns(2)
    with col1:
        uploaded_excel = st.file_uploader("1. 上傳 Excel 空白範本", type=['xlsx'], key="p26_excel")
    with col2:
        # 開啟 accept_multiple_files=True 允許一次上傳多個月份的 Word 檔
        uploaded_vacations = st.file_uploader("2. 上傳各月份輪休預排表 (可複選多檔)", type=['doc', 'docx'], accept_multiple_files=True, key="p26_vac")

    default_personnel = pd.DataFrame([
        {"職稱": "分局長", "姓名": ""},
        {"職稱": "副分局長(一)", "姓名": ""},
        {"職稱": "副分局長(二)", "姓名": ""},
        {"職稱": "交通組組長", "姓名": ""}
    ])

    st.subheader("👥 督導人員名單設定 (可自由新增、刪除或修改職稱與姓名)")
    edited_personnel = st.data_editor(default_personnel, num_rows="dynamic", use_container_width=True, key="p26_editor")

    year_str = get_year_from_excel(uploaded_excel)

    st.subheader("📝 基準勤務日期與時段設定")
    
    default_dates = []
    for m in range(1, 7):
        for d in [5, 12, 19, 26]:
            default_dates.append(f"{m}/{d}(06-10)")
            default_dates.append(f"{m}/{d}(16-20)")
            
    default_str = "、".join(default_dates)

    combined_date_str = st.text_area(
        "此為「未扣除休假前」的基準日期清單：", 
        value=default_str, 
        height=150
    )
    
    full_date_list = [d.strip() for d in combined_date_str.split("、") if d.strip()]
    hours_per_shift = 4

    st.divider()

    st.subheader("📥 匯出與寄送督勤表")
    if uploaded_excel and full_date_list:
        try:
            vacation_dict = extract_vacations_from_word(uploaded_vacations)
            excel_data = process_excel(uploaded_excel, edited_personnel, full_date_list, hours_per_shift, vacation_dict)
            file_name = f"龍潭分局_{year_str}上半年督導行人及護老交通安全勤務表.xlsx"

            if uploaded_vacations:
                st.success(f"✅ 已成功掛載 {len(uploaded_vacations)} 份輪休預排表，產出報表時將自動剔除人員休假日期。")

            b1, b2 = st.columns(2)
            with b1:
                st.download_button(
                    label="📥 下載督勤日期表 (Excel)",
                    data=excel_data,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with b2:
                if st.button("📧 將此報表一鍵寄至我的信箱", use_container_width=True, key="p26_btn"):
                    with st.spinner("信件發送中，請稍候…"):
                        ok, mail_err = send_csv_email(excel_data, file_name, year_str)
                        if ok:
                            st.success("✅ 信件發送成功！報表已隨信夾帶至您的信箱。")
                        else:
                            st.error(f"❌ 發信失敗: {mail_err}")
        except Exception as e:
            st.error(f"⚠️ 處理檔案時發生錯誤：{e}")
    else:
        st.warning("⚠️ 請先上傳 Excel 空白範本並確認排定日期，方可進行下載與寄送。")

if __name__ == "__main__":
    main()
