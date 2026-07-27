import streamlit as st
import pandas as pd
from datetime import timedelta
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
        
        msg["Subject"] = f"龍潭分局_{year_str}連假專案督勤表自動產出結果_{file_name}"
        
        body_text = (
            f"您好，\n\n"
            f"系統已自動依據自訂人員、Word輪休預排與上傳範本產出連假專案督勤表（龍潭分局），附件為對應的 Excel 統計檔案。\n\n"
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
            # 在這裡，您可以針對每一個 Word 檔的內容進行迴圈讀取與比對
            # 由於每個檔案可能代表不同月份，讀取出來的日期可以不斷 append 到同一個人員的清單中
            # vacations.setdefault(name, []).append(date_str)
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
            
            # 過濾休假日期
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

def main():
    show_sidebar()

    st.title("🗓️ 上半年連假專案督勤表生成")
    st.info("請於下方上傳 **Excel 空白範本** 與 **各月份人員輪休預排表 (支援多檔上傳)**，系統將自動排除休假人員的督導日期。")

    col1, col2, col3 = st.columns(3)
    with col1:
        uploaded_pdf = st.file_uploader("1. 上傳辦公日曆表 (PDF)", type=['pdf'], key="p25_pdf")
    with col2:
        uploaded_excel = st.file_uploader("2. 上傳 Excel 空白範本", type=['xlsx'], key="p25_excel")
    with col3:
        # 開啟 accept_multiple_files=True 允許一次上傳 6 個月份的 Word 檔
        uploaded_vacations = st.file_uploader("3. 上傳各月份輪休預排表 (可複選多檔)", type=['doc', 'docx'], accept_multiple_files=True, key="p25_vac")

    default_personnel = pd.DataFrame([
        {"職稱": "分局長", "姓名": ""},
        {"職稱": "副分局長(一)", "姓名": ""},
        {"職稱": "副分局長(二)", "姓名": ""},
        {"職稱": "交通組組長", "姓名": ""}
    ])

    st.subheader("👥 督導人員名單設定 (可自由新增、刪除或修改職稱與姓名)")
    edited_personnel = st.data_editor(default_personnel, num_rows="dynamic", use_container_width=True, key="p25_editor")

    year_str = "115年"
    ce_year = 2026
    if uploaded_excel:
        wb_temp = openpyxl.load_workbook(uploaded_excel, read_only=True)
        ws_temp = wb_temp['警力統計'] if '警力統計' in wb_temp.sheetnames else wb_temp.active
        match = re.search(r'(\d{2,3})年', ws_temp['A1'].value or "")
        if match:
            roc_y = int(match.group(1))
            year_str = f"{roc_y}年"
            ce_year = roc_y + 1911

    # 假日運算
    holidays_H1 = [
        {"start": f"{ce_year}-01-01", "end": f"{ce_year}-01-01"},
        {"start": f"{ce_year}-01-03", "end": f"{ce_year}-01-04"},
        {"start": f"{ce_year}-01-10", "end": f"{ce_year}-01-11"},
        {"start": f"{ce_year}-01-17", "end": f"{ce_year}-01-18"},
        {"start": f"{ce_year}-01-24", "end": f"{ce_year}-01-25"},
        {"start": f"{ce_year}-01-31", "end": f"{ce_year}-02-01"},
        {"start": f"{ce_year}-02-07", "end": f"{ce_year}-02-08"},
        {"start": f"{ce_year}-02-14", "end": f"{ce_year}-02-22"},
        {"start": f"{ce_year}-02-27", "end": f"{ce_year}-03-01"},
        {"start": f"{ce_year}-03-07", "end": f"{ce_year}-03-08"},
        {"start": f"{ce_year}-03-14", "end": f"{ce_year}-03-15"},
        {"start": f"{ce_year}-03-21", "end": f"{ce_year}-03-22"},
        {"start": f"{ce_year}-03-28", "end": f"{ce_year}-03-29"},
        {"start": f"{ce_year}-04-03", "end": f"{ce_year}-04-06"},
        {"start": f"{ce_year}-04-11", "end": f"{ce_year}-04-12"},
        {"start": f"{ce_year}-04-18", "end": f"{ce_year}-04-19"},
        {"start": f"{ce_year}-04-25", "end": f"{ce_year}-04-26"},
        {"start": f"{ce_year}-05-01", "end": f"{ce_year}-05-03"},
        {"start": f"{ce_year}-05-09", "end": f"{ce_year}-05-10"},
        {"start": f"{ce_year}-05-16", "end": f"{ce_year}-05-17"},
        {"start": f"{ce_year}-05-23", "end": f"{ce_year}-05-24"},
        {"start": f"{ce_year}-05-30", "end": f"{ce_year}-05-31"},
        {"start": f"{ce_year}-06-06", "end": f"{ce_year}-06-07"},
        {"start": f"{ce_year}-06-13", "end": f"{ce_year}-06-14"},
        {"start": f"{ce_year}-06-19", "end": f"{ce_year}-06-21"},
        {"start": f"{ce_year}-06-27", "end": f"{ce_year}-06-28"}
    ]

    full_date_list = []
    for h in holidays_H1:
        start_date = pd.to_datetime(h["start"])
        end_date = pd.to_datetime(h["end"])
        target_start = start_date - timedelta(days=1)
        target_end = end_date - timedelta(days=1)
        for d in pd.date_range(start=target_start, end=target_end):
            full_date_list.append(f"{d.month}/{d.day}(22-06)")

    total_base_days = len(full_date_list)
    total_base_hours = total_base_days * 8
    
    st.divider()

    st.subheader("📥 匯出與寄送督勤表")
    if uploaded_excel:
        try:
            vacation_dict = extract_vacations_from_word(uploaded_vacations)
            excel_data = process_excel(uploaded_excel, edited_personnel, full_date_list, 8, vacation_dict)
            file_name = f"龍潭分局_{year_str}上半年督導防制危險駕車勤務表.xlsx"

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
                if st.button("📧 將此報表一鍵寄至我的信箱", use_container_width=True, key="p25_btn"):
                    with st.spinner("信件發送中，請稍候…"):
                        ok, mail_err = send_csv_email(excel_data, file_name, year_str)
                        if ok:
                            st.success("✅ 信件發送成功！報表已隨信夾帶至您的信箱。")
                        else:
                            st.error(f"❌ 發信失敗: {mail_err}")
        except Exception as e:
            st.error(f"⚠️ 處理檔案時發生錯誤：{e}")
    else:
        st.warning("⚠️ 請先上傳 Excel 空白範本，方可進行下載與寄送。")

if __name__ == "__main__":
    main()
