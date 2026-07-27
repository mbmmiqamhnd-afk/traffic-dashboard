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

def clean_text(text):
    """清除字串中的所有空白與換行，方便比對"""
    if not text:
        return ""
    return re.sub(r'\s+', '', str(text))

def extract_vacations_from_word(uploaded_vacations_list):
    vacations = {}
    if not uploaded_vacations_list:
        return vacations
        
    for uploaded_vacation in uploaded_vacations_list:
        try:
            doc = docx.Document(uploaded_vacation)
            
            # 從檔名自動抓取月份 (例如 "115-1月分局長輪休預排表範本" -> 抓取 "1")
            month_match = re.search(r'(\d+)月', uploaded_vacation.name)
            month = month_match.group(1) if month_match else "1"

            for table in doc.tables:
                col_names = {}
                start_row_idx = 0
                
                # 尋找表頭 (過濾掉空白，尋找 "分局長")
                for i, row in enumerate(table.rows):
                    cells_text = [clean_text(cell.text) for cell in row.cells]
                    if any("分局長" in c for c in cells_text):
                        for col_idx, text in enumerate(cells_text):
                            if text and "日期" not in text and "星期" not in text and "輪休" not in text:
                                col_names[col_idx] = text
                        start_row_idx = i + 1
                        break
                
                if not col_names:
                    continue
                
                # 從表頭的下一列開始讀取日期與記號
                for i in range(start_row_idx, len(table.rows)):
                    cells_text = [clean_text(cell.text) for cell in table.rows[i].cells]
                    
                    # 確保第一欄是日期數字
                    if not cells_text or not cells_text[0].isdigit():
                        continue
                        
                    day = int(cells_text[0])
                    date_val = f"{month}/{day}"
                    
                    for col_idx, officer_name in col_names.items():
                        if col_idx < len(cells_text):
                            mark = cells_text[col_idx]
                            # 如果格子裡有任何非空白字元 (例如 ●)，就視為休假
                            if mark and mark.strip() != "":
                                vacations.setdefault(officer_name, []).append(date_val)
                                
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
        title = clean_text(row.get("職稱", ""))
        name = clean_text(row.get("姓名", ""))
        
        # 寫入 Excel 時保留原有的顯示格式 (含空白)
        raw_title = str(row.get("職稱", "")).strip()
        raw_name = str(row.get("姓名", "")).strip()
        
        if raw_title or raw_name:
            ws.cell(row=current_row, column=1, value=raw_title)
            ws.cell(row=current_row, column=2, value=raw_name)
            
            person_dates = []
            user_vacations = []
            
            # 智慧對位邏輯
            for v_name, dates in vacation_dict.items():
                is_match = False
                # 1. 職稱完全相符 (例如 分局長)
                if title and title == v_name:
                    is_match = True
                # 2. 姓名第一個字出現在 Word 欄位名 (例如 "何" 出現在 "何副分局長")
                elif name and name[0] in v_name:
                    is_match = True
                # 3. 如果介面只有打職稱，沒有名字，確保不要抓錯人
                elif title == "分局長" and "副" not in v_name and "分局長" in v_name:
                    is_match = True
                    
                if is_match:
                    user_vacations.extend(dates)
            
            # 過濾休假日期
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
    st.info("請於下方上傳 **Excel 空白範本** 與 **各月份人員輪休預排表 (.docx)**，系統將自動排除休假人員的督導日期。")

    col1, col2, col3 = st.columns(3)
    with col1:
        uploaded_pdf = st.file_uploader("1. 上傳辦公日曆表 (PDF)", type=['pdf'], key="p25_pdf")
    with col2:
        uploaded_excel = st.file_uploader("2. 上傳 Excel 空白範本", type=['xlsx'], key="p25_excel")
    with col3:
        uploaded_vacations = st.file_uploader("3. 上傳各月份輪休預排表 (請上傳 .docx 格式，可複選)", type=['docx'], accept_multiple_files=True, key="p25_vac")

    default_personnel = pd.DataFrame([
        {"職稱": "分局長", "姓名": ""},
        {"職稱": "副分局長(一)", "姓名": ""},
        {"職稱": "副分局長(二)", "姓名": ""},
        {"職稱": "交通組組長", "姓名": ""}
    ])

    st.subheader("👥 督導人員名單設定 (可自由新增、刪除或修改職稱與姓名)")
    st.markdown("💡 **關鍵提醒**：為了讓系統成功辨識並排除副分局長的休假，請務必在下方「姓名」欄填寫長官的姓氏（例如：`何XX` 或 `蔡XX`）。")
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
            
    st.divider()

    st.subheader("📥 匯出與寄送督勤表")
    if uploaded_excel:
        try:
            vacation_dict = extract_vacations_from_word(uploaded_vacations)
            excel_data = process_excel(uploaded_excel, edited_personnel, full_date_list, 8, vacation_dict)
            file_name = f"龍潭分局_{year_str}上半年督導防制危險駕車勤務表.xlsx"

            if uploaded_vacations:
                st.success(f"✅ 已成功掛載 {len(uploaded_vacations)} 份輪休預排表！系統已自動剔除排定之休假日。")
                with st.expander("🔍 點此查看系統抓取到的「各長官休假清單」（若有缺漏請檢查上方姓名是否輸入正確）"):
                    st.write(vacation_dict)

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
