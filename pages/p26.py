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
import pdfplumber  # 處理 PDF 檔案

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
            f"系統已自動依據自訂人員、輪休預排與上傳範本產出行人及護老專案督勤表（龍潭分局），附件為對應的 Excel 統計檔案。\n\n"
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
    if not text:
        return ""
    return re.sub(r'\s+', '', str(text))

def extract_vacations_from_files(uploaded_vacations_list):
    vacations = {}
    if not uploaded_vacations_list:
        return vacations
        
    for uploaded_file in uploaded_vacations_list:
        try:
            month = "1"
            month_match = re.search(r'(\d+)月', uploaded_file.name)
            if month_match:
                month = str(int(month_match.group(1)))
                
            ext = uploaded_file.name.lower().split('.')[-1]

            # ========== PDF 解析引擎 ==========
            if ext == "pdf":
                with pdfplumber.open(uploaded_file) as pdf:
                    for page in pdf.pages:
                        text = page.extract_text() or ""
                        m_match = re.search(r'年\s*(\d+)\s*月', text)
                        if m_match:
                            month = str(int(m_match.group(1)))
                            
                        tables = page.extract_tables()
                        for table in tables:
                            col_names = {}
                            start_row_idx = 0
                            
                            for i, row in enumerate(table):
                                cells_text = [clean_text(c) if c else "" for c in row]
                                if any("分局長" in c for c in cells_text):
                                    for col_idx, txt in enumerate(cells_text):
                                        if txt and "日期" not in txt and "星期" not in txt and "輪休" not in txt:
                                            col_names[col_idx] = txt
                                    start_row_idx = i + 1
                                    break
                            
                            if not col_names: continue
                            
                            for i in range(start_row_idx, len(table)):
                                cells_text = [clean_text(c) if c else "" for c in table[i]]
                                if not cells_text or not cells_text[0].isdigit():
                                    continue
                                day = int(cells_text[0])
                                date_val = f"{month}/{day}"
                                for col_idx, officer_name in col_names.items():
                                    if col_idx < len(cells_text):
                                        mark = cells_text[col_idx]
                                        if mark and mark.strip() != "":
                                            vacations.setdefault(officer_name, []).append(date_val)

            # ========== Word 解析引擎 ==========
            elif ext in ["docx", "doc"]:
                doc = docx.Document(uploaded_file)
                if doc.paragraphs:
                    m_match = re.search(r'年\s*(\d+)\s*月', doc.paragraphs[0].text)
                    if m_match:
                        month = str(int(m_match.group(1)))
                        
                for table in doc.tables:
                    col_names = {}
                    start_row_idx = 0
                    for i, row in enumerate(table.rows):
                        cells_text = [clean_text(cell.text) for cell in row.cells]
                        if any("分局長" in c for c in cells_text):
                            for col_idx, txt in enumerate(cells_text):
                                if txt and "日期" not in txt and "星期" not in txt and "輪休" not in txt:
                                    col_names[col_idx] = txt
                            start_row_idx = i + 1
                            break
                    
                    if not col_names: continue
                    
                    for i in range(start_row_idx, len(table.rows)):
                        cells_text = [clean_text(cell.text) for cell in table.rows[i].cells]
                        if not cells_text or not cells_text[0].isdigit():
                            continue
                        day = int(cells_text[0])
                        date_val = f"{month}/{day}"
                        for col_idx, officer_name in col_names.items():
                            if col_idx < len(cells_text):
                                mark = cells_text[col_idx]
                                if mark and mark.strip() != "":
                                    vacations.setdefault(officer_name, []).append(date_val)
                                    
        except Exception as e:
            st.warning(f"休假表「{uploaded_file.name}」解析失敗：{e}")
            
    for k in vacations:
        vacations[k] = list(set(vacations[k]))
        
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
        
        raw_title = str(row.get("職稱", "")).strip()
        raw_name = str(row.get("姓名", "")).strip()
        
        if raw_title or raw_name:
            ws.cell(row=current_row, column=1, value=raw_title)
            ws.cell(row=current_row, column=2, value=raw_name)
            
            person_dates = []
            user_vacations = []
            
            for v_name, dates in vacation_dict.items():
                is_match = False
                if title and title == v_name:
                    is_match = True
                elif name and name[0] in v_name:
                    is_match = True
                elif title == "分局長" and "副" not in v_name and "分局長" in v_name:
                    is_match = True
                    
                if is_match:
                    user_vacations.extend(dates)
            
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
    st.info("請於下方上傳 **Excel 空白範本** 與 **各月份輪休預排表 (支援 PDF/Word)**，系統將自動排除休假人員。")

    col1, col2 = st.columns(2)
    with col1:
        uploaded_excel = st.file_uploader("1. 上傳 Excel 空白範本", type=['xlsx'], key="p26_excel")
    with col2:
        # 已開放支援 PDF 格式上傳！
        uploaded_vacations = st.file_uploader("2. 上傳各月份輪休預排表 (可複選多檔)", type=['pdf', 'docx', 'doc'], accept_multiple_files=True, key="p26_vac")

    default_personnel = pd.DataFrame([
        {"職稱": "分局長", "姓名": ""},
        {"職稱": "副分局長(一)", "姓名": ""},
        {"職稱": "副分局長(二)", "姓名": ""},
        {"職稱": "交通組組長", "姓名": ""}
    ])

    st.subheader("👥 督導人員名單設定")
    st.markdown("💡 **提醒**：為了精準對位，請在下方「姓名」欄填寫長官的姓氏（例如：`何XX` 或 `蔡XX`）。")
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
            vacation_dict = extract_vacations_from_files(uploaded_vacations)
            excel_data = process_excel(uploaded_excel, edited_personnel, full_date_list, hours_per_shift, vacation_dict)
            file_name = f"龍潭分局_{year_str}上半年督導行人及護老交通安全勤務表.xlsx"

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
