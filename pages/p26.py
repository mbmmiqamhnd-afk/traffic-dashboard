import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Font
import io
import os

# 匯入系統原本的側邊欄設定
try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        pass

def generate_excel_file(combined_date_str, total_hours):
    # 載入指定的行人及護老專案上傳範本檔案
    file_path = '376431843C_1150087037_ATTACH4.xlsx'
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到範本檔案 {file_path}")

    wb = openpyxl.load_workbook(file_path)
    ws = wb['警力統計']

    # 直接寫入標題列(第2列)下方的 4 列儲存格：第 3, 4, 5, 6 列
    for row in range(3, 7):
        # C 欄：寫入合併後的日期清單
        cell_c = ws.cell(row=row, column=3)
        cell_c.value = combined_date_str
        cell_c.alignment = Alignment(wrapText=True, vertical='center', horizontal='left')
        
        # D 欄：寫入合計時數
        cell_d = ws.cell(row=row, column=4)
        cell_d.value = total_hours
        cell_d.alignment = Alignment(vertical='center', horizontal='center')
        cell_d.font = Font(name='微軟正黑體', size=12)

    # 存入 BytesIO 以供 Streamlit 下載
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    # 載入系統專屬側邊欄
    show_sidebar()

    st.title("👵 行人及護老專案督勤表生成")
    st.info("此系統專為「行人及護老交通安全專案」設計。請直接貼上排定的督勤日期字串，系統將自動為您結算總時數並匯出報表。")

    # 1. 提供彈性的輸入介面
    st.subheader("📝 勤務日期與時段設定")
    
    # 預設提供一個範例字串，您可以直接刪除並貼上新的
    default_str = "1/5(08-10)、1/12(08-10)、1/19(16-18)"
    combined_date_str = st.text_area(
        "請貼上或輸入督勤日期清單 (請以頓號「、」分隔)：", 
        value=default_str, 
        height=100
    )

    # 讓使用者彈性選擇該專案單次勤務的時數 (通常護老勤務為 2 小時或 4 小時)
    hours_per_shift = st.number_input("單次勤務時數 (小時)：", min_value=1, max_value=12, value=2, step=1)

    # 2. 自動計算天數與時數
    # 透過計算頓號「、」的數量來反推總天數 (空字串防呆)
    if combined_date_str.strip() == "":
        total_days = 0
    else:
        total_days = len(combined_date_str.split("、"))
        
    total_hours = total_days * hours_per_shift

    # 3. 顯示運算結果供使用者預覽
    st.subheader("📋 運算結果預覽")
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="合計督勤次數", value=f"{total_days} 次")
    with col2:
        st.metric(label="總計時數 (D欄輸出值)", value=f"{total_hours} 小時")

    st.divider()

    # 4. 提供 Excel 下載按鈕
    st.subheader("📥 匯出督勤表")
    
    # 只有當有輸入資料時才允許下載
    if total_days > 0:
        try:
            excel_data = generate_excel_file(combined_date_str, total_hours)
            st.download_button(
                label="下載督勤日期表 (Excel)",
                data=excel_data,
                file_name="115年上半年督導行人及護老交通安全勤務表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except FileNotFoundError as fnf_err:
            st.error(f"⚠️ 錯誤：{fnf_err}。請確認 `376431843C_1150087037_ATTACH4.xlsx` 檔案是否存在於主目錄中。")
        except Exception as e:
            st.error(f"⚠️ 發生未知的錯誤：{e}")
    else:
        st.warning("請先於上方輸入督勤日期，方可匯出表單。")

if __name__ == "__main__":
    main()
