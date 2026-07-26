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
    st.info("此系統專為「行人及護老交通安全專案」設計。預設為每月排定 4 日，每日涵蓋 06-10 及 16-20 兩個時段。")

    st.subheader("📝 勤務日期與時段設定")
    
    # 預設自動產生上半年 (1~6月)，每月 4 天 (暫以 5, 12, 19, 26 日為例)，分為 06-10 與 16-20 兩時段
    default_dates = []
    for m in range(1, 7):
        for d in [5, 12, 19, 26]: # 預設每月4日的佔位日期，可於介面中直接修改
            default_dates.append(f"{m}/{d}(06-10)")
            default_dates.append(f"{m}/{d}(16-20)")
            
    default_str = "、".join(default_dates)

    # 提供文字框讓使用者核對與修改日期
    combined_date_str = st.text_area(
        "請確認或修改督勤日期清單 (各時段請以頓號「、」分隔)：", 
        value=default_str, 
        height=250
    )

    # 每個時段為 4 小時 (06-10 或 16-20)
    hours_per_shift = st.number_input("單一時段勤務時數 (小時)：", min_value=1, max_value=12, value=4, step=1)

    # 自動計算班次、天數與總時數
    if combined_date_str.strip() == "":
        total_shifts = 0
    else:
        total_shifts = len(combined_date_str.split("、"))
        
    total_hours = total_shifts * hours_per_shift
    total_days = total_shifts // 2 # 每日2班，換算為天數供參考

    st.subheader("📋 運算結果預覽")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="合計排定天數", value=f"{total_days} 天")
    with col2:
        st.metric(label="合計督勤班次", value=f"{total_shifts} 班")
    with col3:
        st.metric(label="總計時數 (D欄輸出值)", value=f"{total_hours} 小時")

    st.divider()

    st.subheader("📥 匯出督勤表")
    
    if total_shifts > 0:
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
