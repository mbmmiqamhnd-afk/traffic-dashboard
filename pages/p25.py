import streamlit as st
import pandas as pd
from datetime import timedelta
import openpyxl
from openpyxl.styles import Alignment
import io
import os

def generate_excel_file(combined_date_str):
    # 載入指定的上傳範本檔案
    file_path = '376431843C_1150087034_ATTACH3.xlsx'
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到範本檔案 {file_path}")

    wb = openpyxl.load_workbook(file_path)
    ws = wb['警力統計']

    # 直接寫入標題列(第2列)下方的 4 列儲存格：C3, C4, C5, C6
    for row in range(3, 7):
        cell = ws.cell(row=row, column=3)
        cell.value = combined_date_str
        # 開啟自動換行，讓長字串能在儲存格內完整顯示
        cell.alignment = Alignment(wrapText=True, vertical='center', horizontal='left')

    # 存入 BytesIO 以供 Streamlit 下載
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    st.title("🗓️ 115年上半年連假專案督勤表生成")
    st.info("此系統將自動運算 115 年上半年各週末與連續假期（含收假前一日），並匯出符合公文書格式之督勤日期表。")

    # 1. 設定假日區間 (包含所有國定連假與一般週末)
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

    # 2. 運算「起始日-1 至 結束日-1」的日期資料，並合併為單一字串
    date_list = []
    for h in holidays_115_H1:
        start_date = pd.to_datetime(h["start"])
        end_date = pd.to_datetime(h["end"])
        
        target_start = start_date - timedelta(days=1)
        target_end = end_date - timedelta(days=1)
        
        date_range = pd.date_range(start=target_start, end=target_end)
        for d in date_range:
            # 依據格式要求，時段轉換為帶有冒號的 22:00-06:00
            date_list.append(f"{d.month}/{d.day}(22:00-06:00)")

    # 將所有日期使用頓號「、」合併
    combined_date_str = "、".join(date_list)

    # 3. 顯示運算結果供使用者預覽
    st.subheader("📋 運算結果預覽")
    st.text_area("即將寫入 C 欄的日期清單：", value=combined_date_str, height=200)

    st.divider()

    # 4. 提供 Excel 下載按鈕
    st.subheader("📥 匯出督勤表")
    st.write("點擊下方按鈕，下載自動排版完成的 Excel 範本。")
    
    try:
        excel_data = generate_excel_file(combined_date_str)
        st.download_button(
            label="下載督勤日期表 (Excel)",
            data=excel_data,
            file_name="115年上半年督導防制危險駕車勤務表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except FileNotFoundError as fnf_err:
        st.error(f"⚠️ 錯誤：{fnf_err}。請確認 `376431843C_1150087034_ATTACH3.xlsx` 檔案是否存在於主目錄中。")
    except Exception as e:
        st.error(f"⚠️ 發生未知的錯誤：{e}")

if __name__ == "__main__":
    main()
