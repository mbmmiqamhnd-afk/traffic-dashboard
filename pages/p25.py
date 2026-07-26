import streamlit as st
import pandas as pd
from datetime import timedelta
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font
import io

def generate_excel_file(records):
    # 載入指定的上傳範本檔案
    # 確保 376431843C_1150087034_ATTACH3.xlsx 與 p25.py 在可被相對路徑存取的位置，
    # 若您的 app.py 在外層目錄，這裡的路徑通常寫相對於 app.py 的位置。
    wb = openpyxl.load_workbook('376431843C_1150087034_ATTACH3.xlsx')
    ws = wb['警力統計']

    # 解除注意事項的合併儲存格
    ws.unmerge_cells('A11:A15')
    ws.unmerge_cells('B11:D15')
    ws.unmerge_cells('A16:D17')
    ws.unmerge_cells('A18:D20')

    # 插入 15 列，擴充表格空間
    ws.insert_rows(11, 15)

    # 重新合併注意事項 (往下位移 15 列)
    ws.merge_cells('A26:A30')
    ws.merge_cells('B26:D30')
    ws.merge_cells('A31:D32')
    ws.merge_cells('A33:D35')

    # 設定邊框與字體
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))

    # 開始寫入資料 (從 A4 開始)
    current_row = 4
    for r in records:
        ws.cell(row=current_row, column=1, value=r["name"])
        ws.cell(row=current_row, column=2, value="")
        ws.cell(row=current_row, column=3, value=r["date"])
        ws.cell(row=current_row, column=4, value=4)
        
        for col in range(1, 5):
            cell = ws.cell(row=current_row, column=col)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(name='微軟正黑體', size=12)
            
        current_row += 1

    # 存入 BytesIO 以供 Streamlit 下載
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    st.title("🗓️ 115年上半年連假專案督勤表生成")
    st.info("此系統將自動運算 115 年上半年各連續假期（含收假前一日），並匯出符合公文書格式之督勤日期表。")

    # 1. 設定假日區間
    holidays_115_H1 = [
        {"name": "元旦", "start": "2026-01-01", "end": "2026-01-01"},
        {"name": "農曆春節", "start": "2026-02-14", "end": "2026-02-22"},
        {"name": "228和平紀念日", "start": "2026-02-27", "end": "2026-03-01"},
        {"name": "兒童節與清明節", "start": "2026-04-03", "end": "2026-04-06"},
        {"name": "勞動節", "start": "2026-05-01", "end": "2026-05-03"},
        {"name": "端午節", "start": "2026-06-19", "end": "2026-06-21"}
    ]

    # 2. 運算「起始日-1 至 結束日-1」的日期資料
    records = []
    for h in holidays_115_H1:
        start_date = pd.to_datetime(h["start"])
        end_date = pd.to_datetime(h["end"])
        
        target_start = start_date - timedelta(days=1)
        target_end = end_date - timedelta(days=1)
        
        date_range = pd.date_range(start=target_start, end=target_end)
        for d in date_range:
            formatted_date = f"{d.month}/{d.day}(08:00-12:00)"
            records.append({
                "name": h["name"],
                "date": formatted_date
            })

    # 3. 顯示運算結果供使用者預覽
    st.subheader("📋 運算結果預覽")
    df_display = pd.DataFrame([{"專案名稱": r["name"], "督勤日期與時段": r["date"]} for r in records])
    st.dataframe(df_display, use_container_width=True)

    st.divider()

    # 4. 提供 Excel 下載按鈕
    st.subheader("📥 匯出督勤表")
    st.write("點擊下方按鈕，下載自動排版完成的 Excel 範本。")
    
    try:
        excel_data = generate_excel_file(records)
        st.download_button(
            label="下載督勤日期表 (Excel)",
            data=excel_data,
            file_name="115年上半年督導防制危險駕車勤務表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except FileNotFoundError:
        st.error("⚠️ 錯誤：找不到範本檔案 `376431843C_1150087034_ATTACH3.xlsx`。請確認檔案是否存在於主目錄中。")
    except Exception as e:
        st.error(f"⚠️ 發生未知的錯誤：{e}")

if __name__ == "__main__":
    main()
