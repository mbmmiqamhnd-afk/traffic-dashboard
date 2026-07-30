import streamlit as st
import pandas as pd
import re
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

# ==========================================
# 0. 嘗試載入主程式的 Google Sheets 連線工具
# ==========================================
try:
    from app import get_gsheet_connection, get_or_create_ws, _ws_clear, _ws_update, _sh_batch_update
    HAS_GSHEET = True
except ImportError:
    HAS_GSHEET = False

# ==========================================
# 1. 頁面基本設定與側邊欄
# ==========================================
st.set_page_config(page_title="舉發績效結算", page_icon="👮", layout="wide")

st.title("⚡ 員警交通違規舉發績效結算")
st.markdown("本頁面專門處理**員警個人績效配分**與**門檻結算**。匯出的報表將保留原始格式，並自動填入配分與黃色警示。")

st.sidebar.header("⚙️ 結算參數設定")
unit_type = st.sidebar.radio(
    "🏢 選擇單位與基準", 
    ["龍潭交通分隊 (基準800)", "一般單位 (基準400)"],
    index=0 
)
quota = 800 if "龍潭" in unit_type else 400
threshold_7x = quota * 7

st.sidebar.markdown("---")
st.sidebar.subheader("📂 步驟 1：上傳配分表")
db_file = st.sidebar.file_uploader("外部配分表 (如：檔案 B)", type=["xlsx"], key="db_file")

st.sidebar.subheader("📂 步驟 2：上傳本期舉發資料")
data_files = st.sidebar.file_uploader("批次選擇多個員警的半年期 Excel 檔案", type=["xlsx", "xls"], accept_multiple_files=True, key="data_files")

# ==========================================
# 2. 核心結算邏輯
# ==========================================
if db_file and data_files:
    with st.spinner("🔄 正在讀取與比對資料..."):
        try:
            df_db = pd.read_excel(db_file)
            if '違規條款' not in df_db.columns:
                st.error("❌ 配分表缺少『違規條款』欄位！請檢查檔案格式。")
                st.stop()
        except Exception as e:
            st.error(f"❌ 讀取配分表失敗：{e}")
            st.stop()

        # 建立配分查找字典，提升比對速度
        db_map = {}
        for _, row in df_db.iterrows():
            rule = str(row.get('違規條款', '')).strip()
            if rule:
                s = pd.to_numeric(row.get('攔舉配分', 0), errors='coerce')
                d = pd.to_numeric(row.get('逕舉配分', 0), errors='coerce')
                db_map[rule] = {
                    'stop': 0 if pd.isna(s) else int(s),
                    'dir': 0 if pd.isna(d) else int(d)
                }

        processed_sheets = []
        
        def extract_officer_name(df_head):
            for r_idx, row in df_head.iterrows():
                for c_idx, val in enumerate(row.values):
                    val_str = str(val).strip()
                    if "舉發員警" in val_str:
                        clean = re.sub(r'舉發員警[:：]?', '', val_str).strip()
                        if clean: return clean
                        if c_idx + 1 < len(row.values):
                            next_val = str(row.values[c_idx + 1]).strip()
                            if next_val and next_val.lower() != 'nan':
                                return next_val
            return ""

        for f in data_files:
            f.seek(0)
            try:
                xls = pd.ExcelFile(f)
                for sheet_name in xls.sheet_names:
                    # 讀取整張表並將 NaN 補成空字串，以利維持原始樣貌
                    raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    raw_df = raw_df.fillna("")
                    
                    header_idx = -1
                    officer_name = extract_officer_name(raw_df.head(20))
                    if not officer_name:
                        officer_name = sheet_name.strip()
                    
                    # 尋找資料標題列
                    for idx, row in raw_df.iterrows():
                        if "違規條款" in row.astype(str).str.replace(" ", "").values:
                            header_idx = idx
                            break
                    
                    if header_idx == -1:
                        st.warning(f"檔案 {f.name} 的工作表『{sheet_name}』找不到違規條款標題，已略過。")
                        continue

                    # 定位欄位索引
                    col_rule = col_s_score = col_d_score = col_s_cnt = col_d_cnt = -1
                    for c, val in enumerate(raw_df.iloc[header_idx]):
                        val_str = str(val).strip().replace(" ", "")
                        if val_str == "違規條款": col_rule = c
                        if val_str == "攔舉配分": col_s_score = c
                        if val_str == "逕舉配分": col_d_score = c
                        if val_str == "攔停數": col_s_cnt = c
                        if val_str == "逕舉數": col_d_cnt = c
                        
                    if col_rule == -1 or col_s_score == -1 or col_d_score == -1:
                        st.warning(f"工作表『{sheet_name}』缺少配分或條款欄位，略過。")
                        continue

                    grand_total = 0
                    yellow_cells = []
                    
                    # 遍歷資料列計算分數與配分
                    for r in range(header_idx + 1, len(raw_df)):
                        rule = str(raw_df.iloc[r, col_rule]).strip()
                        if not rule or "合計" in rule or "製表" in rule or "舉發單張數" in rule:
                            continue
                            
                        # 安全讀取計數
                        def safe_int(val):
                            try: return int(float(str(val).replace(",", "")))
                            except: return 0
                                
                        stop_cnt = safe_int(raw_df.iloc[r, col_s_cnt]) if col_s_cnt != -1 else 0
                        dir_cnt = safe_int(raw_df.iloc[r, col_d_cnt]) if col_d_cnt != -1 else 0
                        
                        # 查表
                        if rule in db_map:
                            s_score = db_map[rule]['stop']
                            d_score = db_map[rule]['dir']
                        else:
                            s_score = 0
                            d_score = 0
                            
                        # 更新原始 DataFrame 的儲存格內容
                        raw_df.iat[r, col_s_score] = s_score
                        raw_df.iat[r, col_d_score] = d_score
                        
                        # 記錄需要標黃色的儲存格 (轉換為 1-based index 供 openpyxl 使用)
                        if s_score == 0 and d_score == 0:
                            yellow_cells.append((r + 1, col_s_score + 1))
                            yellow_cells.append((r + 1, col_d_score + 1))
                            
                        grand_total += (s_score * stop_cnt) + (d_score * dir_cnt)
                        
                    processed_sheets.append({
                        "officer": officer_name.replace(" ", ""),
                        "df": raw_df,
                        "yellow_cells": yellow_cells,
                        "grand_total": grand_total,
                        "col_d_score": col_d_score
                    })
                    
            except Exception as e:
                st.error(f"❌ 處理 {f.name} 時發生錯誤：{e}")

        # ==========================================
        # 3. 結算彙整與報表產出
        # ==========================================
        if processed_sheets:
            # 將所有工作表的結果進行群組加總 (計算總表用)
            df_raw_summary = pd.DataFrame([{
                "員警姓名": s["officer"],
                "本期總分": s["grand_total"]
            } for s in processed_sheets])
            
            df_summary = df_raw_summary.groupby('員警姓名', as_index=False)['本期總分'].sum()
            
            # --- 上半年歷史分數 ---
            # 實務上這裡可以透過讀取外部資料庫動態獲取
            df_summary['上半年分數'] = 5800  
            
            def calc_rem(score):
                if score >= threshold_7x: return score - threshold_7x
                return score % quota
                
            df_summary['上半年剩餘分數'] = df_summary['上半年分數'].apply(calc_rem)
            df_summary['最終總分'] = df_summary['本期總分'] + df_summary['上半年剩餘分數']

            st.success("✅ 所有員警檔案結算與原格式重建完畢！")
            st.subheader("📊 員警績效結算總表")
            st.dataframe(df_summary, use_container_width=True, hide_index=True)

            # --- 產出 Excel 下載 (含總表與所有個人明細表) ---
            output = io.BytesIO()
            wb = Workbook()
            
            # 1. 建立總表
            ws_summary = wb.active
            ws_summary.title = "績效結算總表"
            ws_summary['A1'] = "交通違規舉發績效結算表"
            ws_summary['A1'].font = Font(size=14, bold=True)
            ws_summary['A2'] = f"結算單位：{unit_type}"
            
            header = list(df_summary.columns)
            header_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
            for c_idx, title in enumerate(header, 1):
                cell = ws_summary.cell(row=4, column=c_idx, value=title)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
                ws_summary.column_dimensions[get_column_letter(c_idx)].width = 16
                
            for r_idx, row_data in enumerate(df_summary.values, 5):
                for c_idx, val in enumerate(row_data, 1):
                    ws_summary.cell(row=r_idx, column=c_idx, value=val)

            # 2. 建立各員警的原始明細表 (附帶四層結算結構)
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            sheet_name_counts = {}
            
            for s in processed_sheets:
                # 確保工作表名稱不重複且不超過長度限制
                base_name = s["officer"][:25]
                if base_name not in sheet_name_counts:
                    sheet_name_counts[base_name] = 1
                    final_name = base_name
                else:
                    sheet_name_counts[base_name] += 1
                    final_name = f"{base_name}({sheet_name_counts[base_name]})"
                    
                ws_officer = wb.create_sheet(title=final_name)
                
                # 寫入保留原格式的資料
                for r_idx, row in s["df"].iterrows():
                    ws_officer.append([val if val != "" else None for val in row.tolist()])
                    
                # 標示無配分的黃色警示區塊
                for r, c in s["yellow_cells"]:
                    ws_officer.cell(row=r, column=c).fill = yellow_fill
                    
                # 取得該員警在總表中的結算數據
                officer_summary = df_summary[df_summary['員警姓名'] == s['officer']].iloc[0]
                
                # 在表尾寫入 4 層結算結構
                footer_start_row = ws_officer.max_row + 2
                write_col = s["col_d_score"] + 1 if s["col_d_score"] != -1 else 5
                
                def write_footer(row_offset, title, val, color="000000"):
                    c_title = ws_officer.cell(row=footer_start_row + row_offset, column=write_col - 1, value=title)
                    c_title.font = Font(bold=True)
                    c_val = ws_officer.cell(row=footer_start_row + row_offset, column=write_col, value=val)
                    c_val.font = Font(bold=True, color=color)
                
                write_footer(0, "本期分數：", int(officer_summary["本期總分"]), "0000FF")
                write_footer(1, "上半年分數：", int(officer_summary["上半年分數"]), "0000FF")
                write_footer(2, "上半年剩餘分數：", int(officer_summary["上半年剩餘分數"]), "0000FF")
                write_footer(3, "總分：", int(officer_summary["最終總分"]), "FF0000")
                
                # 適度調整欄寬讓資料好讀
                for col_idx in range(1, len(s["df"].columns) + 1):
                    ws_officer.column_dimensions[get_column_letter(col_idx)].width = 12

            wb.save(output)
            
            col1, col2 = st.columns([1, 3])
            with col1:
                st.download_button(
                    label="📥 下載完整報表 (含個人明細)",
                    data=output.getvalue(),
                    file_name="舉發績效結算與個人明細表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            # --- 同步總表至 Google Sheets ---
            if HAS_GSHEET:
                with col2:
                    if st.button("☁️ 同步總表至 Google Sheets", use_container_width=True):
                        with st.spinner("正在寫入雲端..."):
                            try:
                                sh = get_gsheet_connection()
                                if sh:
                                    ws_name = "績效結算總表"
                                    ws = get_or_create_ws(sh, ws_name, rows=max(30, len(df_summary) + 10), cols=10)
                                    _ws_clear(ws)
                                    
                                    title_text = f"交通違規舉發績效結算表 ({unit_type})"
                                    _ws_update(ws, 'A1', [[title_text], df_summary.columns.tolist()] + df_summary.values.tolist())
                                    
                                    reqs = [
                                        {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": len(df_summary.columns)}, "mergeType": "MERGE_ALL"}},
                                        {"repeatCell": {
                                            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1},
                                            "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 14}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}},
                                            "fields": "userEnteredFormat.textFormat.bold,userEnteredFormat.textFormat.fontSize,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment"
                                        }},
                                        {"repeatCell": {
                                            "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": len(df_summary.columns)},
                                            "cell": {"userEnteredFormat": {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8}}},
                                            "fields": "userEnteredFormat.textFormat.bold,userEnteredFormat.backgroundColor"
                                        }}
                                    ]
                                    _sh_batch_update(sh, {"requests": reqs})
                                    st.success("✅ 雲端資料庫已同步！")
                            except Exception as e:
                                st.error(f"雲端同步失敗：{e}")
