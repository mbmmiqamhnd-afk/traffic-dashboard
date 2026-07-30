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
st.markdown("本頁面專門處理**員警個人績效配分**與**門檻結算**。匯出的報表將保留原始格式，若原始檔案缺少配分欄位，系統將**自動動態安插**並填入分數與黃色警示。")

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

st.sidebar.subheader("📂 步驟 2：上傳原始舉發資料")
data_files = st.sidebar.file_uploader("批次選擇多個員警的半年期 Excel 檔案 (支援原始無配分欄位格式)", type=["xlsx", "xls"], accept_multiple_files=True, key="data_files")

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
                    raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    raw_df = raw_df.astype('object')
                    
                    header_idx = -1
                    officer_name = extract_officer_name(raw_df.head(20))
                    if not officer_name:
                        officer_name = sheet_name.strip()
                    
                    for idx, row in raw_df.iterrows():
                        if "違規條款" in row.astype(str).str.replace(" ", "").values:
                            header_idx = idx
                            break
                    
                    if header_idx == -1: continue

                    header_row = [str(x).strip().replace(" ", "") for x in raw_df.iloc[header_idx]]
                    col_rule = header_row.index("違規條款") if "違規條款" in header_row else -1
                    col_s_cnt = header_row.index("攔停數") if "攔停數" in header_row else -1
                    col_d_cnt = header_row.index("逕舉數") if "逕舉數" in header_row else -1
                    
                    col_s_score = header_row.index("攔舉配分") if "攔舉配分" in header_row else -1
                    col_d_score = header_row.index("逕舉配分") if "逕舉配分" in header_row else -1
                    col_subtotal = header_row.index("小計") if "小計" in header_row else -1
                    
                    if col_rule == -1 or col_s_cnt == -1 or col_d_cnt == -1: continue

                    if col_s_score == -1:
                        idx_s_score = col_s_cnt + 1
                        raw_df.insert(idx_s_score, f'new_{idx_s_score}', None)
                        raw_df.iat[header_idx, idx_s_score] = "攔舉配分"
                        col_s_score = idx_s_score
                        
                        if col_d_cnt >= idx_s_score: col_d_cnt += 1
                        if col_subtotal >= idx_s_score: col_subtotal += 1
                        
                        idx_d_score = col_d_cnt + 1
                        raw_df.insert(idx_d_score, f'new_{idx_d_score}', None)
                        raw_df.iat[header_idx, idx_d_score] = "逕舉配分"
                        col_d_score = idx_d_score
                        
                        if col_subtotal >= idx_d_score: col_subtotal += 1
                        
                        idx_subtotal = idx_d_score + 1
                        raw_df.insert(idx_subtotal, f'new_{idx_subtotal}', None)
                        raw_df.iat[header_idx, idx_subtotal] = "小計"
                        col_subtotal = idx_subtotal
                        raw_df.columns = range(raw_df.shape[1])

                    grand_total = 0
                    yellow_cells = []
                    
                    for r in range(header_idx + 1, len(raw_df)):
                        rule = str(raw_df.iloc[r, col_rule]).strip()
                        if not rule or "合計" in rule or "製表" in rule or "舉發單張數" in rule or rule.lower() == 'nan':
                            continue
                            
                        def safe_int(val):
                            try: return int(float(str(val).replace(",", "")))
                            except: return 0
                                
                        stop_cnt = safe_int(raw_df.iloc[r, col_s_cnt])
                        dir_cnt = safe_int(raw_df.iloc[r, col_d_cnt])
                        
                        if rule in db_map:
                            s_score = db_map[rule]['stop']
                            d_score = db_map[rule]['dir']
                        else:
                            s_score = 0
                            d_score = 0
                            
                        raw_df.iat[r, col_s_score] = s_score if s_score != 0 else 0
                        raw_df.iat[r, col_d_score] = d_score if d_score != 0 else 0
                        
                        row_subtotal = (s_score * stop_cnt) + (d_score * dir_cnt)
                        if col_subtotal != -1:
                            raw_df.iat[r, col_subtotal] = row_subtotal
                        
                        if s_score == 0 and d_score == 0:
                            yellow_cells.append((r + 1, col_s_score + 1))
                            yellow_cells.append((r + 1, col_d_score + 1))
                            
                        grand_total += row_subtotal
                        
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
            df_raw_summary = pd.DataFrame([{
                "員警姓名": s["officer"],
                "本期總分": s["grand_total"]
            } for s in processed_sheets])
            
            df_summary = df_raw_summary.groupby('員警姓名', as_index=False)['本期總分'].sum()
            df_summary['上半年分數'] = 5800  
            
            def calc_rem(score):
                if score >= threshold_7x: return score - threshold_7x
                return score % quota
                
            df_summary['上半年剩餘分數'] = df_summary['上半年分數'].apply(calc_rem)
            df_summary['最終總分'] = df_summary['本期總分'] + df_summary['上半年剩餘分數']

            st.success("✅ 所有原始報表已自動擴充欄位、結算並重建完畢！")
            st.subheader("📊 員警績效結算總表")
            st.dataframe(df_summary, use_container_width=True, hide_index=True)

            # --- 產出 Excel 下載 ---
            output = io.BytesIO()
            wb = Workbook()
            
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

            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            sheet_name_counts = {}
            
            for s in processed_sheets:
                base_name = s["officer"][:25]
                if base_name not in sheet_name_counts:
                    sheet_name_counts[base_name] = 1
                    final_name = base_name
                else:
                    sheet_name_counts[base_name] += 1
                    final_name = f"{base_name}({sheet_name_counts[base_name]})"
                    
                ws_officer = wb.create_sheet(title=final_name)
                
                for r_idx, row in s["df"].iterrows():
                    cleaned_row = [val if pd.notna(val) else None for val in row.tolist()]
                    ws_officer.append(cleaned_row)
                    
                for r, c in s["yellow_cells"]:
                    ws_officer.cell(row=r, column=c).fill = yellow_fill
                    
                officer_summary = df_summary[df_summary['員警姓名'] == s['officer']].iloc[0]
                
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
                
                for col_idx in range(1, len(s["df"].columns) + 1):
                    ws_officer.column_dimensions[get_column_letter(col_idx)].width = 12

            wb.save(output)
            
            col1, col2 = st.columns([1, 3])
            with col1:
                st.download_button(
                    label="📥 下載完整報表 (自動擴充格式)",
                    data=output.getvalue(),
                    file_name="舉發績效結算與個人明細表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            # --- 同步至 Google Sheets ---
            if HAS_GSHEET:
                with col2:
                    if st.button("☁️ 同步總表與明細至 Google Sheets", use_container_width=True):
                        with st.spinner("正在將總表與所有員警明細寫入雲端 (請稍候)..."):
                            try:
                                sh = get_gsheet_connection()
                                if sh:
                                    # 1. 寫入總表
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
                                    
                                    # 2. 寫入各員警獨立分頁 (包含原始數據、安插欄位、表尾分數計算與顏色)
                                    sheet_name_counts_gs = {}
                                    for s in processed_sheets:
                                        base_name_gs = s["officer"][:25]
                                        if base_name_gs not in sheet_name_counts_gs:
                                            sheet_name_counts_gs[base_name_gs] = 1
                                            final_name_gs = base_name_gs
                                        else:
                                            sheet_name_counts_gs[base_name_gs] += 1
                                            final_name_gs = f"{base_name_gs}({sheet_name_counts_gs[base_name_gs]})"
                                            
                                        ws_off = get_or_create_ws(sh, final_name_gs, rows=len(s["df"]) + 15, cols=len(s["df"].columns) + 2)
                                        _ws_clear(ws_off)
                                        
                                        data_2d = s["df"].fillna("").values.tolist()
                                        _ws_update(ws_off, 'A1', data_2d)
                                        
                                        officer_summary = df_summary[df_summary['員警姓名'] == s['officer']].iloc[0]
                                        footer_start_row_gs = len(data_2d) + 2
                                        write_col_idx = s["col_d_score"] + 1 if s["col_d_score"] != -1 else 5
                                        
                                        footer_data = [
                                            ["本期分數：", int(officer_summary["本期總分"])],
                                            ["上半年分數：", int(officer_summary["上半年分數"])],
                                            ["上半年剩餘分數：", int(officer_summary["上半年剩餘分數"])],
                                            ["總分：", int(officer_summary["最終總分"])]
                                        ]
                                        
                                        # 定位並寫入四層分數 (自動對齊逕舉配分下方)
                                        start_cell_gs = f"{get_column_letter(write_col_idx - 1)}{footer_start_row_gs}"
                                        _ws_update(ws_off, start_cell_gs, footer_data)
                                        
                                        sheet_id = ws_off.id
                                        col_idx_0based = write_col_idx - 2
                                        
                                        reqs.append({
                                            "repeatCell": {
                                                "range": {"sheetId": sheet_id, "startRowIndex": footer_start_row_gs - 1, "endRowIndex": footer_start_row_gs + 2, "startColumnIndex": col_idx_0based, "endColumnIndex": col_idx_0based + 2},
                                                "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 1.0}}}},
                                                "fields": "userEnteredFormat.textFormat"
                                            }
                                        })
                                        reqs.append({
                                            "repeatCell": {
                                                "range": {"sheetId": sheet_id, "startRowIndex": footer_start_row_gs + 2, "endRowIndex": footer_start_row_gs + 3, "startColumnIndex": col_idx_0based, "endColumnIndex": col_idx_0based + 2},
                                                "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}}}},
                                                "fields": "userEnteredFormat.textFormat"
                                            }
                                        })
                                        
                                        # 套用黃色警示背景色
                                        for r, c in s["yellow_cells"]:
                                            reqs.append({
                                                "repeatCell": {
                                                    "range": {"sheetId": sheet_id, "startRowIndex": r - 1, "endRowIndex": r, "startColumnIndex": c - 1, "endColumnIndex": c},
                                                    "cell": {"userEnteredFormat": {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 0.0}}},
                                                    "fields": "userEnteredFormat.backgroundColor"
                                                }
                                            })
                                            
                                    _sh_batch_update(sh, {"requests": reqs})
                                    st.success("✅ 雲端資料庫已完整同步 (包含總表與所有員警個人明細表)！")
                            except Exception as e:
                                st.error(f"雲端同步失敗：{e}")
