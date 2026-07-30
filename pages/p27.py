import streamlit as st
import pandas as pd
import re
import io
import time
import traceback
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

# 嘗試從主程式匯入 Google Sheets 連線工具 (請依據您的主程式檔名修改，例如 app 或 main)
# 這樣可以共用同一組 GCP_CREDS 與連線快取
try:
    from app import get_gsheet_connection, get_or_create_ws, _ws_clear, _ws_update, _sh_batch_update
    HAS_GSHEET = True
except ImportError:
    HAS_GSHEET = False

st.set_page_config(page_title="舉發績效結算", page_icon="👮", layout="wide")

st.title("⚡ 員警交通違規舉發績效結算")
st.markdown("本頁面專門處理**員警個人績效配分**與**門檻結算**。請上傳當月資料與配分表。")

# ==========================================
# 1. 側邊欄：參數設定與檔案上傳
# ==========================================
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

st.sidebar.subheader("📂 步驟 2：上傳本月舉發資料")
data_files = st.sidebar.file_uploader("批次選擇多個員警的 Excel 檔案", type=["xlsx", "xls"], accept_multiple_files=True, key="data_files")

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

        all_results = []
        
        for f in data_files:
            f.seek(0)
            try:
                # 尋找標頭列 (包含 '違規條款')
                raw_df = pd.read_excel(f, header=None)
                header_idx = -1
                for idx, row in raw_df.iterrows():
                    if "違規條款" in row.astype(str).values:
                        header_idx = idx
                        break
                
                if header_idx != -1:
                    f.seek(0)
                    df_officer = pd.read_excel(f, header=header_idx)
                else:
                    st.warning(f"檔案 {f.name} 找不到『違規條款』標題列，已略過。")
                    continue

                # 查表合併配分
                df_merged = pd.merge(df_officer, df_db[['違規條款', '攔舉配分', '逕舉配分']], on='違規條款', how='left')

                # 數值處理與計算
                for col in ['攔停數', '逕舉數', '攔舉配分_y', '逕舉配分_y', '攔舉配分', '逕舉配分']:
                    if col in df_merged.columns:
                        df_merged[col] = pd.to_numeric(df_merged[col], errors='coerce').fillna(0)
                        
                p_stop = df_merged.get('攔舉配分_y', df_merged.get('攔舉配分', 0))
                p_dir = df_merged.get('逕舉配分_y', df_merged.get('逕舉配分', 0))
                
                df_merged['單項總分'] = (df_merged.get('攔停數', 0) * p_stop) + (df_merged.get('逕舉數', 0) * p_dir)
                monthly_total = df_merged['單項總分'].sum()
                
                # 從檔名萃取姓名
                officer_name = re.sub(r'\.[a-zA-Z0-9]+$', '', f.name)
                all_results.append({
                    "員警姓名": officer_name,
                    "當月總分": monthly_total
                })
                
            except Exception as e:
                st.error(f"❌ 處理 {f.name} 時發生錯誤：{e}")

        # ==========================================
        # 3. 結算彙整與報表產出
        # ==========================================
        if all_results:
            df_summary = pd.DataFrame(all_results)
            
            # TODO: 實務上您可以串接另一個資料庫或讀取雲端硬碟，取得真正的上半年分數
            # 目前以 5800 作為模擬範例
            df_summary['上半年分數'] = 5800  
            
            def calc_rem(score):
                if score >= threshold_7x: return score - threshold_7x
                return score % quota
                
            df_summary['上半年剩餘分數'] = df_summary['上半年分數'].apply(calc_rem)
            df_summary['最終總分'] = df_summary['當月總分'] + df_summary['上半年剩餘分數']

            st.success("✅ 所有檔案結算完畢！")
            
            # --- 畫面展示 ---
            st.subheader("📊 績效結算結果")
            st.dataframe(df_summary, use_container_width=True, hide_index=True)

            # --- 產出 Excel 下載 (檔案流模式核心) ---
            output = io.BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.title = "績效結算報表"
            
            # 設定大標題 (A1, A2)
            ws['A1'] = "交通違規舉發績效結算表"
            ws['A1'].font = Font(size=14, bold=True)
            ws['A2'] = f"結算單位：{unit_type}"
            
            # 寫入標題與資料 (強制從 A4 開始寫入)
            header = list(df_summary.columns)
            header_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
            for c_idx, title in enumerate(header, 1):
                cell = ws.cell(row=4, column=c_idx, value=title)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
                ws.column_dimensions[cell.column_letter].width = 16
                
            for r_idx, row_data in enumerate(df_summary.values, 5):
                for c_idx, val in enumerate(row_data, 1):
                    ws.cell(row=r_idx, column=c_idx, value=val)

            wb.save(output)
            
            col1, col2 = st.columns([1, 3])
            with col1:
                st.download_button(
                    label="📥 下載 Excel 報表",
                    data=output.getvalue(),
                    file_name="舉發績效結算表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            # --- 同步至 Google Sheets (可選) ---
            if HAS_GSHEET:
                with col2:
                    if st.button("☁️ 同步至 Google Sheets", use_container_width=True):
                        with st.spinner("正在寫入雲端..."):
                            try:
                                sh = get_gsheet_connection()
                                if sh:
                                    ws_name = "績效結算總表"
                                    ws = get_or_create_ws(sh, ws_name, rows=max(30, len(df_summary) + 10), cols=10)
                                    _ws_clear(ws)
                                    
                                    title_text = f"交通違規舉發績效結算表 ({unit_type})"
                                    _ws_update(ws, 'A1', [[title_text], df_summary.columns.tolist()] + df_summary.values.tolist())
                                    
                                    # 基本格式設定
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
