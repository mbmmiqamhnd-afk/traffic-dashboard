import streamlit as st
import pandas as pd
import re
import io

# ==========================================
# 輔助函式：車號標準化 (去空白、特殊符號、轉大寫)
# ==========================================
def normalize_plate(plate):
    if pd.isna(plate):
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(plate)).upper()

# ==========================================
# 輔助函式：讀取檔案 (支援 Excel 指定工作表與 CSV 編碼容錯)
# ==========================================
def load_data(file, sheet_name=None):
    file.seek(0) 
    
    if file.name.endswith('.xlsx'):
        return pd.read_excel(file, sheet_name=sheet_name, engine='openpyxl')
    else:
        try:
            return pd.read_csv(file, encoding='utf-8-sig')
        except UnicodeDecodeError:
            file.seek(0)
            return pd.read_csv(file, encoding='big5')

# ==========================================
# 輔助函式：自動尋找預設工作表索引
# ==========================================
def get_default_sheet_index(sheet_names, keywords):
    """根據關鍵字自動尋找最相符的工作表，若無則預設回傳第一個(0)"""
    for i, sheet_name in enumerate(sheet_names):
        for kw in keywords:
            if kw in sheet_name:
                return i
    return 0

# ==========================================
# 主程式
# ==========================================
st.set_page_config(page_title="噪音改裝車輛嘉獎統計系統", layout="wide")
st.title("🚓 噪音改裝車輛嘉獎次數統計系統 (自動偵測版)")

st.markdown("""
💡 **系統會自動判斷統計模式與工作表：**
*   **自動抓取工作表**：上傳 Excel 後，系統會自動尋找並鎖定對應的資料表 (如：受理明細、靜桃)，您無須手動點選。
*   **上半年統計**：只需上傳前兩個檔案，第三個留空即可。
*   **下半年統計**：上傳全部三個檔案，系統會自動將「前期明細」納入合計。
""")

# --- 側邊欄設定區 ---
st.sidebar.header("⚙️ 參數設定")
st.sidebar.markdown("請設定各檔案要從**第幾列**開始讀取資料（預設為 2，即跳過第 1 列標題）")
start_row_src1 = st.sidebar.number_input("[靜桃清冊] 起始列", min_value=2, value=2, step=1)
start_row_tgt = st.sidebar.number_input("[受理明細] 起始列", min_value=2, value=2, step=1)
start_row_src2 = st.sidebar.number_input("[前期明細] 起始列 (下半年專用)", min_value=2, value=2, step=1)

# --- 檔案上傳區與工作表選擇 ---
st.markdown("### 📥 上傳資料檔案")
col1, col2, col3 = st.columns(3)

# 1. 受理明細 (目標表)
with col1: 
    file_tgt = st.file_uploader("1. 上傳 [受理明細] (必填)", type=['csv', 'xlsx'])
    sheet_tgt = None
    if file_tgt and file_tgt.name.endswith('.xlsx'):
        xls_tgt = pd.ExcelFile(file_tgt, engine='openpyxl')
        # 自動尋找包含「受理明細」的工作表
        default_idx = get_default_sheet_index(xls_tgt.sheet_names, ['受理明細'])
        sheet_tgt = st.selectbox("📂 選擇工作表 (已自動辨識)", xls_tgt.sheet_names, index=default_idx, key="sheet_tgt")

# 2. 靜桃清冊 (本期來源)
with col2: 
    file_src1 = st.file_uploader("2. 上傳 [靜桃清冊] (必填)", type=['csv', 'xlsx'])
    sheet_src1 = None
    if file_src1 and file_src1.name.endswith('.xlsx'):
        xls_src1 = pd.ExcelFile(file_src1, engine='openpyxl')
        # 自動尋找包含「靜桃」的工作表
        default_idx = get_default_sheet_index(xls_src1.sheet_names, ['靜桃'])
        sheet_src1 = st.selectbox("📂 選擇工作表 (已自動辨識)", xls_src1.sheet_names, index=default_idx, key="sheet_src1")

# 3. 前期明細 (前期來源)
with col3: 
    file_src2 = st.file_uploader("3. 上傳 [前期明細] (上半年請留空)", type=['csv', 'xlsx'])
    sheet_src2 = None
    if file_src2 and file_src2.name.endswith('.xlsx'):
        xls_src2 = pd.ExcelFile(file_src2, engine='openpyxl')
        # 自動尋找包含「嘉獎」或「明細」的工作表
        default_idx = get_default_sheet_index(xls_src2.sheet_names, ['嘉獎', '明細'])
        sheet_src2 = st.selectbox("📂 選擇工作表 (已自動辨識)", xls_src2.sheet_names, index=default_idx, key="sheet_src2")

# --- 執行統計區塊 ---
if file_tgt and file_src1:
    if st.button("🚀 開始執行統計", type="primary"):
        with st.spinner('資料讀取與處理中...'):
            try:
                is_second_half = file_src2 is not None

                df_tgt = load_data(file_tgt, sheet_tgt)
                df_src1 = load_data(file_src1, sheet_src1)

                df_tgt_filtered = df_tgt.iloc[start_row_tgt - 2:]
                df_src1_filtered = df_src1.iloc[start_row_src1 - 2:]

                # -------------------------------------------
                # 步驟 1：建立 [車號 -> 通報人] 對照表
                # -------------------------------------------
                plate_to_reporter = {}
                for _, row in df_src1_filtered.iterrows():
                    if len(row) > 6:
                        plate = normalize_plate(row.iloc[4])
                        name = str(row.iloc[6]).strip()
                        if plate and name and name != 'nan':
                            plate_to_reporter[plate] = name

                # -------------------------------------------
                # 步驟 2：計算本期件數 (篩選「龍警分交字」)
                # -------------------------------------------
                current_counts = {}
                for _, row in df_tgt_filtered.iterrows():
                    if len(row) > 1:
                        doc_num = str(row.iloc[0])
                        plate = normalize_plate(row.iloc[1])
                        
                        if "龍警分交字" in doc_num and plate in plate_to_reporter:
                            reporter = plate_to_reporter[plate]
                            current_counts[reporter] = current_counts.get(reporter, 0) + 1

                # -------------------------------------------
                # 步驟 3：讀取前期資料
                # -------------------------------------------
                history_map = {}
                if is_second_half:
                    df_src2_data = load_data(file_src2, sheet_src2)
                    df_src2_filtered = df_src2_data.iloc[start_row_src2 - 2:]
                    for _, row in df_src2_filtered.iterrows():
                        if len(row) > 4:
                            h_name = str(row.iloc[0]).strip()
                            h_val = row.iloc[4]
                            if h_name and h_name != 'nan' and pd.notna(h_val):
                                try:
                                    history_map[h_name] = int(float(h_val))
                                except ValueError:
                                    pass

                # -------------------------------------------
                # 步驟 4：整合計算商數與輸出設定
                # -------------------------------------------
                output_data = []
                for name, count_current in current_counts.items():
                    if is_second_half:
                        count_history = history_map.get(name, 0)
                        count_total = count_current + count_history
                        reward_count = count_total // 6
                        output_data.append([name, count_current, count_history, count_total, reward_count])
                    else:
                        count_total = count_current
                        reward_count = count_total // 6
                        output_data.append([name, count_current, count_total, reward_count])

                if is_second_half:
                    cols = ['通報人(A)', '本期件數(B)', '前期件數(C)', '合計件數(D)', '嘉獎數(E)']
                    sort_col = '嘉獎數(E)'
                    mode_name = "下半年"
                else:
                    cols = ['通報人(A)', '本期件數(B)', '合計件數(C)', '嘉獎數(D)']
                    sort_col = '嘉獎數(D)'
                    mode_name = "上半年"

                df_result = pd.DataFrame(output_data, columns=cols)
                df_result = df_result.sort_values(by=sort_col, ascending=False).reset_index(drop=True)

                st.success(f"✅ 統計完成！已自動採用「{mode_name}模式」，共計算 {len(df_result)} 位通報人。")
                st.dataframe(df_result, use_container_width=True)

                csv_buffer = io.StringIO()
                df_result.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                filename = f"{mode_name}嘉獎次數統計結果.csv"
                
                st.download_button(
                    label=f"📥 下載統計結果 ({filename})",
                    data=csv_buffer.getvalue(),
                    file_name=filename,
                    mime="text/csv",
                    type="primary"
                )

            except Exception as e:
                st.error(f"❌ 發生錯誤，請檢查檔案格式或設定的起始列。\n詳細錯誤訊息：{e}")
else:
    st.info("請至少上傳「受理明細」與「靜桃清冊」兩個檔案，以啟動統計按鈕。支援 CSV 與 Excel 格式。")
