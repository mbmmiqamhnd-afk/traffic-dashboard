import streamlit as st
import pandas as pd
import re
import io
import smtplib
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

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
    for i, sheet_name in enumerate(sheet_names):
        for kw in keywords:
            if kw in sheet_name:
                return i
    return 0

# ==========================================
# 輔助函式：寄送 Email 備份
# ==========================================
def send_csv_email(df, mode_name):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        
        date_str = datetime.now().strftime('%Y%m%d')
        msg["Subject"] = f"龍潭分局_{mode_name}噪音改裝車輛嘉獎統計結果_{date_str}"
        
        body_text = (
            f"您好，\n\n"
            f"附件為系統自動產生的「{mode_name}」噪音改裝車輛嘉獎次數統計結果（CSV格式），請查收。\n\n"
            f"本信件由交通執法自動化分析引擎發送。"
        )
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        filename = f"{mode_name}嘉獎次數統計結果.csv"
        csv_str = df.to_csv(index=False, encoding='utf-8-sig')
        
        part = MIMEBase("application", "csv")
        part.set_payload(csv_str.encode('utf-8-sig'))
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(filename)}")
        msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
            
        return True, None
    except Exception as e:
        return False, str(e)

# ==========================================
# 主程式
# ==========================================
st.set_page_config(page_title="噪音改裝車輛嘉獎統計系統", layout="wide")
st.title("🚓 噪音改裝車輛嘉獎次數統計系統 (內容偵測版)")

st.markdown("""
💡 **系統具備全方位自動偵測能力：**
*   **資料年度辨識**：自動掃描上傳的資料內容，擷取「年度」(如114年、115年)；透過有無第三個檔案判斷「上下半年」。
*   **工作表鎖定**：上傳 Excel 後，系統會自動尋找並鎖定對應的資料表。
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

with col1: 
    file_tgt = st.file_uploader("1. 上傳 [受理明細] (必填)", type=['csv', 'xlsx'])
    sheet_tgt = None
    if file_tgt and file_tgt.name.endswith('.xlsx'):
        xls_tgt = pd.ExcelFile(file_tgt, engine='openpyxl')
        default_idx = get_default_sheet_index(xls_tgt.sheet_names, ['受理明細'])
        sheet_tgt = st.selectbox("📂 選擇工作表 (已自動辨識)", xls_tgt.sheet_names, index=default_idx, key="sheet_tgt")

with col2: 
    file_src1 = st.file_uploader("2. 上傳 [靜桃清冊] (必填)", type=['csv', 'xlsx'])
    sheet_src1 = None
    if file_src1 and file_src1.name.endswith('.xlsx'):
        xls_src1 = pd.ExcelFile(file_src1, engine='openpyxl')
        default_idx = get_default_sheet_index(xls_src1.sheet_names, ['靜桃'])
        sheet_src1 = st.selectbox("📂 選擇工作表 (已自動辨識)", xls_src1.sheet_names, index=default_idx, key="sheet_src1")

with col3: 
    file_src2 = st.file_uploader("3. 上傳 [前期明細] (上半年請留空)", type=['csv', 'xlsx'])
    sheet_src2 = None
    if file_src2 and file_src2.name.endswith('.xlsx'):
        xls_src2 = pd.ExcelFile(file_src2, engine='openpyxl')
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
                # 自動從「受理明細」資料內容中擷取年度
                # -------------------------------------------
                auto_year = "115" # 預設防呆值
                found_year = False
                # 掃描前 50 筆資料，尋找包含「數字+年」的格式 (例如 114年, 115年)
                for _, row in df_tgt_filtered.head(50).iterrows():
                    row_content = " ".join([str(val) for val in row if pd.notna(val)])
                    match = re.search(r'(\d{2,3})年', row_content)
                    if match:
                        auto_year = match.group(1)
                        found_year = True
                        break

                plate_to_reporter = {}
                for _, row in df_src1_filtered.iterrows():
                    if len(row) > 6:
                        plate = normalize_plate(row.iloc[4])
                        name = str(row.iloc[6]).strip()
                        if plate and name and name != 'nan':
                            plate_to_reporter[plate] = name

                current_counts = {}
                for _, row in df_tgt_filtered.iterrows():
                    if len(row) > 1:
                        doc_num = str(row.iloc[0])
                        plate = normalize_plate(row.iloc[1])
                        
                        if "龍警分交字" in doc_num and plate in plate_to_reporter:
                            reporter = plate_to_reporter[plate]
                            current_counts[reporter] = current_counts.get(reporter, 0) + 1

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

                # 將從資料抓取到的年度套用到最終名稱
                year_str = f"{auto_year}年"
                
                if is_second_half:
                    cols = ['通報人(A)', '本期件數(B)', '前期件數(C)', '合計件數(D)', '嘉獎數(E)']
                    sort_col = '嘉獎數(E)'
                    mode_name = f"{year_str}下半年"
                else:
                    cols = ['通報人(A)', '本期件數(B)', '合計件數(C)', '嘉獎數(D)']
                    sort_col = '嘉獎數(D)'
                    mode_name = f"{year_str}上半年"

                df_result = pd.DataFrame(output_data, columns=cols)
                df_result = df_result.sort_values(by=sort_col, ascending=False).reset_index(drop=True)
                
                st.session_state['df_result'] = df_result
                st.session_state['mode_name'] = mode_name
                st.session_state['auto_year'] = auto_year
                st.session_state['calc_done'] = True

            except Exception as e:
                st.error(f"❌ 發生錯誤，請檢查檔案格式或設定的起始列。\n詳細錯誤訊息：{e}")

# --- 結果顯示與後續操作區塊 ---
if st.session_state.get('calc_done', False):
    df_result = st.session_state['df_result']
    mode_name = st.session_state['mode_name']
    auto_year = st.session_state['auto_year']
    
    st.info(f"🔎 系統已從資料內容自動偵測出年度為：**{auto_year} 年**")
    st.success(f"✅ 統計完成！已自動採用「{mode_name}模式」，共計算 {len(df_result)} 位通報人。")
    st.dataframe(df_result, use_container_width=True)

    st.markdown("### 💾 儲存與備份")
    d_col1, d_col2 = st.columns(2)
    
    with d_col1:
        csv_buffer = io.StringIO()
        df_result.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
        filename = f"{mode_name}嘉獎次數統計結果.csv"
        
        st.download_button(
            label=f"📥 下載統計結果 ({filename})",
            data=csv_buffer.getvalue(),
            file_name=filename,
            mime="text/csv",
            use_container_width=True
        )

    with d_col2:
        if st.button("📧 將統計結果寄至我的信箱 (備份)", use_container_width=True):
            with st.spinner("信件寄送中，請稍候…"):
                ok, mail_err = send_csv_email(df_result, mode_name)
                if ok:
                    st.success("✅ 信件發送成功！請檢查您的信箱。")
                else:
                    st.error(f"❌ 發信失敗: {mail_err}")
else:
    if not (file_tgt and file_src1):
        st.info("請至少上傳「受理明細」與「靜桃清冊」兩個檔案，以啟動統計按鈕。支援 CSV 與 Excel 格式。")
