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
# 固定參數設定 (隱藏於後台執行)
# ==========================================
START_ROW_TGT = 2   # [受理明細] 起始列
START_ROW_SRC1 = 2  # [靜桃清冊] 起始列
START_ROW_SRC2 = 2  # [前期明細] 起始列

COL_DATE = 1        # 通報時間 所在欄位 (B欄)
COL_PLATE = 4       # 車號 所在欄位 (E欄)
COL_NAME = 6        # 通報人 所在欄位 (G欄)
COL_UNIT = 7        # 單位 所在欄位 (H欄)

# ==========================================
# 輔助函式：車號標準化
# ==========================================
def normalize_plate(plate):
    if pd.isna(plate):
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(plate)).upper()

# ==========================================
# 輔助函式：年度與月份智慧解析 (嚴格雙重過濾版)
# ==========================================
def extract_year_month(date_val):
    if pd.isna(date_val): return None, None
    
    if isinstance(date_val, (datetime, pd.Timestamp)): 
        y = date_val.year
        if y > 1911: y -= 1911
        return y, date_val.month
    
    s = str(date_val).strip()
    
    parts = re.split(r'[/.-]', s.split(' ')[0])
    if len(parts) >= 2:
        try:
            y = int(parts[0])
            m = int(parts[1])
            if y > 1911: y -= 1911
            if 1 <= m <= 12: 
                return y, m
        except: pass
    
    match = re.search(r'(\d{2,4})年(\d{1,2})月', s)
    if match:
        try:
            y = int(match.group(1))
            m = int(match.group(2))
            if y > 1911: y -= 1911
            if 1 <= m <= 12: 
                return y, m
        except: pass
        
    return None, None

# ==========================================
# 輔助函式：讀取檔案 
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

def get_default_sheet_index(sheet_names, keywords):
    for i, sheet_name in enumerate(sheet_names):
        for kw in keywords:
            if kw in sheet_name:
                return i
    return 0

# ==========================================
# 輔助函式：寄送 Email 備份
# ==========================================
def send_csv_email(df_person, df_unit, mode_name):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        
        date_str = datetime.now().strftime('%Y%m%d')
        msg["Subject"] = f"龍潭分局_{mode_name}噪音改裝車輛統計結果_{date_str}"
        
        body_text = (
            f"您好，\n\n"
            f"附件為系統自動產生的「{mode_name}」噪音改裝車輛統計結果，包含：\n"
            f"1. 個人嘉獎次數統計 (僅計算成案)\n"
            f"2. 各單位通報件數統計 (已嚴格剔除跨年度與非本期資料)\n\n"
            f"本信件由交通執法自動化分析引擎發送。"
        )
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        filename_p = f"{mode_name}個人嘉獎統計.csv"
        csv_str_p = df_person.to_csv(index=False, encoding='utf-8-sig')
        part1 = MIMEBase("application", "csv")
        part1.set_payload(csv_str_p.encode('utf-8-sig'))
        encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(filename_p)}")
        msg.attach(part1)

        filename_u = f"{mode_name}單位通報件數統計.csv"
        csv_str_u = df_unit.to_csv(index=False, encoding='utf-8-sig')
        part2 = MIMEBase("application", "csv")
        part2.set_payload(csv_str_u.encode('utf-8-sig'))
        encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(filename_u)}")
        msg.attach(part2)

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

# --- 側邊欄：導覽功能 (保留側邊欄給主頁面切換使用) ---
st.sidebar.header("🏠 系統導覽")

# 這裡為您預留了返回主頁的按鈕，如果您有特定的主頁檔名 (例如 app.py 或 Home.py)
# 只要把下面的註解拿掉，並改成您的主頁檔名就可以一鍵跳轉了！
# if st.sidebar.button("⬅️ 回到系統主頁", use_container_width=True):
#     st.switch_page("app.py")

# 提示訊息，讓側邊欄不會太空曠
st.sidebar.info("💡 統計參數已於系統背景固定，無需手動設定即可直接執行統計。")
st.sidebar.markdown("---")

st.title("🚓 噪音改裝車輛嘉獎與績效統計系統")

st.markdown("""
💡 **系統已進入全自動極簡模式：**
只需依序上傳檔案並點擊執行，系統將自動偵測年度、過濾歷史無效資料，並產出個人與單位績效雙報表。
""")

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

                df_tgt_filtered = df_tgt.iloc[START_ROW_TGT - 2:]
                df_src1_filtered = df_src1.iloc[START_ROW_SRC1 - 2:]

                # -------------------------------------------
                # 自動擷取基準年度 (從受理明細)
                # -------------------------------------------
                auto_year = "115"
                for _, row in df_tgt_filtered.head(50).iterrows():
                    row_content = " ".join([str(val) for val in row if pd.notna(val)])
                    match = re.search(r'(\d{2,3})年', row_content)
                    if match:
                        auto_year = match.group(1)
                        break
                
                target_months = [7, 8, 9, 10, 11, 12] if is_second_half else [1, 2, 3, 4, 5, 6]

                # -------------------------------------------
                # 掃描靜桃清冊：嚴格雙重過濾
                # -------------------------------------------
                plate_to_reporter = {}
                unit_counts = {}
                max_needed_col = max(COL_PLATE, COL_NAME, COL_UNIT, COL_DATE)
                
                for _, row in df_src1_filtered.iterrows():
                    if len(row) > max_needed_col:
                        plate = normalize_plate(row.iloc[COL_PLATE])
                        name = str(row.iloc[COL_NAME]).strip()
                        unit = str(row.iloc[COL_UNIT]).strip()
                        
                        # 1. 建立嘉獎比對名單
                        if plate and name and name != 'nan':
                            plate_to_reporter[plate] = name

                        # 2. 單位件數：年度與月份雙重吻合才計入
                        raw_date = row.iloc[COL_DATE]
                        y, m = extract_year_month(raw_date)
                        
                        is_valid_time = False
                        if y is not None and m is not None:
                            if str(y) == str(auto_year) and m in target_months:
                                is_valid_time = True

                        if is_valid_time:
                            if unit and unit != 'nan' and unit.strip() != "":
                                unit_counts[unit] = unit_counts.get(unit, 0) + 1

                # -------------------------------------------
                # 掃描受理明細：計算個人成案數
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
                # 讀取前期資料 (個人嘉獎跨期用)
                # -------------------------------------------
                history_map = {}
                if is_second_half:
                    df_src2_data = load_data(file_src2, sheet_src2)
                    df_src2_filtered = df_src2_data.iloc[START_ROW_SRC2 - 2:]
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
                # 整合資料
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

                year_str = f"{auto_year}年"
                
                if is_second_half:
                    cols = ['通報人(A)', '本期件數(成案)', '前期件數(成案)', '合計件數(成案)', '嘉獎數(E)']
                    sort_col = '嘉獎數(E)'
                    mode_name = f"{year_str}下半年"
                else:
                    cols = ['通報人(A)', '本期件數(成案)', '合計件數(成案)', '嘉獎數(D)']
                    sort_col = '嘉獎數(D)'
                    mode_name = f"{year_str}上半年"

                df_person = pd.DataFrame(output_data, columns=cols)
                df_person = df_person.sort_values(by=sort_col, ascending=False).reset_index(drop=True)

                unit_data = [[u, c] for u, c in unit_counts.items()]
                df_unit = pd.DataFrame(unit_data, columns=['單位', '本期通報總數(不限成案)'])
                df_unit = df_unit.sort_values(by='本期通報總數(不限成案)', ascending=False).reset_index(drop=True)
                
                st.session_state['df_person'] = df_person
                st.session_state['df_unit'] = df_unit
                st.session_state['mode_name'] = mode_name
                st.session_state['auto_year'] = auto_year
                st.session_state['calc_done'] = True

            except Exception as e:
                st.error(f"❌ 發生錯誤，請檢查檔案格式。\n詳細錯誤訊息：{e}")

# --- 結果顯示與後續操作區塊 ---
if st.session_state.get('calc_done', False):
    df_person = st.session_state['df_person']
    df_unit = st.session_state['df_unit']
    mode_name = st.session_state['mode_name']
    auto_year = st.session_state['auto_year']
    
    st.info(f"🔎 系統基準年度鎖定為：**{auto_year} 年**，已自動濾除所有跨年度歷史資料。")
    st.success(f"✅ 統計完成！已自動採用「{mode_name}模式」。")

    tab1, tab2 = st.tabs(["👮 個人嘉獎次數統計 (僅計算成案)", "🏢 各單位通報統計 (含所有通報)"])
    
    with tab1:
        st.dataframe(df_person, use_container_width=True)
        csv_buffer_p = io.StringIO()
        df_person.to_csv(csv_buffer_p, index=False, encoding='utf-8-sig')
        filename_p = f"{mode_name}個人嘉獎統計.csv"
        st.download_button(
            label=f"📥 下載個人統計 ({filename_p})",
            data=csv_buffer_p.getvalue(),
            file_name=filename_p,
            mime="text/csv",
            use_container_width=True
        )

    with tab2:
        st.dataframe(df_unit, use_container_width=True)
        csv_buffer_u = io.StringIO()
        df_unit.to_csv(csv_buffer_u, index=False, encoding='utf-8-sig')
        filename_u = f"{mode_name}單位通報件數統計.csv"
        st.download_button(
            label=f"📥 下載單位統計 ({filename_u})",
            data=csv_buffer_u.getvalue(),
            file_name=filename_u,
            mime="text/csv",
            use_container_width=True
        )

    st.markdown("---")
    if st.button("📧 將以上【兩份】統計結果一併寄至我的信箱 (備份)", use_container_width=True):
        with st.spinner("信件與附件處理中，請稍候…"):
            ok, mail_err = send_csv_email(df_person, df_unit, mode_name)
            if ok:
                st.success("✅ 信件發送成功！兩份報表已隨信夾帶，請檢查您的信箱。")
            else:
                st.error(f"❌ 發信失敗: {mail_err}")
else:
    if not (file_tgt and file_src1):
        st.info("請至少上傳「受理明細」與「靜桃清冊」兩個檔案，以啟動統計按鈕。支援 CSV 與 Excel 格式。")
