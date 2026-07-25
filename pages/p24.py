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
from collections import Counter

# ==========================================
# 輔助函式：車號標準化
# ==========================================
def normalize_plate(plate):
    if pd.isna(plate):
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(plate)).upper()

# ==========================================
# 輔助函式：智慧解析年度與月份 (相容公部門常見格式)
# ==========================================
def extract_year_month(date_val):
    if pd.isna(date_val): return None, None
    
    # 處理 Excel 原生日期格式
    if isinstance(date_val, datetime) or isinstance(date_val, pd.Timestamp):
        y = date_val.year
        if y > 1911: y -= 1911 # 轉換為民國年
        return y, date_val.month
    
    s = str(date_val).strip()
    
    # 嘗試解析以符號分隔的日期 (例: 115/05/12, 115-06-21, 115.8.1)
    parts = re.split(r'[/.-]', s.split(' ')[0]) # 若有時間則先切掉空白後面的時間
    if len(parts) >= 2:
        try:
            y = int(parts[0])
            m = int(parts[1])
            if y > 1911: y -= 1911
            if 1 <= m <= 12: return y, m
        except: pass
    
    # 嘗試解析純文字 (例: 115年5月12日)
    match = re.search(r'(\d{2,4})\s*年\s*(\d{1,2})\s*月', s)
    if match:
        try:
            y = int(match.group(1))
            m = int(match.group(2))
            if y > 1911: y -= 1911
            if 1 <= m <= 12: return y, m
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
            f"2. 各單位通報件數統計 (含所有通報紀錄)\n\n"
            f"本信件由交通執法自動化分析引擎發送。"
        )
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        csv_person = df_person.to_csv(index=False, encoding='utf-8-sig')
        part1 = MIMEBase("application", "csv")
        part1.set_payload(csv_person.encode('utf-8-sig'))
        encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(f'{mode_name}個人嘉獎統計.csv')}")
        msg.attach(part1)

        csv_unit = df_unit.to_csv(index=False, encoding='utf-8-sig')
        part2 = MIMEBase("application", "csv")
        part2.set_payload(csv_unit.encode('utf-8-sig'))
        encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(f'{mode_name}單位通報件數統計.csv')}")
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
st.title("🚓 噪音改裝車輛績效與嘉獎統計系統 (全自動感知版)")

st.markdown("""
💡 **系統具備全方位自動偵測能力：**
*   **時間智慧感知**：系統會自動讀取「靜桃清冊」的通報時間，並以此決定當前統計為上半年或下半年。
*   **精準過濾**：自動剔除不在該半年度區間的通報紀錄。
""")

# --- 側邊欄設定區 ---
st.sidebar.header("⚙️ 參數設定")
st.sidebar.markdown("請設定各檔案要從**第幾列**開始讀取資料：")
start_row_src1 = st.sidebar.number_input("[靜桃清冊] 起始列", min_value=2, value=2, step=1)
start_row_tgt = st.sidebar.number_input("[受理明細] 起始列", min_value=2, value=2, step=1)
start_row_src2 = st.sidebar.number_input("[前期明細] 起始列 (選填，用於跨期合併)", min_value=2, value=2, step=1)

st.sidebar.markdown("---")
st.sidebar.markdown("請確認 [靜桃清冊] 的**欄位位置** (A欄=0, B欄=1, C欄=2...)：")
col_date = st.sidebar.number_input("通報時間 所在欄位", value=1, step=1)    # 預設 B 欄
col_plate = st.sidebar.number_input("車號 所在欄位", value=4, step=1)    # 預設 E 欄
col_name = st.sidebar.number_input("通報人 所在欄位", value=6, step=1)  # 預設 G 欄
col_unit = st.sidebar.number_input("單位 所在欄位", value=7, step=1)    # 預設 H 欄

# --- 檔案上傳區與工作表選擇 ---
st.markdown("### 📥 上傳資料檔案")
col1, col2, col3 = st.columns(3)

with col1: 
    file_tgt = st.file_uploader("1. 上傳 [受理明細] (必填)", type=['csv', 'xlsx'])
    sheet_tgt = None
    if file_tgt and file_tgt.name.endswith('.xlsx'):
        xls_tgt = pd.ExcelFile(file_tgt, engine='openpyxl')
        sheet_tgt = st.selectbox("📂 選擇工作表", xls_tgt.sheet_names, index=get_default_sheet_index(xls_tgt.sheet_names, ['受理明細']), key='tgt')

with col2: 
    file_src1 = st.file_uploader("2. 上傳 [靜桃清冊] (必填)", type=['csv', 'xlsx'])
    sheet_src1 = None
    if file_src1 and file_src1.name.endswith('.xlsx'):
        xls_src1 = pd.ExcelFile(file_src1, engine='openpyxl')
        sheet_src1 = st.selectbox("📂 選擇工作表", xls_src1.sheet_names, index=get_default_sheet_index(xls_src1.sheet_names, ['靜桃']), key='src1')

with col3: 
    file_src2 = st.file_uploader("3. 上傳 [前期明細] (選填)", type=['csv', 'xlsx'])
    sheet_src2 = None
    if file_src2 and file_src2.name.endswith('.xlsx'):
        xls_src2 = pd.ExcelFile(file_src2, engine='openpyxl')
        sheet_src2 = st.selectbox("📂 選擇工作表", xls_src2.sheet_names, index=get_default_sheet_index(xls_src2.sheet_names, ['嘉獎', '明細']), key='src2')

# --- 執行統計區塊 ---
if file_tgt and file_src1:
    if st.button("🚀 開始執行績效與嘉獎統計", type="primary"):
        with st.spinner('資料讀取與處理中...'):
            try:
                has_history_file = file_src2 is not None

                df_tgt = load_data(file_tgt, sheet_tgt)
                df_src1 = load_data(file_src1, sheet_src1)

                df_tgt_filtered = df_tgt.iloc[start_row_tgt - 2:]
                df_src1_filtered = df_src1.iloc[start_row_src1 - 2:]

                # ----------------------------------------------------
                # 1. 智慧掃描靜桃清冊：自動偵測年度與上下半年
                # ----------------------------------------------------
                detected_years = []
                detected_halfs = []
                
                valid_scan_count = 0
                for _, row in df_src1_filtered.iterrows():
                    if len(row) > col_date:
                        y, m = extract_year_month(row.iloc[col_date])
                        if y and m:
                            detected_years.append(y)
                            detected_halfs.append(1 if m <= 6 else 2)
                            valid_scan_count += 1
                            if valid_scan_count >= 100: # 取前100筆有效資料進行判斷已足夠精準
                                break
                
                auto_year = "115"
                is_second_half_mode = False
                if detected_years:
                    auto_year = str(Counter(detected_years).most_common(1)[0][0])
                    most_common_half = Counter(detected_halfs).most_common(1)[0][0]
                    is_second_half_mode = (most_common_half == 2)
                
                target_months = [7, 8, 9, 10, 11, 12] if is_second_half_mode else [1, 2, 3, 4, 5, 6]
                mode_name = f"{auto_year}年下半年" if is_second_half_mode else f"{auto_year}年上半年"

                # ----------------------------------------------------
                # 2. 建立車號對應字典 & 單位通報件數統計
                # ----------------------------------------------------
                plate_info = {}
                unit_counts = {}
                max_needed_col = max(col_plate, col_name, col_unit, col_date)
                
                for _, row in df_src1_filtered.iterrows():
                    if len(row) > max_needed_col:
                        # 日期嚴格過濾：依據自動判定出的半年度進行剔除
                        _, m = extract_year_month(row.iloc[col_date])
                        if m is not None and m not in target_months:
                            continue

                        plate = normalize_plate(row.iloc[col_plate])
                        name = str(row.iloc[col_name]).strip()
                        unit = str(row.iloc[col_unit]).strip()
                        
                        if unit and unit != 'nan':
                            unit_counts[unit] = unit_counts.get(unit, 0) + 1
                        
                        if plate and name and name != 'nan':
                            plate_info[plate] = {'name': name, 'unit': unit}

                # ----------------------------------------------------
                # 3. 計算個人本期「成案」件數
                # ----------------------------------------------------
                current_counts = {}
                for _, row in df_tgt_filtered.iterrows():
                    if len(row) > 1:
                        doc_num = str(row.iloc[0])
                        plate = normalize_plate(row.iloc[1])
                        
                        if "龍警分交字" in doc_num and plate in plate_info:
                            reporter = plate_info[plate]['name']
                            current_counts[reporter] = current_counts.get(reporter, 0) + 1

                # ----------------------------------------------------
                # 4. 處理前期資料 (如果有上傳的話)
                # ----------------------------------------------------
                history_map = {}
                if has_history_file:
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

                # ----------------------------------------------------
                # 5. 輸出個人嘉獎資料表
                # ----------------------------------------------------
                person_data = []
                for name, count_current in current_counts.items():
                    count_history = history_map.get(name, 0) if has_history_file else 0
                    count_total = count_current + count_history
                    reward_count = count_total // 6
                    
                    if has_history_file:
                        person_data.append([name, count_current, count_history, count_total, reward_count])
                    else:
                        person_data.append([name, count_current, count_total, reward_count])

                if has_history_file:
                    cols_p = ['通報人(A)', '本期件數(成案)', '前期件數(成案)', '合計件數(成案)', '嘉獎數']
                else:
                    cols_p = ['通報人(A)', '本期件數(成案)', '合計件數(成案)', '嘉獎數']

                df_person = pd.DataFrame(person_data, columns=cols_p)
                df_person = df_person.sort_values(by='嘉獎數', ascending=False).reset_index(drop=True)

                # ----------------------------------------------------
                # 6. 輸出各單位績效資料表
                # ----------------------------------------------------
                unit_data = [[u, c] for u, c in unit_counts.items()]
                df_unit = pd.DataFrame(unit_data, columns=['單位', '本期通報總數(不限成案)'])
                df_unit = df_unit.sort_values(by='本期通報總數(不限成案)', ascending=False).reset_index(drop=True)
                
                st.session_state['df_person'] = df_person
                st.session_state['df_unit'] = df_unit
                st.session_state['mode_name'] = mode_name
                st.session_state['calc_done'] = True

            except Exception as e:
                st.error(f"❌ 發生錯誤，請確認側邊欄的設定是否正確。\n詳細錯誤訊息：{e}")

# --- 結果顯示與後續操作區塊 ---
if st.session_state.get('calc_done', False):
    df_person = st.session_state['df_person']
    df_unit = st.session_state['df_unit']
    mode_name = st.session_state['mode_name']
    
    st.success(f"✅ 系統自動偵測通報時間為「{mode_name}」，已為您精準過濾並計算完成！")
    
    tab1, tab2 = st.tabs(["👮 個人嘉獎次數統計 (僅計算成案)", "🏢 各單位通報統計 (含所有通報)"])
    
    with tab1:
        st.dataframe(df_person, use_container_width=True)
        csv_p = io.StringIO()
        df_person.to_csv(csv_p, index=False, encoding='utf-8-sig')
        st.download_button(
            label=f"📥 下載個人統計 ({mode_name}個人嘉獎統計.csv)",
            data=csv_p.getvalue(),
            file_name=f"{mode_name}個人嘉獎統計.csv",
            mime="text/csv"
        )
        
    with tab2:
        st.dataframe(df_unit, use_container_width=True)
        csv_u = io.StringIO()
        df_unit.to_csv(csv_u, index=False, encoding='utf-8-sig')
        st.download_button(
            label=f"📥 下載單位統計 ({mode_name}單位通報件數統計.csv)",
            data=csv_u.getvalue(),
            file_name=f"{mode_name}單位通報件數統計.csv",
            mime="text/csv"
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
        st.info("請上傳「受理明細」與「靜桃清冊」檔案，以啟動統計按鈕。第三個檔案為選填。")
