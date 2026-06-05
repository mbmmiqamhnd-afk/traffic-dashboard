import streamlit as st
import pandas as pd
import io
import re
import smtplib
import urllib.parse as _ul
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from menu import show_sidebar

# --- 1. 頁面配置 ---
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

# 呼叫側邊欄導航
show_sidebar()

# --- 2. 郵件發送功能 ---
def send_stats_email(filename, summary_df, detail_df, report_prefix):
    try:
        if "email" not in st.secrets:
            return False, "未偵測到郵件設定"
        sender = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        msg["Subject"] = f"【系統備份】{filename.replace('.xlsx', '')}"
        
        body = f"您好：\n\n附件為{report_prefix}通過「分單位獨立課表過濾」後產出的交通疏導統計報表。\n發送時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        msg.attach(MIMEText(body, "plain", "utf-8"))
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary_df.to_excel(writer, index=False, sheet_name='月彙整總表')
            detail_df.to_excel(writer, index=False, sheet_name='人員核銷明細')
        
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(output.getvalue())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(filename)}")
        msg.attach(part)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# --- 3. 主程式邏輯 ---
def run_app():
    st.title("⏱️ 交通疏導勤務時數彙整系統 (分單位獨立課表版)")
    st.markdown("---")

    # A. 檔案上傳區
    uploaded_files = st.file_uploader("📂 請一次選取並上傳勤務明細檔 (可單一單位，也可全分局批次)", accept_multiple_files=True, type=['csv', 'xlsx'])

    # B. 分單位規則矩陣表
    st.subheader("🎯 1. 分單位精準核銷規則矩陣")
    st.info("💡 說明：每個單位都有自己的專屬規則！如果某單位假日不排交通崗，請直接把該單位的「假日核銷時段」內容清空，系統就會自動忽略該所假日的守望。")
    
    # 建立預設的分局各單位規則表
    default_rules = pd.DataFrame([
        {"單位": "龍潭派出所", "平日核銷時段": "06-07, 07-08, 16-17, 17-18", "假日核銷時段": "10-11, 11-12, 16-17, 17-18", "專屬番號(白名單)": ""},
        {"單位": "中興派出所", "平日核銷時段": "06-07, 07-08, 16-17, 17-18", "假日核銷時段": "10-11, 11-12, 16-17, 17-18", "專屬番號(白名單)": ""},
        {"單位": "聖亭派出所", "平日核銷時段": "06-07, 07-08, 16-17, 17-18", "假日核銷時段": "", "專屬番號(白名單)": ""},
        {"單位": "石門派出所", "平日核銷時段": "06-07, 07-08, 16-17, 17-18", "假日核銷時段": "10-11, 11-12, 16-17, 17-18", "專屬番號(白名單)": ""},
        {"單位": "高平派出所", "平日核銷時段": "06-07, 07-08, 16-17, 17-18", "假日核銷時段": "10-11, 11-12, 16-17, 17-18", "專屬番號(白名單)": ""},
        {"單位": "三和派出所", "平日核銷時段": "06-07, 07-08, 16-17, 17-18", "假日核銷時段": "", "專屬番號(白名單)": ""},
        {"單位": "交通分隊",   "平日核銷時段": "06-07, 07-08, 16-17, 17-18", "假日核銷時段": "10-11, 11-12, 16-17, 17-18", "專屬番號(白名單)": ""}
    ])

    edited_rules_df = st.data_editor(default_rules, use_container_width=True, hide_index=True, key="rules_editor")
    
    st.divider()
    
    exclude_input = st.text_input("🛑 全域番號黑名單 (只要是這些番號的守望，各單位一律剔除)", value="A, B, C, XA, XB")
    ex_list = [i.strip().upper() for i in exclude_input.split(',') if i.strip()]

    if uploaded_files:
        all_processed_records = []
        rules_dict = edited_rules_df.set_index('單位').to_dict('index')

        for file in uploaded_files:
            try:
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=None, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, header=None, encoding='cp950')
                else:
                    df = pd.read_excel(file, header=None)
                
                # 單位名稱精準辨識
                match = re.search(r'(龍潭|中興|石門|高平|三和|聖亭|交通)(派出所|分隊|所)?', file.name)
                if match:
                    base_name = match.group(1)
                    u_name = "交通分隊" if base_name == '交通' else base_name + "派出所"
                else:
                    temp = re.sub(r'\d+', '', file.name)
                    temp = re.sub(r'(交通|疏導|勤務|明細|彙整|統計|執行|時數|工作|紀錄|表|年|月|日|\.xlsx|\.csv)', '', temp)
                    u_name = temp.strip(' _-()（）')
                    if not u_name: u_name = "未知單位"

                # 智慧動態判定日期與平假日
                date_digits = "".join(re.findall(r'\d+', file.name))
                is_weekend = False
                current_date_str = ""
                
                if len(date_digits) >= 4:
                    try:
                        md = date_digits[-4:]
                        dt = datetime.strptime(f"2026{md}", "%Y%m%d")
                        current_date_str = dt.strftime("%m月%d日")
                        if dt.weekday() in [5, 6]: 
                            is_weekend = True
                    except:
                        pass
                
                # 調用該單位的專屬規則
                unit_rule = rules_dict.get(u_name)
                if unit_rule:
                    wd_str = str(unit_rule['平日核銷時段']) if pd.notna(unit_rule['平日核銷時段']) else ""
                    we_str = str(unit_rule['假日核銷時段']) if pd.notna(unit_rule['假日核銷時段']) else ""
                    in_str = str(unit_rule['專屬番號(白名單)']) if pd.notna(unit_rule['專屬番號(白名單)']) else ""
                else:
                    wd_str, we_str, in_str = "", "", ""

                wd_whitelist = [t.strip().replace(' ', '') for t in wd_str.split(',') if t.strip()]
                we_whitelist = [t.strip().replace(' ', '') for t in we_str.split(',') if t.strip()]
                in_list = [i.strip().upper() for i in in_str.split(',') if i.strip()]

                if is_weekend:
                    active_whitelist = we_whitelist
                    is_whitelisted_empty = (len(we_str.strip()) == 0)
                    day_type_label = "🔴 假日崗哨"
                else:
                    active_whitelist = wd_whitelist
                    is_whitelisted_empty = (len(wd_str.strip()) == 0)
                    day_type_label = "🔵 平日崗哨"

                # 定位結構起點
                header_row_idx = 0
                data_start_idx = 2
                for i in range(len(df)):
                    row_str = "".join(df.iloc[i].astype(str).tolist())
                    if "姓名" in row_str or "-" in row_str or ":" in row_str or "|" in row_str:
                        header_row_idx = i
                        data_start_idx = i + 1
                        break

                target_columns = []
                
                if not is_whitelisted_empty:
                    header_row_list = df.iloc[header_row_idx].astype(str).tolist()
                    for c_idx, cell in enumerate(header_row_list):
                        cell_clean = cell.replace(' ', '').replace('\n', '').replace('\r', '')
                        cell_clean = cell_clean.replace('|', '-').replace('~', '-').replace('～', '-').strip()
                        
                        if any(t in cell_clean for t in active_whitelist):
                            if c_idx not in target_columns:
                                target_columns.append(c_idx)
                    
                    if not target_columns and len(active_whitelist) > 0:
                        target_columns = [2, 12]
                else:
                    pass

                for r_idx in range(data_start_idx, len(df)):
                    row = df.iloc[r_idx]
                    if pd.isna(row[0]) or pd.isna(row[1]): continue
                    
                    s_code = str(row[0]).strip().upper()
                    if s_code in ex_list: continue 
                    if in_list and (s_code not in in_list): continue

                    name = str(row[1]).replace('\n', '').replace(' ', '')
                    if name in ['nan', 'None', '', '姓名', '合計', '總計', '重疊']: continue
                    
                    h_count = 0
                    for c_idx in target_columns:
                        if c_idx < len(row):
                            cell_value = str(row[c_idx]).replace('\n', '').replace(' ', '')
                            if "守望" in cell_value: 
                                h_count += 1
                    
                    if h_count > 0:
                        all_processed_records.append({
                            "單位": u_name, 
                            "日期": current_date_str if current_date_str else "未識別",
                            "類型": day_type_label,
                            "番號": s_code,
                            "姓名": name, 
                            "核銷時數": h_count,
                            "原始檔名": file.name
                        })
            except Exception as e:
                st.error(f"檔案 {file.name} 解析失敗，原因: {str(e)}")

        if all_processed_records:
            full_raw_df = pd.DataFrame(all_processed_records)
            
            st.divider()
            tab1, tab2 = st.tabs(["🏆 月彙整總表 (造冊專用)", "📝 每日審核明細區 (可手動剔除雜訊)"])
            
            with tab2:
                st.subheader("📝 每日尖峰交通疏導明細")
                detail_display_df = full_raw_df[["單位", "日期", "類型", "番號", "姓名", "核銷時數", "原始檔名"]]
                edited_df = st.data_editor(detail_display_df, use_container_width=True, num_rows="dynamic", key="global_data_editor")

            with tab1:
                st.subheader("📊 各單位人員月時數核銷總表")
                
                if not edited_df.empty:
                    summary = edited_df.groupby(['單位', '姓名'])['核銷時數'].sum().reset_index()
                    summary.columns = ['單位', '姓名', '總計疏導時數']
                    
                    summary = summary.sort_values(
                        by=['單位', '姓名'], 
                        ascending=[True, True],
                        key=lambda col: col.map(lambda x: str(x).encode('big5', errors='ignore'))
                    ).reset_index(drop=True)
                    
                    col_result, col_action = st.columns([3, 2])
                    
                    with col_result:
                        st.dataframe(summary, use_container_width=True, hide_index=True)

                    with col_action:
                        today = datetime.now().strftime('%m%d')
                        
                        # --- 【動態命名邏輯】自動判斷目前畫面上處理了幾個單位 ---
                        unique_units = summary['單位'].unique()
                        if len(unique_units) == 1:
                            report_prefix = unique_units[0] # 如果只有1家，字首為該單位名稱 (例: 聖亭派出所)
                        else:
                            report_prefix = "全分局" # 如果大於1家，字首自動變為全分局
                            
                        fname = f"{report_prefix}_交通疏導彙整統計_{today}.xlsx"
                        
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            summary.to_excel(writer, index=False, sheet_name='月彙整總表')
                            edited_df.to_excel(writer, index=False, sheet_name='人員明細')
                        
                        st.write("### 📥 輸出報表")
                        st.download_button(f"📥 下載 {report_prefix} 總表 Excel", output.getvalue(), fname, use_container_width=True)
                        
                        st.write("---")
                        if st.button(f"📧 寄送 {report_prefix} 報表至信箱", use_container_width=True):
                            with st.spinner("報表發送中..."):
                                ok, err = send_stats_email(fname, summary, edited_df, report_prefix)
                                if ok: st.success(f"✅ {report_prefix} 精準核銷總表已送達信箱！")
                                else: st.error(f"❌ 寄送失敗: {err}")
                else:
                    st.warning("⚠️ 明細已被全數刪除。")
        else:
            st.warning("⚠️ 依目前配置規則，未偵測到任何符合規定的守望紀錄。")

if __name__ == "__main__":
    run_app()
