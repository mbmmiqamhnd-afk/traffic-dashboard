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
def send_stats_email(filename, summary_df, detail_df):
    try:
        if "email" not in st.secrets:
            return False, "未偵測到郵件設定"
        sender = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        msg["Subject"] = f"【全分局時數總彙整備份】{filename.replace('.xlsx', '')}"
        
        body = f"您好：\n\n附件為全分局通過「平假日尖峰時段 + 番號精準雙重過濾」後產出的交通疏導統計總報表。\n發送時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        msg.attach(MIMEText(body, "plain", "utf-8"))
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary_df.to_excel(writer, index=False, sheet_name='分局月彙整總表')
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
    st.title("⏱️ 交通疏導勤務時數彙整系統 (平假日尖峰+番號雙重精準過濾版)")
    st.markdown("---")

    # A. 檔案上傳區
    uploaded_files = st.file_uploader("📂 請一次選取並上傳『全分局所有單位』的當月勤務明細表", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files:
        st.subheader("🎯 1. 交通疏導核銷與過濾條件設定")
        st.info("💡 說明：為了精準剔除同時段「非交通疏導」的守望，除了時間之外，您可以利用【番號白名單】指定只有哪些勤務代號算疏導時數；或利用【番號黑名單】排除無效守望。")
        
        col_wd, col_we = st.columns(2)
        with col_wd:
            wd_input = st.text_input("平常日(週一至五)核銷時段關鍵字", value="06-08, 07-09, 16-18, 17-19")
        with col_we:
            we_input = st.text_input("例假日(週六、日)核銷時段關鍵字", value="10-12, 16-18, 15-17, 11-13")
            
        col_white, col_black = st.columns(2)
        with col_white:
            # 白名單：如果派出所交通崗有固定番號，填在這裡 (若留空則代表不限制，全部番號都檢查)
            include_input = st.text_input("番號白名單 (只有這些番號才算交通疏導，留空代表不限制番號)", value="", placeholder="例如: 01, 02, 03")
        with col_black:
            # 黑名單：明確知道不是交通疏導的番號填在這裡
            exclude_input = st.text_input("番號黑名單 (這些番號的守望一律自動剔除)", value="A, B, C, XA, XB")
            
        wd_whitelist = [t.strip() for t in wd_input.split(',') if t.strip()]
        we_whitelist = [t.strip() for t in we_input.split(',') if t.strip()]
        in_list = [i.strip().upper() for i in include_input.split(',') if i.strip()]
        ex_list = [i.strip().upper() for i in exclude_input.split(',') if i.strip()]

        all_processed_records = []

        # 批次循環處理上傳檔案
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

                # 智慧動態判定該檔案日期為平常日還是假日
                date_digits = "".join(re.findall(r'\d+', file.name))
                is_weekend = False
                current_date_str = ""
                
                if len(date_digits) >= 4:
                    try:
                        md = date_digits[-4:]
                        this_year = datetime.now().year
                        dt = datetime.strptime(f"{this_year}{md}", "%Y%m%d")
                        current_date_str = dt.strftime("%m月%d日")
                        if dt.weekday() in [5, 6]:
                            is_weekend = True
                    except:
                        pass
                
                active_whitelist = we_whitelist if is_weekend else wd_whitelist
                day_type_label = "🔴 假日崗哨" if is_weekend else "🔵 平日崗哨"

                # 智慧定位時段欄位索引
                target_columns = []
                for r in [0, 1, 2]:
                    if r < len(df):
                        header_row = df.iloc[r].astype(str).tolist()
                        for c_idx, cell in enumerate(header_row):
                            cell_clean = cell.replace(' ', '').replace('\n', '')
                            if any(t in cell_clean for t in time_whitelist):
                                if c_idx not in target_columns:
                                    target_columns.append(c_idx)
                
                if not target_columns:
                    target_columns = [2, 12]

                # 讀取每位同仁當天的勤務內容
                for r_idx in range(2, len(df)):
                    row = df.iloc[r_idx]
                    s_code = str(row[0]).strip().upper()
                    
                    # --- 【過濾核心 1：番號黑名單過濾】 ---
                    if s_code in ex_list: continue 
                    
                    # --- 【過濾核心 2：番號白名單限制】 ---
                    # 如果有設定白名單，且該員番號不在白名單內，直接跳過不抓
                    if in_list and (s_code not in in_list): continue

                    name = str(row[1]).replace('\n', '').replace(' ', '')
                    if name in ['nan', 'None', '', '姓名', '合計', '總計']: continue
                    
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
                            "姓名": name, 
                            "核銷時數": h_count,
                            "番號": s_code,
                            "原始檔名": file.name
                        })
            except:
                st.error(f"檔案 {file.name} 讀取失敗")

        # D. 成果輸出與校對區
        if all_processed_records:
            full_raw_df = pd.DataFrame(all_processed_records)
            
            st.divider()
            tab1, tab2 = st.tabs(["🏆 全分局月彙整總表 (造冊專用)", "📝 每日審核明細區 (可手動剔除雜訊)"])
            
            with tab2:
                st.subheader("📝 全分局每日尖峰交通疏導明細")
                st.warning("⚠️ 提示：若發現同時間有非交通疏導的同仁被抓進來，請點選該列最左側的序號，直接按鍵盤 `Delete` 鍵整列刪除，第一頁的月總表會即時扣除時數！")
                
                # 重新排列明細欄位方便承辦人肉眼對帳
                detail_display_df = full_raw_df[["單位", "日期", "類型", "番號", "姓名", "核銷時數", "原始檔名"]]
                edited_df = st.data_editor(detail_display_df, use_container_width=True, num_rows="dynamic", key="global_data_editor")

            with tab1:
                st.subheader("📊 各單位人員月時數核銷總表")
                
                if not edited_df.empty:
                    summary = edited_df.groupby(['單位', '姓名'])['核銷時數'].sum().reset_index()
                    summary.columns = ['單位', '姓名', '總計疏導時數']
                    
                    # 依單位筆畫、姓名筆畫排序 A 到 Z
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
                        fname = f"分局交通疏導總彙整統計_{today}.xlsx"
                        
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            summary.to_excel(writer, index=False, sheet_name='分局月彙整總表')
                            edited_df.to_excel(writer, index=False, sheet_name='人員明細')
                        
                        st.write("### 📥 輸出報表")
                        st.download_button("📥 下載分局月彙整總表 Excel", output.getvalue(), fname, use_container_width=True)
                        
                        st.write("---")
                        if st.button("📧 寄送審核結果至我的信箱", use_container_width=True):
                            with st.spinner("報表發送中..."):
                                ok, err = send_stats_email(fname, summary, edited_df)
                                if ok: st.success("✅ 全分局精準核銷總表已送達信箱！")
                                else: st.error(f"❌ 寄送失敗: {err}")
                else:
                    st.warning("⚠️ 明細已被全數刪除。")
        else:
            st.warning("⚠️ 依目前平假日尖峰與番號規則，未偵測到任何符合規定的守望紀錄。")

if __name__ == "__main__":
    run_app()
