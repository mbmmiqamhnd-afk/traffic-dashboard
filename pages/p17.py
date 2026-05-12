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
        msg["Subject"] = f"【系統備份】{filename.replace('.xlsx', '')}"
        
        body = f"您好：\n\n附件為修正後的交通疏導統計報表。\n發送時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
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
    st.title("⏱️ 交通疏導勤務時數彙整系統")
    st.markdown("---")

    # A. 檔案上傳區
    uploaded_files = st.file_uploader("📂 請上傳各單位勤務明細檔 (可批次拖入整個月份的檔案)", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files:
        all_raw_data = []
        units_found = set()
        
        for file in uploaded_files:
            try:
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=None, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, header=None, encoding='cp950')
                else:
                    df = pd.read_excel(file, header=None)
                
                # --- 精準提取單位名稱邏輯 ---
                # 1. 移除日期與數字
                temp_name = re.sub(r'\d+', '', file.name)
                # 2. 移除副檔名
                temp_name = temp_name.replace('.xlsx', '').replace('.csv', '')
                # 3. 移除常見冗餘字眼 (保持單位欄位純淨)
                temp_name = re.sub(r'(交通|勤務|明細|彙整|統計|執行|時數|工作|紀錄)', '', temp_name)
                # 4. 去除多餘空格與特殊符號
                u_name = temp_name.strip(' _-()（）')
                
                if not u_name: u_name = "未知單位"
                units_found.add(u_name)
                
                all_raw_data.append({"unit": u_name, "df": df, "filename": file.name})
            except:
                st.error(f"檔案 {file.name} 讀取失敗")

        st.divider()
        
        # B. 主畫面設定區塊
        st.subheader("🏢 1. 選擇單位與設定規則")
        
        col_unit, col_ex, col_am, col_pm = st.columns([2, 2, 1, 1])
        
        with col_unit:
            target_unit = st.selectbox("🎯 請選擇要校對的單位", sorted(list(units_found)))
        with col_ex:
            u_exclude = st.text_input(f"排除番號 ({target_unit})", value="A, B, C", key=f"ex_{target_unit}")
        with col_am:
            u_am = st.text_input("上午尖峰索引", value="2", key=f"am_{target_unit}")
        with col_pm:
            u_pm = st.text_input("下午尖峰索引", value="12", key=f"pm_{target_unit}")
        
        ex_list = [i.strip().upper() for i in u_exclude.split(',') if i.strip()]
        try:
            p_indices = [int(i.strip()) for i in (u_am + "," + u_pm).split(',') if i.strip()]
        except:
            p_indices = [2, 12]

        # C. 解析資料
        processed_records = []
        for item in all_raw_data:
            if item["unit"] == target_unit:
                df = item["df"]
                for r_idx in range(2, len(df)):
                    row = df.iloc[r_idx]
                    s_code = str(row[0]).strip().upper()
                    if s_code in ex_list: continue
                    
                    name = str(row[1]).replace('\n', '').replace(' ', '')
                    if name in ['nan', 'None', '', '姓名', '合計', '總計']: continue
                    
                    h_count = 0
                    for c_idx in p_indices:
                        if c_idx < len(row):
                            if "守望" in str(row[c_idx]).replace('\n', ''): h_count += 1
                    
                    if h_count > 0:
                        processed_records.append({
                            "單位": item["unit"], 
                            "姓名": name, 
                            "當日尖峰時數": h_count,
                            "番號": s_code, 
                            "日期來源": item["filename"]
                        })

        # D. 展示、編輯與輸出
        st.divider()
        st.subheader(f"📝 2. {target_unit} - 人員核銷明細")
        
        if processed_records:
            final_raw_df = pd.DataFrame(processed_records)
            st.caption("💡 點選列首後按 `Delete` 可刪除多排人員。")
            
            edited_df = st.data_editor(final_raw_df, use_container_width=True, num_rows="dynamic", key=f"editor_{target_unit}")

            if not edited_df.empty:
                summary = edited_df.groupby(['單位', '姓名'])['當日尖峰時數'].sum().reset_index()
                summary.columns = ['單位', '姓名', '總計尖峰時數']
                
                col_result, col_action = st.columns([3, 2])
                
                with col_result:
                    st.dataframe(summary.sort_values('總計尖峰時數', ascending=False), use_container_width=True, hide_index=True)

                with col_action:
                    today = datetime.now().strftime('%m%d')
                    fname = f"{target_unit}_交通疏導統計_{today}.xlsx"
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        summary.to_excel(writer, index=False, sheet_name='月彙整')
                        edited_df.to_excel(writer, index=False, sheet_name='明細')
                    
                    st.write("### 📥 輸出報表")
                    st.download_button(f"📥 下載 {target_unit} 統計 Excel", output.getvalue(), fname, use_container_width=True)
                    
                    st.write("---")
                    if st.button(f"📧 寄送 {target_unit} 報表至信箱", use_container_width=True):
                        with st.spinner("發送中..."):
                            ok, err = send_stats_email(fname, summary, edited_df)
                            if ok: st.success("✅ 郵件成功寄出")
                            else: st.error(f"❌ 寄送失敗: {err}")
        else:
            st.warning(f"⚠️ 在目前規則下，找不到『{target_unit}』的守望紀錄。")

if __name__ == "__main__":
    run_app()
