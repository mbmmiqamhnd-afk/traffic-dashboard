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

# --- 1. 頁面配置 (必須在最頂端) ---
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

# --- 2. 郵件發送功能 ---
def send_stats_email(filename, summary_df, detail_df):
    try:
        if "email" not in st.secrets:
            return False, "未偵測到郵件設定 (secrets.toml)"
        
        sender = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender
        msg["Subject"] = f"【系統發送】{filename.replace('.xlsx', '')}"
        
        body = f"您好：\n\n附件為交通疏導時數統計報表。\n發送時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
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
    st.title("⏱️ 交通疏導勤務時數彙整 (分單位終極版)")
    st.markdown("---")

    uploaded_files = st.file_uploader("請上傳『所有單位』的勤務明細檔", accept_multiple_files=True, type=['csv', 'xlsx'])

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
                
                u_name = re.split(r'\d+', file.name)[0].strip()
                if not u_name: u_name = "未定義單位"
                units_found.add(u_name)
                all_raw_data.append({"unit": u_name, "df": df, "filename": file.name})
            except:
                st.error(f"檔案 {file.name} 讀取失敗")

        st.sidebar.header("🏢 單位別規則設定")
        target_unit = st.sidebar.selectbox("請選擇要設定與校對的單位", sorted(list(units_found)))
        
        st.sidebar.divider()
        st.sidebar.subheader(f"📍 {target_unit} 專屬設定")
        
        u_exclude = st.sidebar.text_input(f"排除番號 ({target_unit})", value="A, B, C", key=f"ex_{target_unit}")
        u_am = st.sidebar.text_input(f"上午尖峰欄位索引 ({target_unit})", value="2, 3", key=f"am_{target_unit}")
        u_pm = st.sidebar.text_input(f"下午尖峰欄位索引 ({target_unit})", value="12, 13", key=f"pm_{target_unit}")
        
        ex_list = [i.strip().upper() for i in u_exclude.split(',') if i.strip()]
        try:
            p_indices = [int(i.strip()) for i in (u_am + "," + u_pm).split(',') if i.strip()]
        except:
            p_indices = [2, 3, 12, 13]

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
                            if "守望" in str(row[c_idx]).replace('\n', ''): 
                                h_count += 1
                    
                    if h_count > 0:
                        processed_records.append({
                            "單位": item["unit"], "姓名": name, "當日尖峰時數": h_count,
                            "番號": s_code, "日期來源": item["filename"]
                        })

        if processed_records:
            final_raw_df = pd.DataFrame(processed_records)
            st.subheader(f"📝 {target_unit} - 人員明細校對")
            edited_df = st.data_editor(final_raw_df, use_container_width=True, num_rows="dynamic", key=f"editor_{target_unit}")

            if not edited_df.empty:
                summary = edited_df.groupby(['單位', '姓名'])['當日尖峰時數'].sum().reset_index()
                summary.columns = ['單位', '姓名', '總計尖峰時數']
                
                st.divider()
                st.subheader(f"📊 {target_unit} - 月彙整結果")
                st.dataframe(summary.sort_values('總計尖峰時數', ascending=False), use_container_width=True, hide_index=True)

                today = datetime.now().strftime('%m%d')
                fname = f"{target_unit}_交通疏導統計_{today}.xlsx"
                
                st.divider()
                col1, col2 = st.columns(2)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月彙整總表')
                    edited_df.to_excel(writer, index=False, sheet_name='人員核銷明細')
                
                col1.download_button(f"📥 下載 {fname}", output.getvalue(), fname, use_container_width=True)
                if col2.button(f"📧 寄送 {target_unit} 結果至我的信箱", use_container_width=True):
                    with st.spinner("報表寄送中..."):
                        ok, err = send_stats_email(fname, summary, edited_df)
                        if ok: st.success("✅ 郵件發送成功！")
                        else: st.error(f"❌ 郵件失敗: {err}")
        else:
            st.warning(f"⚠️ 單位『{target_unit}』找不到符合條件的守望資料。")

    # --- 修正後的返回導航區 ---
    st.divider()
    col_back, _ = st.columns([1, 2])
    with col_back:
        if st.button("🏠 返回系統主選單", use_container_width=True, type="primary"):
            # 嘗試多種可能的主頁面檔案名稱
            for main_file in ["main.py", "Home.py", "app.py", "首頁.py"]:
                try:
                    st.switch_page(main_file)
                    break
                except:
                    continue
            else:
                # 若程式碼導航失敗，則給予提示
                st.info("💡 請點擊左側邊欄頂部的首頁選項返回。")

if __name__ == "__main__":
    run_app()
