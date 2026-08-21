import streamlit as st
import pandas as pd
import numpy as np
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import urllib.parse as _ul

# 載入自訂側邊欄
try:
    import menu
except ImportError:
    pass

# ==========================================
# 0. 輔助函式：發送單一檔案 Email
# ==========================================
def send_single_file_email(file_bytes, file_name, mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
    """使用 st.secrets 設定檔發送夾帶報表的電子郵件"""
    try:
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender  # 寄給自己
        msg["Subject"] = f"🔎 偽變造車牌專案敘獎統計 - {file_name}"
        
        body_text = f"長官您好，\n\n系統已自動產出相關檔案，附件為最新的【{file_name}】。\n\n本信件由交通執法自動化分析引擎發送。"
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        # 解析 MIME 類型
        main_type, sub_type = mime_type.split('/') if '/' in mime_type else ("application", "octet-stream")
        part = MIMEBase(main_type, sub_type)
        part.set_payload(file_bytes.getvalue())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(file_name)}")
        msg.attach(part)

        # 透過 SMTP 發送
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# ==========================================
# 1. 資料處理與統計邏輯
# ==========================================
def process_traffic_data(file):
    """讀取並清洗自選匯出 Excel 資料 (動態尋找標題列，防呆升級版)"""
    try:
        df = pd.read_excel(file, sheet_name="案件明細", header=None)
        
        header_row_index = None
        for idx, row in df.iterrows():
            row_values = [str(val).strip() for val in row.values]
            if '單號' in row_values and '舉發員警1' in row_values:
                header_row_index = idx
                break
                
        if header_row_index is None:
            st.error("找不到資料標題列！請確認上傳的檔案工作表【案件明細】中是否包含『單號』與『舉發員警1』。")
            return None
            
        df.columns = df.iloc[header_row_index]
        df = df.iloc[header_row_index + 1:].reset_index(drop=True)
        
        if df.empty:
            st.warning("系統判定此檔案中沒有任何案件明細資料。")
            return None
        
        cols_to_keep = ['單號', '簡式車種名稱', '違規法條1', '違規事實1', '入案日', '舉發員警1']
        missing_cols = [col for col in cols_to_keep if col not in df.columns]
        if missing_cols:
            st.error(f"上傳的檔案缺少以下必要欄位：{', '.join(missing_cols)}，請確認自選匯出時是否有勾選。")
            return None
            
        df = df[cols_to_keep].dropna(how='all')
        return df
    
    except Exception as e:
        st.error(f"檔案解析失敗，錯誤訊息：{str(e)}")
        return None

def calculate_merits_for_officer(group):
    """計算單一員警的預估嘉獎次數，並嚴格依照獎勵名目拆分"""
    group = group.sort_values(by='入案日')
    
    fake_plate_cases = 0   
    fake_plate_merits = 0  
    
    other_plate_cases = 0  
    other_plate_merits = 0 
    
    tickets = []
    
    for idx, row in group.iterrows():
        violation = str(row['違規事實1'])
        vehicle = str(row['簡式車種名稱'])
        tickets.append(str(row['單號']))
        
        current_total_merits = fake_plate_merits + other_plate_merits
        multiplier = 2 if current_total_merits >= 9 else 1
        
        if '偽造' in violation or '變造' in violation:
            fake_plate_cases += 1
            if '汽車' in vehicle:
                fake_plate_merits += 2 * multiplier
            else:
                fake_plate_merits += 1 * multiplier
                
        elif '他車' in violation:
            other_plate_cases += 1
            if other_plate_cases % 2 == 0:
                other_plate_merits += 1 * multiplier
                
    return pd.Series({
        '【偽變造車牌】件數': fake_plate_cases,
        '【偽變造車牌】嘉獎數': fake_plate_merits,
        '【懸掛他車號牌】件數': other_plate_cases,
        '【懸掛他車號牌】嘉獎數': other_plate_merits,
        '總件數合計': fake_plate_cases + other_plate_cases,
        '總嘉獎合計': fake_plate_merits + other_plate_merits,
        '舉發單號明細': ", ".join(tickets)
    })

# ==========================================
# 2. 主程式介面
# ==========================================
def main():
    st.set_page_config(page_title="偽變造車牌專案 敘獎統計", layout="wide")
    
    try:
        menu.show_sidebar()
    except Exception as e:
        st.sidebar.error("無法載入側邊欄，請確認根目錄下有 menu.py")
        
    st.title("🔎 偽變造車牌專案 - 自動敘獎統計系統")
    st.markdown("本模組專為計算「加強取締查緝偽(變)造車牌及違法(規)權利車」專案期間出力人員敘獎所設計。")
    st.divider()
    
    uploaded_file = st.file_uploader("請上傳『自選匯出.xlsx』(資料來源需包含簡式車種名稱與違規事實)", type=["xlsx"])
    
    if uploaded_file is not None:
        with st.spinner("資料處理中，請稍候..."):
            df = process_traffic_data(uploaded_file)
            
        if df is not None:
            with st.expander("📄 檢視原始案件明細", expanded=False):
                st.dataframe(df, use_container_width=True)
            
            st.subheader("📊 員警專案敘獎統計表 (依名目區分)")
            
            merit_stats = df.groupby('舉發員警1').apply(calculate_merits_for_officer).reset_index()
            merit_stats = merit_stats.sort_values(by=['總嘉獎合計', '總件數合計'], ascending=[False, False]).reset_index(drop=True)
            
            styled_df = (merit_stats.style
                         .background_gradient(subset=['【偽變造車牌】嘉獎數'], cmap='Reds')
                         .background_gradient(subset=['【懸掛他車號牌】嘉獎數'], cmap='Blues'))
            
            st.dataframe(styled_df, use_container_width=True)
            
            # 建立記憶體中的 Excel 檔案
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                merit_stats.to_excel(writer, index=False, sheet_name='敘獎統計表')
                df_sorted = df.sort_values(by=['舉發員警1', '入案日']).reset_index(drop=True)
                df_sorted.to_excel(writer, index=False, sheet_name='案件明細')
                
            excel_data = output.getvalue()
            excel_filename = '偽變造車牌專案敘獎統計含明細.xlsx'
            
            st.divider()
            col_dl, col_mail = st.columns(2)
            
            with col_dl:
                st.download_button(
                    label="📥 下載敘獎名冊及明細 (Excel)",
                    data=excel_data,
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            
            with col_mail:
                if st.button("📧 將此統計表一鍵寄至我的信箱", use_container_width=True):
                    with st.spinner("信件發送中，請稍候…"):
                        output.seek(0)
                        ok, mail_err = send_single_file_email(output, excel_filename)
                        if ok:
                            st.success("✅ 信件發送成功！統計報表 Excel 已夾帶至您的信箱。")
                        else:
                            st.error(f"❌ 發信失敗，請檢查系統信箱設定。錯誤訊息: {mail_err}")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.info("💡 **系統計算標準：**\n"
                    "1. **名目拆分**：系統已將【偽變造車牌】與【懸掛他車號牌】獨立計算，方便對接敘獎事由。\n"
                    "2. **偽變造車牌**：汽車記嘉獎2次，機車記嘉獎1次。\n"
                    "3. **懸掛他車號牌**：每 2 件記嘉獎1次。\n"
                    "4. **獎勵加倍**：專案累計達 9 次嘉獎門檻後，後續入案之件數獎勵自動加倍計算。")

if __name__ == "__main__":
    main()
