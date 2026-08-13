import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import urllib.parse as _ul

# ==========================================
# 💡 匯入系統原本的側邊欄設定
# ==========================================
try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        pass

# ==========================================
# 0. 輔助函式：發送單一檔案 Email
# ==========================================
def send_single_file_email(file_bytes, file_name, mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
    try:
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender  # 寄給自己
        msg["Subject"] = f"🚗 無照駕駛移置保管敘獎統計 - {file_name}"
        
        body_text = f"長官您好，\n\n系統已自動產出相關檔案，附件為最新的【{file_name}】。\n\n本信件由交通執法自動化分析引擎發送。"
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        main_type, sub_type = mime_type.split('/') if '/' in mime_type else ("application", "octet-stream")
        part = MIMEBase(main_type, sub_type)
        part.set_payload(file_bytes.getvalue())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(file_name)}")
        msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# ==========================================
# 核心邏輯
# ==========================================
# 解析年齡字串提取數字
def parse_age(age_str):
    try:
        if pd.isna(age_str):
            return np.nan
        return int(re.sub(r'\D', '', str(age_str)))
    except:
        return np.nan

# 判斷適用案件類別
def categorize_case(row):
    fact = str(row['違規事實1'])
    law = str(row['違規法條1'])
    age = row['age_num']
    
    if pd.isna(age):
        return '不採計'
        
    # 未成年 (未滿18歲) 適用任何無照
    if age < 18:
        return '未成年'
        
    # 成年人 (18歲含以上) 僅適用小型車 (或法條內之汽車)
    if age >= 18:
        is_small_car = '小型車' in fact or '汽車駕駛人' in fact
        is_motorcycle = '機車' in fact
        
        if is_small_car and not is_motorcycle:
            if law.startswith('21101') or law.startswith('21104') or law.startswith('21105') or law.startswith('212'):
                return '成年'
            
    return '不採計'

# ==========================================
# 主程式執行區塊
# ==========================================
def main():
    st.set_page_config(page_title="無照駕駛移置保管敘獎統計", page_icon="🚗", layout="wide")
    show_sidebar()

    st.title("🚗 無照駕駛車輛移置保管敘獎統計系統")
    st.markdown("""
    依據最新公文 (桃警交字第1140031633號) 邏輯設計：
    * **未成年**：每 **5** 件嘉獎一次，每季每人上限二次。
    * **成年人 (小型車/汽車)**：每 **3** 件嘉獎一次，每季每人上限二次。
    * **跨季累計**：當季未達標件數可保留至下一季 (不得跨年)。
    """)

    st.info("💡 請上傳系統產出的「自選匯出.xlsx」(需包含『案件明細』工作表)")

    uploaded_file = st.file_uploader("📂 上傳自選匯出 Excel", type=["xlsx"])

    if uploaded_file:
        with st.spinner("資料處理中，請稍候..."):
            try:
                xls = pd.ExcelFile(uploaded_file)
                if '案件明細' not in xls.sheet_names:
                    st.error("❌ 找不到『案件明細』工作表，請確認您上傳的是正確的自選匯出報表。")
                    st.stop()
                    
                df_raw = pd.read_excel(xls, sheet_name='案件明細', header=None)
                
                # 定位欄位名稱列
                header_idx = df_raw[df_raw.iloc[:, 1] == '違規法條1'].index[0]
                
                df = df_raw.iloc[header_idx+1:].copy()
                df.columns = ['單號', '違規法條1', '違規事實1', '入案日', '舉發員警1', '違規人年齡']
                df = df.dropna(subset=['單號', '舉發員警1'])
                
                df['age_num'] = df['違規人年齡'].apply(parse_age)
                df['案件類別'] = df.apply(categorize_case, axis=1)
                
                df_valid = df[df['案件類別'] != '不採計'].copy()
                
                def get_quarter(date_str):
                    try:
                        month = int(str(date_str)[3:5])
                        if month <= 3: return 1
                        elif month <= 6: return 2
                        elif month <= 9: return 3
                        else: return 4
                    except:
                        return 1
                df_valid['季別'] = df_valid['入案日'].apply(get_quarter)
                
                summary = df_valid.pivot_table(
                    index='舉發員警1', 
                    columns=['案件類別', '季別'], 
                    aggfunc='size', 
                    fill_value=0
                )
                
                results = []
                for officer, row in summary.iterrows():
                    # Q1 計算
                    juv_q1 = row.get(('未成年', 1), 0)
                    adult_q1 = row.get(('成年', 1), 0)
                    
                    juv_merit_q1 = min(juv_q1 // 5, 2)
                    juv_carry_to_q2 = juv_q1 - (juv_merit_q1 * 5)
                    
                    adult_merit_q1 = min(adult_q1 // 3, 2)
                    adult_carry_to_q2 = adult_q1 - (adult_merit_q1 * 3)
                    
                    # Q2 計算
                    juv_q2_raw = row.get(('未成年', 2), 0)
                    adult_q2_raw = row.get(('成年', 2), 0)
                    
                    juv_total_q2 = juv_q2_raw + juv_carry_to_q2
                    juv_merit_q2 = min(juv_total_q2 // 5, 2)
                    
                    adult_total_q2 = adult_q2_raw + adult_carry_to_q2
                    adult_merit_q2 = min(adult_total_q2 // 3, 2)
                    
                    total_merit = juv_merit_q1 + adult_merit_q1 + juv_merit_q2 + adult_merit_q2
                    
                    results.append({
                        '舉發員警': officer,
                        'Q1 未成年(件)': juv_q1,
                        'Q1 成年(件)': adult_q1,
                        'Q1 核算嘉獎數': juv_merit_q1 + adult_merit_q1,
                        
                        'Q2 未成年(含Q1保留)': juv_total_q2,
                        'Q2 成年(含Q1保留)': adult_total_q2,
                        'Q2 核算嘉獎數': juv_merit_q2 + adult_merit_q2,
                        
                        '總嘉獎數': total_merit
                    })
                    
                res_df = pd.DataFrame(results)
                
                if res_df.empty:
                    st.warning("⚠️ 沒有計算出任何符合敘獎資格的資料，請確認入案日與案件是否吻合條件。")
                else:
                    st.success("✅ 計算完成！")
                    
                    with st.expander("🔍 檢視有效案件明細 (供除錯與核對用)"):
                        st.dataframe(df_valid[['單號', '入案日', '舉發員警1', '違規人年齡', '案件類別', '違規事實1', '違規法條1']])
                    
                    st.subheader("🏆 敘獎結算結果")
                    st.dataframe(res_df.style.highlight_max(subset=['總嘉獎數'], color='lightgreen'))
                    
                    # 建立 Excel 檔案
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        res_df.to_excel(writer, sheet_name='敘獎結算表', index=False)
                        df_valid.to_excel(writer, sheet_name='採計案件明細', index=False)
                    output.seek(0)
                    
                    st.divider()
                    col_dl, col_mail = st.columns(2)
                    excel_filename = "無照駕駛車輛移置_敘獎統計表.xlsx"
                    
                    with col_dl:
                        st.download_button(
                            label="📥 下載完整統計報表 (Excel)",
                            data=output,
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
                                    st.error(f"❌ 發信失敗，請檢查系統信箱設定。錯誤代碼: {mail_err}")
                    
            except Exception as e:
                st.error(f"❌ 檔案處理發生錯誤，請確認欄位格式是否與標準資料庫一致。\n\n詳細錯誤訊息：{e}")

if __name__ == "__main__":
    main()
