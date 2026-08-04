import streamlit as st
import pandas as pd
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
        msg["To"] = sender
        msg["Subject"] = f"🚨 危險駕車與改裝車輛取締 - {file_name}"
        
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
# 主程式執行區塊
# ==========================================
def main():
    # 頁面設定
    st.set_page_config(page_title="危險駕車取締獎勵統計", layout="wide", page_icon="🚨")
    show_sidebar()

    st.title("🚨 危險駕車與改裝車輛取締 - 獎勵統計儀表板")
    st.markdown("""
    **💡 獎勵計算標準：**
    * **【Group A：每 2 輛嘉獎一次】**：違反道交條例第13條第1款、第18條第1項及第43條第3項。
    * **【Group B：每 4 輛嘉獎一次】**：
        * 第16條第1項第1款（限定機車改裝排氣管）
        * 第16條第1項第2款（限定標的為排氣管及消音器設備）
        * 第43條第1項第1、3、4、5款
    * *註：累積達一大功 (9次嘉獎) 後，以倍數計算並以此類推。*
    """)
    st.divider()

    # ==========================================
    # 檔案上傳區塊
    # ==========================================
    uploaded_file = st.file_uploader("📂 請上傳舉發違規資料檔 (支援格式：xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            with st.spinner('資料處理中，請稍候...'):
                # 讀取資料
                df = pd.read_excel(uploaded_file, sheet_name="list2")
                
                # 確保欄位為字串型態，避免比對出錯及 NaN 問題
                df['條款1'] = df['條款1'].astype(str).str.strip()
                df['違規事實1'] = df['違規事實1'].astype(str).fillna('')
                df['車種'] = df['車種'].astype(str).fillna('')
                df['舉發員警'] = df['舉發員警'].astype(str).str.strip()
                
                # ==========================================
                # 條件篩選邏輯區
                # ==========================================
                # Group A (2件 = 1嘉獎): 第13-1, 18-1, 43-3 條
                mask_A = df['條款1'].str.startswith(('131', '181', '433'))
                
                # Group B (4件 = 1嘉獎): 
                # 1-1. 第16條第1項第1款 (16101開頭)：限定機車(包含機車或重型)
                mask_16_1_1 = df['條款1'].str.startswith('16101') & df['車種'].str.contains('機車|重型|輕型', na=False)
                
                # 1-2. 第16條第1項第2款 (16102開頭)：限定排氣管或消音器
                mask_16_1_2 = df['條款1'].str.startswith('16102') & df['違規事實1'].str.contains('排氣管|消音器', na=False)
                
                # 2. 第43-1-1, 43-1-3, 43-1-4, 43-1-5
                mask_43_1 = df['條款1'].str.startswith(('43101', '43103', '43104', '43105'))
                
                mask_B = mask_16_1_1 | mask_16_1_2 | mask_43_1
                
                # 切分出符合條件的 DataFrame
                df_A = df[mask_A]
                df_B = df[mask_B]
                
                # ==========================================
                # 獎勵核算區
                # ==========================================
                # 計算各員警件數
                counts_A = df_A['舉發員警'].value_counts().rename("GroupA_件數")
                counts_B = df_B['舉發員警'].value_counts().rename("GroupB_件數")
                
                # 以外部合併確保所有員警名單
                reward_df = pd.concat([counts_A, counts_B], axis=1).fillna(0).astype(int)
                
                if 'nan' in reward_df.index:
                    reward_df = reward_df.drop('nan')
                    
                # 計算嘉獎與大功數
                reward_df['嘉獎次數'] = (reward_df['GroupA_件數'] // 2) + (reward_df['GroupB_件數'] // 4)
                reward_df['大功數'] = reward_df['嘉獎次數'] // 9
                reward_df['剩餘嘉獎'] = reward_df['嘉獎次數'] % 9
                
                # 排序與重命名
                reward_df = reward_df.sort_values(by=["嘉獎次數", "GroupA_件數"], ascending=[False, False]).reset_index()
                if '舉發員警' in reward_df.columns:
                    reward_df = reward_df.rename(columns={"舉發員警": "舉發員警名稱"})
                elif 'index' in reward_df.columns:
                    reward_df = reward_df.rename(columns={"index": "舉發員警名稱"})
                
                reward_df = reward_df[(reward_df['GroupA_件數'] > 0) | (reward_df['GroupB_件數'] > 0)]
                reward_df = reward_df[['舉發員警名稱', 'GroupA_件數', 'GroupB_件數', '嘉獎次數', '大功數', '剩餘嘉獎']]

            # ==========================================
            # 畫面呈現區
            # ==========================================
            st.success("✅ 資料處理完成！")
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("適用【2件1嘉獎】", f"{len(df_A)} 件")
            col2.metric("適用【4件1嘉獎】", f"{len(df_B)} 件")
            col3.metric("核發嘉獎總數", f"{reward_df['嘉獎次數'].sum()} 次")
            col4.metric("達成大功總數", f"{reward_df['大功數'].sum()} 次")
            
            st.divider()
            
            # 顯示核算結果表
            st.subheader("🏆 員警獎勵核算總表")
            st.dataframe(reward_df, use_container_width=True, hide_index=True)
            
            # 建立 Excel 檔案與下載/寄送按鈕
            if not reward_df.empty:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    reward_df.to_excel(writer, index=False, sheet_name='獎勵結算表')
                    df_A.to_excel(writer, index=False, sheet_name='GroupA_明細')
                    df_B.to_excel(writer, index=False, sheet_name='GroupB_明細')
                output.seek(0)
                
                # 將下載按鈕與寄信按鈕並排
                col_dl, col_mail = st.columns(2)
                
                excel_filename = "危險駕車與改裝車輛取締_獎勵結算表.xlsx"
                
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
                            # 將指標移回開頭確保讀取正常
                            output.seek(0)
                            ok, mail_err = send_single_file_email(output, excel_filename)
                            if ok:
                                st.success("✅ 信件發送成功！統計報表 Excel 已夾帶至您的信箱。")
                            else:
                                st.error(f"❌ 發信失敗，請檢查系統信箱設定。錯誤代碼: {mail_err}")

            st.divider()
            
            # 原始資料驗證區塊
            st.subheader("🔍 案件明細檢核區")
            tab1, tab2 = st.tabs(["📌 Group A：2件1嘉獎 (第13-1, 18-1, 43-3條)", "📌 Group B：4件1嘉獎 (第16-1, 43-1條)"])
            
            display_columns = ['單號', '車牌', '車種', '違規日期', '條款1', '違規事實1', '舉發員警', '舉發單位']
            
            with tab1:
                if not df_A.empty:
                    st.dataframe(df_A[display_columns], use_container_width=True, hide_index=True)
                else:
                    st.info("查無符合 Group A 條件之案件。")
                    
            with tab2:
                if not df_B.empty:
                    st.dataframe(df_B[display_columns], use_container_width=True, hide_index=True)
                else:
                    st.info("查無符合 Group B 條件之案件。")

        except Exception as e:
            st.error(f"❌ 檔案處理發生錯誤，請確認欄位格式是否與標準資料庫一致。\n\n詳細錯誤訊息：{e}")

# 執行點
if __name__ == "__main__":
    main()
