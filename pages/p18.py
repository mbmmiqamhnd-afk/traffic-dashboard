import streamlit as st
import pandas as pd
import io
import sys
import os
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# 自動將上層目錄加入路徑
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
try:
    from app import show_sidebar
except ImportError:
    def show_sidebar():
        pass 

def send_email_with_attachment(to_email, file_data, filename, year, month):
    """
    發送電子郵件功能的核心函數
    """
    # 🌟 請在此設定您的寄件資訊
    sender_email = "您的Gmail地址@gmail.com"
    sender_password = "您的Google應用程式密碼" # 非登入密碼，是16位元的 App Password
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['Subject'] = f"【自動發送】龍潭分局 {year}年{month}月份 獎勵金點數統計表"
    
    body = f"郭同仁您好：\n\n系統已完成 {year}年{month}月份的獎勵金點數彙整。\n附件為最新產出的統計報表，請查收。\n\n(本郵件由交通戰情室自動化系統發送)"
    msg.attach(MIMEText(body, 'plain'))
    
    # 加入附件
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(file_data)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {filename}")
    msg.attach(part)
    
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"郵件發送失敗：{str(e)}")
        return False

def p18_page():
    show_sidebar()

    st.title("💰 龍潭分局 - 獎勵金點數統計表產生器 (含自動郵件)")
    st.info("系統將自動裁剪報表、修正小計，並在完成後自動將 Excel 寄送到您的信箱。")

    # 1. 點數權重設定
    with st.expander("⚙️ 點數權重設定", expanded=False):
        col1, col2, col3 = st.columns(3)
        p_a2 = col1.number_input("A2 點數/件", value=10.0, step=1.0)
        p_a3 = col2.number_input("A3 點數/件", value=5.0, step=1.0)
        p_traf = col3.number_input("交整點數/小時", value=5.0, step=1.0)

    # 2. 檔案上傳
    st.subheader("📂 原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)", type=['xlsx'], accept_multiple_files=True)

    # 3. 接收郵件設定
    target_email = st.text_input("📧 請輸入您的接收郵件地址", value="您的電子信箱@gmail.com")

    if st.button("🚀 執行彙整並自動寄件", type="primary"):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方資料皆已完成上傳！")
            return

        with st.spinner("正在重新計算數據、產生總表並發送郵件..."):
            try:
                # --- [數據處理邏輯維持不變，包含修正小計與權重] ---
                df_acc = pd.read_excel(file_acc, header=4)
                df_acc['姓名'] = df_acc['姓名'].astype(str).str.strip()
                dict_acc = df_acc.groupby('姓名')[['A2類', 'A3類']].sum().to_dict(orient='index')

                df_traf_all = pd.concat([pd.read_excel(f, sheet_name='月彙整總表') for f in file_traf_list])
                df_traf_all['姓名'] = df_traf_all['姓名'].astype(str).str.strip()
                dict_traf = df_traf_all.groupby('姓名')['總計尖峰時數'].sum().to_dict()

                dfs_raw = pd.read_excel(file_template, sheet_name=None, header=None)
                
                # 偵測日期
                ext_year, ext_month = "115", "4"
                for _, df_scan in dfs_raw.items():
                    for r in range(min(15, len(df_scan))):
                        for c in range(min(10, len(df_scan.columns))):
                            v = str(df_scan.iloc[r, c])
                            m = re.search(r'開單日期[：:\s]*(\d{3})(\d{2})', v)
                            if m:
                                ext_year, ext_month = m.group(1), str(int(m.group(2)))
                                break

                # --- 裁剪、計算、產生小計 (邏輯與先前修正版相同) ---
                final_sheets = {}
                summary_rows = []
                g_cite, g_acc, g_traf, g_all = 0, 0, 0, 0

                for sheet_name, df in dfs_raw.items():
                    if '總表' in sheet_name: continue
                    start_r, start_c = None, None
                    for r_idx, row in df.iterrows():
                        row_str = [str(x).strip() for x in row.values]
                        if '員警姓名' in row_str:
                            start_r, start_c = r_idx, row_str.index('員警姓名')
                            break
                    if start_r is not None:
                        df_members = df.iloc[start_r+1:, start_c:].copy()
                        df_members.columns = [str(c).strip() for c in df.iloc[start_r, start_c:]]
                        df_members = df_members[~df_members['員警姓名'].astype(str).str.contains('小計|總計|合計', na=False)]
                        df_members = df_members.dropna(subset=['員警姓名']).astype(object)
                        
                        s_cite_sub, s_acc_sub, s_traf_sub = 0, 0, 0
                        for idx, row in df_members.iterrows():
                            name = str(row['員警姓名']).strip()
                            a2 = dict_acc.get(name, {}).get('A2類', 0)
                            a3 = dict_acc.get(name, {}).get('A3類', 0)
                            th = dict_traf.get(name, 0)
                            ap, tp = a2 * p_a2 + a3 * p_a3, th * p_traf
                            cp = pd.to_numeric(row['取締點數'], errors='coerce') or 0
                            
                            df_members.at[idx, 'A2件數'] = a2 if a2 > 0 else ""
                            df_members.at[idx, 'A3件數'] = a3 if a3 > 0 else ""
                            df_members.at[idx, '事故點數'] = ap if ap > 0 else ""
                            df_members.at[idx, '交整時數'] = th if th > 0 else ""
                            df_members.at[idx, '交整點數'] = tp if tp > 0 else ""
                            df_members.at[idx, '個人總點數'] = cp + ap + tp
                            s_acc_sub += ap; s_traf_sub += tp; s_cite_sub += cp

                        # 建立小計列
                        sub_row = {c: "" for c in df_members.columns}
                        sub_row['員警姓名'] = '小計'
                        for cn in df_members.columns:
                            if cn in ['員警姓名', '蓋章']: continue
                            val = pd.to_numeric(df_members[cn], errors='coerce').sum()
                            sub_row[cn] = val if val > 0 else 0
                        
                        df_final = pd.concat([df_members, pd.DataFrame([sub_row])], ignore_index=True)
                        if '蓋章' in df_final.columns: df_final = df_final.drop(columns=['蓋章'])
                        final_sheets[sheet_name] = df_final
                        summary_rows.append([sheet_name, s_cite_sub, s_acc_sub, s_traf_sub, s_cite_sub + s_acc_sub + s_traf_sub])
                        g_cite += s_cite_sub; g_acc += s_acc_sub; g_traf += s_traf_sub; g_all += (s_cite_sub + s_acc_sub + s_traf_sub)

                # 建立總表
                df_summary = pd.DataFrame([['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']] + summary_rows + [['合計', g_cite, g_acc, g_traf, g_all]])

                # 封裝 Excel 資料
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_summary.to_excel(writer, sheet_name='總表', header=False, index=False)
                    for sn, dff in final_sheets.items():
                        dff.to_excel(writer, sheet_name=sn, index=False)
                
                excel_data = output.getvalue()
                final_filename = f"桃園市政府警察局龍潭分局{ext_year}年{ext_month}月份處理道路交通安全人員獎勵金點數統計表.xlsx"

                # --- 🌟 執行自動寄信 ---
                if send_email_with_attachment(target_email, excel_data, final_filename, ext_year, ext_month):
                    st.success(f"🎊 彙整成功！報表已自動寄送至：{target_email}")
                else:
                    st.warning("⚠️ 報表已產生，但郵件發送失敗，請檢查 Gmail 應用程式密碼設定。")

                st.download_button(label="📥 同時下載一份到電腦", data=excel_data, file_name=final_filename)

            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")

if __name__ == "__main__":
    p18_page()
