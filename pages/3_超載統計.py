import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

st.set_page_config(page_title="超載統計", layout="wide", page_icon="🚛")
st.title("🚛 超載 (stoneCnt) 自動統計")

st.markdown("""
### 📝 使用說明
1. 請上傳 **3 個** `stoneCnt` 系列的 Excel 檔案。
2. **上傳後自動分析**。
3. 支援一鍵寄信。
""")

# ==========================================
# 1. 參數設定 (移到最上方，避免錯誤)
# ==========================================
TARGETS = {
    '聖亭所': 24, '龍潭所': 32, '中興所': 24, '石門所': 19,
    '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30
}

UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所',
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所',
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'
}

UNIT_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 2. 寄信函數
# ==========================================
def send_email(recipient, subject, body, file_bytes, filename):
    try:
        if "email" not in st.secrets:
            st.error("❌ 未設定 Secrets！")
            return False
        sender = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]

        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(file_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=Header(filename, 'utf-8').encode())
        msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, recipient, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"❌ 寄信失敗: {e}")
        return False

# ==========================================
# 3. 資料解析函數
# ==========================================
def parse_stone(f):
    if not f: return {}
    counts = {}
    try:
        xls = pd.ExcelFile(f)
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            curr = None
            for _, row in df.iterrows():
                # 轉字串搜尋關鍵字
                s = row.astype(str).str.cat(sep=' ')
                
                # 抓取單位
                if "舉發單位：" in s:
                    m = re.search(r"舉發單位：(\S+)", s)
                    if m: curr = m.group(1).strip()
                
                # 抓取數值
                if "總計" in s and curr:
                    # 抓取該行所有數字
                    nums = [float(x) for x in row if str(x).replace('.','',1).isdigit()]
                    if nums:
                        # 使用全域變數 UNIT_MAP
                        short = UNIT_MAP.get(curr, curr)
                        # 取最後一個數字作為總計
                        counts[short] = counts.get(short, 0) + int(nums[-1])
                        curr = None
        return counts
    except Exception as e:
        st.error(f"解析檔案錯誤: {e}")
        return {}

# ==========================================
# 4. 主程式執行
# ==========================================
uploaded_files = st.file_uploader("請拖曳 3 個 stoneCnt 檔案至此", accept_multiple_files=True, type=['xlsx', 'xls'])

if uploaded_files:
    if len(uploaded_files) < 3:
        st.warning("⏳ 檔案不足 3 個，請繼續上傳...")
    else:
        try:
            files_config = {"Week": None, "YTD": None, "Last_YTD": None}
            for f in uploaded_files:
                if "(1)" in f.name: files_config["YTD"] = f
                elif "(2)" in f.name: files_config["Last_YTD"] = f
                else: files_config["Week"] = f
            
            # 開始解析
            d_wk = parse_stone(files_config["Week"])
            d_yt = parse_stone(files_config["YTD"])
            d_ly = parse_stone(files_config["Last_YTD"])

            rows = []
            for u in UNIT_ORDER:
                rows.append({
                    '單位': u, 
                    '本期': d_wk.get(u,0), 
                    '本年累計': d_yt.get(u,0), 
                    '去年累計': d_ly.get(u,0), 
                    '目標值': TARGETS.get(u,0)
                })
            
            df = pd.DataFrame(rows)
            
            # 計算合計 (排除警備隊)
            df_calc = df.copy()
            mask_guard = df_calc['單位'] == '警備隊'
            df_calc.loc[mask_guard, ['本期', '本年累計', '去年累計', '目標值']] = 0
            
            total = df_calc[['本期', '本年累計', '去年累計', '目標值']].sum().to_dict()
            total['單位'] = '合計'
            
            df_final = pd.concat([pd.DataFrame([total]), df], ignore_index=True)
            
            # 計算比較與達成率
            df_final['本年與去年同期比較'] = df_final['本年累計'] - df_final['去年累計']
            df_final['達成率'] = df_final.apply(lambda x: f"{x['本年累計']/x['目標值']:.2%}" if x['目標值']>0 else "—", axis=1)
            
            # 警備隊特殊顯示
            df_final.loc[df_final['單位']=='警備隊', ['本年與去年同期比較', '目標值', '達成率']] = "—"
            
            cols = ['單位', '本期', '本年累計', '去年累計', '本年與去年同期比較', '目標值', '達成率']
            df_final = df_final[cols]
            
            # 顯示結果
            st.success("✅ 分析完成！")
            st.dataframe(df_final, use_container_width=True, hide_index=True)
            
            # 轉為 Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='超載統計')
            
            excel_data = output.getvalue()
            file_name_out = '超載統計表.xlsx'

            # --- 寄信區塊 ---
            st.markdown("---")
            st.subheader("📧 發送結果")
            col1, col2 = st.columns([3, 1])
            with col1:
                default_mail = st.secrets["email"]["user"] if "email" in st.secrets else ""
                email_receiver = st.text_input("收件信箱", value=default_mail)
            with col2:
                st.write(""); st.write("")
                if st.button("📤 立即寄出", type="primary"):
                    if not email_receiver: st.warning("請輸入信箱！")
                    else:
                        with st.spinner("寄送中..."):
                            if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", "附件為超載統計報表。", excel_data, file_name_out):
                                st.balloons(); st.success(f"已發送至 {email_receiver}")

            st.download_button(label="📥 下載 Excel", data=excel_data, file_name=file_name_out, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e:
            st.error(f"錯誤：{e}")
