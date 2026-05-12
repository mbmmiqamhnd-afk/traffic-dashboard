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
        # 從 st.secrets 讀取郵件帳密 (需先在 .streamlit/secrets.toml 設定)
        if "email" not in st.secrets:
            return False, "未偵測到郵件設定 (secrets.toml)"
        
        sender = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender  # 寄給自己備份
        msg["Subject"] = f"【系統發送】{filename.replace('.xlsx', '')}"
        
        body = f"您好：\n\n附件為您剛才在交通疏導系統中修正後的時數統計報表。\n發送時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        msg.attach(MIMEText(body, "plain", "utf-8"))
        
        # 產生 Excel 檔案資料流
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary_df.to_excel(writer, index=False, sheet_name='月彙整總表')
            detail_df.to_excel(writer, index=False, sheet_name='人員核銷明細')
        
        # 附件封裝
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(output.getvalue())
        encoders.encode_base64(part)
        # 處理中文檔名編碼
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(filename)}")
        msg.attach(part)
        
        # SMTP 傳送 (使用 Gmail SSL)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# --- 3. 主程式邏輯 ---
def run_app():
    st.title("⏱️ 交通疏導勤務時數彙整 (完整終極版)")
    st.markdown("---")

    # --- 側邊欄設定 ---
    st.sidebar.header("⚙️ 篩選規則設定")
    
    # A. 排除番號
    exclude_input = st.sidebar.text_input("要排除的番號 (A欄內容)", value="A, B, C")
    exclude_list = [i.strip().upper() for i in exclude_input.split(',') if i.strip()]
    
    st.sidebar.divider()

    # B. 尖峰時段欄位座標 (C欄=2, D欄=3...)
    am_cols_input = st.sidebar.text_input("上午尖峰欄位索引 (C欄起, 逗號隔開)", value="2, 3")
    pm_cols_input = st.sidebar.text_input("下午尖峰欄位索引 (逗號隔開)", value="12, 13")
    
    try:
        peak_col_indices = [int(i.strip()) for i in (am_cols_input + "," + pm_cols_input).split(',') if i.strip()]
    except:
        peak_col_indices = [2, 3, 12, 13]

    # --- 主介面：上傳 ---
    uploaded_files = st.file_uploader("請選取並上傳當月所有勤務明細檔", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files:
        all_records = []
        detected_units = set()
        
        for file in uploaded_files:
            try:
                # 讀取檔案
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=None, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, header=None, encoding='cp950')
                else:
                    df = pd.read_excel(file, header=None)

                # 提取單位名稱 (數字前的文字)
                unit_name = re.split(r'\d+', file.name)[0].strip()
                if unit_name:
                    detected_units.add(unit_name)

                # 掃描資料 (從第三列開始)
                for r_idx in range(2, len(df)):
                    row = df.iloc[r_idx]
                    
                    # 番號過濾 (A欄)
                    shift_code = str(row[0]).strip().upper()
                    if shift_code in exclude_list:
                        continue
                    
                    # 姓名取得 (B欄)
                    name = str(row[1]).replace('\n', '').replace(' ', '')
                    if name in ['nan', 'None', '', '姓名', '合計', '總計']: 
                        continue
                    
                    # 統計該員當天尖峰時數
                    daily_hours = 0
                    for c_idx in peak_col_indices:
                        if c_idx < len(row):
                            cell_content = str(row[c_idx]).replace('\n', '')
                            if "守望" in cell_content:
                                daily_hours += 1
                    
                    if daily_hours > 0:
                        all_records.append({
                            "單位": unit_name,
                            "姓名": name,
                            "當日尖峰時數": daily_hours,
                            "番號": shift_code,
                            "日期來源": file.name
                        })
            except Exception as e:
                st.error(f"解析 {file.name} 失敗：{e}")

        if all_records:
            raw_person_df = pd.DataFrame(all_records)
            
            st.divider()
            st.subheader("📝 第一步：確認每日人員名單 (可點選列後按 Delete 刪除)")
            
            # 使用人員明細編輯器 (人/天為單位)
            edited_df = st.data_editor(
                raw_person_df,
                use_container_width=True,
                num_rows="dynamic",
                key="person_editor_final",
                hide_index=False
            )

            if not edited_df.empty:
                # 即時重新彙總
                summary = edited_df.groupby(['單位', '姓名'])['當日尖峰時數'].sum().reset_index()
                summary.columns = ['單位', '姓名', '總計尖峰時數']
                summary = summary.sort_values(['單位', '總計尖峰時數'], ascending=[True, False])
                
                st.divider()
                st.subheader("📊 第二步：自動更新之彙整結果")
                st.dataframe(summary, use_container_width=True, hide_index=True)

                # 檔名產生邏輯
                unit_list = sorted(list(detected_units))
                filename_prefix = "_".join(unit_list[:2]) + (f"等{len(unit_list)}單位" if len(unit_list)>2 else "")
                final_filename = f"{filename_prefix}_交通疏導統計_{datetime.now().strftime('%m%d')}.xlsx"

                # --- 執行按鈕區 ---
                st.divider()
                col_dl, col_mail = st.columns(2)
                
                # 下載按鈕
                output_data = io.BytesIO()
                with pd.ExcelWriter(output_data, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月彙整總表')
                    edited_df.to_excel(writer, index=False, sheet_name='人員核銷明細')
                
                col_dl.download_button(
                    label=f"📥 下載 Excel 報表",
                    data=output_data.getvalue(),
                    file_name=final_filename,
                    use_container_width=True
                )

                # 寄信按鈕
                if col_mail.button("📧 寄送最終結果到我的信箱", use_container_width=True):
                    with st.spinner("寄送中..."):
                        ok, err = send_stats_email(final_filename, summary, edited_df)
                        if ok: st.success("✅ 郵件發送成功！")
                        else: st.error(f"❌ 郵件失敗: {err}")
            else:
                st.warning("⚠️ 名單已全數刪除。")
        else:
            st.warning("⚠️ 找不到符合條件的資料。")

if __name__ == "__main__":
    run_app()
