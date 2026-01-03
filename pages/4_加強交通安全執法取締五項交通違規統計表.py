import streamlit as st
import pandas as pd
import io
import numpy as np  # 新增 numpy 用於處理空值
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# 設定頁面資訊
st.set_page_config(page_title="五項交通違規統計 (彈性版)", layout="wide", page_icon="🚦")
st.title("🚦 加強交通安全執法取締統計表")

# --- 側邊欄設定 ---
with st.sidebar:
    st.header("⚙️ 設定")
    auto_email = st.checkbox("分析完成後自動寄信", value=True, help="若關閉，分析後需手動點擊按鈕才會寄出。")
    st.markdown("---")
    st.markdown("""
    ### 📝 操作說明
    1. 拖曳上傳檔案 (不限數量)。
    2. 系統依檔名自動辨識：
       - `(1)` → 本年
       - `(2)` → 去年
       - `footman`/`行人` → 行人
    3. 缺少的檔案數值自動補 0。
    """)

# --- 寄信函數 ---
def send_email(recipient, subject, body, file_bytes, filename):
    try:
        if "email" not in st.secrets:
            st.error("❌ 未設定 Secrets！請在 Streamlit Cloud 設定 email 資訊。")
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

# --- 智慧讀取函數 ---
def smart_read(fobj, fname):
    try:
        fobj.seek(0)
        # 判斷是否為 Excel
        if fname.endswith(('.xls', '.xlsx')): 
            try:
                df_temp = pd.read_excel(fobj, header=None, nrows=20)
            except:
                fobj.seek(0)
                df_temp = pd.read_excel(fobj, header=None, nrows=20, engine='openpyxl')

            header_idx = -1
            for i, row in df_temp.iterrows():
                row_str = row.astype(str).values
                if '單位' in row_str:
                    header_idx = i
                    break
            if header_idx == -1: header_idx = 3 
            
            fobj.seek(0)
            df = pd.read_excel(fobj, header=header_idx)
        else:
            # CSV 處理
            try:
                df_temp = pd.read_csv(fobj, header=None, nrows=20, encoding='utf-8')
            except:
                fobj.seek(0)
                df_temp = pd.read_csv(fobj, header=None, nrows=20, encoding='cp950')

            header_idx = -1
            for i, row in df_temp.iterrows():
                if '單位' in row.astype(str).values:
                    header_idx = i
                    break
            if header_idx == -1: header_idx = 3
            
            fobj.seek(0)
            try:
                df = pd.read_csv(fobj, header=header_idx, encoding='utf-8')
            except:
                fobj.seek(0)
                df = pd.read_csv(fobj, header=header_idx, encoding='cp950')
        
        # 欄位與單位清洗
        df.columns = [str(c).strip() for c in df.columns]
        if '單位' not in df.columns:
            match = [c for c in df.columns if '單位' in c]
            if match: df.rename(columns={match[0]: '單位'}, inplace=True)
        
        return df
    except Exception as e: 
        return pd.DataFrame(columns=['單位'])

# --- 主程式 ---
uploaded_files = st.file_uploader("請將報表檔案拖曳至此 (支援 Excel/CSV)", accept_multiple_files=True)

if uploaded_files:
    # 1. 檔案分類與識別
    file_map = {}
    
    for f in uploaded_files:
        name = f.name
        is_foot = 'footman' in name.lower() or '行人' in name
        
        if '(2)' in name: period = 'last'   # 去年
        elif '(1)' in name: period = 'curr' # 本年
        else: period = 'week'               # 本期
        
        type_key = 'foot' if is_foot else 'gen'
        key = f"{period}_{type_key}"
        file_map[key] = {'file': f, 'name': name}
    
    # 顯示識別狀態
    expected_keys = {
        'week_gen': '本期_一般', 'week_foot': '本期_行人',
        'curr_gen': '本年_一般', 'curr_foot': '本年_行人',
        'last_gen': '去年_一般', 'last_foot': '去年_行人'
    }
    found_keys = file_map.keys()
    missing_files = [label for k, label in expected_keys.items() if k not in found_keys]
    
    if missing_files:
        st.warning(f"⚠️ 未偵測到以下檔案 (將以 0 計算): {', '.join(missing_files)}")
    else:
        st.info("✅ 所有預期檔案皆已上傳")

    try:
        # 2. 核心處理邏輯
        def process_data(key_gen, key_foot, suffix):
            if key_gen not in file_map: 
                return pd.DataFrame(columns=['單位'])
            
            df = smart_read(file_map[key_gen]['file'], file_map[key_gen]['name'])
            
            # 基礎清洗
            df = df[~df['單位'].isin(['合計', '總計', '小計', 'nan'])].dropna(subset=['單位']).copy()
            df['單位'] = df['單位'].astype(str).str.strip()
            
            def clean_num(x):
                try: return float(str(x).replace(',', '').replace('nan', '0'))
                except: return 0.0

            for c in df.columns:
                if c != '單位': df[c] = df[c].apply(clean_num)

            cols = df.columns
            def get_sum(keyword_list):
                matched_cols = []
                for k in keyword_list:
                    matches = [c for c in cols if str(c) == k or str(c).startswith(k)]
                    matched_cols.extend(matches)
                if not matched_cols: return 0
                return df[matched_cols].sum(axis=1)

            res = pd.DataFrame()
            res['單位'] = df['單位']
            res[f'酒駕_{suffix}'] = get_sum(['35條', '73條2項', '73條3項'])
            res[f'闖紅燈_{suffix}'] = get_sum(['53條'])
            res[f'嚴重超速_{suffix}'] = get_sum(['43條'])
            res[f'車不讓人_{suffix}'] = get_sum(['44條', '48條'])
            
            if key_foot in file_map:
                foot = smart_read(file_map[key_foot]['file'], file_map[key_foot]['name'])
                if '單位' in foot.columns:
                    foot = foot[~foot['單位'].isin(['合計', '總計', '小計', 'nan'])].copy()
                    foot['單位'] = foot['單位'].astype(str).str.strip()
                    
                    ped_cols = [c for c in foot.columns if '78' in str(c) or '行人' in str(c)]
                    if ped_cols:
                        target_col = ped_cols[0]
                        foot[target_col] = foot[target_col].apply(clean_num)
                        res = res.merge(foot[['單位', target_col]], on='單位', how='left')
                        res.rename(columns={target_col: f'行人違規_{suffix}'}, inplace=True)
            
            target_col_name = f'行人違規_{suffix}'
            if target_col_name not in res.columns: 
                res[target_col_name] = 0
            res[target_col_name] = res[target_col_name].fillna(0)
            
            return res

        # 執行運算
        df_w = process_data('week_gen', 'week_foot', '本期')
        df_c = process_data('curr_gen', 'curr_foot', '本年')
        df_l = process_data('last_gen', 'last_foot', '去年')

        # 合併
        all_units = pd.concat([df_w['單位'], df_c['單位'], df_l['單位']]).unique()
        base_df = pd.DataFrame({'單位': all_units})
        base_df = base_df[base_df['單位'].notna() & (base_df['單位'] != '')]

        full = base_df.merge(df_c, on='單位', how='left') \
                      .merge(df_l, on='單位', how='left') \
                      .merge(df_w, on='單位', how='left') \
                      .fillna(0)
        
        # 單位對照
        u_map = {
            '龍潭交通分隊': '交通分隊', '交通組': '科技執法', 
            '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', 
            '中興派出所': '中興所', '石門派出所': '石門所', 
            '高平派出所': '高平所', '三和派出所': '三和所'
        }
        full['Target_Unit'] = full['單位'].map(u_map)
        final = full[full['Target_Unit'].notna()].copy()

        if final.empty: 
            st.error("❌ 無法對應到有效單位，請確認報表內容。")
        else:
            # 計算比較
            cats = ['酒駕', '闖紅燈', '嚴重超速', '車不讓人', '行人違規']
            for c in cats: 
                col_curr = f'{c}_本年'
                col_last = f'{c}_去年'
                val_curr = final[col_curr] if col_curr in final.columns else 0
                val_last = final[col_last] if col_last in final.columns else 0
                final[f'{c}_比較'] = val_curr - val_last

            num_cols = [c for c in final.columns if c not in ['單位', 'Target_Unit']]
            total_row = final[num_cols].sum().to_frame().T
            total_row['Target_Unit'] = '合計'
            
            result = pd.concat([total_row, final], ignore_index=True)

            # 排序
            order = ['合計', '科技執法', '交通分隊', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
            result['Target_Unit'] = pd.Categorical(result['Target_Unit'], categories=order, ordered=True)
            result.sort_values('Target_Unit', inplace=True)

            cols_out = ['Target_Unit']
            for p in ['本期', '本年', '去年', '比較']:
                for c in cats: 
                    col_name = f'{c}_{p}'
                    if col_name in result.columns:
                        cols_out.append(col_name)
                    else:
                        result[col_name] = 0
                        cols_out.append(col_name)
            
            final_table = result[cols_out].copy()
            final_table.rename(columns={'Target_Unit': '取締項目'}, inplace=True)
            
            # --- 🔥 調整：先轉整數，再新增「統計期間」列 ---
            # 1. 先將數字部分轉為 Int，去除小數點 (e.g. 10.0 -> 10)
            try: 
                final_table.iloc[:, 1:] = final_table.iloc[:, 1:].astype(int)
            except: 
                pass
            
            # 2. 建立新的一列 (全空字串)
            period_row = pd.DataFrame([[""] * len(final_table.columns)], columns=final_table.columns)
            # 3. 設定第一欄標題
            period_row.iloc[0, 0] = "統計期間"
            
            # 4. 合併：將統計期間列放在最上方
            final_table = pd.concat([period_row, final_table], ignore_index=True)

            # 5. 確保空值顯示為空字串，而不是 NaN
            final_table = final_table.fillna("")

            st.success("✅ 分析完成！")
            st.dataframe(final_table, use_container_width=True)
            
            # 輸出 Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_table.to_excel(writer, index=False, sheet_name='交通違規統計')
                worksheet = writer.sheets['交通違規統計']
                worksheet.set_column(0, len(final_table.columns)-1, 12)
            
            excel_data = output.getvalue()
            file_name_out = '交通違規統計表.xlsx'

            # 寄信邏輯
            email_receiver = st.secrets["email"]["user"] if "email" in st.secrets else "尚未設定"
            
            if auto_email:
                if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
                file_ids = ",".join(sorted([f.name for f in uploaded_files]))
                
                if file_ids not in st.session_state["sent_cache"]:
                    with st.spinner(f"正在自動寄送報表至 {email_receiver}..."):
                        if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", "附件為交通違規統計報表。", excel_data, file_name_out):
                            st.balloons()
                            st.success(f"✅ 郵件已發送至 {email_receiver}")
                            st.session_state["sent_cache"].add(file_ids)
                else:
                    st.info(f"✅ 此份報表剛才已自動發送過。")
            else:
                if st.button("📧 立即發送郵件"):
                    with st.spinner(f"正在寄送報表至 {email_receiver}..."):
                        if send_email(email_receiver, f"📊 [手動發送] {file_name_out}", "附件為交通違規統計報表。", excel_data, file_name_out):
                            st.success(f"✅ 郵件已發送至 {email_receiver}")

            st.download_button(label="📥 下載 Excel", data=excel_data, file_name=file_name_out, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        st.error(f"發生系統錯誤：{e}")
        st.write("建議：請檢查上傳檔案格式是否為標準警方匯出報表。")
