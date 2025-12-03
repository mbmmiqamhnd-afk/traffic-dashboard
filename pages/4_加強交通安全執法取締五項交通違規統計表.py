import streamlit as st
import pandas as pd
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

st.set_page_config(page_title="五項交通違規統計", layout="wide", page_icon="🚦")
st.title("🚦 加強交通安全執法取締統計表")

st.markdown("""
### 📝 操作說明
1. 請上傳 **6 個檔案** (本期/本年/去年 的「自選匯出」與「footman」)。
2. **上傳後自動分析** 並 **自動寄出**。
3. 系統會自動修正行人違規欄位識別問題。
""")

# --- 寄信函數 ---
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

# --- 主程式 ---
uploaded_files = st.file_uploader("請將 6 個檔案拖曳至此", accept_multiple_files=True)

if uploaded_files:
    if len(uploaded_files) < 6:
        st.warning("⏳ 檔案不足 6 個，請繼續上傳...")
    else:
        try:
            # 1. 檔案分類
            file_map = {}
            for f in uploaded_files:
                name = f.name
                # 檔名判斷邏輯
                is_foot = 'footman' in name.lower() or '行人' in name
                
                if '(2)' in name: period = 'last'
                elif '(1)' in name: period = 'curr'
                else: period = 'week'
                
                key = f"{period}_{'foot' if is_foot else 'gen'}"
                file_map[key] = {'file': f, 'name': name}

            # 2. 智慧讀取函數
            def smart_read(fobj, fname):
                try:
                    fobj.seek(0)
                    # 判斷是否為 Excel
                    if fname.endswith(('.xls', '.xlsx')): 
                        # 先讀前 20 行找表頭
                        try:
                            df_temp = pd.read_excel(fobj, header=None, nrows=20)
                        except:
                            # 如果預設讀取失敗，嘗試別的引擎
                            fobj.seek(0)
                            df_temp = pd.read_excel(fobj, header=None, nrows=20, engine='openpyxl')

                        header_idx = -1
                        for i, row in df_temp.iterrows():
                            row_str = row.astype(str).values
                            if '單位' in row_str:
                                header_idx = i
                                break
                        if header_idx == -1: header_idx = 3 # 預設 fallback
                        
                        fobj.seek(0)
                        df = pd.read_excel(fobj, header=header_idx)
                    else:
                        # CSV 處理
                        try:
                            df_temp = pd.read_csv(fobj, header=None, nrows=20, encoding='utf-8')
                        except:
                            fobj.seek(0)
                            df_temp = pd.read_csv(fobj, header=None, nrows=20, encoding='cp950') # 嘗試 big5/cp950

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
                    
                    # 欄位清理
                    df.columns = [str(c).strip() for c in df.columns]
                    # 找單位欄
                    if '單位' not in df.columns:
                        match = [c for c in df.columns if '單位' in c]
                        if match: df.rename(columns={match[0]: '單位'}, inplace=True)
                    
                    return df
                except Exception as e: 
                    # st.error(f"讀取 {fname} 失敗: {e}")
                    return pd.DataFrame(columns=['單位'])

            # 3. 核心處理邏輯
            def process_data(key_gen, key_foot, suffix):
                if key_gen not in file_map: return pd.DataFrame(columns=['單位'])
                
                # --- 處理一般報表 ---
                df = smart_read(file_map[key_gen]['file'], file_map[key_gen]['name'])
                
                # 清洗
                df = df[~df['單位'].isin(['合計', '總計', '小計', 'nan'])].dropna(subset=['單位']).copy()
                df['單位'] = df['單位'].astype(str).str.strip()
                
                # 轉數值工具
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
                
                # --- 處理行人報表 (強化版) ---
                if key_foot in file_map:
                    foot = smart_read(file_map[key_foot]['file'], file_map[key_foot]['name'])
                    
                    if '單位' in foot.columns:
                        foot = foot[~foot['單位'].isin(['合計', '總計', '小計', 'nan'])].copy()
                        foot['單位'] = foot['單位'].astype(str).str.strip() # 強力去空白
                        
                        # 🔥 關鍵修正：同時搜尋 '78' 和 '行人'，增加命中率
                        ped_cols = [c for c in foot.columns if '78' in str(c) or '行人' in str(c)]
                        
                        if ped_cols:
                            target_col = ped_cols[0] # 取第一個符合的
                            foot[target_col] = foot[target_col].apply(clean_num)
                            
                            # 合併
                            res = res.merge(foot[['單位', target_col]], on='單位', how='left')
                            res.rename(columns={target_col: f'行人違規_{suffix}'}, inplace=True)
                
                # 補零 (Fill NA)
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
            full = df_c.merge(df_l, on='單位', how='outer').merge(df_w, on='單位', how='left').fillna(0)
            
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
                st.error("❌ 錯誤：找不到對應單位。請確認上傳檔案的「單位」欄位名稱是否正確。")
            else:
                # 計算比較與合計
                cats = ['酒駕', '闖紅燈', '嚴重超速', '車不讓人', '行人違規']
                for c in cats: 
                    final[f'{c}_比較'] = final[f'{c}_本年'] - final[f'{c}_去年']

                num_cols = final.columns.drop(['單位', 'Target_Unit'])
                total_row = final[num_cols].sum().to_frame().T
                total_row['Target_Unit'] = '合計'
                
                result = pd.concat([total_row, final], ignore_index=True)

                # 排序
                order = ['合計', '科技執法', '交通分隊', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
                result['Target_Unit'] = pd.Categorical(result['Target_Unit'], categories=order, ordered=True)
                result.sort_values('Target_Unit', inplace=True)

                cols_out = ['Target_Unit']
                for p in ['本期', '本年', '去年', '比較']:
                    for c in cats: cols_out.append(f'{c}_{p}')
                
                final_table = result[cols_out].copy()
                final_table.rename(columns={'Target_Unit': '單位'}, inplace=True)
                try: final_table.iloc[:, 1:] = final_table.iloc[:, 1:].astype(int)
                except: pass

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

                # 自動寄信
                if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
                file_ids = ",".join(sorted([f.name for f in uploaded_files]))
                email_receiver = st.secrets["email"]["user"]
                
                if file_ids not in st.session_state["sent_cache"]:
                    with st.spinner(f"正在自動寄送報表至 {email_receiver}..."):
                        if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", "附件為交通違規統計報表。", excel_data, file_name_out):
                            st.balloons()
                            st.success(f"✅ 郵件已發送至 {email_receiver}")
                            st.session_state["sent_cache"].add(file_ids)
                else:
                    st.info(f"✅ 報表已於剛才發送至 {email_receiver}")

                st.download_button(label="📥 下載 Excel", data=excel_data, file_name=file_name_out, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e: st.error(f"發生錯誤：{e}")
