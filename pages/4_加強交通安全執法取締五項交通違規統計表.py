import streamlit as st
import pandas as pd
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# 設定頁面資訊
st.set_page_config(page_title="五項交通違規統計 (自動日期版)", layout="wide", page_icon="🚦")
# 隱藏預設標題，改用自訂 HTML 標題
# st.title("🚦 加強交通安全執法取締統計表")

# --- 側邊欄設定 ---
with st.sidebar:
    st.header("⚙️ 設定")
    auto_email = st.checkbox("分析完成後自動寄信", value=True)
    st.markdown("---")
    st.markdown("""
    ### 📝 操作說明
    1. 拖曳上傳檔案 (不限數量)。
    2. 系統依檔名自動辨識：
       - `(1)` → 本年
       - `(2)` → 去年
       - `footman`/`行人` → 行人
    3. **自動抓取日期**：依「入案日期」欄位自動填入統計區間。
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
        if fname.endswith(('.xls', '.xlsx')): 
            try: df_temp = pd.read_excel(fobj, header=None, nrows=20)
            except: 
                fobj.seek(0)
                df_temp = pd.read_excel(fobj, header=None, nrows=20, engine='openpyxl')
            
            header_idx = -1
            for i, row in df_temp.iterrows():
                if '單位' in row.astype(str).values:
                    header_idx = i
                    break
            if header_idx == -1: header_idx = 3 
            
            fobj.seek(0)
            df = pd.read_excel(fobj, header=header_idx)
        else:
            try: df_temp = pd.read_csv(fobj, header=None, nrows=20, encoding='utf-8')
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
            try: df = pd.read_csv(fobj, header=header_idx, encoding='utf-8')
            except: 
                fobj.seek(0)
                df = pd.read_csv(fobj, header=header_idx, encoding='cp950')
        
        df.columns = [str(c).strip() for c in df.columns]
        if '單位' not in df.columns:
            match = [c for c in df.columns if '單位' in c]
            if match: df.rename(columns={match[0]: '單位'}, inplace=True)
        return df
    except Exception as e: 
        return pd.DataFrame(columns=['單位'])

# --- 🔥 新增：日期範圍計算函數 ---
def get_date_range_str(df, default_str=""):
    """
    從 DataFrame 中尋找 '入案日期' 或 '違規日期'，
    並回傳 'MMDD~MMDD' 格式的字串。
    """
    target_col = None
    for col in df.columns:
        if '入案日期' in str(col) or '違規日期' in str(col):
            target_col = col
            break
    
    if target_col and not df[target_col].dropna().empty:
        try:
            # 轉字串並移除可能的符號
            dates = df[target_col].astype(str).apply(lambda x: x.replace('/', '').replace('-', '').replace('.', '').strip())
            # 過濾非數字 (避免標題列殘留)
            dates = dates[dates.str.isnumeric()]
            
            if not dates.empty:
                min_date = dates.min()
                max_date = dates.max()
                
                # 取後四碼 (MMDD)
                min_mmdd = min_date[-4:] if len(min_date) >= 4 else min_date
                max_mmdd = max_date[-4:] if len(max_date) >= 4 else max_date
                
                return f"({min_mmdd}~{max_mmdd})"
        except:
            pass
            
    return default_str

# --- 主程式 ---
uploaded_files = st.file_uploader("請將報表檔案拖曳至此 (支援 Excel/CSV)", accept_multiple_files=True)

if uploaded_files:
    file_map = {}
    for f in uploaded_files:
        name = f.name
        is_foot = 'footman' in name.lower() or '行人' in name
        if '(2)' in name: period = 'last'
        elif '(1)' in name: period = 'curr'
        else: period = 'week'
        type_key = 'foot' if is_foot else 'gen'
        key = f"{period}_{type_key}"
        file_map[key] = {'file': f, 'name': name}
    
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
        # 儲存計算出的日期字串
        date_ranges = {
            'week': "", 
            'curr': "", 
            'last': ""
        }

        def process_data(key_gen, key_foot, suffix, range_key):
            df_gen = pd.DataFrame(columns=['單位'])
            
            # 1. 處理一般報表
            if key_gen in file_map:
                df_gen = smart_read(file_map[key_gen]['file'], file_map[key_gen]['name'])
                
                # 🔥 計算日期 (優先使用一般報表的日期)
                if date_ranges[range_key] == "":
                    date_ranges[range_key] = get_date_range_str(df_gen)

            # 2. 處理資料表內容
            df = df_gen.copy()
            if '單位' in df.columns:
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
            if not df.empty:
                res['單位'] = df['單位']
                res[f'酒駕_{suffix}'] = get_sum(['35條', '73條2項', '73條3項'])
                res[f'闖紅燈_{suffix}'] = get_sum(['53條'])
                res[f'嚴重超速_{suffix}'] = get_sum(['43條'])
                res[f'車不讓人_{suffix}'] = get_sum(['44條', '48條'])
            else:
                res = pd.DataFrame(columns=['單位']) # 空白防呆

            # 3. 處理行人報表
            if key_foot in file_map:
                foot = smart_read(file_map[key_foot]['file'], file_map[key_foot]['name'])
                
                # 🔥 如果一般報表沒抓到日期，嘗試從行人報表抓
                if date_ranges[range_key] == "":
                    date_ranges[range_key] = get_date_range_str(foot)

                if '單位' in foot.columns:
                    foot = foot[~foot['單位'].isin(['合計', '總計', '小計', 'nan'])].copy()
                    foot['單位'] = foot['單位'].astype(str).str.strip()
                    ped_cols = [c for c in foot.columns if '78' in str(c) or '行人' in str(c)]
                    if ped_cols:
                        target_col = ped_cols[0]
                        foot[target_col] = foot[target_col].apply(clean_num)
                        
                        if res.empty: # 如果只有行人資料
                             res = foot[['單位', target_col]].copy()
                             res.rename(columns={target_col: f'行人違規_{suffix}'}, inplace=True)
                        else:
                            res = res.merge(foot[['單位', target_col]], on='單位', how='left')
                            res.rename(columns={target_col: f'行人違規_{suffix}'}, inplace=True)
            
            target_col_name = f'行人違規_{suffix}'
            if target_col_name not in res.columns: res[target_col_name] = 0
            res[target_col_name] = res[target_col_name].fillna(0)
            return res

        # 執行運算並同時抓取日期
        df_w = process_data('week_gen', 'week_foot', '本期', 'week')
        df_c = process_data('curr_gen', 'curr_foot', '本年', 'curr')
        df_l = process_data('last_gen', 'last_foot', '去年', 'last')

        # 合併邏輯
        unit_sources = []
        if not df_w.empty and '單位' in df_w.columns: unit_sources.append(df_w['單位'])
        if not df_c.empty and '單位' in df_c.columns: unit_sources.append(df_c['單位'])
        if not df_l.empty and '單位' in df_l.columns: unit_sources.append(df_l['單位'])
        
        if unit_sources:
            all_units = pd.concat(unit_sources).unique()
            base_df = pd.DataFrame({'單位': all_units})
            base_df = base_df[base_df['單位'].notna() & (base_df['單位'] != '')]

            full = base_df.merge(df_c, on='單位', how='left') \
                          .merge(df_l, on='單位', how='left') \
                          .merge(df_w, on='單位', how='left') \
                          .fillna(0)
        else:
            full = pd.DataFrame(columns=['單位']) # 完全沒資料的情況

        u_map = {
            '龍潭交通分隊': '交通分隊', '交通組': '科技執法', 
            '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', 
            '中興派出所': '中興所', '石門派出所': '石門所', 
            '高平派出所': '高平所', '三和派出所': '三和所'
        }
        
        if '單位' in full.columns:
            full['Target_Unit'] = full['單位'].map(u_map)
            final = full[full['Target_Unit'].notna()].copy()
        else:
            final = pd.DataFrame()

        if final.empty: 
            st.error("❌ 無法對應到有效單位。請確認上傳檔案內容。")
        else:
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

            order = ['合計', '科技執法', '交通分隊', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
            result['Target_Unit'] = pd.Categorical(result['Target_Unit'], categories=order, ordered=True)
            result.sort_values('Target_Unit', inplace=True)

            cols_out = ['Target_Unit']
            for p in ['本期', '本年', '去年', '比較']:
                for c in cats: 
                    col_name = f'{c}_{p}'
                    if col_name in result.columns: cols_out.append(col_name)
                    else: result[col_name] = 0; cols_out.append(col_name)
            
            final_table = result[cols_out].copy()
            try: final_table.iloc[:, 1:] = final_table.iloc[:, 1:].astype(int)
            except: pass

            st.success("✅ 分析完成！")
            
            # --- 產生顯示用的標題字串 ---
            label_week = f"本期 {date_ranges['week']}"
            label_curr = f"本年累計 {date_ranges['curr']}"
            label_last = f"去年累計 {date_ranges['last']}"
            label_comp = "本年與去年同期比較"

            # --- 網頁預覽 ---
            st.markdown("""
                <h2 style='text-align: center; color: blue; font-family: "Microsoft JhengHei", sans-serif;'>
                    加強交通安全執法取締五項交通違規統計表
                </h2>
            """, unsafe_allow_html=True)

            display_df = final_table.copy()
            new_columns = []
            
            for col in display_df.columns:
                if col == 'Target_Unit':
                    new_columns.append(('統計期間', '取締項目'))
                elif '本期' in col:
                    item_name = col.replace('_本期', '')
                    new_columns.append((label_week, item_name))
                elif '本年' in col:
                    item_name = col.replace('_本年', '')
                    new_columns.append((label_curr, item_name))
                elif '去年' in col:
                    item_name = col.replace('_去年', '')
                    new_columns.append((label_last, item_name))
                elif '比較' in col:
                    item_name = col.replace('_比較', '')
                    new_columns.append((label_comp, item_name))
                else:
                    new_columns.append(('', col))

            display_df.columns = pd.MultiIndex.from_tuples(new_columns)
            st.dataframe(display_df, use_container_width=True)

            # --- Excel 輸出 ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_table.to_excel(writer, index=False, header=False, startrow=3, sheet_name='交通違規統計')
                
                workbook = writer.book
                worksheet = writer.sheets['交通違規統計']
                
                fmt_title = workbook.add_format({
                    'bold': True, 'font_size': 20, 'font_color': 'blue', 
                    'align': 'center', 'valign': 'vcenter'
                })
                fmt_period_red = workbook.add_format({
                    'bold': True, 'font_color': 'red', 'align': 'center', 
                    'valign': 'vcenter', 'border': 1
                })
                fmt_period_black = workbook.add_format({
                    'bold': True, 'font_color': 'black', 'align': 'center', 
                    'valign': 'vcenter', 'border': 1
                })
                fmt_header = workbook.add_format({
                    'bold': True, 'align': 'center', 'valign': 'vcenter', 
                    'border': 1, 'text_wrap': True
                })
                fmt_label = workbook.add_format({
                    'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1
                })
                
                # 繪製標題與動態日期
                worksheet.merge_range('A1:U1', '加強交通安全執法取締五項交通違規統計表', fmt_title)
                
                worksheet.write('A2', '統計期間', fmt_label)
                # 使用變數填入日期
                worksheet.merge_range('B2:F2', label_week, fmt_period_red)
                worksheet.merge_range('G2:K2', label_curr, fmt_period_red)
                worksheet.merge_range('L2:P2', label_last, fmt_period_red)
                worksheet.merge_range('Q2:U2', label_comp, fmt_period_black)
                
                headers = ['取締項目'] + ['酒駕', '闖紅燈', '嚴重\n超速', '車不\n讓人', '行人\n違規'] * 4
                worksheet.write_row('A3', headers, fmt_header)
                
                worksheet.set_column('A:A', 15)
                worksheet.set_column('B:U', 9)

            excel_data = output.getvalue()
            file_name_out = '交通違規統計表.xlsx'

            email_receiver = st.secrets["email"]["user"] if "email" in st.secrets else "尚未設定"
            if auto_email:
                if "sent_cache" not in st.session_state: st.session_state["sent_cache"] = set()
                file_ids = ",".join(sorted([f.name for f in uploaded_files]))
                if file_ids not in st.session_state["sent_cache"]:
                    with st.spinner(f"正在自動寄送報表至 {email_receiver}..."):
                        if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", "附件為交通違規統計報表。", excel_data, file_name_out):
                            st.balloons(); st.success(f"✅ 郵件已發送"); st.session_state["sent_cache"].add(file_ids)
                else: st.info(f"✅ 報表已發送過。")
            else:
                if st.button("📧 立即發送郵件"):
                    if send_email(email_receiver, f"📊 [手動發送] {file_name_out}", "附件", excel_data, file_name_out): st.success("✅ 發送成功")

            st.download_button("📥 下載 Excel", excel_data, file_name_out, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        st.error(f"系統錯誤：{e}")
