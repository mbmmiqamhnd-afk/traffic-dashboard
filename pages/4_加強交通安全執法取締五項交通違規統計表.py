import streamlit as st
import pandas as pd
import io
import smtplib
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# 設定頁面資訊
st.set_page_config(page_title="五項交通違規統計 (修復版)", layout="wide", page_icon="🚦")

# --- 側邊欄 ---
with st.sidebar:
    st.header("⚙️ 設定")
    auto_email = st.checkbox("分析完成後自動寄信", value=True)
    st.markdown("---")
    st.markdown("""
    ### 📝 操作說明
    1. 拖曳上傳檔案。
    2. 系統自動辨識年份與類別。
    3. **修復錯誤**：
       - 解決 `Unknown format code 'f' for object of type 'str'` 錯誤。
       - 確保文字欄位不被當作數字處理。
    """)
    status_container = st.container()

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

# --- 精確日期提取 ---
def extract_header_date(file_obj, filename):
    try:
        file_obj.seek(0)
        if filename.endswith(('.xls', '.xlsx')):
            try: df_head = pd.read_excel(file_obj, header=None, nrows=5)
            except: 
                file_obj.seek(0)
                df_head = pd.read_excel(file_obj, header=None, nrows=5, engine='openpyxl')
        else:
            try: df_head = pd.read_csv(file_obj, header=None, nrows=5, encoding='utf-8')
            except: 
                file_obj.seek(0)
                df_head = pd.read_csv(file_obj, header=None, nrows=5, encoding='cp950')
        
        if len(df_head) > 2:
            row_vals = df_head.iloc[2].astype(str).values
            row_text = " ".join(row_vals)
            clean_text = re.sub(r'[/\-\~\.\s]', '', row_text)
            matches = re.findall(r'(\d{6,7})', clean_text)
            valid_dates = [m for m in matches if len(m) >= 6]
            if len(valid_dates) >= 2:
                s_mmdd = valid_dates[0][-4:]
                e_mmdd = valid_dates[1][-4:]
                return f"({s_mmdd}~{e_mmdd})"
        return ""
    except: return ""

# --- 精確讀取函數 ---
def smart_read(fobj, fname):
    try:
        fobj.seek(0)
        header_idx = 3
        if fname.endswith(('.xls', '.xlsx')): 
            try: df = pd.read_excel(fobj, header=header_idx)
            except: 
                fobj.seek(0)
                df = pd.read_excel(fobj, header=header_idx, engine='openpyxl')
        else:
            try: df = pd.read_csv(fobj, header=header_idx, encoding='utf-8')
            except: 
                fobj.seek(0)
                df = pd.read_csv(fobj, header=header_idx, encoding='cp950')
        
        df.columns = [str(c).strip().replace('\n', '').replace(' ', '') for c in df.columns]
        if '單位' not in df.columns:
            match = [c for c in df.columns if '單位' in c]
            if match: df.rename(columns={match[0]: '單位'}, inplace=True)
        return df
    except Exception as e:
        return pd.DataFrame(columns=['單位'])

# --- 主程式 ---
uploaded_files = st.file_uploader("請將報表檔案拖曳至此", accept_multiple_files=True)

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
    
    date_labels = {'week': "", 'curr': "", 'last': ""}

    u_map = {
        '龍潭交通分隊': '交通分隊', '交通分隊': '交通分隊',
        '交通組': '科技執法', '科技執法': '科技執法',
        '聖亭派出所': '聖亭所', '聖亭所': '聖亭所',
        '龍潭派出所': '龍潭所', '龍潭所': '龍潭所',
        '中興派出所': '中興所', '中興所': '中興所',
        '石門派出所': '石門所', '石門所': '石門所',
        '高平派出所': '高平所', '高平所': '高平所',
        '三和派出所': '三和所', '三和所': '三和所'
    }
    def map_unit_name(raw_name):
        raw = str(raw_name)
        for key, val in u_map.items():
            if key in raw: return val
        return None

    try:
        def process_data(key_gen, key_foot, suffix, period_key):
            df_gen = pd.DataFrame(columns=['單位'])
            if key_gen in file_map:
                f_obj = file_map[key_gen]['file']
                f_name = file_map[key_gen]['name']
                if date_labels[period_key] == "":
                    date_labels[period_key] = extract_header_date(f_obj, f_name)
                df_gen = smart_read(f_obj, f_name)

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
                    for c in cols:
                        if k in c or c.startswith(k): matched_cols.append(c)
                matched_cols = list(set(matched_cols))
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
                res = pd.DataFrame(columns=['單位'])

            if key_foot in file_map:
                f_obj = file_map[key_foot]['file']
                f_name = file_map[key_foot]['name']
                if date_labels[period_key] == "":
                    date_labels[period_key] = extract_header_date(f_obj, f_name)
                
                foot = smart_read(f_obj, f_name)
                if '單位' not in foot.columns:
                     try:
                         f_obj.seek(0)
                         foot = pd.read_excel(f_obj, header=None)
                         for i, row in foot.iterrows():
                             if '單位' in str(row.values):
                                 f_obj.seek(0)
                                 foot = pd.read_excel(f_obj, header=i)
                                 foot.columns = [str(c).strip() for c in foot.columns]
                                 break
                     except: pass

                if '單位' in foot.columns:
                    foot = foot[~foot['單位'].isin(['合計', '總計', '小計', 'nan'])].copy()
                    foot['單位'] = foot['單位'].astype(str).str.strip()
                    ped_cols = [c for c in foot.columns if '78' in str(c) or '行人' in str(c)]
                    if ped_cols:
                        target_col = ped_cols[0]
                        foot[target_col] = foot[target_col].apply(clean_num)
                        if res.empty: 
                            res = foot[['單位', target_col]].copy()
                            res.rename(columns={target_col: f'行人違規_{suffix}'}, inplace=True)
                        else:
                            res = res.merge(foot[['單位', target_col]], on='單位', how='left')
                            res.rename(columns={target_col: f'行人違規_{suffix}'}, inplace=True)
            
            target_col_name = f'行人違規_{suffix}'
            if target_col_name not in res.columns: res[target_col_name] = 0
            res[target_col_name] = res[target_col_name].fillna(0)
            return res

        df_w = process_data('week_gen', 'week_foot', '本期', 'week')
        df_c = process_data('curr_gen', 'curr_foot', '本年', 'curr')
        df_l = process_data('last_gen', 'last_foot', '去年', 'last')

        with status_container:
            st.info(f"📅 日期偵測：\n本期 {date_labels['week']} | 本年 {date_labels['curr']} | 去年 {date_labels['last']}")

        unit_sources = []
        for d in [df_w, df_c, df_l]:
            if not d.empty and '單位' in d.columns: unit_sources.append(d['單位'])
        
        if unit_sources:
            all_units = pd.concat(unit_sources).unique()
            base_df = pd.DataFrame({'單位': all_units})
            base_df = base_df[base_df['單位'].notna() & (base_df['單位'] != '')]
            full = base_df.merge(df_c, on='單位', how='left') \
                          .merge(df_l, on='單位', how='left') \
                          .merge(df_w, on='單位', how='left') \
                          .fillna(0)
        else:
            full = pd.DataFrame(columns=['單位'])

        if '單位' in full.columns:
            full['Target_Unit'] = full['單位'].apply(map_unit_name)
            final = full[full['Target_Unit'].notna()].copy()
        else:
            final = pd.DataFrame()

        if final.empty: 
            st.error("❌ 找不到有效單位。")
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

            order = ['合計', '科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '交通分隊']
            result['Target_Unit'] = pd.Categorical(result['Target_Unit'], categories=order, ordered=True)
            result.sort_values('Target_Unit', inplace=True)

            cols_out = ['Target_Unit']
            for p in ['本期', '本年', '去年', '比較']:
                for c in cats: 
                    col_name = f'{c}_{p}'
                    if col_name in result.columns: cols_out.append(col_name)
                    else: result[col_name] = 0; cols_out.append(col_name)
            
            final_table = result[cols_out].copy()
            # 確保數值欄位是數字
            try: final_table.iloc[:, 1:] = final_table.iloc[:, 1:].astype(float).fillna(0)
            except: pass

            st.success("✅ 分析完成！")
            
            txt_week = f"本期 {date_labels['week']}"
            txt_curr = f"本年累計 {date_labels['curr']}"
            txt_last = f"去年累計 {date_labels['last']}"
            txt_comp = "本年與去年同期比較"

            # --- 網頁預覽 ---
            st.markdown("""
                <h2 style='text-align: center; color: blue; font-family: "Microsoft JhengHei", sans-serif;'>
                    加強交通安全執法取締五項交通違規統計表
                </h2>
            """, unsafe_allow_html=True)
            
            display_df = final_table.copy()
            new_columns = []
            for col in display_df.columns:
                if col == 'Target_Unit': new_columns.append(('統計期間', '取締項目'))
                elif '本期' in col: new_columns.append((txt_week, col.replace('_本期', '')))
                elif '本年' in col: new_columns.append((txt_curr, col.replace('_本年', '')))
                elif '去年' in col: new_columns.append((txt_last, col.replace('_去年', '')))
                elif '比較' in col: new_columns.append((txt_comp, col.replace('_比較', '')))
                else: new_columns.append(('', col))
            display_df.columns = pd.MultiIndex.from_tuples(new_columns)
            
            def highlight_negative_red(val):
                # 增加防呆，確保只有數值才判斷
                if isinstance(val, (int, float)) and val < 0:
                    return 'color: red'
                return 'color: black'

            # 找出數值欄位 (排除第一欄 'Target_Unit')
            # 這裡要注意 MultiIndex 的情況，第一欄的 Level 1 是 '取締項目'
            numeric_cols = display_df.columns[1:]
            
            # 使用 subset 參數來避開文字欄位
            styled_df = display_df.style\
                .map(highlight_negative_red, subset=numeric_cols)\
                .format("{:.0f}", subset=numeric_cols)
            
            st.dataframe(styled_df, use_container_width=True)

            # --- Excel 輸出 ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_table.to_excel(writer, index=False, header=False, startrow=3, sheet_name='交通違規統計')
                workbook = writer.book
                worksheet = writer.sheets['交通違規統計']
                
                fmt_base = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
                fmt_data = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0;[Red]-0'})
                fmt_title = workbook.add_format({'bold': True, 'font_size': 20, 'font_color': 'blue', 'align': 'center', 'valign': 'vcenter'})
                fmt_black = workbook.add_format({'bold': True, 'color': 'black'})
                fmt_red = workbook.add_format({'bold': True, 'color': 'red'})
                fmt_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
                fmt_label = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
                
                worksheet.merge_range('A1:U1', '加強交通安全執法取締五項交通違規統計表', fmt_title)
                worksheet.write('A2', '統計期間', fmt_label)
                
                def write_mixed_cell(r, c1, c2, text):
                    worksheet.merge_range(r, c1, r, c2, "", fmt_base)
                    if "(" in text:
                        parts = text.split("(", 1)
                        worksheet.write_rich_string(r, c1, fmt_black, parts[0], fmt_red, "(" + parts[1], fmt_base)
                    else:
                        fmt_bold_black = workbook.add_format({'bold': True, 'color': 'black', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                        worksheet.write(r, c1, text, fmt_bold_black)

                write_mixed_cell(1, 1, 5, txt_week)
                write_mixed_cell(1, 6, 10, txt_curr)
                write_mixed_cell(1, 11, 15, txt_last)
                write_mixed_cell(1, 16, 20, txt_comp)
                
                headers = ['取締項目'] + ['酒駕', '闖紅燈', '嚴重\n超速', '車不\n讓人', '行人\n違規'] * 4
                worksheet.write_row('A3', headers, fmt_header)
                worksheet.set_column('A:A', 15, fmt_base)
                worksheet.set_column('B:U', 9, fmt_data)

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
