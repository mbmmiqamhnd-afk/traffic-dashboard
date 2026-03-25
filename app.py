import streamlit as st
import pandas as pd
import io
import re
import gspread
import smtplib
import calendar
import traceback
import os
from datetime import datetime, timedelta, date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email import encoders
from email.header import Header
from pdf2image import convert_from_bytes
from pptx import Presentation

# ==========================================
# 0. 系統初始化與格式套件
# ==========================================
st.set_page_config(page_title="龍潭分局交通智慧戰情室", page_icon="🚓", layout="wide")

try:
    from gspread_formatting import *
    HAS_FORMATTING = True
except ImportError:
    HAS_FORMATTING = False

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

# ==========================================
# 1. 各模組運算大腦 (100% 保留您的原始代碼)
# ==========================================

def process_tech_enforcement(files):
    """📸 科技執法原始邏輯"""
    TO_EMAIL = "mbmmiqamhnd@gmail.com"
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    MY_EMAIL = st.secrets["email"]["user"]
    MY_PASSWORD = st.secrets["email"]["password"]
    GCP_CREDS = st.secrets["gcp_service_account"]

    def get_col_name(df, possible_names):
        clean_cols = [str(c).strip() for c in df.columns]
        for name in possible_names:
            if name in clean_cols: return df.columns[clean_cols.index(name)]
        return None

    def format_roc_date_range_to_yesterday():
        yesterday = datetime.now() - timedelta(days=1)
        return f"{yesterday.year - 1911}年1月1日至{yesterday.year - 1911}年{yesterday.month}月{yesterday.day}日"

    def create_formatted_excel(df_loc, date_range_text, total_count):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            ws = workbook.add_worksheet('科技執法成效統計')
            blue_title_fmt = workbook.add_format({'bold': True, 'font_size': 24, 'color': 'blue'})
            red_title_fmt = workbook.add_format({'bold': True, 'font_size': 24, 'color': 'red'}) 
            header_fmt = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1, 'bold': True, 'align': 'center'})
            data_fmt = workbook.add_format({'border': 1, 'align': 'center'})
            total_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFFFCC', 'align': 'center'})

            ws.write_rich_string('A1', blue_title_fmt, '科技執法成效 ', red_title_fmt, f'({date_range_text})')
            ws.write('A2', '統計期間', workbook.add_format({'align': 'center', 'border': 1}))
            ws.write('B2', date_range_text, workbook.add_format({'border': 1, 'color': 'red', 'align': 'center'}))
            ws.write('A3', '路口名稱', header_fmt)
            ws.write('B3', '舉發件數', header_fmt)
            
            curr_row = 3
            for _, row in df_loc.iterrows():
                ws.write(curr_row, 0, row['路段名稱'], data_fmt)
                ws.write(curr_row, 1, row['舉發件數'], data_fmt)
                curr_row += 1
            
            ws.write(curr_row, 0, '舉發總數', total_fmt)
            ws.write(curr_row, 1, total_count, total_fmt)
            
            chart = workbook.add_chart({'type': 'bar'})
            chart.add_series({'name': '舉發件數', 'categories': ['科技執法成效統計', 3, 0, curr_row - 1, 0], 'values': ['科技執法成效統計', 3, 1, curr_row - 1, 1], 'data_labels': {'value': True}})
            chart.set_title({'name': '違規路段排行'})
            ws.insert_chart('D2', chart, {'x_scale': 1.5, 'y_scale': 1.5})
        return output

    uploaded_file = files[0]
    if uploaded_file.name.endswith('.csv'):
        try: df = pd.read_csv(uploaded_file)
        except: 
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding='cp950')
    else: 
        df = pd.read_excel(uploaded_file)
    
    df.columns = [str(c).strip() for c in df.columns]
    loc_col = get_col_name(df, ['違規地點', '路口名稱', '地點'])
    
    if not loc_col:
        st.error("❌ 找不到『地點』相關欄位！請確認檔案格式。")
        return

    df[loc_col] = df[loc_col].astype(str).str.replace('桃園市', '', regex=False).str.replace('龍潭區', '', regex=False).str.strip()
    date_range_str = format_roc_date_range_to_yesterday()
    loc_summary = df[loc_col].value_counts().head(10).reset_index()
    loc_summary.columns = ['路段名稱', '舉發件數']

    st.markdown(f"### 📅 統計期間：:blue[科技執法成效 ]:red[({date_range_str})]")
    st.dataframe(loc_summary, hide_index=True)

    excel_data = create_formatted_excel(loc_summary, date_range_str, len(df))
    try:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        sheet_name = "科技執法-路段排行"
        try: ws = sh.worksheet(sheet_name)
        except: ws = sh.add_worksheet(title=sheet_name, rows="100", cols="20")
        ws.clear()
        title_text = f"科技執法成效 ({date_range_str})"
        update_data = [[title_text, ""], ["路段名稱", "舉發件數"]] + loc_summary.values.tolist() + [["舉發總數", len(df)]]
        ws.update(values=update_data)

        start_index_of_red = len("科技執法成效 ") 
        requests = {"requests": [{"updateCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1},
            "rows": [{"values": [{"userEnteredValue": {"stringValue": title_text},
            "textFormatRuns": [
                {"startIndex": 0, "format": {"foregroundColor": {"red": 0.0, "green": 0.0, "blue": 1.0}, "bold": True, "fontSize": 24}},
                {"startIndex": start_index_of_red, "format": {"foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}, "bold": True, "fontSize": 24}}
            ]}]}], "fields": "userEnteredValue,textFormatRuns"}}]}
        sh.batch_update(requests)
        st.success("✅ Google 試算表『路段排行』同步成功")
    except Exception as e: 
        st.warning(f"⚠️ 雲端同步失敗: {e}")

    try:
        msg = MIMEMultipart()
        msg['From'], msg['To'] = MY_EMAIL, TO_EMAIL
        msg['Subject'] = f"科技執法統計報告({date_range_str})"
        msg.attach(MIMEText(f"長官好，科技執法路段排行報表已完成。\n\n統計期間：{date_range_str}\n舉發總件數：{len(df)} 件", 'plain'))
        part = MIMEApplication(excel_data.getvalue(), Name="Tech_Enforcement.xlsx")
        part.add_header('Content-Disposition', 'attachment', filename="Tech_Enforcement.xlsx")
        msg.attach(part)
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
            s.starttls(); s.login(MY_EMAIL, MY_PASSWORD); s.send_message(msg)
        st.success(f"✅ 報表已寄送")
    except Exception as e: 
        st.error(f"❌ 郵件寄送失敗：{e}")


def process_overload(files):
    """🚛 超載違規原始邏輯"""
    TARGETS = {'聖亭所': 20, '龍潭所': 27, '中興所': 20, '石門所': 16, '高平所': 14, '三和所': 8, '警備隊': 0, '交通分隊': 22}
    UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'}
    UNIT_DATA_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

    f_wk, f_yt, f_ly = None, None, None
    for f in files:
        if "(1)" in f.name: f_yt = f
        elif "(2)" in f.name: f_ly = f
        else: f_wk = f

    def parse_report(f):
        if not f: return {}, "0000000", "0000000"
        counts, s, e = {}, "0000000", "0000000"
        f.seek(0)
        df_top = pd.read_excel(f, header=None, nrows=15)
        text_block = df_top.to_string()
        m = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', text_block)
        if m: s, e = m.group(1), m.group(2)
        f.seek(0)
        xls = pd.ExcelFile(f)
        for sn in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sn, header=None)
            u = None
            for _, r in df.iterrows():
                rs = " ".join(r.astype(str))
                if "舉發單位：" in rs:
                    m2 = re.search(r"舉發單位：(\S+)", rs)
                    if m2: u = m2.group(1).strip()
                if "總計" in rs and u:
                    nums = [float(str(x).replace(',','')) for x in r if str(x).replace('.','',1).isdigit()]
                    if nums:
                        short = UNIT_MAP.get(u, u)
                        if short in UNIT_DATA_ORDER: counts[short] = counts.get(short, 0) + int(nums[-1])
                        u = None
        return counts, s, e

    d_wk, s_wk, e_wk = parse_report(f_wk)
    d_yt, s_yt, e_yt = parse_report(f_yt)
    d_ly, s_ly, e_ly = parse_report(f_ly)

    raw_wk = f"本期 ({s_wk[-4:]}~{e_wk[-4:]})"
    raw_yt = f"本年累計 ({s_yt[-4:]}~{e_yt[-4:]})"
    raw_ly = f"去年累計 ({s_ly[-4:]}~{e_ly[-4:]})"

    body = []
    for u in UNIT_DATA_ORDER:
        yv, tv = d_yt.get(u, 0), TARGETS.get(u, 0)
        body.append({'統計期間': u, raw_wk: d_wk.get(u, 0), raw_yt: yv, raw_ly: d_ly.get(u, 0), '本年與去年同期比較': yv - d_ly.get(u, 0), '目標值': tv, '達成率': f"{yv/tv:.0%}" if tv > 0 else "—"})
    
    df_body = pd.DataFrame(body)
    sum_v = df_body[df_body['統計期間'] != '警備隊'][[raw_wk, raw_yt, raw_ly, '目標值']].sum()
    total_row = pd.DataFrame([{'統計期間': '合計', raw_wk: sum_v[raw_wk], raw_yt: sum_v[raw_yt], raw_ly: sum_v[raw_ly], '本年與去年同期比較': sum_v[raw_yt] - sum_v[raw_ly], '目標值': sum_v['目標值'], '達成率': f"{sum_v[raw_yt]/sum_v['目標值']:.0%}" if sum_v['目標值'] > 0 else "0%"}])
    df_final = pd.concat([total_row, df_body], ignore_index=True)

    y, m, d = int(e_yt[:3])+1911, int(e_yt[3:5]), int(e_yt[5:])
    prog_str = f"{((date(y, m, d) - date(y, 1, 1)).days + 1) / (366 if calendar.isleap(y) else 365):.1%}"
    f_plain = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt[:3]}年{e_yt[3:5]}月{e_yt[5:]}日 (入案日期)應達成率為{prog_str}"

    st.write("📊 取締超載違規件數統計表")
    st.dataframe(df_final, hide_index=True)

    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
    sh = gc.open_by_url(GOOGLE_SHEET_URL)
    ws = sh.get_worksheet(1) 
    
    clean_cols = ['統計期間', raw_wk, raw_yt, raw_ly, '本年與去年同期比較', '目標值', '達成率']
    footer_row_idx = 2 + len(df_final) + 1
    ws.update(range_name='A1', values=[['取締超載違規件數統計表']])
    ws.update(range_name='A2', values=[clean_cols] + df_final.values.tolist())
    ws.update(range_name=f'A{footer_row_idx}', values=[[f_plain]])
    st.write("✅ 試算表數據已更新")

    df_sync = df_final.copy()
    df_sync.columns = clean_cols
    df_excel_buffer = io.BytesIO()
    with pd.ExcelWriter(df_excel_buffer, engine='xlsxwriter') as writer:
        df_sync.to_excel(writer, index=False, startrow=1, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        title_format = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'font_color': 'blue'})
        worksheet.merge_range('A1:G1', '取締超載違規件數統計表', title_format)
        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:G', 12)

    user = st.secrets["email"]["user"]
    pwd = st.secrets["email"]["password"]
    msg = MIMEMultipart()
    msg['Subject'] = Header(f"🚛 超載報表 - {e_yt} ({prog_str})", 'utf-8').encode()
    msg['From'] = user
    msg['To'] = "mbmmiqamhnd@gmail.com"
    msg.attach(MIMEText("自動產生的超載報表已同步，請查閱附件。", 'plain'))
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(df_excel_buffer.getvalue())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="Overload_Report.xlsx"')
    msg.attach(part)
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(user, pwd)
    server.send_message(msg)
    server.quit()
    st.write("✅ 電子郵件自動寄送成功")


def process_major(files):
    """🚨 重大違規原始邏輯 (完全保留 A1 及黑紅格式)"""
    UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
    TARGETS = {'聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479, '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006}
    FOOTNOTE_TEXT = "重大交通違規指：「酒駕」、「闖紅燈」、「嚴重超速」、「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」及「不暫停讓行人」"

    def get_standard_unit(raw_name):
        name = str(raw_name).strip()
        if '分隊' in name: return '交通分隊'
        if '科技' in name or '交通組' in name: return '科技執法'
        if '警備' in name: return '警備隊'
        if '聖亭' in name: return '聖亭所'
        if '龍潭' in name: return '龍潭所'
        if '中興' in name: return '中興所'
        if '石門' in name: return '石門所'
        if '高平' in name: return '高平所'
        if '三和' in name: return '三和所'
        return None

    def sync_to_specified_sheet(df):
        try:
            gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
            sh = gc.open_by_url(GOOGLE_SHEET_URL)
            ws = sh.get_worksheet(0)
            
            # 1. 準備數據 (包含兩層 Header, 數據, 腳註)
            col_tuples = df.columns.tolist()
            top_row = [t[0] for t in col_tuples]
            bottom_row = [t[1] for t in col_tuples]
            data_body = df.values.tolist() 
            
            data_list = [top_row, bottom_row] + data_body
            
            # 2. 從 A2 開始寫入，保留 A1 總標題格式
            ws.update(range_name='A2', values=data_list)
            
            # 3. 處理內容顏色 (括號紅字與負值紅字)
            if HAS_FORMATTING:
                data_rows_end_idx = len(data_list) + 1
                red_color = {"red": 1.0, "green": 0.0, "blue": 0.0}
                black_color = {"red": 0.0, "green": 0.0, "blue": 0.0}
                
                requests = []
                # 標題括號紅字 (Row Index 1)
                for i, text in enumerate(top_row):
                    if "(" in text:
                        p_start = text.find("(")
                        requests.append({
                            "updateCells": {
                                "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": i, "endColumnIndex": i+1},
                                "rows": [{ "values": [{ "textFormatRuns": [
                                    {"startIndex": 0, "format": {"foregroundColor": black_color}},
                                    {"startIndex": p_start, "format": {"foregroundColor": red_color}}
                                ], "userEnteredValue": {"stringValue": text} }] }],
                                "fields": "userEnteredValue,textFormatRuns"
                            }
                        })

                # 負值紅字規則 (H 欄)
                requests.append({
                    "addConditionalFormatRule": {
                        "rule": {
                            "ranges": [{"sheetId": ws.id, "startRowIndex": 3, "endRowIndex": data_rows_end_idx - 1, "startColumnIndex": 7, "endColumnIndex": 8}],
                            "booleanRule": {
                                "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                                "format": {"textFormat": {"foregroundColor": red_color}}
                            }
                        }, "index": 0
                    }
                })
                sh.batch_update({"requests": requests})
            return True
        except Exception as e:
            st.error(f"同步出錯：{e}")
            return False

    def parse_excel_data(uploaded_file, sheet_keyword, col_indices):
        try:
            content = uploaded_file.getvalue()
            xl = pd.ExcelFile(io.BytesIO(content))
            target_sheet = next((s for s in xl.sheet_names if sheet_keyword in s), xl.sheet_names[0])
            df = pd.read_excel(xl, sheet_name=target_sheet, header=None)
            
            date_display = ""
            try:
                row_3 = "".join(df.iloc[2].astype(str))
                match = re.search(r'(\d{7})([至\-~])(\d{7})', row_3)
                if match:
                    date_display = f"{match.group(1)[3:]}-{match.group(3)[3:]}"
            except:
                date_display = ""
                
            unit_data = {}
            for _, row in df.iterrows():
                u = get_standard_unit(row.iloc[0])
                if u and "合計" not in str(row.iloc[0]):
                    def clean(v):
                        if pd.isna(v) or str(v).strip().lower() == 'nan': return 0
                        try: return int(float(str(v).replace(',', '').strip()))
                        except: return 0
                    unit_data[u] = {'stop': clean(row.iloc[col_indices[0]]), 'cit': clean(row.iloc[col_indices[1]])}
            return unit_data, date_display
        except:
            return None, ""

    file_period, file_year = None, None
    for f in files:
        if "(1)" in f.name: file_year = f
        else: file_period = f
    if not file_year and len(files)>1: file_year = files[1]
    if not file_period and len(files)>0: file_period = files[0]

    d_week, date_w = parse_excel_data(file_period, "重點違規統計表", [15, 16])
    d_year, date_y = parse_excel_data(file_year, "(1)", [15, 16])
    d_last, _ = parse_excel_data(file_year, "(1)", [18, 19])
    
    if d_week and d_year:
        rows = []
        t = {k: 0 for k in ['ws', 'wc', 'ys', 'yc', 'ls', 'lc', 'diff', 'tgt']}
        for u in UNIT_ORDER:
            w, y, l = d_week.get(u, {'stop':0, 'cit':0}), d_year.get(u, {'stop':0, 'cit':0}), d_last.get(u, {'stop':0, 'cit':0})
            ys_sum, ls_sum = y['stop'] + y['cit'], l['stop'] + l['cit']
            tgt = TARGETS.get(u, 0)
            diff = int(ys_sum - ls_sum)
            rate = f"{(ys_sum/tgt):.1%}" if tgt > 0 else "0%"
            
            if u != '警備隊':
                t['diff'] += diff; t['tgt'] += tgt
            
            rows.append([u, w['stop'], w['cit'], y['stop'], y['cit'], l['stop'], l['cit'], diff if u != '警備隊' else "—", tgt, rate if u != '警備隊' else "—"])
            t['ws']+=w['stop']; t['wc']+=w['cit']; t['ys']+=y['stop']; t['yc']+=y['cit']; t['ls']+=l['stop']; t['lc']+=l['cit']
        
        total_rate = f"{((t['ys']+t['yc'])/t['tgt']):.1%}" if t['tgt']>0 else "0%"
        rows.insert(0, ['合計', t['ws'], t['wc'], t['ys'], t['yc'], t['ls'], t['lc'], t['diff'], t['tgt'], total_rate])
        rows.append([FOOTNOTE_TEXT] + [""] * 9)
        
        label_w = f"本期({date_w})" if date_w else "本期"
        label_y = f"本年累計({date_y})" if date_y else "本年累計"
        label_l = f"去年累計({date_y})" if date_y else "去年累計" 
        
        header_top = ['統計期間', label_w, label_w, label_y, label_y, label_l, label_l, '本年與去年同期比較', '目標值', '達成率']
        header_bottom = ['取締方式', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '', '', '']
        
        df_final = pd.DataFrame(rows, columns=pd.MultiIndex.from_arrays([header_top, header_bottom]))
        
        st.write("📊 報表預覽")
        st.dataframe(df_final, use_container_width=True)

        if sync_to_specified_sheet(df_final):
            st.success("✅ 同步完成！已保留雲端首尾格式，僅更新數據內容。")


def process_project(files):
    """🔥 強化專案原始邏輯"""
    PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"
    TARGET_CONFIG = {
        '聖亭所': [5, 115, 5, 16, 7, 10], '龍潭所': [6, 145, 7, 20, 9, 12],
        '中興所': [5, 115, 5, 16, 7, 10], '石門所': [3, 80, 4, 11, 5, 7],
        '高平所': [3, 80, 4, 11, 5, 7], '三和所': [2, 40, 2, 6, 2, 5],
        '交通分隊': [5, 115, 4, 16, 6, 8], '交通組': [0, 0, 0, 0, 0, 0], '警備隊': [0, 0, 0, 0, 0, 0]
    }
    CATS = ["酒後駕車", "闖紅燈", "嚴重超速", "車不讓人", "行人違規", "大型車違規"]
    LAW_MAP = {"酒後駕車": ["35條", "73條2項", "73條3項"], "闖紅燈": ["53條"], "嚴重超速": ["43條", "40條"], "車不讓人": ["44條", "48條"], "行人違規": ["78條"]}

    def map_unit_name(raw_name):
        raw = str(raw_name).strip()
        if '交通分隊' in raw:
            if '龍潭' in raw: return '交通分隊'
            if not any(ex in raw for ex in ['楊梅', '大溪', '平鎮', '中壢', '八德', '蘆竹', '龜山', '大園', '桃園']): return '交通分隊'
        if '交通組' in raw: return '交通組'
        if '警備隊' in raw: return '警備隊'
        for k in ['聖亭', '中興', '石門', '高平', '三和']: 
            if k in raw: return k + '所'
        if '龍潭派出所' in raw or raw in ['龍潭', '龍潭所']: return '龍潭所'
        return None

    def make_columns_unique(df):
        cols = pd.Series(df.columns.map(str))
        for dup in cols[cols.duplicated()].unique(): cols[cols == dup] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
        df.columns = cols
        return df

    def get_counts(df, unit, categories_list):
        df_c = df.reset_index(drop=True)
        if '單位' not in df_c.columns: return {cat: 0 for cat in categories_list}
        rows = df_c[df_c['單位'].apply(map_unit_name) == unit].copy()
        counts = {}
        for cat in categories_list:
            keywords = LAW_MAP.get(cat, [])
            matched = [c for c in df_c.columns if any(k in str(c) for k in keywords)]
            counts[cat] = int(rows[matched].sum().sum()) if not rows.empty else 0
        return counts

    f1_active = next((f for f in files if "強化" in f.name), None)
    f2_active = [f for f in files if "r17" in f.name.lower() or "砂石車" in f.name]

    def smart_read(f, **kwargs):
        fname = f.name
        f.seek(0)
        if fname.endswith('.csv'):
            try: return pd.read_csv(f, **kwargs)
            except: 
                f.seek(0); return pd.read_csv(f, encoding='cp950', **kwargs)
        return pd.read_excel(f, **kwargs)

    date_range_str = "未知期間"
    df1_h = smart_read(f1_active, nrows=10, header=None)
    for _, r in df1_h.iterrows():
        for cell in r.values:
            if '統計期間' in str(cell):
                raw = str(cell).replace('(入案日)', '').split('：')[-1].split(':')[-1].strip()
                m = re.search(r'([0-9年月日\-至\s]+)', raw)
                if m: date_range_str = m.group(1).replace('115', '').strip()

    df1 = make_columns_unique(smart_read(f1_active, skiprows=3)).reset_index(drop=True)
    df2_list = []
    for f in f2_active:
        df_t = smart_read(f, header=None)
        h_idx = None
        for i, row in df_t.head(30).iterrows():
            row_vals = [str(x).strip() for x in row.values]
            if '單位' in row_vals and '舉發總數' in row_vals:
                h_idx = i; break
        if h_idx is not None:
            df_c = df_t.iloc[h_idx+1:].copy()
            df_c.columns = [str(x).strip() for x in df_t.iloc[h_idx].values]
            df_c = make_columns_unique(df_c).reset_index(drop=True)
            df_c['來源檔名'] = str(f.name)
            df2_list.append(df_c)

    df2_all = pd.concat(df2_list, ignore_index=True)
    for c in ['舉發總數', '違反管制規定', '其他違規']:
        if c not in df2_all.columns: df2_all[c] = 0
        df2_all[c] = pd.to_numeric(df2_all[c], errors='coerce').fillna(0)

    df2_all['大型車純違規'] = (df2_all['舉發總數'] - df2_all['違反管制規定'] - df2_all['其他違規']).clip(lower=0)

    final_rows = []
    for unit in TARGET_CONFIG.keys():
        d15 = get_counts(df1, unit, CATS[:5])
        if unit == '交通分隊': u_rows = df2_all[(df2_all['來源檔名'].str.contains('大隊|交大', na=False)) & (df2_all['單位'].str.contains('龍潭', na=False))]
        else: u_rows = df2_all[(df2_all['單位'].apply(map_unit_name) == unit) & (~df2_all['來源檔名'].str.contains('大隊|交大', na=False))]
        h_sum = int(u_rows['大型車純違規'].sum()) if not u_rows.empty else 0
        res = [unit]
        for i, cat in enumerate(CATS):
            cnt = d15.get(cat, 0) if cat != "大型車違規" else h_sum
            tgt = TARGET_CONFIG[unit][i]
            res.extend([cnt, tgt, f"{(cnt/tgt*100):.1f}%" if tgt > 0 else "0.0%"])
        final_rows.append(res)

    headers = ["單位"]
    for cat in CATS: headers.extend([f"{cat}_取締件數", f"{cat}_目標值", f"{cat}_達成率"])
    df_f = pd.DataFrame(final_rows, columns=headers)

    total = ["合計"]
    for i in range(1, len(headers), 3):
        cs, ts = df_f.iloc[:, i].sum(), df_f.iloc[:, i+1].sum()
        total.extend([int(cs), int(ts), f"{(cs/ts*100):.1f}%" if ts > 0 else "0.0%"])
    df_f = pd.concat([pd.DataFrame([total], columns=headers), df_f], ignore_index=True)

    st.write(f"📊 {PROJECT_NAME}")
    st.dataframe(df_f, use_container_width=True, hide_index=True)

    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
    sh = gc.open_by_url(GOOGLE_SHEET_URL)
    ws = sh.worksheet(PROJECT_NAME)
    
    full_t = f"{PROJECT_NAME} (統計期間：{date_range_str})"
    ws.clear()
    ws.update(values=[
        [full_t] + [""] * 18,
        [""] + [c for c in CATS for _ in range(3)],
        ["單位"] + ["取締件數", "目標值", "達成率"] * 6
    ] + df_f.values.tolist())

    reqs = [
        {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 19}, "mergeType": "MERGE_ALL"}},
        {"updateCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1},
            "rows": [{"values": [{"userEnteredValue": {"stringValue": full_t}, "textFormatRuns": [
                {"startIndex": 0, "format": {"foregroundColor": {"red": 0.0, "green": 0.0, "blue": 1.0}, "bold": True, "fontSize": 16}},
                {"startIndex": len(PROJECT_NAME), "format": {"foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}, "bold": True, "fontSize": 16}}
            ]}]}], "fields": "userEnteredValue,textFormatRuns"}},
        {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 3, "startColumnIndex": 0, "endColumnIndex": 19}, "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment"}}
    ]
    sh.batch_update({"requests": reqs})
    st.write("✅ 雲端試算表已自動完成同步與美化！")


def process_accident(all_files):
    """🚑 交通事故原始邏輯 (保留 A1 及表頭紅字設定)"""
    def get_gsheet_rich_text_req(sheet_id, row_idx, col_idx, text):
        text = str(text)
        pattern = r'([0-9\(\)\/\-]+)'
        tokens = re.split(pattern, text)
        runs = []
        current_pos = 0
        for token in tokens:
            if not token: continue
            color = {"red": 1, "green": 0, "blue": 0} if re.match(pattern, token) else {"red": 0, "green": 0, "blue": 0}
            runs.append({"startIndex": current_pos, "format": {"foregroundColor": color, "bold": True}})
            current_pos += len(token)
        return {
            "updateCells": {
                "rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}],
                "fields": "userEnteredValue,textFormatRuns",
                "range": {"sheetId": sheet_id, "startRowIndex": row_idx, "endRowIndex": row_idx + 1, "startColumnIndex": col_idx, "endColumnIndex": col_idx + 1}
            }
        }

    def clean_traffic_data(df_raw):
        df_raw[0] = df_raw[0].astype(str)
        df_data = df_raw[df_raw[0].str.contains("所|總計|合計", na=False)].copy()
        cols = {0: "Station", 5: "A1_Deaths", 9: "A2_Injuries"}
        df_data = df_data.rename(columns=cols)
        for c in [5, 9]:
            target = cols[c]
            df_data[target] = pd.to_numeric(df_data[target].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
        df_data['Station_Short'] = df_data['Station'].str.replace('派出所', '所').str.replace('總計', '合計').str.strip()
        return df_data

    def build_traffic_table(df_wk, df_prev, df_cur, df_lst, stations, col_name, labels, is_a2=False):
        m = pd.merge(df_wk[['Station_Short', col_name]], df_prev[['Station_Short', col_name]], on='Station_Short', suffixes=('_wk', '_prev'))
        m = pd.merge(m, df_cur[['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_cur'})
        m = pd.merge(m, df_lst[['Station_Short', col_name]], on='Station_Short').rename(columns={col_name: col_name+'_lst'})
        m = m[m['Station_Short'].isin(stations)].copy()
        m['Station_Short'] = pd.Categorical(m['Station_Short'], categories=stations, ordered=True)
        m = m.sort_values('Station_Short')
        total = m.select_dtypes(include='number').sum().to_dict()
        total['Station_Short'] = '合計'
        m = pd.concat([pd.DataFrame([total]), m], ignore_index=True)
        m['Diff'] = m[col_name+'_cur'] - m[col_name+'_lst']
        if is_a2:
            m['Pct'] = m.apply(lambda x: f"{(x['Diff']/x[col_name+'_lst']):.2%}" if x[col_name+'_lst'] != 0 else "0.00%", axis=1)
            res = m[['Station_Short', col_name+'_wk', col_name+'_prev', col_name+'_cur', col_name+'_lst', 'Diff', 'Pct']]
            res.columns = ['統計期間', f'本期({labels["wk"]})', f'前期({labels["prev"]})', f'本年累計({labels["cur"]})', f'去年累計({labels["lst"]})', '本年與去年同期比較', '增減比例']
        else:
            res = m[['Station_Short', col_name+'_wk', col_name+'_cur', col_name+'_lst', 'Diff']]
            res.columns = ['統計期間', f'本期({labels["wk"]})', f'本年累計({labels["cur"]})', f'去年累計({labels["lst"]})', '本年與去年同期比較']
        return res

    meta = []
    for f in all_files:
        f.seek(0)
        try: df_raw = pd.read_csv(f, header=None)
        except: 
            f.seek(0)
            df_raw = pd.read_excel(f, header=None)
        
        sample_text = str(df_raw.iloc[:5, :5].values)
        dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', sample_text)
        if len(dates) >= 2:
            d_range = f"{int(dates[0][1]):02d}{int(dates[0][2]):02d}-{int(dates[1][1]):02d}{int(dates[1][2]):02d}"
            meta.append({
                'df': clean_traffic_data(df_raw),
                'year': int(dates[1][0]),
                'start_day': int(dates[0][1])*100 + int(dates[0][2]),
                'range': d_range,
                'is_cumu': (int(dates[0][1]) == 1 and int(dates[0][2]) == 1)
            })

    this_year = max(m['year'] for m in meta)
    f_lst = sorted([f for f in meta if f['year'] < this_year], key=lambda x: x['year'])[-1]
    f_cur = [f for f in meta if f['year'] == this_year and f['is_cumu']][0]
    period_files = sorted([f for f in meta if f['year'] == this_year and not f['is_cumu']], key=lambda x: x['start_day'])
    f_prev, f_wk = period_files[0], period_files[1]

    labels = {"wk": f_wk['range'], "prev": f_prev['range'], "cur": f_cur['range'], "lst": f_lst['range']}
    stations = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
    
    a1_res = build_traffic_table(f_wk['df'], f_prev['df'], f_cur['df'], f_lst['df'], stations, 'A1_Deaths', labels)
    a2_res = build_traffic_table(f_wk['df'], f_prev['df'], f_cur['df'], f_lst['df'], stations, 'A2_Injuries', labels, is_a2=True)

    st.subheader(f"📅 分析期間：{labels['wk']}")
    col1, col2 = st.columns(2)
    col1.write("A1 死亡人數統計")
    col1.dataframe(a1_res, hide_index=True)
    col2.write("A2 受傷人數統計")
    col2.dataframe(a2_res, hide_index=True)

    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
    sh = gc.open_by_url(GOOGLE_SHEET_URL)
    
    # 完全使用您的原始設計，僅更新 A2 以降，並針對 A2 表頭作數字轉紅格式
    for ws_idx, df in zip([2, 3], [a1_res, a2_res]):
        ws = sh.get_worksheet(ws_idx)
        ws.batch_clear(["A2:G20"])
        
        reqs = []
        for c_idx, c_name in enumerate(df.columns):
            reqs.append(get_gsheet_rich_text_req(ws.id, 1, c_idx, c_name))
        sh.batch_update({"requests": reqs})
        
        data_rows = [[int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in row] for row in df.values.tolist()]
        ws.update('A3', data_rows)

    st.write("✅ 交通事故統計完成，雲端紅字格式已更新！")


# ==========================================
# 4. 戰情室首頁與排程器
# ==========================================
with st.sidebar:
    st.title("🚓 龍潭分局戰情室")
    app_mode = st.selectbox("功能模組", ["🏠 智慧批次處理中心", "📂 PDF 轉 PPTX 工具"])
    st.info("💡 只要將所有報表全選、拖入，系統即自動分類並依序處理所有業務！")

if app_mode == "🏠 智慧批次處理中心":
    st.header("📈 交通數據全自動批次處理中心")
    st.markdown("請將您從警政系統匯出的所有報表（不限順序、不限數量）**一次全部拖入下方**。")
    
    uploads = st.file_uploader("📂 拖入所有報表檔案", type=["xlsx", "csv", "xls"], accept_multiple_files=True)
    
    if uploads:
        file_hash = sum([f.size for f in uploads]) + len(uploads)
        if st.session_state.get("last_processed_hash") == file_hash:
            st.success("✅ 目前上傳的檔案皆已全自動處理完畢！")
            st.info("💡 若要處理新報表，請直接點擊上方『X』刪除舊檔並拖入新檔案。")
        else:
            cat_files = {"科技執法": [], "重大違規": [], "超載統計": [], "強化專案": [], "交通事故": []}
            for f in uploads:
                name = f.name.lower()
                if "list" in name or "地點" in name or "科技" in name: cat_files["科技執法"].append(f)
                elif "stone" in name or "超載" in name: cat_files["超載統計"].append(f)
                elif "重大" in name: cat_files["重大違規"].append(f)
                elif "強化" in name or "專案" in name or "砂石車" in name or "r17" in name: cat_files["強化專案"].append(f)
                elif "a1" in name or "a2" in name or "事故" in name or "案件統計" in name: cat_files["交通事故"].append(f)
            
            st.divider()
            st.subheader("🚀 啟動全自動批次作業")
            
            try:
                if cat_files["科技執法"]:
                    with st.status(f"📸 自動處理【科技執法】({len(cat_files['科技執法'])} 份)...", expanded=True) as status:
                        process_tech_enforcement(cat_files["科技執法"])
                        status.update(label="✅ 科技執法處理完成！", state="complete")
                        
                if cat_files["超載統計"]:
                    with st.status(f"🚛 自動處理【超載統計】({len(cat_files['超載統計'])} 份)...", expanded=True) as status:
                        process_overload(cat_files["超載統計"])
                        status.update(label="✅ 超載統計處理完成！", state="complete")
                        
                if cat_files["重大違規"]:
                    with st.status(f"🚨 自動處理【重大交通違規】({len(cat_files['重大違規'])} 份)...", expanded=True) as status:
                        process_major(cat_files["重大違規"])
                        status.update(label="✅ 重大交通違規處理完成！", state="complete")
                        
                if cat_files["強化專案"]:
                    with st.status(f"🔥 自動處理【強化交通安全專案】({len(cat_files['強化專案'])} 份)...", expanded=True) as status:
                        process_project(cat_files["強化專案"])
                        status.update(label="✅ 強化交通安全專案處理完成！", state="complete")
                        
                if cat_files["交通事故"]:
                    with st.status(f"🚑 自動處理【交通事故】({len(cat_files['交通事故'])} 份)...", expanded=True) as status:
                        process_accident(cat_files["交通事故"])
                        status.update(label="✅ 交通事故處理完成！", state="complete")
                
                st.session_state["last_processed_hash"] = file_hash
                st.balloons()
                
            except Exception as e:
                st.error(f"⚠️ 批次處理發生例外錯誤：{e}")
                st.write(traceback.format_exc())

elif app_mode == "📂 PDF 轉 PPTX 工具":
    st.header("📂 PDF 行政文書轉 PPTX 簡報")
    pdf_file = st.file_uploader("上傳 PDF 檔案", type=["pdf"])
    if pdf_file and st.button("🚀 開始轉換"):
        with st.spinner("轉換中..."):
            try:
                images, prs = convert_from_bytes(pdf_file.read(), dpi=200), Presentation()
                for img in images:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format='PNG')
                    slide.shapes.add_picture(io.BytesIO(img_byte_arr.getvalue()), 0, 0, width=prs.slide_width, height=prs.slide_height)
                pptx_io = io.BytesIO()
                prs.save(pptx_io)
                st.success("✅ 轉換完成！")
                st.download_button("📥 下載 PPTX", data=pptx_io.getvalue(), file_name=f"{pdf_file.name.replace('.pdf', '')}_轉換.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            except Exception as e:
                st.error(f"錯誤：{e}")
