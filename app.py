import streamlit as st
import pandas as pd
import io
import re
import gspread
import smtplib
import calendar
import traceback
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

# ==========================================
# 1. 全局常數與設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
TO_EMAIL = "mbmmiqamhnd@gmail.com"

try:
    MY_EMAIL = st.secrets.get("email", {}).get("user", "")
    MY_PASSWORD = st.secrets.get("email", {}).get("password", "")
    GCP_CREDS = dict(st.secrets.get("gcp_service_account", {}))
except:
    MY_EMAIL, MY_PASSWORD, GCP_CREDS = "", "", None

# --- [重大違規常數] ---
MAJOR_UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
MAJOR_TARGETS = {'聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479, '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006}
MAJOR_FOOTNOTE = "重大交通違規指：「酒駕」、「闖紅燈」、「嚴重超速」、「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」及「不暫停讓行人」"

# --- [超載統計常數] ---
OVERLOAD_TARGETS = {'聖亭所': 20, '龍潭所': 27, '中興所': 20, '石門所': 16, '高平所': 14, '三和所': 8, '警備隊': 0, '交通分隊': 22}
OVERLOAD_UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'}
OVERLOAD_UNIT_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# --- [強化專案常數] ---
PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"
PROJECT_TARGETS = {
    '聖亭所': [5, 115, 5, 16, 7, 10], '龍潭所': [6, 145, 7, 20, 9, 12],
    '中興所': [5, 115, 5, 16, 7, 10], '石門所': [3, 80, 4, 11, 5, 7],
    '高平所': [3, 80, 4, 11, 5, 7], '三和所': [2, 40, 2, 6, 2, 5],
    '交通分隊': [5, 115, 4, 16, 6, 8], '交通組': [0, 0, 0, 0, 0, 0], '警備隊': [0, 0, 0, 0, 0, 0]
}
PROJECT_CATS = ["酒後駕車", "闖紅燈", "嚴重超速", "車不讓人", "行人違規", "大型車違規"]
PROJECT_LAW_MAP = {"酒後駕車": ["35條", "73條2項", "73條3項"], "闖紅燈": ["53條"], "嚴重超速": ["43條", "40條"], "車不讓人": ["44條", "48條"], "行人違規": ["78條"]}

# ==========================================
# 2. 🌟 完全恢復的格式化邏輯 🌟
# ==========================================
def sync_to_specified_sheet(df):
    """重大違規：保留 A1 總標題，僅更新 A2 以降並標紅字"""
    try:
        gc = gspread.service_account_from_dict(GCP_CREDS)
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
            # 標題括號紅字 (Row Index 1 = Google Sheets 裡的 Row 2)
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

def get_gsheet_rich_text_req(sheet_id, row_idx, col_idx, text):
    """交通事故：第二列表頭數字轉紅"""
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


# ==========================================
# 3. 業務邏輯處理
# ==========================================

def process_tech_enforcement(files):
    f = files[0]
    f.seek(0)
    df = pd.read_csv(f, encoding='cp950') if f.name.endswith('.csv') else pd.read_excel(f)
    df.columns = [str(c).strip() for c in df.columns]
    
    loc_col = next((c for c in df.columns if c in ['違規地點', '路口名稱', '地點']), None)
    if not loc_col:
        st.error("❌ 找不到『地點』相關欄位！")
        return
        
    df[loc_col] = df[loc_col].astype(str).str.replace('桃園市', '').str.replace('龍潭區', '').str.strip()
    yesterday = datetime.now() - timedelta(days=1)
    date_range_str = f"{yesterday.year - 1911}年1月1日至{yesterday.year - 1911}年{yesterday.month}月{yesterday.day}日"
    
    loc_summary = df[loc_col].value_counts().head(10).reset_index()
    loc_summary.columns = ['路段名稱', '舉發件數']
    
    st.write("📊 **科技執法路段排行：**")
    st.dataframe(loc_summary, hide_index=True)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook, ws = writer.book, writer.book.add_worksheet('科技執法成效統計')
        ws.write_rich_string('A1', workbook.add_format({'bold': True, 'font_size': 24, 'color': 'blue'}), '科技執法成效 ', 
                             workbook.add_format({'bold': True, 'font_size': 24, 'color': 'red'}), f'({date_range_str})')
        ws.write('A2', '統計期間', workbook.add_format({'align': 'center', 'border': 1}))
        ws.write('B2', date_range_str, workbook.add_format({'border': 1, 'color': 'red', 'align': 'center'}))
        header_fmt = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1, 'bold': True, 'align': 'center'})
        data_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        ws.write('A3', '路口名稱', header_fmt); ws.write('B3', '舉發件數', header_fmt)
        
        for i, row in loc_summary.iterrows():
            ws.write(i+3, 0, row['路段名稱'], data_fmt); ws.write(i+3, 1, row['舉發件數'], data_fmt)
        ws.write(len(loc_summary)+3, 0, '舉發總數', workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFFFCC', 'align': 'center'}))
        ws.write(len(loc_summary)+3, 1, len(df), workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFFFCC', 'align': 'center'}))
        
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({'name': '舉發件數', 'categories': ['科技執法成效統計', 3, 0, len(loc_summary)+2, 0], 'values': ['科技執法成效統計', 3, 1, len(loc_summary)+2, 1], 'data_labels': {'value': True}})
        ws.insert_chart('D2', chart, {'x_scale': 1.5, 'y_scale': 1.5})
        
    if GCP_CREDS:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        ws = sh.worksheet("科技執法-路段排行") if "科技執法-路段排行" in [s.title for s in sh.worksheets()] else sh.add_worksheet(title="科技執法-路段排行", rows="100", cols="20")
        ws.clear()
        title_text = f"科技執法成效 ({date_range_str})"
        ws.update(values=[[title_text, ""], ["路段名稱", "舉發件數"]] + loc_summary.values.tolist() + [["舉發總數", len(df)]])
        reqs = {"requests": [{"updateCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1},
                "rows": [{"values": [{"userEnteredValue": {"stringValue": title_text},
                "textFormatRuns": [{"startIndex": 0, "format": {"foregroundColor": {"red": 0.0, "green": 0.0, "blue": 1.0}, "bold": True, "fontSize": 24}},
                                   {"startIndex": len("科技執法成效 "), "format": {"foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}, "bold": True, "fontSize": 24}}]}]}],
                "fields": "userEnteredValue,textFormatRuns"}}]}
        sh.batch_update(reqs)
        
    if MY_EMAIL and MY_PASSWORD:
        msg = MIMEMultipart()
        msg['From'], msg['To'], msg['Subject'] = MY_EMAIL, TO_EMAIL, f"科技執法統計報告({date_range_str})"
        msg.attach(MIMEText(f"長官好，科技執法路段排行報表已完成。\n\n統計期間：{date_range_str}\n舉發總件數：{len(df)} 件", 'plain'))
        part = MIMEApplication(output.getvalue(), Name="Tech_Enforcement.xlsx")
        part.add_header('Content-Disposition', 'attachment', filename="Tech_Enforcement.xlsx")
        msg.attach(part)
        with smtplib.SMTP('smtp.gmail.com', 587) as s:
            s.starttls()
            s.login(MY_EMAIL, MY_PASSWORD)
            s.send_message(msg)

def process_overload(files):
    f_wk, f_yt, f_ly = None, None, None
    for f in files:
        if "(1)" in f.name: f_yt = f
        elif "(2)" in f.name: f_ly = f
        else: f_wk = f
        
    def parse_rpt(f):
        f.seek(0)
        counts, s, e = {}, "0000000", "0000000"
        text_block = pd.read_excel(f, header=None, nrows=15).to_string()
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
                        short = OVERLOAD_UNIT_MAP.get(u, u)
                        if short in OVERLOAD_UNIT_ORDER: counts[short] = counts.get(short, 0) + int(nums[-1])
                        u = None
        return counts, s, e

    d_wk, s_wk, e_wk = parse_rpt(f_wk); d_yt, s_yt, e_yt = parse_rpt(f_yt); d_ly, s_ly, e_ly = parse_rpt(f_ly)
    raw_wk, raw_yt, raw_ly = f"本期 ({s_wk[-4:]}~{e_wk[-4:]})", f"本年累計 ({s_yt[-4:]}~{e_yt[-4:]})", f"去年累計 ({s_ly[-4:]}~{e_ly[-4:]})"
    
    body = []
    for u in OVERLOAD_UNIT_ORDER:
        yv, tv = d_yt.get(u, 0), OVERLOAD_TARGETS.get(u, 0)
        body.append({'統計期間': u, raw_wk: d_wk.get(u, 0), raw_yt: yv, raw_ly: d_ly.get(u, 0), '本年與去年同期比較': yv - d_ly.get(u, 0), '目標值': tv, '達成率': f"{yv/tv:.0%}" if tv > 0 else "—"})
    df_body = pd.DataFrame(body)
    sum_v = df_body[df_body['統計期間'] != '警備隊'][[raw_wk, raw_yt, raw_ly, '目標值']].sum()
    total_row = pd.DataFrame([{'統計期間': '合計', raw_wk: sum_v[raw_wk], raw_yt: sum_v[raw_yt], raw_ly: sum_v[raw_ly], '本年與去年同期比較': sum_v[raw_yt] - sum_v[raw_ly], '目標值': sum_v['目標值'], '達成率': f"{sum_v[raw_yt]/sum_v['目標值']:.0%}" if sum_v['目標值'] > 0 else "0%"}])
    df_final = pd.concat([total_row, df_body], ignore_index=True)
    
    st.write("📊 **超載統計結果：**")
    st.dataframe(df_final, hide_index=True)

    y, m, d = int(e_yt[:3])+1911, int(e_yt[3:5]), int(e_yt[5:])
    prog_str = f"{((date(y, m, d) - date(y, 1, 1)).days + 1) / (366 if calendar.isleap(y) else 365):.1%}"
    f_plain = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt[:3]}年{e_yt[3:5]}月{e_yt[5:]}日 (入案日期)應達成率為{prog_str}"

    if GCP_CREDS:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        ws = gc.open_by_url(GOOGLE_SHEET_URL).get_worksheet(1)
        ws.update(range_name='A1', values=[['取締超載違規件數統計表']])
        ws.update(range_name='A2', values=[df_final.columns.tolist()] + df_final.values.tolist())
        ws.update(range_name=f'A{2 + len(df_final) + 1}', values=[[f_plain]])
        
    if MY_EMAIL:
        df_excel_buffer = io.BytesIO()
        with pd.ExcelWriter(df_excel_buffer, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, startrow=1, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            title_format = writer.book.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'font_color': 'blue'})
            worksheet.merge_range('A1:G1', '取締超載違規件數統計表', title_format)
            worksheet.set_column('A:A', 15); worksheet.set_column('B:G', 12)
            
        msg = MIMEMultipart()
        msg['Subject'] = Header(f"🚛 超載報表 - {e_yt} ({prog_str})", 'utf-8').encode()
        msg['From'], msg['To'] = MY_EMAIL, TO_EMAIL
        msg.attach(MIMEText("自動產生的超載報表已同步，請查閱附件。", 'plain'))
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(df_excel_buffer.getvalue())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="Overload_Report.xlsx"')
        msg.attach(part)
        with smtplib.SMTP('smtp.gmail.com', 587) as s:
            s.starttls(); s.login(MY_EMAIL, MY_PASSWORD); s.send_message(msg)

def process_major(files):
    f_period, f_year = None, None
    for f in files:
        if "(1)" in f.name: f_year = f
        else: f_period = f
    if not f_year: f_year = files[1]
    if not f_period: f_period = files[0]

    def parse_ex(f, sheet_kw, cols):
        f.seek(0)
        xl = pd.ExcelFile(f)
        df = pd.read_excel(xl, sheet_name=next((s for s in xl.sheet_names if sheet_kw in s), xl.sheet_names[0]), header=None)
        dt = ""
        try:
            m = re.search(r'(\d{7})([至\-~])(\d{7})', "".join(df.iloc[2].astype(str)))
            if m: dt = f"{m.group(1)[3:]}-{m.group(3)[3:]}"
        except: pass
        
        def get_u(n):
            n = str(n).strip()
            if '分隊' in n: return '交通分隊'
            if '科技' in n or '交通組' in n: return '科技執法'
            if '警備' in n: return '警備隊'
            for k in ['聖亭', '龍潭', '中興', '石門', '高平', '三和']:
                if k in n: return k + '所'
            return None
            
        def clean_val(v):
            if pd.isna(v) or str(v).strip().lower() == 'nan': return 0
            try: return int(float(str(v).replace(',', '').strip() or 0))
            except: return 0

        udata = {}
        for _, r in df.iterrows():
            u = get_u(r.iloc[0])
            if u and "合計" not in str(r.iloc[0]):
                udata[u] = {'stop': clean_val(r.iloc[cols[0]]), 'cit': clean_val(r.iloc[cols[1]])}
        return udata, dt

    d_wk, date_w = parse_ex(f_period, "重點違規統計表", [15, 16])
    d_year, date_y = parse_ex(f_year, "(1)", [15, 16])
    d_last, _ = parse_ex(f_year, "(1)", [18, 19])

    rows = []
    t = {k: 0 for k in ['ws', 'wc', 'ys', 'yc', 'ls', 'lc', 'diff', 'tgt']}
    for u in MAJOR_UNIT_ORDER:
        w, y, l = d_wk.get(u, {'stop':0, 'cit':0}), d_year.get(u, {'stop':0, 'cit':0}), d_last.get(u, {'stop':0, 'cit':0})
        ys_sum, ls_sum = y['stop'] + y['cit'], l['stop'] + l['cit']
        tgt = MAJOR_TARGETS.get(u, 0)
        diff = int(ys_sum - ls_sum)
        rate = f"{(ys_sum/tgt):.1%}" if tgt > 0 else "0%"
        if u != '警備隊':
            t['diff'] += diff; t['tgt'] += tgt
        rows.append([u, w['stop'], w['cit'], y['stop'], y['cit'], l['stop'], l['cit'], diff if u != '警備隊' else "—", tgt, rate if u != '警備隊' else "—"])
        t['ws']+=w['stop']; t['wc']+=w['cit']; t['ys']+=y['stop']; t['yc']+=y['cit']; t['ls']+=l['stop']; t['lc']+=l['cit']

    total_rate = f"{((t['ys']+t['yc'])/t['tgt']):.1%}" if t['tgt']>0 else "0%"
    rows.insert(0, ['合計', t['ws'], t['wc'], t['ys'], t['yc'], t['ls'], t['lc'], t['diff'], t['tgt'], total_rate])
    rows.append([MAJOR_FOOTNOTE] + [""] * 9)
    
    label_w = f"本期({date_w})" if date_w else "本期"
    label_y = f"本年累計({date_y})" if date_y else "本年累計"
    label_l = f"去年累計({date_y})" if date_y else "去年累計" 
    
    header_top = ['統計期間', label_w, label_w, label_y, label_y, label_l, label_l, '本年與去年同期比較', '目標值', '達成率']
    header_bottom = ['取締方式', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '', '', '']
    
    df_final = pd.DataFrame(rows, columns=pd.MultiIndex.from_arrays([header_top, header_bottom]))
    st.write("📊 **重大違規統計結果：**")
    st.dataframe(df_final, use_container_width=True)

    if GCP_CREDS:
        sync_to_specified_sheet(df_final)

def process_project(files):
    f1 = next((f for f in files if "強化" in f.name), None)
    f2_list = [f for f in files if "r17" in f.name.lower() or "砂石車" in f.name]
    
    def s_read(f, **kwargs):
        f.seek(0)
        if f.name.endswith('.csv'):
            try: return pd.read_csv(f, **kwargs)
            except: 
                f.seek(0); return pd.read_csv(f, encoding='cp950', **kwargs)
        return pd.read_excel(f, **kwargs)

    date_str = "未知期間"
    df1_h = s_read(f1, nrows=10, header=None)
    for _, r in df1_h.iterrows():
        for c in r.values:
            if '統計期間' in str(c):
                m = re.search(r'([0-9年月日\-至\s]+)', str(c).replace('(入案日)', '').split('：')[-1].split(':')[-1].strip())
                if m: date_str = m.group(1).replace('115', '').strip()

    def m_uniq(df):
        cols = pd.Series(df.columns.map(str))
        for d in cols[cols.duplicated()].unique(): cols[cols == d] = [f"{d}_{i}" if i!=0 else d for i in range(sum(cols == d))]
        df.columns = cols; return df
        
    df1 = m_uniq(s_read(f1, skiprows=3)).reset_index(drop=True)
    df2_all = []
    for f in f2_list:
        df_t = s_read(f, header=None)
        h_idx = next((i for i, r in df_t.head(30).iterrows() if '單位' in [str(x).strip() for x in r.values] and '舉發總數' in [str(x).strip() for x in r.values]), None)
        if h_idx is not None:
            df_c = df_t.iloc[h_idx+1:].copy()
            df_c.columns = [str(x).strip() for x in df_t.iloc[h_idx].values]
            df_c = m_uniq(df_c).reset_index(drop=True)
            df_c['來源檔名'] = str(f.name)
            df2_all.append(df_c)
            
    df2 = pd.concat(df2_all, ignore_index=True)
    for c in ['舉發總數', '違反管制規定', '其他違規']: df2[c] = pd.to_numeric(df2.get(c, 0), errors='coerce').fillna(0)
    df2['大型車純違規'] = (df2['舉發總數'] - df2['違反管制規定'] - df2['其他違規']).clip(lower=0)

    def get_unit(raw):
        raw = str(raw).strip()
        if '交通分隊' in raw: return '交通分隊' if '龍潭' in raw or not any(x in raw for x in ['楊梅','大溪','平鎮','中壢','八德','蘆竹','龜山','大園','桃園']) else None
        if '交通組' in raw: return '交通組'
        if '警備隊' in raw: return '警備隊'
        for k in ['聖亭', '中興', '石門', '高平', '三和']: 
            if k in raw: return k + '所'
        if '龍潭派出所' in raw or raw in ['龍潭', '龍潭所']: return '龍潭所'
        return None

    def get_c(unit):
        r = df1[df1.get('單位', pd.Series()).apply(get_unit) == unit]
        return {cat: int(r[[c for c in df1.columns if any(k in str(c) for k in PROJECT_LAW_MAP.get(cat, []))]].sum().sum()) if not r.empty else 0 for cat in PROJECT_CATS[:5]}

    final_rows = []
    for u, tgts in PROJECT_TARGETS.items():
        d15 = get_c(u)
        u_r = df2[(df2['來源檔名'].str.contains('大隊|交大', na=False)) & (df2['單位'].str.contains('龍潭', na=False))] if u == '交通分隊' else df2[(df2['單位'].apply(get_unit) == u) & (~df2['來源檔名'].str.contains('大隊|交大', na=False))]
        h_sum = int(u_r['大型車純違規'].sum()) if not u_r.empty else 0
        res = [u]
        for i, cat in enumerate(PROJECT_CATS):
            cnt = d15.get(cat, 0) if cat != "大型車違規" else h_sum
            res.extend([cnt, tgts[i], f"{(cnt/tgts[i]*100):.1f}%" if tgts[i] > 0 else "0.0%"])
        final_rows.append(res)

    headers = ["單位"] + [f"{cat}_{x}" for cat in PROJECT_CATS for x in ["取締件數", "目標值", "達成率"]]
    df_f = pd.DataFrame(final_rows, columns=headers)
    
    t_row = ["合計"]
    for i in range(1, len(headers), 3):
        cs, ts = df_f.iloc[:, i].sum(), df_f.iloc[:, i+1].sum()
        t_row.extend([int(cs), int(ts), f"{(cs/ts*100):.1f}%" if ts > 0 else "0.0%"])
    df_f = pd.concat([pd.DataFrame([t_row], columns=headers), df_f], ignore_index=True)
    
    st.write(f"📊 **{PROJECT_NAME} 統計結果：**")
    st.dataframe(df_f, hide_index=True)

    if GCP_CREDS:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        ws = gc.open_by_url(GOOGLE_SHEET_URL).worksheet(PROJECT_NAME)
        full_t = f"{PROJECT_NAME} (統計期間：{date_str})"
        ws.clear()
        ws.update(values=[[full_t] + [""] * 18, [""] + [c for c in PROJECT_CATS for _ in range(3)], ["單位"] + ["取締件數", "目標值", "達成率"] * 6] + df_f.values.tolist())
        
        reqs = [
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 19}, "mergeType": "MERGE_ALL"}},
            {"updateCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1},
                "rows": [{"values": [{"userEnteredValue": {"stringValue": full_t}, "textFormatRuns": [
                {"startIndex": 0, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 1}, "bold": True, "fontSize": 16}},
                {"startIndex": len(PROJECT_NAME), "format": {"foregroundColor": {"red": 1, "green": 0, "blue": 0}, "bold": True, "fontSize": 16}}]}]}], "fields": "userEnteredValue,textFormatRuns"}},
            {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 3, "startColumnIndex": 0, "endColumnIndex": 19}, "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment"}}
        ]
        ws.spreadsheet.batch_update({"requests": reqs})

def process_accident(files):
    meta = []
    for f in files:
        f.seek(0)
        df_raw = pd.read_csv(f, header=None) if f.name.endswith('.csv') else pd.read_excel(f, header=None)
        dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', str(df_raw.iloc[:5, :5].values))
        if len(dates) >= 2:
            df_raw[0] = df_raw[0].astype(str)
            df_data = df_raw[df_raw[0].str.contains("所|總計|合計", na=False)].rename(columns={0: "Station", 5: "A1_Deaths", 9: "A2_Injuries"})
            for c in ["A1_Deaths", "A2_Injuries"]: df_data[c] = pd.to_numeric(df_data[c].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
            df_data['Station_Short'] = df_data['Station'].str.replace('派出所', '所').str.replace('總計', '合計').str.strip()
            
            meta.append({'df': df_data, 'year': int(dates[1][0]), 'start_day': int(dates[0][1])*100 + int(dates[0][2]), 
                         'range': f"{int(dates[0][1]):02d}{int(dates[0][2]):02d}-{int(dates[1][1]):02d}{int(dates[1][2]):02d}", 'is_cumu': (int(dates[0][1]) == 1 and int(dates[0][2]) == 1)})
                         
    this_year = max(m['year'] for m in meta)
    f_lst = sorted([f for f in meta if f['year'] < this_year], key=lambda x: x['year'])[-1]
    f_cur = next(f for f in meta if f['year'] == this_year and f['is_cumu'])
    period_files = sorted([f for f in meta if f['year'] == this_year and not f['is_cumu']], key=lambda x: x['start_day'])
    f_prev, f_wk = period_files[0], period_files[1]

    labels = {"wk": f_wk['range'], "prev": f_prev['range'], "cur": f_cur['range'], "lst": f_lst['range']}
    stations = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
    
    def bld_tbl(c_name, is_a2=False):
        m = pd.merge(f_wk['df'][['Station_Short', c_name]], f_prev['df'][['Station_Short', c_name]], on='Station_Short', suffixes=('_wk', '_prev'))
        m = pd.merge(pd.merge(m, f_cur['df'][['Station_Short', c_name]].rename(columns={c_name: c_name+'_cur'}), on='Station_Short'), f_lst['df'][['Station_Short', c_name]].rename(columns={c_name: c_name+'_lst'}), on='Station_Short')
        m = m[m['Station_Short'].isin(stations)].copy()
        m['Station_Short'] = pd.Categorical(m['Station_Short'], categories=stations, ordered=True)
        m = pd.concat([pd.DataFrame([dict(m.select_dtypes(include='number').sum().to_dict(), Station_Short='合計')]), m.sort_values('Station_Short')], ignore_index=True)
        m['Diff'] = m[c_name+'_cur'] - m[c_name+'_lst']
        if is_a2:
            m['Pct'] = m.apply
