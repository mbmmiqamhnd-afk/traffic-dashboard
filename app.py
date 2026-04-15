import streamlit as st
import pandas as pd
import io
import re
import gspread
import traceback
from datetime import datetime, timedelta
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

try:
    GCP_CREDS = dict(st.secrets.get("gcp_service_account", {}))
except:
    GCP_CREDS = None

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
PROJECT_LAW_MAP = {
    "酒後駕車": ["35條", "73條2項", "73條3項"], 
    "闖紅燈": ["53條"], 
    "嚴重超速": ["43條", "40條"],  
    "車不讓人": ["44條", "48條"], 
    "行人違規": ["78條"]
}

# ==========================================
# 2. 輔助工具區
# ==========================================
def get_gsheet_rich_text_req(sheet_id, row_idx, col_idx, text):
    """🚑交通事故專用：Google Sheets 標題括號與數字轉紅字"""
    text = str(text)
    pattern = r'([0-9\(\)\/\-]+)'
    tokens = re.split(pattern, text)
    runs = []
    current_pos = 0
    for token in tokens:
        if not token: continue
        color = {"red": 1.0, "green": 0.0, "blue": 0.0} if re.match(pattern, token) else {"red": 0.0, "green": 0.0, "blue": 0.0}
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
# 3. 業務邏輯處理區
# ==========================================

# ----------------- [1. 科技執法] -----------------
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
    
    if GCP_CREDS:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        ws_name = "科技執法-路段排行"
        ws = sh.worksheet(ws_name) if ws_name in [s.title for s in sh.worksheets()] else sh.add_worksheet(title=ws_name, rows="100", cols="20")
        ws.clear()
        title_text = f"科技執法成效 ({date_range_str})"
        ws.update(range_name='A1', values=[[title_text, ""], ["路段名稱", "舉發件數"]] + loc_summary.values.tolist() + [["舉發總數", len(df)]])
        
        reqs = {"requests": [{"updateCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1},
                "rows": [{"values": [{"userEnteredValue": {"stringValue": title_text},
                "textFormatRuns": [{"startIndex": 0, "format": {"foregroundColor": {"red": 0.0, "green": 0.0, "blue": 1.0}, "bold": True, "fontSize": 24}},
                                   {"startIndex": len("科技執法成效 "), "format": {"foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}, "bold": True, "fontSize": 24}}]}]}],
                "fields": "userEnteredValue,textFormatRuns"}}]}
        sh.batch_update(reqs)

# ----------------- [2. 超載統計] -----------------
def process_overload(files):
    f_wk, f_yt, f_ly = None, None, None
    for f in files:
        if "(1)" in f.name: f_yt = f
        elif "(2)" in f.name: f_ly = f
        else: f_wk = f
        
    def parse_rpt(f):
        if not f: return {}, "0000000", "0000000"
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
                rs = " ".join([str(x) for x in r.values])
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

    if GCP_CREDS:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        ws = sh.get_worksheet(1)
        ws.update(range_name='A1', values=[['取締超載違規件數統計表']])
        ws.update(range_name='A2', values=[df_final.columns.tolist()] + df_final.values.tolist())
        
        requests = []
        for i, col_name in enumerate(df_final.columns):
            if "(" in col_name:
                p_start = col_name.find("(")
                requests.append({
                    "updateCells": {
                        "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": i, "endColumnIndex": i+1},
                        "rows": [{ "values": [{ "textFormatRuns": [
                            {"startIndex": 0, "format": {"foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}, "bold": True}},
                            {"startIndex": p_start, "format": {"foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}, "bold": True}}
                        ], "userEnteredValue": {"stringValue": col_name} }] }],
                        "fields": "userEnteredValue,textFormatRuns"
                    }
                })
        if requests: sh.batch_update({"requests": requests})

# ----------------- [3. 重大交通違規] -----------------
def process_major(files):
    if len(files) < 2:
        st.error("❌ 需要至少上傳『本期』與『(1)累計』報表才能進行重大違規比對。")
        return

    sorted_files = sorted(files, key=lambda x: x.size)
    f_wk, f_year = sorted_files[0], sorted_files[1]
    if "(1)" in f_wk.name: f_wk, f_year = f_year, f_wk

    def get_robust_date(df):
        try:
            raw_cells = [str(val) for val in df.head(10).values.flatten() if pd.notna(val)]
            clean_text = re.sub(r'\s+', '', "".join(raw_cells))
            match = re.search(r'1\d{2}(\d{4})[至\-~]1\d{2}(\d{4})', clean_text)
            if match: return f"{match.group(1)}-{match.group(2)}"
            dates = re.findall(r'(?<!\d)1\d{6}(?!\d)', clean_text)
            if len(dates) >= 2: return f"{dates[0][-4:]}-{dates[1][-4:]}"
            return ""
        except: return ""

    def parse_major_data(f, sheet_kw, col_pair):
        f.seek(0)
        xl = pd.ExcelFile(f)
        sn = next((s for s in xl.sheet_names if sheet_kw in s), xl.sheet_names[0])
        df = pd.read_excel(xl, sheet_name=sn, header=None)
        dt_str = get_robust_date(df)
        
        def clean_unit(n):
            n = str(n).strip()
            if '分隊' in n: return '交通分隊'
            if any(k in n for k in ['科技', '交通組']): return '科技執法'
            if '警備' in n: return '警備隊'
            for k in ['聖亭', '龍潭', '中興', '石門', '高平', '三和']:
                if k in n: return k + '所'
            return None

        def to_i(v):
            try: return int(float(str(v).replace(',', '').strip()))
            except: return 0

        res = {}
        for _, r in df.iterrows():
            u = clean_unit(r.iloc[0])
            if u and "合計" not in str(r.iloc[0]):
                res[u] = {'stop': to_i(r.iloc[col_pair[0]]), 'cit': to_i(r.iloc[col_pair[1]])}
        return res, dt_str

    d_wk, date_wk = parse_major_data(f_wk, "重點違規", [15, 16])
    d_year, date_yr = parse_major_data(f_year, "(1)", [15, 16])
    d_last, _ = parse_major_data(f_year, "(1)", [18, 19])

    table_rows = []
    summary = {k: 0 for k in ['ws', 'wc', 'ys', 'yc', 'ls', 'lc', 'diff', 'tgt']}
    
    for u in MAJOR_UNIT_ORDER:
        w_data = d_wk.get(u, {'stop':0, 'cit':0})
        y_data = d_year.get(u, {'stop':0, 'cit':0})
        l_data = d_last.get(u, {'stop':0, 'cit':0})
        
        y_total = y_data['stop'] + y_data['cit']
        l_total = l_data['stop'] + l_data['cit']
        tgt = MAJOR_TARGETS.get(u, 0)
        diff = int(y_total - l_total)
        
        rate = f"{(y_total/tgt):.1%}" if tgt > 0 else "0%"
        
        if u != '警備隊':
            summary['diff'] += diff; summary['tgt'] += tgt
            
        table_rows.append([u, w_data['stop'], w_data['cit'], y_data['stop'], y_data['cit'], l_data['stop'], l_data['cit'], diff if u != '警備隊' else "—", tgt, rate if u != '警備隊' else "—"])
        
        summary['ws'] += w_data['stop']; summary['wc'] += w_data['cit']
        summary['ys'] += y_data['stop']; summary['yc'] += y_data['cit']
        summary['ls'] += l_data['stop']; summary['lc'] += l_data['cit']

    total_rate = f"{((summary['ys']+summary['yc'])/summary['tgt']):.1%}" if summary['tgt'] > 0 else "0%"
    table_rows.insert(0, ['合計', summary['ws'], summary['wc'], summary['ys'], summary['yc'], summary['ls'], summary['lc'], summary['diff'], summary['tgt'], total_rate])
    table_rows.append([MAJOR_FOOTNOTE] + [""] * 9)
    
    h_wk = f"本期({date_wk})" if date_wk else "本期"
    h_yr = f"本年累計({date_yr})" if date_yr else "本年累計"
    h_ls = f"去年累計({date_yr})" if date_yr else "去年累計"
    
    header_1 = ['統計期間', h_wk, h_wk, h_yr, h_yr, h_ls, h_ls, '本年與去年同期比較', '目標值', '達成率']
    header_2 = ['取締方式', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '', '', '']
    
    df_result = pd.DataFrame(table_rows, columns=pd.MultiIndex.from_arrays([header_1, header_2]))
    st.write("📊 **重大違規統計結果：**")
    st.dataframe(df_result, use_container_width=True)

    if GCP_CREDS:
        try:
            gc = gspread.service_account_from_dict(GCP_CREDS)
            sh = gc.open_by_url(GOOGLE_SHEET_URL)
            ws = sh.get_worksheet(0)
            
            titles = df_result.columns.tolist()
            top_row = [t[0] for t in titles]
            bottom_row = [t[1] for t in titles]
            data_body = df_result.values.tolist()
            data_list = [top_row, bottom_row] + data_body
            ws.update(range_name='A2', values=data_list)
            
            red_color = {"red": 1.0, "green": 0.0, "blue": 0.0}
            black_color = {"red": 0.0, "green": 0.0, "blue": 0.0}
            requests = []
            
            for i, text in enumerate(top_row):
                if "(" in text:
                    p_start = text.find("(")
                    requests.append({
                        "updateCells": {
                            "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": i, "endColumnIndex": i+1},
                            "rows": [{ "values": [{ "textFormatRuns": [
                                {"startIndex": 0, "format": {"foregroundColor": black_color, "bold": True}},
                                {"startIndex": p_start, "format": {"foregroundColor": red_color, "bold": True}}
                            ], "userEnteredValue": {"stringValue": text} }] }],
                            "fields": "userEnteredValue,textFormatRuns"
                        }
                    })

            red_fmt = {"textFormat": {"foregroundColor": red_color}}
            black_fmt = {"textFormat": {"foregroundColor": black_color}}
            for r_idx, row_vals in enumerate(data_body):
                val = row_vals[7]
                target_row = 3 + r_idx
                is_negative = isinstance(val, (int, float)) and val < 0
                fmt = red_fmt if is_negative else black_fmt
                requests.append({
                    "repeatCell": {
                        "range": {"sheetId": ws.id, "startRowIndex": target_row, "endRowIndex": target_row + 1, "startColumnIndex": 7, "endColumnIndex": 8},
                        "cell": {"userEnteredFormat": fmt},
                        "fields": "userEnteredFormat.textFormat.foregroundColor"
                    }
                })
                
            if requests:
                sh.batch_update({"requests": requests})
            st.write("✅ 重大違規雲端格式與數據同步完成")
        except Exception as e:
            st.error(f"雲端同步出錯：{e}")

# ----------------- [4. 強化專案] -----------------
def process_project(files):
    f1 = next((f for f in files if any(k in f.name for k in ["強化", "法條", "自選匯出"])), None)
    f2_list = [f for f in files if any(k in f.name.upper() for k in ["R17", "砂石", "大貨"])]

    if not f1 or not f2_list:
        st.error("❌ 找不到強化專案報表！需包含法條與R17大型車資料。")
        return

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
        u_r = df2[df2['單位'].apply(get_unit) == u]
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
        ws.update(range_name='A1', values=[[full_t] + [""] * 18, [""] + [c for c in PROJECT_CATS for _ in range(3)], ["單位"] + ["取締件數", "目標值", "達成率"] * 6] + df_f.values.tolist())
        
        red_cells = []
        for c_idx, cat in enumerate(PROJECT_CATS):
            rate_col_idx = 3 + c_idx * 3
            valid_rates = []
            for row_idx, row in df_f.iterrows():
                unit = row['單位']
                if unit in ['合計', '警備隊', '交通組']: continue
                target_val = row[f"{cat}_目標值"]
                if target_val > 0:
                    try:
                        rate_val = float(str(row[f"{cat}_達成率"]).replace('%', ''))
                        valid_rates.append((row_idx, rate_val))
                    except: pass
            
            if valid_rates:
                valid_rates.sort(key=lambda x: x[1])
                threshold = valid_rates[1][1] if len(valid_rates) > 1 else valid_rates[0][1]
                for row_idx, rate_val in valid_rates:
                    if rate_val <= threshold and rate_val < 100.0:
                        red_cells.append((3 + row_idx, rate_col_idx))

        reqs = [
            {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 3, "endRowIndex": 20, "startColumnIndex": 0, "endColumnIndex": 19}, "cell": {"userEnteredFormat": {"textFormat": {"foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}, "bold": False}}}, "fields": "userEnteredFormat.textFormat.foregroundColor,userEnteredFormat.textFormat.bold"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 19}, "mergeType": "MERGE_ALL"}},
            {"updateCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1}, "rows": [{"values": [{"userEnteredValue": {"stringValue": full_t}, "textFormatRuns": [{"startIndex": 0, "format": {"foregroundColor": {"red": 0.0, "green": 0.0, "blue": 1.0}, "bold": True, "fontSize": 16}}, {"startIndex": len(PROJECT_NAME), "format": {"foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}, "bold": True, "fontSize": 16}}]}]}], "fields": "userEnteredValue,textFormatRuns"}},
            {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 3, "startColumnIndex": 0, "endColumnIndex": 19}, "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}}, "fields": "userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment"}}
        ]
        
        red_format = {"textFormat": {"foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}, "bold": True}}
        for r, c in red_cells:
            reqs.append({"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": r, "endRowIndex": r+1, "startColumnIndex": c, "endColumnIndex": c+1}, "cell": {"userEnteredFormat": red_format}, "fields": "userEnteredFormat.textFormat.foregroundColor,userEnteredFormat.textFormat.bold"}})

        ws.spreadsheet.batch_update({"requests": reqs})
        st.write("✅ 強化專案雲端同步完成 (未達100%自動標示紅字)")

# ----------------- [5. 交通事故] -----------------
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
            m['Pct'] = m.apply(lambda x: f"{(x['Diff']/x[c_name+'_lst']):.2%}" if x[c_name+'_lst'] != 0 else "0.00%", axis=1)
            res = m[['Station_Short', c_name+'_wk', c_name+'_prev', c_name+'_cur', c_name+'_lst', 'Diff', 'Pct']]
            res.columns = ['統計期間', f'本期({labels["wk"]})', f'前期({labels["prev"]})', f'本年累計({labels["cur"]})', f'去年累計({labels["lst"]})', '本年與去年同期比較', '增減比例']
        else:
            res = m[['Station_Short', c_name+'_wk', c_name+'_cur', c_name+'_lst', 'Diff']]
            res.columns = ['統計期間', f'本期({labels["wk"]})', f'本年累計({labels["cur"]})', f'去年累計({labels["lst"]})', '本年與去年同期比較']
        return res

    a1_res, a2_res = bld_tbl('A1_Deaths'), bld_tbl('A2_Injuries', True)
    
    c1, c2 = st.columns(2)
    c1.write("📊 **A1 死亡人數統計**"); c1.dataframe(a1_res, hide_index=True)
    c2.write("📊 **A2 受傷人數統計**"); c2.dataframe(a2_res, hide_index=True)

    if GCP_CREDS:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        RED_FMT = {"textFormat": {"foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}}}
        BLACK_FMT = {"textFormat": {"foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}}}

        for ws_idx, df in zip([2, 3], [a1_res, a2_res]):
            ws = sh.get_worksheet(ws_idx)
            ws.batch_clear(["A2:G20"]) 
            
            reqs = []
            for c_idx, c_name in enumerate(df.columns):
                reqs.append(get_gsheet_rich_text_req(ws.id, 1, c_idx, c_name))
            sh.batch_update({"requests": reqs})
            
            data_rows = [[int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in row] for row in df.values.tolist()]
            ws.update(range_name='A3', values=data_rows)
            
            diff_col = 4 if ws_idx == 2 else 5
            color_reqs = []
            for r_idx, row_vals in enumerate(df.values):
                val = row_vals[diff_col]
                target_r = 2 + r_idx 
                fmt = RED_FMT if isinstance(val, (int, float)) and val > 0 else BLACK_FMT
                color_reqs.append({"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": target_r, "endRowIndex": target_r + 1, "startColumnIndex": diff_col, "endColumnIndex": diff_col + 1}, "cell": {"userEnteredFormat": fmt}, "fields": "userEnteredFormat.textFormat.foregroundColor"}})
            sh.batch_update({"requests": color_reqs})
            
        st.write("✅ 交通事故雲端已更新")

# ----------------- [6. 靜桃計畫] -----------------
def process_jing_tao(files):
    """
    靜桃計畫大執法專案統計
    修正重點：為雙層表頭(MultiIndex)實作雙色(黑紅)格式化寫入
    """
    df = None

    # 1. 尋找與讀取目標檔案
    for f in files:
        f.seek(0)
        is_excel_file = f.name.lower().endswith(('.xlsx', '.xls'))

        if is_excel_file:
            try:
                xls = pd.ExcelFile(f)
                target_sheet = next((s for s in xls.sheet_names if '靜桃' in s), None)
                if not target_sheet and len(xls.sheet_names) > 1:
                    target_sheet = xls.sheet_names[1]
                elif not target_sheet:
                    target_sheet = xls.sheet_names[0]

                df_temp = pd.read_excel(xls, sheet_name=target_sheet, header=None, nrows=50)
                for idx, row in df_temp.iterrows():
                    row_str = " ".join([str(x) for x in row if pd.notna(x)])
                    if '通報日期' in row_str:
                        f.seek(0)
                        df = pd.read_excel(xls, sheet_name=target_sheet, skiprows=idx)
                        break
                if df is not None:
                    break
            except Exception:
                pass

        if df is None:
            f.seek(0)
            raw_bytes = f.read()
            for enc in ['utf-8-sig', 'utf-8', 'cp950', 'big5']:
                try:
                    text = raw_bytes.decode(enc, errors='ignore')
                    lines = text.splitlines()
                    for idx, line in enumerate(lines[:50]):
                        if '通報日期' in line:
                            f.seek(0)
                            df = pd.read_csv(f, encoding=enc, skiprows=idx, engine='python', on_bad_lines='skip')
                            break
                    if df is not None: break
                except:
                    continue
        if df is not None: break

    if df is None:
        st.error("❌ 找不到包含『通報日期』欄位的清冊檔案！請確認上傳了正確的檔案。")
        return

    # 清理欄位名稱
    df.columns = [str(c).strip().replace('\u3000', '').replace('\n', '') for c in df.columns]

    # ── 動態綁定關鍵欄位 ──
    date_col  = next((c for c in df.columns if '通報日期' in c), None)
    unit_col  = next((c for c in df.columns if '所別' in c or ('單位' in c and '舉發單位' not in c)), None)

    col_22 = next((c for c in df.columns if re.search(r'22.{0,3}0?6|夜間|深夜', c)), None)
    col_06 = next((c for c in df.columns if re.search(r'0?6.{0,3}22|日間|白天', c)), None)

    if not date_col or not unit_col:
        st.error(f"❌ 缺乏關鍵欄位！目前偵測到的欄位如下：{list(df.columns)}")
        return

    if not col_22 and not col_06:
        st.warning("⚠️ 在該工作表中找不到「22-06 / 06-22」時段欄位，日夜間數據將顯示為 0，但系統已為您自動計算「總計」筆數！")

    def parse_roc_date(val):
        if pd.isna(val): return pd.NaT
        s = str(val).strip().split(' ')[0]
        parts = re.split(r'[/\-]', s)
        if len(parts) == 3:
            try: return pd.Timestamp(year=int(parts[0]) + 1911, month=int(parts[1]), day=int(parts[2]))
            except: pass
        return pd.NaT

    df['_date'] = df[date_col].apply(parse_roc_date)

    # 本期區間：以昨日為終點，往前7天
    today    = datetime.now()
    end_dt   = today - timedelta(days=1)
    start_dt = end_dt - timedelta(days=6)
    period_str = f"{start_dt.strftime('%m%d')}-{end_dt.strftime('%m%d')}"

    df_period = df[(df['_date'] >= start_dt) & (df['_date'] <= end_dt)]

    # 累計區間：抓取最早日期，並強制以本期終點(end_dt)作為最晚日期
    valid_dates = df['_date'].dropna()
    if not valid_dates.empty:
        min_dt = valid_dates.min()
        cumu_str = f"({min_dt.year - 1911}{min_dt.strftime('%m%d')}-{end_dt.year - 1911}{end_dt.strftime('%m%d')})"
    else:
        cumu_str = ""

    def count_v(data, col):
        if col is None or col not in data.columns: return 0
        return data[col].astype(str).str.strip().str.upper().str.contains(r'^V$', regex=True, na=False).sum()

    stations      = ['聖亭', '龍潭', '中興', '石門', '高平', '三和', '警備', '交通']
    station_names = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

    results = []
    t_p_22 = t_p_06 = t_a_22 = t_a_06 = t_total = 0

    for kw, name in zip(stations, station_names):
        mask_all    = df[unit_col].astype(str).str.contains(kw, na=False)
        mask_period = df_period[unit_col].astype(str).str.contains(kw, na=False)

        p_22 = count_v(df_period[mask_period], col_22)
        p_06 = count_v(df_period[mask_period], col_06)
        a_22 = count_v(df[mask_all], col_22)
        a_06 = count_v(df[mask_all], col_06)
        
        if not col_22 and not col_06:
            total = len(df[mask_all])
        else:
            total = a_22 + a_06

        results.append([name, p_22, p_06, a_22, a_06, total])
        t_p_22 += p_22; t_p_06 += p_06
        t_a_22 += a_22; t_a_06 += a_06
        t_total += total

    results.insert(0, ['合計', t_p_22, t_p_06, t_a_22, t_a_06, t_total])

    col_22_label = col_22 if col_22 else "22-6時"
    col_06_label = col_06 if col_06 else "6-22時"

    # 建立雙層表頭 (MultiIndex)
    header_1 = ['統計期間', f'本期({period_str})', f'本期({period_str})', f'累計{cumu_str}', f'累計{cumu_str}', '總計']
    header_2 = ['', col_22_label, col_06_label, col_22_label, col_06_label, '']
    
    df_res = pd.DataFrame(results, columns=pd.MultiIndex.from_arrays([header_1, header_2]))

    st.write("📊 **「靜桃計畫」大執法專案統計表：**")
    st.dataframe(df_res, use_container_width=True)

    if GCP_CREDS:
        try:
            gc = gspread.service_account_from_dict(GCP_CREDS)
            sh = gc.open_by_url(GOOGLE_SHEET_URL)
            ws_name = "靜桃計畫"
            existing = [s.title for s in sh.worksheets()]
            ws = sh.worksheet(ws_name) if ws_name in existing else sh.add_worksheet(title=ws_name, rows="30", cols="10")

            ws.clear()
            
            # 將雙層表頭拆解為兩列文字準備寫入 Google Sheets
            titles = df_res.columns.tolist()
            top_row = [t[0] for t in titles]
            bottom_row = [t[1] for t in titles]
            data_body = df_res.values.tolist()
            
            ws.update(range_name='A1', values=[['「靜桃計畫」大執法專案統計表']])
            ws.update(range_name='A2', values=[top_row, bottom_row] + data_body)

            # 加入 Google Sheets 自動合併儲存格與置中的排版設定
            reqs = [
                # 標題大字跨欄置中 (A1:F1)
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 6}, "mergeType": "MERGE_ALL"}},
                # 統計期間跨列置中 (A2:A3)
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 0, "endColumnIndex": 1}, "mergeType": "MERGE_ALL"}},
                # 本期跨欄置中 (B2:C2)
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 1, "endColumnIndex": 3}, "mergeType": "MERGE_ALL"}},
                # 累計跨欄置中 (D2:E2)
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 3, "endColumnIndex": 5}, "mergeType": "MERGE_ALL"}},
                # 總計跨列置中 (F2:F3)
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 5, "endColumnIndex": 6}, "mergeType": "MERGE_ALL"}},
                # 將所有表頭 (第一、二、三列) 設定為粗體且垂直/水平置中
                {"repeatCell": {
                    "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 3, "startColumnIndex": 0, "endColumnIndex": 6},
                    "cell": {"userEnteredFormat": {"textFormat": {"bold": True}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}},
                    "fields": "userEnteredFormat.textFormat.bold,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment"
                }}
            ]

            # --- 新增：動態分色格式化 (Rich Text Format) ---
            black_color = {"red": 0.0, "green": 0.0, "blue": 0.0}
            red_color = {"red": 1.0, "green": 0.0, "blue": 0.0}
            
            for i, text in enumerate(top_row):
                if "(" in text:
                    p_start = text.find("(")
                    reqs.append({
                        "updateCells": {
                            "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": i, "endColumnIndex": i+1},
                            "rows": [{ "values": [{ "textFormatRuns": [
                                {"startIndex": 0, "format": {"foregroundColor": black_color, "bold": True}},
                                {"startIndex": p_start, "format": {"foregroundColor": red_color, "bold": True}}
                            ], "userEnteredValue": {"stringValue": text} }] }],
                            "fields": "userEnteredValue,textFormatRuns"
                        }
                    })

            sh.batch_update({"requests": reqs})
            st.write("✅ 靜桃計畫數據已同步至 Google Sheets（並完成自動排版與雙色處理）")
        except Exception as e:
            st.error(f"雲端同步出錯：{e}")
            st.write(traceback.format_exc())


# ==========================================
# 4. 戰情室首頁與排程器 (自動分流架構)
# ==========================================
with st.sidebar:
    st.title("🚓 龍潭分局戰情室")
    app_mode = st.selectbox("功能模組", ["🏠 智慧批次處理中心", "📂 PDF 轉 PPTX 工具"])

if app_mode == "🏠 智慧批次處理中心":
    st.header("📈 交通數據全自動批次處理中心")
    st.info("💡 請將所需報表全選後，直接拖曳至下方區域即可自動分流處理。")

    uploads = st.file_uploader("📂 拖入所有報表檔案", type=["xlsx", "csv", "xls"], accept_multiple_files=True)
    st.divider()
    st.subheader("🚀 啟動全自動批次作業")

    if uploads:
        file_hash = sum([f.size for f in uploads]) + len(uploads)
        if st.session_state.get("last_processed_hash") == file_hash:
            st.success("✅ 目前上傳的檔案皆已全自動處理完畢！")
            st.info("💡 若要處理新報表，請重新整理頁面或拖入新檔案。")
        else:
            cat_files = {"科技執法": [], "重大違規": [], "超載統計": [], "強化專案": [], "交通事故": [], "靜桃計畫": []}

            for f in uploads:
                name = f.name.lower()
                if any(k in name for k in ["list", "地點", "科技"]):
                    cat_files["科技執法"].append(f)
                elif any(k in name for k in ["stone", "超載"]):
                    cat_files["超載統計"].append(f)
                elif any(k in name for k in ["重大", "重點"]):
                    cat_files["重大違規"].append(f)
                elif any(k in name for k in ["強化", "專案", "砂石", "大貨", "r17", "法條", "自選匯出"]):
                    cat_files["強化專案"].append(f)
                elif any(k in name for k in ["a1", "a2", "事故", "案件統計"]):
                    cat_files["交通事故"].append(f)
                elif any(k in name for k in ["靜桃", "噪音", "改裝車", "總表", "詳細資料"]):
                    cat_files["靜桃計畫"].append(f)

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

                if cat_files["靜桃計畫"]:
                    with st.status(f"🤫 自動處理【靜桃計畫】({len(cat_files['靜桃計畫'])} 份)...", expanded=True) as status:
                        process_jing_tao(cat_files["靜桃計畫"])
                        status.update(label="✅ 靜桃計畫處理完成！", state="complete")

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
