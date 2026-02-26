import streamlit as st
import pandas as pd
import re
import io
import gspread
from gspread_formatting import *

# ==========================================
# 0. 初始化
# ==========================================
st.set_page_config(page_title="交通統計系統", layout="wide")
st.title("🚔 交通統計自動化系統")

try:
    from gspread_formatting import *
    HAS_FORMATTING = True
except ImportError:
    HAS_FORMATTING = False

# ==========================================
# 1. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
TARGETS = {'聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479, '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006}

# 網頁顯示用文字
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

# ==========================================
# 2. 雲端同步邏輯 (首尾格式完全保護)
# ==========================================
def sync_to_specified_sheet(df):
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        ws = sh.get_worksheet(0)
        
        # 1. 準備數據 (僅包含兩層標題與所隊數據，不含最後一列腳註)
        col_tuples = df.columns.tolist()
        top_row = [t[0] for t in col_tuples]
        bottom_row = [t[1] for t in col_tuples]
        
        # 排除 df 的最後一列 (那是腳註文字)，只取純數據部分
        data_body = df.values.tolist()[:-1] 
        
        # 最終寫入清單：[標題1, 標題2, 數據1, 數據2, ...]
        data_list = [top_row, bottom_row] + data_body
        
        # 2. 【核心】從 A2 開始寫入，不使用 ws.clear()
        # 這會填滿第 2 列到倒數第 2 列，完全不碰第 1 列 (總標題) 與 最後一列 (腳註)
        ws.update(range_name='A2', values=data_list)
        
        # 3. 處理必要的顏色邏輯 (僅針對內容，不更動合併結構)
        if HAS_FORMATTING:
            data_rows_end_idx = len(data_list) + 1 # 數據區結束的 Row Index
            red_color = {"red": 1.0, "green": 0.0, "blue": 0.0}
            black_color = {"red": 0.0, "green": 0.0, "blue": 0.0}
            
            requests = []
            
            # 處理標題括號紅字 (Row Index 1)
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

            # 負值紅字規則 (資料區從 Row Index 3 開始，到腳註前一列結束)
            requests.append({
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": ws.id, "startRowIndex": 3, "endRowIndex": data_rows_end_idx, "startColumnIndex": 7, "endColumnIndex": 8}],
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

# ==========================================
# 3. 解析與主介面邏輯 (略，與前版相同)
# ==========================================
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
        except: date_display = ""
            
        unit_data = {}
        for _, row in df.iterrows():
            u = get_standard_unit(row.iloc[0])
            if u and "合計" not in str(row.iloc[0]):
                def clean(v):
                    try: return int(float(str(v).replace(',', '').strip()))
                    except: return 0
                unit_data[u] = {'stop': clean(row.iloc[col_indices[0]]), 'cit': clean(row.iloc[col_indices[1]])}
        return unit_data, date_display
    except: return None, ""

# --- UI 程式碼 ---
col1, col2 = st.columns(2)
with col1:
    file_period = st.file_uploader("📂 上傳「本期」檔案", type=['xlsx'])
with col2:
    file_year = st.file_uploader("📂 上傳「累計」檔案", type=['xlsx'])

if file_period and file_year:
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
        
        df_final = pd.DataFrame(rows, columns=pd.MultiIndex.from_arrays([
            ['統計期間', label_w, label_w, label_y, label_y, label_l, label_l, '本年與去年同期比較', '目標值', '達成率'],
            ['取締方式', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '', '', '']
        ]))
        
        st.dataframe(df_final, use_container_width=True)

        if st.button("🚀 同步數據", type="primary"):
            if sync_to_specified_sheet(df_final):
                st.success("✅ 同步完成！(首尾格式已鎖定，僅更新中間數據)")
