import streamlit as st
import pandas as pd
import re
import io
import gspread
from gspread_formatting import *

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
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

# --- 2. 雲端同步功能 (新增總標題與 A-J 合併) ---
def sync_to_specified_sheet(df):
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        ws = sh.get_worksheet(0)
        
        # 1. 準備資料列 (插入總標題)
        col_tuples = df.columns.tolist()
        top_row = [t[0] for t in col_tuples]
        bottom_row = [t[1] for t in col_tuples]
        
        # 第一列為總標題，其餘 9 欄填空
        main_title_row = ["取締重大交通違規統計表"] + [""] * 9
        data_list = [main_title_row, top_row, bottom_row] + df.values.tolist()
        
        # 2. 寫入基礎數據
        ws.clear()
        ws.update(range_name='A1', values=data_list)
        
        data_rows_count = len(data_list)
        footnote_row_idx = data_rows_count - 1
        red_color = {"red": 1.0, "green": 0.0, "blue": 0.0}
        black_color = {"red": 0.0, "green": 0.0, "blue": 0.0}
        
        # 3. 格式化請求
        requests = [
            # 解除原本所有合併，避免衝突
            {"unmergeCells": {"range": {"sheetId": ws.id}}},
            
            # 【新功能】合併第 1 列 A-J (Index 0)
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}},
            
            # 統計期間標題合併 (Row Index 1 & 2)
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 0, "endColumnIndex": 1}, "mergeType": "MERGE_ALL"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 1, "endColumnIndex": 3}, "mergeType": "MERGE_ALL"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 3, "endColumnIndex": 5}, "mergeType": "MERGE_ALL"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 5, "endColumnIndex": 7}, "mergeType": "MERGE_ALL"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 7, "endColumnIndex": 8}, "mergeType": "MERGE_ALL"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 8, "endColumnIndex": 9}, "mergeType": "MERGE_ALL"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 9, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}},
            
            # 備註列合併 (最後一列)
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": footnote_row_idx, "endRowIndex": footnote_row_idx + 1, "startColumnIndex": 0, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}},
            
            # 總標題文字格式 (置中、加粗、字體 16)
            {
                "repeatCell": {
                    "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1},
                    "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER", "textFormat": {"bold": True, "fontSize": 16}}},
                    "fields": "userEnteredFormat(horizontalAlignment,textFormat)"
                }
            }
        ]
        
        # 4. 標題雙色邏輯 (現在標題在 Row Index 1)
        for i, text in enumerate(top_row):
            if "(" in text:
                paren_start = text.find("(")
                requests.append({
                    "updateCells": {
                        "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": i, "endColumnIndex": i+1},
                        "rows": [{
                            "values": [{
                                "textFormatRuns": [
                                    {"startIndex": 0, "format": {"foregroundColor": black_color}},
                                    {"startIndex": paren_start, "format": {"foregroundColor": red_color}}
                                ],
                                "userEnteredValue": {"stringValue": text}
                            }]
                        }],
                        "fields": "userEnteredValue,textFormatRuns"
                    }
                })

        # 5. 負值紅字規則 (資料區間從 Row Index 3 開始)
        requests.extend([
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": ws.id, "startRowIndex": 3, "endRowIndex": footnote_row_idx, "startColumnIndex": 7, "endColumnIndex": 8}],
                        "booleanRule": {
                            "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                            "format": {"textFormat": {"foregroundColor": red_color}}
                        }
                    }, "index": 0
                }
            }
        ])
        
        sh.batch_update({"requests": requests})
        return True
    except Exception as e:
        st.error(f"雲端同步失敗: {e}")
        return False

# --- 4. 解析邏輯 (維持先前版本) ---
def parse_excel_with_date_extraction(uploaded_file, sheet_keyword, col_indices):
    try:
        content = uploaded_file.getvalue()
        xl = pd.ExcelFile(io.BytesIO(content))
        target_sheet = next((s for s in xl.sheet_names if sheet_keyword in s), xl.sheet_names[0])
        df = pd.read_excel(xl, sheet_name=target_sheet, header=None)
        date_display = ""
        try:
            row_content = "".join(df.iloc[2].astype(str))
            match = re.search(r'(\d{7})([至\-~])(\d{7})', row_content)
            if match:
                date_display = f"{match.group(1)[3:]}-{match.group(3)[3:]}"
        except:
            date_display = ""
        unit_data = {}
        for _, row in df.iterrows():
            u = get_standard_unit(row.iloc[0])
            if u and "合計" not in str(row.iloc[0]):
                def clean(v):
                    try: return int(float(str(v).replace(',', '').strip())) if str(v).strip() not in ['', 'nan', 'None', '-'] else 0
                    except: return 0
                stop_val = 0 if u == '科技執法' else clean(row.iloc[col_indices[0]])
                cit_val = clean(row.iloc[col_indices[1]])
                if u not in unit_data: unit_data[u] = {'stop': stop_val, 'cit': cit_val}
                else: 
                    unit_data[u]['stop'] += stop_val
                    unit_data[u]['cit'] += cit_val
        return unit_data, date_display
    except: return None, ""

# --- 5. 主介面 ---
st.title("🚔 交通統計自動化系統")

file_period = st.file_uploader("📂 上傳「本期」檔案", type=['xlsx'])
file_year = st.file_uploader("📂 上傳「累計」檔案", type=['xlsx'])

if file_period and file_year:
    d_week, date_w = parse_excel_with_date_extraction(file_period, "重點違規統計表", [15, 16])
    d_year, date_y = parse_excel_with_date_extraction(file_year, "(1)", [15, 16])
    d_last, _ = parse_excel_with_date_extraction(file_year, "(1)", [18, 19])
    
    if d_week and d_year:
        rows = []
        t = {k: 0 for k in ['ws', 'wc', 'ys', 'yc', 'ls', 'lc', 'diff', 'tgt']}
        for u in UNIT_ORDER:
            w, y, l = d_week.get(u, {'stop':0, 'cit':0}), d_year.get(u, {'stop':0, 'cit':0}), d_last.get(u, {'stop':0, 'cit':0})
            ys_sum, ls_sum = y['stop'] + y['cit'], l['stop'] + l['cit']
            tgt = TARGETS.get(u, 0)
            diff_display, rate_display = ("—", "—") if u == '警備隊' else (int(ys_sum - ls_sum), f"{(ys_sum/tgt):.1%}" if tgt > 0 else "0%")
            if u != '警備隊':
                t['diff'] += (ys_sum - ls_sum); t['tgt'] += tgt
            rows.append([u, w['stop'], w['cit'], y['stop'], y['cit'], l['stop'], l['cit'], diff_display, tgt, rate_display])
            t['ws']+=w['stop']; t['wc']+=w['cit']; t['ys']+=y['stop']; t['yc']+=y['cit']; t['ls']+=l['stop']; t['lc']+=l['cit']
        
        total_rate = f"{((t['ys']+t['yc'])/t['tgt']):.1%}" if t['tgt']>0 else "0%"
        rows.insert(0, ['合計', t['ws'], t['wc'], t['ys'], t['yc'], t['ls'], t['lc'], t['diff'], t['tgt'], total_rate])
        rows.append([FOOTNOTE_TEXT] + [""] * 9)
        
        label_week = f"本期({date_w})" if date_w else "本期"
        label_year = f"本年累計({date_y})" if date_y else "本年累計"
        label_last = f"去年累計({date_y})" if date_y else "去年累計" 
        
        df_final = pd.DataFrame(rows, columns=pd.MultiIndex.from_arrays([
            ['統計期間', label_week, label_week, label_year, label_year, label_last, label_last, '本年與去年同期比較', '目標值', '達成率'],
            ['取締方式', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '', '', '']
        ]))
        
        st.dataframe(df_final, use_container_width=True)

        if st.button("🚀 同步雲端", type="primary"):
            if sync_to_specified_sheet(df_final): 
                st.info("☁️ 雲端同步成功！已新增總標題列並合併 A-J 欄。")
