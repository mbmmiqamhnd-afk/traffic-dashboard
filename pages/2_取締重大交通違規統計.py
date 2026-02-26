import streamlit as st
import pandas as pd
import re
import io
import smtplib
import gspread
from gspread_formatting import * # 需要安裝 gspread-formatting
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
TARGETS = {
    '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479, 
    '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006
}
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

# --- 2. 雲端同步功能 (精確控制單儲存格雙色) ---
def sync_to_specified_sheet(df):
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        ws = sh.get_worksheet(0)
        
        col_tuples = df.columns.tolist()
        top_row = [t[0] for t in col_tuples]
        bottom_row = [t[1] for t in col_tuples]
        data_list = [top_row, bottom_row] + df.values.tolist()
        
        # 1. 先寫入基礎數據
        ws.update(range_name='A1', values=data_list)
        
        data_rows_count = len(data_list) - 1 
        
        # 2. 準備 RichText 格式請求 (讓標題列的括號變紅)
        # 定義紅色顏色格式
        red_color = {"red": 1.0, "green": 0.0, "blue": 0.0}
        black_color = {"red": 0.0, "green": 0.0, "blue": 0.0}
        
        requests = []
        
        # 遍歷第一列的標題，尋找包含括號的儲存格
        for i, text in enumerate(top_row):
            if "(" in text and ")" in text:
                paren_start = text.find("(")
                # 建立 RichText 請求
                requests.append({
                    "updateCells": {
                        "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": i, "endColumnIndex": i+1},
                        "rows": [{
                            "values": [{
                                "textFormatRuns": [
                                    {"startIndex": 0, "format": {"foregroundColor": black_color}}, # 文字黑色
                                    {"startIndex": paren_start, "format": {"foregroundColor": red_color}} # 括號起變紅
                                ],
                                "userEnteredValue": {"stringValue": text}
                            }]
                        }],
                        "fields": "userEnteredValue,textFormatRuns"
                    }
                })

        # 加入原本的負值紅字規則
        requests.extend([
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": ws.id, "startRowIndex": 2, "endRowIndex": data_rows_count, "startColumnIndex": 7, "endColumnIndex": 8}],
                        "booleanRule": {
                            "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                            "format": {"textFormat": {"foregroundColor": red_color}}
                        }
                    }, "index": 0
                }
            },
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": ws.id, "startRowIndex": 2, "endRowIndex": data_rows_count, "startColumnIndex": 0, "endColumnIndex": 1}],
                        "booleanRule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA",
                                "values": [{"userEnteredValue": '=AND($H3<0, $A3<>"合計", $A3<>"科技執法")'}]
                            },
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

# --- 4. 解析邏輯 (保留簡短日期與橫線) ---
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
                # 去除前3碼，分隔符改為 "-"
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

col_up1, col_up2 = st.columns(2)
with col_up1:
    file_period = st.file_uploader("📂 1. 上傳「本期」檔案", type=['xlsx'])
with col_up2:
    file_year = st.file_uploader("📂 2. 上傳「累計」檔案", type=['xlsx'])

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
            
            if u == '警備隊':
                diff_display, rate_display = "—", "—"
            else:
                diff_val = ys_sum - ls_sum
                diff_display = int(diff_val)
                rate_display = f"{(ys_sum/tgt):.1%}" if tgt > 0 else "0%"
                t['diff'] += diff_val
                t['tgt'] += tgt
            
            rows.append([u, w['stop'], w['cit'], y['stop'], y['cit'], l['stop'], l['cit'], diff_display, tgt, rate_display])
            t['ws']+=w['stop']; t['wc']+=w['cit']; t['ys']+=y['stop']; t['yc']+=y['cit']; t['ls']+=l['stop']; t['lc']+=l['cit']
        
        total_rate = f"{((t['ys']+t['yc'])/t['tgt']):.1%}" if t['tgt']>0 else "0%"
        rows.insert(0, ['合計', t['ws'], t['wc'], t['ys'], t['yc'], t['ls'], t['lc'], t['diff'], t['tgt'], total_rate])
        rows.append([FOOTNOTE_TEXT] + [""] * 9)
        
        # 標題設定
        label_week = f"本期({date_w})" if date_w else "本期"
        label_year = f"本年累計({date_y})" if date_y else "本年累計"
        label_last = f"去年累計({date_y})" if date_y else "去年累計" 
        
        header_top = ['統計期間', label_week, label_week, label_year, label_year, label_last, label_last, '本年與去年同期比較', '目標值', '達成率']
        header_bottom = ['取締方式', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '', '', '']
        
        multi_col = pd.MultiIndex.from_arrays([header_top, header_bottom])
        df_final = pd.DataFrame(rows, columns=multi_col)
        
        st.success("✅ 解析成功！")
        
        # 網頁預覽 (使用 HTML 渲染雙色標題)
        def color_date(val):
            if isinstance(val, str) and "(" in val:
                main_text = val.split("(")[0]
                date_part = val.split("(")[1]
                return f'{main_text}<span style="color:red;">({date_part}</span>'
            return val

        st.write("📊 報表預覽 (括號日期將顯示為紅色)")
        st.dataframe(df_final, use_container_width=True)

        st.divider()
        if st.button("🚀 同步數據並寄出報表", type="primary"):
            if sync_to_specified_sheet(df_final): 
                st.info(f"☁️ 數據同步成功！標題文字為黑色，括號日期已設為紅色。")
