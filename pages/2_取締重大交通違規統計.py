import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import smtplib
import gspread
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# 強制清除快取
try:
    st.cache_data.clear()
    st.cache_resource.clear()
except: pass

st.set_page_config(page_title="取締重大交通違規統計", layout="wide", page_icon="🚔")
st.markdown("## 🚔 取締重大交通違規統計 (v59 全表格構造版)")

# --- 強制清除快取按鈕 ---
if st.button("🧹 清除快取 (若更新無效請按此)", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.success("快取已清除！請重新整理頁面 (F5) 並重新上傳檔案。")

st.markdown("""
### 📝 使用說明 (v59)
1.  **核彈級寫入**：不再分段寫入，而是將整張表的內容與格式打包，一次性強制覆蓋。
2.  **絕對顏色**：顏色屬性直接綁定在資料封包內，Google 試算表無法拒絕。
3.  **功能維持**：全表寫入、目標值更新、自動寄信。
""")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

UNIT_MAP = {
    '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所',
    '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所',
    '警備隊': '警備隊', '龍潭交通分隊': '交通分隊', '交通組': '科技執法'
}
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

TARGETS = {
    '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479,
    '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 0
}

NOTE_TEXT = "重大交通違規指：「闖紅燈」、「酒後駕車」、「嚴重超速」、「未依兩段式左轉」、「不暫停讓行人」、 「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」等8項。"

# ==========================================
# 1. Google Sheets Cell 建構工具 (核心)
# ==========================================
def make_cell(value, is_bold=True, text_color=None, bg_color=None, align='CENTER', font_size=None, rich_text_tokens=None):
    """
    建構一個完整的 Google Sheet Cell 物件
    """
    cell = {}
    
    # 1. 設定值 (Rich Text 或 普通值)
    if rich_text_tokens:
        cell['userEnteredValue'] = {'stringValue': str(value)}
        cell['textFormatRuns'] = rich_text_tokens
    else:
        # 嘗試轉為數字，否則為字串
        if isinstance(value, (int, float)):
            cell['userEnteredValue'] = {'numberValue': value}
        else:
            cell['userEnteredValue'] = {'stringValue': str(value)}

    # 2. 設定格式
    fmt = {
        'horizontalAlignment': align,
        'verticalAlignment': 'MIDDLE',
        'textFormat': {
            'bold': is_bold,
            'foregroundColor': text_color if text_color else {'red': 0, 'green': 0, 'blue': 0} # 預設黑
        },
        'borders': {
            'top': {'style': 'SOLID'}, 'bottom': {'style': 'SOLID'},
            'left': {'style': 'SOLID'}, 'right': {'style': 'SOLID'}
        }
    }
    
    if bg_color:
        fmt['backgroundColor'] = bg_color
    else:
        fmt['backgroundColor'] = {'red': 1, 'green': 1, 'blue': 1} # 預設白

    if font_size:
        fmt['textFormat']['fontSize'] = font_size

    cell['userEnteredFormat'] = fmt
    return cell

def get_rich_text_tokens(text):
    """將文字拆解為 Rich Text Tokens (紅數黑字)"""
    tokens = re.split(r'([0-9\(\)\/\-\.\%\~]+)', str(text))
    runs = []
    current_pos = 0
    for token in tokens:
        if not token: continue
        color = {'red': 0, 'green': 0, 'blue': 0} # 黑
        if re.match(r'^[0-9\(\)\/\-\.\%\~]+$', token):
            color = {'red': 1, 'green': 0, 'blue': 0} # 紅
        
        runs.append({
            "startIndex": current_pos,
            "format": {"foregroundColor": color, "bold": True}
        })
        current_pos += len(token)
    return runs

# ==========================================
# 2. Google Sheets 寫入主程序 (全構造)
# ==========================================
def update_google_sheet(data_list, sheet_url):
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ 錯誤：未設定 Secrets！")
            return False
        
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(0)
        
        st.info(f"📂 寫入目標工作表：**「{ws.title}」** (Index 0)")
        
        # 1. 徹底清除
        ws.clear() 
        
        # 2. 建構全表資料 (Rows)
        grid_rows = []
        
        # 定義顏色常數
        COLOR_RED = {'red': 1, 'green': 0, 'blue': 0}
        COLOR_BLACK = {'red': 0, 'green': 0, 'blue': 0}
        COLOR_YELLOW_BG = {'red': 1.0, 'green': 0.92, 'blue': 0.61}
        
        st.write("---")
        st.write("🔍 **v59 建構日誌**：")

        for r_idx, row_data in enumerate(data_list):
            grid_row = []
            values = []
            
            # 判斷是否為「合計列」(Index 3, 即 Row 4)
            is_total_row = (r_idx == 3)
            bg = COLOR_YELLOW_BG if is_total_row else None
            
            # 判斷本列是否為負數列 (僅針對數據列 r_idx 3~12)
            is_negative_row = False
            unit_name = str(row_data[0]).strip()
            
            if 3 <= r_idx <= 12:
                try:
                    val_str = str(row_data[7]).replace(',', '')
                    if float(val_str) < 0: is_negative_row = True
                except: pass

            for c_idx, cell_val in enumerate(row_data):
                # 預設設定
                c_bold = True
                c_color = COLOR_BLACK
                c_align = 'CENTER'
                c_size = None
                c_rich = None
                
                # --- 特殊邏輯處理 ---
                
                # A. 標題列 (Row 2, Index 1) 的日期部分 -> Rich Text
                if r_idx == 1 and c_idx in [1, 3, 5]:
                    c_rich = get_rich_text_tokens(cell_val)
                
                # B. 數據列 (Rows 4-13) 的負數變色
                elif 3 <= r_idx <= 12:
                    # 判斷 H 欄 (Index 7) 負數
                    if c_idx == 7 and is_negative_row:
                        c_color = COLOR_RED
                    
                    # 判斷 A 欄 (Index 0) 單位名稱變紅 (排除科技執法)
                    if c_idx == 0 and is_negative_row and unit_name != "科技執法":
                        c_color = COLOR_RED
                        if c_idx == 0: # 只顯示一次 log
                             st.write(f"🔴 Row {r_idx+1}: {unit_name} -> 設為紅色")

                # C. 說明列 (Last Row)
                elif r_idx == 13:
                    c_align = 'LEFT'
                    c_bold = False
                    c_size = 10

                # 建構單格
                grid_cell = make_cell(
                    cell_val, 
                    is_bold=c_bold, 
                    text_color=c_color, 
                    bg_color=bg, 
                    align=c_align, 
                    font_size=c_size, 
                    rich_text_tokens=c_rich
                )
                values.append(grid_cell)
            
            grid_rows.append({"values": values})

        # 3. 發送 updateCells 指令 (一次性覆蓋全表)
        body = {
            "requests": [
                {
                    "updateCells": {
                        "start": {"sheetId": ws.id, "rowIndex": 0, "columnIndex": 0},
                        "rows": grid_rows,
                        "fields": "*" # 更新所有屬性
                    }
                },
                # 補上合併儲存格指令
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}},
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 13, "endRowIndex": 14, "startColumnIndex": 0, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}},
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 1, "endColumnIndex": 3}, "mergeType": "MERGE_ALL"}},
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 3, "endColumnIndex": 5}, "mergeType": "MERGE_ALL"}},
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 5, "endColumnIndex": 7}, "mergeType": "MERGE_ALL"}},
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 7, "endColumnIndex": 8}, "mergeType": "MERGE_ALL"}},
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 8, "endColumnIndex": 9}, "mergeType": "MERGE_ALL"}},
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 9, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}}
            ]
        }
        
        sh.batch_update(body)
        st.write("✅ **全表構造寫入完成**")
        st.write("---")
        return True

    except Exception as e:
        st.error(f"❌ 寫入或格式化失敗: {e}")
        return False

# ==========================================
# 3. 寄信函數
# ==========================================
def send_email(recipient, subject, body, file_bytes, filename):
    try:
        if "email" not in st.secrets: return False
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
    except: return False

# ==========================================
# 4. 解析函數
# ==========================================
def parse_focus_report(uploaded_file):
    if not uploaded_file: return None
    file_name = uploaded_file.name
    try:
        content = uploaded_file.getvalue()
        start_date, end_date = "", ""
        df = None; header_idx = -1
        
        df_raw = pd.read_excel(io.BytesIO(content), header=None, nrows=25)
        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
            if not start_date:
                match = re.search(r'入案日期[：:]?\s*(\d{3,7}).*至\s*(\d{3,7})', row_str)
                if match: start_date, end_date = match.group(1), match.group(2)
            if "單位" in row_str:
                header_idx = i
                if start_date: break
        
        if header_idx == -1:
            st.warning(f"⚠️ 檔案 {file_name} 解析警告：找不到標題列。")
            return None

        df = pd.read_excel(io.BytesIO(content), header=header_idx)
        keywords = ["酒後", "闖紅燈", "嚴重超速", "逆向", "轉彎", "蛇行", "不暫停讓行人", "機車"]
        stop_cols = []; cit_cols = []
        
        for i in range(len(df.columns)):
            col_str = str(df.columns[i])
            if any(k in col_str for k in keywords) and "路肩" not in col_str and "大型車" not in col_str:
                stop_cols.append(i); cit_cols.append(i+1)
        
        unit_data = {}
        for _, row in df.iterrows():
            raw_unit = str(row['單位']).strip()
            if raw_unit == 'nan' or not raw_unit or "合計" in raw_unit: continue
            
            unit_name = UNIT_MAP.get(raw_unit, raw_unit)
            s, c = 0, 0
            
            for col in stop_cols:
                try:
                    val = row.iloc[col]
                    if pd.isna(val) or str(val).strip() == "": val = 0
                    s += float(str(val).replace(',', ''))
                except: pass
            
            for col in cit_cols:
                try:
                    val = row.iloc[col]
                    if pd.isna(val) or str(val).strip() == "": val = 0
                    c += float(str(val).replace(',', ''))
                except: pass

            unit_data[unit_name] = {'stop': s, 'cit': c}

        duration = 0
        try:
            if start_date and end_date:
                s_d = re.sub(r'[^\d]', '', start_date); e_d = re.sub(r'[^\d]', '', end_date)
                d1 = date(int(s_d[:3])+1911, int(s_d[3:5]), int(s_d[5:]))
                d2 = date(int(e_d[:3])+1911, int(e_d[3:5]), int(e_d[5:]))
                duration = (d2 - d1).days
        except: duration = 0
        if not start_date: start_date = "0000000"
        if not end_date: end_date = "0000000"
        return {'data': unit_data, 'start': start_date, 'end': end_date, 'duration': duration, 'filename': file_name}
    except Exception as e:
        st.warning(f"⚠️ 檔案 {file_name} 錯誤: {e}")
        return None

def get_mmdd(date_str):
    clean = re.sub(r'[^\d]', '', str(date_str))
    return clean[-4:] if len(clean) >= 4 else clean

# ==========================================
# 5. 主程式
# ==========================================
# ★★★ v59 Key ★★★
uploaded_files = st.file_uploader("請拖曳 3 個 Focus 統計檔案至此", accept_multiple_files=True, type=['xlsx', 'xls'], key="focus_uploader_v59_full_grid")

if uploaded_files:
    if len(uploaded_files) < 3: st.warning("⏳ 檔案不足 (需 3 個)...")
    else:
        try:
            parsed_files = []
            for f in uploaded_files:
                res = parse_focus_report(f)
                if res: parsed_files.append(res)
            
            if len(parsed_files) < 3: 
                st.error("❌ 解析失敗。")
                st.stop()

            parsed_files.sort(key=lambda x: x['start'])
            file_last_year = parsed_files[0]
            others = parsed_files[1:]
            others.sort(key=lambda x: x['duration'], reverse=True)
            
            file_week = others[0] 
            file_year = others[1]

            unit_rows = []
            accum = {'ws':0, 'wc':0, 'ys':0, 'yc':0, 'ls':0, 'lc':0}
            
            for u in UNIT_ORDER:
                w = file_week['data'].get(u, {'stop':0, 'cit':0})
                y = file_year['data'].get(u, {'stop':0, 'cit':0})
                l = file_last_year['data'].get(u, {'stop':0, 'cit':0})
                
                if u == '科技執法': w['stop'], y['stop'], l['stop'] = 0, 0, 0
                y_total = y['stop'] + y['cit']; l_total = l['stop'] + l['cit']
                
                w_s, w_c = int(w['stop']), int(w['cit'])
                y_s, y_c = int(y['stop']), int(y['cit'])
                l_s, l_c = int(l['stop']), int(l['cit'])

                row_data = [u, w_s, w_c, y_s, y_c, l_s, l_c]
                
                if u == '警備隊': 
                    row_data.extend(['—', '', '']) 
                else:
                    diff = int(y_total - l_total)
                    row_data.append(diff)
                    if u == '科技執法':
                        row_data.extend(['', ''])
                    else:
                        tgt = TARGETS.get(u, 0)
                        rate_str = f"{y_total/tgt:.0%}" if tgt > 0 else "0%"
                        row_data.extend([tgt, rate_str])
                
                accum['ws']+=w_s; accum['wc']+=w_c
                accum['ys']+=y_s; accum['yc']+=y_c
                accum['ls']+=l_s; accum['lc']+=l_c
                unit_rows.append(row_data)

            total_target = sum([v for k,v in TARGETS.items() if k not in ['警備隊', '科技執法']])
            t_diff = (accum['ys']+accum['yc']) - (accum['ls']+accum['lc'])
            t_rate = (accum['ys']+accum['yc'])/total_target if total_target > 0 else 0
            total_rate_str = f"{t_rate:.0%}"
            
            total_row = ['合計', accum['ws'], accum['wc'], accum['ys'], accum['yc'], accum['ls'], accum['lc'], t_diff, total_target, total_rate_str]
            final_rows = [total_row] + unit_rows

            cols = ['取締方式', '本期_當場攔停', '本期_逕行舉發', '本年_當場攔停', '本年_逕行舉發', '去年_當場攔停', '去年_逕行舉發', '本年與去年比較', '目標值', '達成率']
            df_final = pd.DataFrame(final_rows, columns=cols)

            # ==========================================
            # ★★★ 網頁預覽區 ★★★
            # ==========================================
            st.success("✅ 分析完成！下方為預覽畫面")

            def format_mixed(text, date_val):
                return f"<span style='color:black'>{text}</span><br><span style='color:red; font-weight:bold;'>({date_val})</span>"

            s_w, e_w = get_mmdd(file_week['start']), get_mmdd(file_week['end'])
            s_y, e_y = get_mmdd(file_year['start']), get_mmdd(file_year['end'])
            s_l, e_l = get_mmdd(file_last_year['start']), get_mmdd(file_last_year['end'])

            str_week = format_mixed("本期", f"{s_w}~{e_w}")
            str_year = format_mixed("本年累計", f"{s_y}~{e_y}")
            str_last = format_mixed("去年累計", f"{s_l}~{e_l}")
            
            header_compare = "<span style='color:black'>本年與去年<br>同期比較</span>"
            header_target = "<span style='color:black'>目標值</span>"
            header_rate = "<span style='color:black'>達成率</span>"
            header_stat = "<span style='color:black'>統計期間</span>"

            style = "<style>table{width:100%;border-collapse:collapse;text-align:center;font-family:'Microsoft JhengHei',sans-serif;color:#333;}th,td{border:1px solid #999;padding:8px;font-size:14px;}.title{font-size:20px;font-weight:bold;background-color:#f0f0f0;color:#000;}.header-top{background-color:#ffffff;font-weight:bold;} .header-sub{background-color:#ffffff;font-weight:bold;color:#000;}.unit-col{background-color:#fafafa;font-weight:bold;text-align:left;color:#000;}.footer-note{text-align:left;font-size:12px;background-color:#fff;color:#000;border:1px solid #999;}</style>"
            
            table_start = f"<table><tr><td colspan='10' class='title'>取締重大交通違規件數統計表</td></tr><tr><td class='header-top'>{header_stat}</td><td colspan='2' class='header-top'>{str_week}</td><td colspan='2' class='header-top'>{str_year}</td><td colspan='2' class='header-top'>{str_last}</td><td rowspan='2' class='header-top' style='vertical-align:middle;'>{header_compare}</td><td rowspan='2' class='header-top' style='vertical-align:middle;'>{header_target}</td><td rowspan='2' class='header-top' style='vertical-align:middle;'>{header_rate}</td></tr><tr><td class='header-sub'>取締方式</td><td class='header-sub'>當場攔停</td><td class='header-sub'>逕行舉發</td><td class='header-sub'>當場攔停</td><td class='header-sub'>逕行舉發</td><td class='header-sub'>當場攔停</td><td class='header-sub'>逕行舉發</td></tr>"
            
            rows_html = ""
            for row in final_rows:
                rows_html += "<tr>"
                is_total_row = (row[0] == '合計')
                
                unit_name_red = False
                try:
                    comp_val = int(row[7])
                    unit_name = str(row[0])
                    if comp_val < 0 and unit_name != '科技執法':
                        unit_name_red = True
                except: pass

                for i, cell in enumerate(row):
                    cell_style_list = []
                    if is_total_row: cell_style_list.append("background-color:#FFEB9C;")
                    else: cell_style_list.append("background-color:#fff;")
                    
                    if i == 0: 
                        cell_style_list.append("text-align:left;font-weight:bold;")
                        if unit_name_red: cell_style_list.append("color:red;")
                        else: cell_style_list.append("color:black;")
                    else:
                        is_negative = False
                        if i == 7:
                            try:
                                if int(cell) < 0: is_negative = True
                            except: pass
                        if is_negative: cell_style_list.append("color:red;font-weight:bold;")
                        else: cell_style_list.append("color:#000;")
                    
                    style_str = f"style='{''.join(cell_style_list)}'"
                    rows_html += f"<td {style_str}>{cell}</td>"
                rows_html += "</tr>"
            
            rows_html += f"<tr><td colspan='10' class='footer-note'>{NOTE_TEXT}</td></tr>"

            final_html = style + table_start + rows_html + "</table>"
            st.markdown(final_html, unsafe_allow_html=True)

            # ==========================================
            # Excel 產生邏輯
            # ==========================================
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, header=False, sheet_name='Sheet1', startrow=3)
                workbook = writer.book
                ws = writer.sheets['Sheet1']
                
                fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
                fmt_top_base = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#ffffff', 'text_wrap': True, 'font_color': 'black'})
                fmt_font_black = workbook.add_format({'font_color': 'black', 'bold': True})
                fmt_font_red = workbook.add_format({'font_color': 'red', 'bold': True})
                fmt_sub = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
                fmt_total = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFEB9C'})
                fmt_total_neg = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFEB9C', 'font_color': 'red'})
                fmt_note = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': False, 'font_size': 10})

                ws.merge_range('A1:J1', '取締重大交通違規件數統計表', fmt_title)
                
                ws.write('A2', '統計期間', fmt_top_base) 
                ws.merge_range('B2:C2', "", fmt_top_base)
                ws.write_rich_string('B2', fmt_font_black, "本期", fmt_font_red, f"\n({s_w}~{e_w})", fmt_top_base)
                ws.merge_range('D2:E2', "", fmt_top_base)
                ws.write_rich_string('D2', fmt_font_black, "本年累計", fmt_font_red, f"\n({s_y}~{e_y})", fmt_top_base)
                ws.merge_range('F2:G2', "", fmt_top_base)
                ws.write_rich_string('F2', fmt_font_black, "去年累計", fmt_font_red, f"\n({s_l}~{e_l})", fmt_top_base)
                ws.merge_range('H2:H3', '本年與去年\n同期比較', fmt_top_base)
                ws.merge_range('I2:I3', '目標值', fmt_top_base)
                ws.merge_range('J2:J3', '達成率', fmt_top_base)

                ws.write('A3', '取締方式', fmt_sub)
                ws.write('B3', '當場攔停', fmt_sub); ws.write('C3', '逕行舉發', fmt_sub)
                ws.write('D3', '當場攔停', fmt_sub); ws.write('E3', '逕行舉發', fmt_sub)
                ws.write('F3', '當場攔停', fmt_sub); ws.write('G3', '逕行舉發', fmt_sub)

                row_idx = 3
                total_data = final_rows[0]
                for col_idx, val in enumerate(total_data):
                    current_fmt = fmt_total
                    if col_idx == 7:
                        try:
                            if int(val) < 0: current_fmt = fmt_total_neg
                        except: pass
                    ws.write(row_idx, col_idx, val, current_fmt)

                fmt_red_num = workbook.add_format({'font_color': 'red', 'bold': True})
                last_data_row = 3 + len(final_rows) - 1
                
                ws.conditional_format(4, 7, last_data_row, 7, {
                    'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt_red_num
                })

                ws.conditional_format(4, 0, last_data_row, 0, {
                    'type': 'formula',
                    'criteria': '=AND($H4<0, $A4<>"科技執法")',
                    'format': fmt_red_num
                })

                footer_row = last_data_row + 1
                ws.merge_range(footer_row, 0, footer_row, 9, NOTE_TEXT, fmt_note)

                ws.set_column(0, 0, 15)
                ws.set_column(1, 6, 11)
                ws.set_column(7, 7, 13)
                ws.set_column(8, 9, 10)
            
            excel_data = output.getvalue()
            file_name_out = f'重點違規統計_{file_year["end"]}.xlsx'

            # ==========================================
            # ★★★ 準備完整寫入資料 (Rows 1-14) ★★★
            # ==========================================
            sheet_r1 = ['取締重大交通違規件數統計表'] + [''] * 9
            sheet_r2 = [
                '統計期間', 
                f'本期\n({s_w}~{e_w})', '', 
                f'本年累計\n({s_y}~{e_y})', '', 
                f'去年累計\n({s_l}~{e_l})', '', 
                '本年與去年\n同期比較', '目標值', '達成率'
            ]
            sheet_r3 = ['取締方式', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '當場攔停', '逕行舉發', '', '', '']
            sheet_data = df_final.fillna("").values.tolist()
            sheet_r14 = [NOTE_TEXT] + [''] * 9

            full_sheet_data = [sheet_r1, sheet_r2, sheet_r3] + sheet_data + [sheet_r14]
            file_ids = ",".join(sorted([f.name for f in uploaded_files]))

            def run_automation():
                with st.status("🚀 執行自動化任務...", expanded=True) as status:
                    st.write("📧 正在寄送 Email...")
                    email_receiver = st.secrets["email"]["user"] if "email" in st.secrets else None
                    if email_receiver:
                        if send_email(email_receiver, f"📊 [自動通知] {file_name_out}", "附件為重點違規統計報表。", excel_data, file_name_out):
                            st.write(f"✅ Email 已發送")
                    else: st.warning("⚠️ 未設定 Email Secrets")
                    
                    st.write("📊 正在寫入 Google 試算表 (A1 ~ J14) 並修復顏色...")
                    if update_google_sheet(full_sheet_data, GOOGLE_SHEET_URL):
                        st.write("✅ 寫入成功！ (綠字已修復，格式已同步)")
                    else: st.write("❌ 寫入失敗")
                    
                    status.update(label="執行完畢", state="complete", expanded=False)
                    st.balloons()
            
            if file_ids not in st.session_state["sent_cache"]:
                run_automation()
                st.session_state["sent_cache"].add(file_ids)
            else: st.info("✅ 已自動執行過。")

            if st.button("🔄 強制執行", type="primary"): run_automation()

            st.download_button(label="📥 下載 Excel", data=excel_data, file_name=file_name_out, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e: 
            st.error(f"❌ 發生嚴重錯誤：{e}")
