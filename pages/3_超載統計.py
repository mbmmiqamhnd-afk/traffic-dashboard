import streamlit as st
import pandas as pd
import re
import io
import gspread
from datetime import date
import calendar

# 強制清除快取
try:
    st.cache_data.clear()
    st.cache_resource.clear()
except: pass

st.set_page_config(page_title="超載統計", layout="wide", page_icon="🚛")
st.title("🚛 超載自動統計 (v37 精準格式連動版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 
TARGETS = {'聖亭所': 24, '龍潭所': 32, '中興所': 24, '石門所': 19, '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30}
UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'}
UNIT_DATA_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. 核心函數：富文本格式處理 (Google Sheets API)
# ==========================================
def apply_rich_text_format(ws, row_idx, col_idx, text):
    """
    row_idx, col_idx 均為 1-based (Excel 習慣)
    """
    # 定義需要標紅的字符集合
    red_chars = set("0123456789年月日~().%")
    
    runs = []
    current_is_red = None
    
    for i, char in enumerate(text):
        is_red = char in red_chars
        if is_red != current_is_red:
            format_run = {"startIndex": i}
            if is_red:
                format_run["format"] = {"foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}, "bold": True}
            else:
                format_run["format"] = {"foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}, "bold": False}
            runs.append(format_run)
            current_is_red = is_red

    # 構造 Google Sheets API 請求
    request = {
        "updateCells": {
            "rows": [{
                "values": [{
                    "userEnteredValue": {"stringValue": text},
                    "textFormatRuns": runs
                }]
            }],
            "fields": "userEnteredValue,textFormatRuns",
            "range": {
                "sheetId": ws.id,
                "startRowIndex": row_idx - 1,
                "endRowIndex": row_idx,
                "startColumnIndex": col_idx - 1,
                "endColumnIndex": col_idx
            }
        }
    }
    return request

def sync_to_google_sheets(df, footer_text):
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        ws = sh.get_worksheet(1) # 分頁 2
        
        # 1. 寫入基本數據
        clean_cols = [re.sub(r'<[^>]+>', '', c) for c in df.columns]
        payload = [clean_cols] + df.values.tolist()
        ws.update(range_name='A2', values=payload)
        
        # 2. 準備批次更新請求 (格式化)
        requests = []
        
        # A. 格式化標題列日期 (B2, C2, D2)
        for i, col_content in enumerate(clean_cols[1:4], start=2): # B=2, C=3, D=4
            requests.append(apply_rich_text_format(ws, 2, i, col_content))
            
        # B. 寫入並格式化末端說明列
        footer_row_idx = 2 + len(df) + 1
        ws.update_cell(footer_row_idx, 1, footer_text)
        requests.append(apply_rich_text_format(ws, footer_row_idx, 1, footer_text))
        
        # 3. 發送 API 請求
        sh.batch_update({"requests": requests})
        return True
    except Exception as e:
        st.error(f"❌ 格式同步失敗: {e}")
        return False

# ==========================================
# 2. 解析與介面
# ==========================================
def parse_stone_report(f):
    if not f: return {}, "0000000", "0000000"
    unit_counts, s_str, e_str = {}, "0000000", "0000000"
    try:
        f.seek(0)
        text = pd.read_excel(f, header=None, nrows=15).to_string()
        m = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', text)
        if m: s_str, e_str = m.group(1), m.group(2)
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
                        if short in UNIT_DATA_ORDER: unit_counts[short] = unit_counts.get(short, 0) + int(nums[-1])
                        u = None
        return unit_counts, s_str, e_str
    except: return {}, "0000000", "0000000"

files = st.file_uploader("上傳 3 個 stoneCnt 報表", accept_multiple_files=True, type=['xlsx', 'xls'])

if files and len(files) >= 3:
    try:
        f_wk, f_yt, f_ly = None, None, None
        for f in files:
            if "(1)" in f.name: f_yt = f
            elif "(2)" in f.name: f_ly = f
            else: f_wk = f
        
        d_wk, s_wk, e_wk = parse_stone_report(f_wk)
        d_yt, s_yt, e_yt = parse_stone_report(f_yt)
        d_ly, s_ly, e_ly = parse_stone_report(f_ly)

        # 網頁顯示用的 HTML (局部紅)
        r_s, r_e = "<span style='color:red; font-weight:bold;'>", "</span>"
        c_wk = f"本期 {r_s}({s_wk[-4:]}~{e_wk[-4:]}){r_e}"
        c_yt = f"本年累計 {r_s}({s_yt}~{e_yt}){r_e}"
        c_ly = f"去年累計 {r_s}({s_ly}~{e_ly}){r_e}"

        body = []
        for u in UNIT_DATA_ORDER:
            yv, tv = d_yt.get(u, 0), TARGETS.get(u, 0)
            body.append({
                '統計期間': u, c_wk: d_wk.get(u, 0), c_yt: yv, c_ly: d_ly.get(u, 0),
                '本年與去年同期比較': yv - d_ly.get(u, 0), '目標值': tv, '達成率': f"{yv/tv:.0%}" if tv > 0 else "—"
            })
        
        df_final = pd.concat([pd.DataFrame([{'統計期間': '合計', c_wk: 0, c_yt: 0, c_ly: 0, '本年與去年同期比較': 0, '目標值': 0, '達成率': '0%'}]), pd.DataFrame(body)], ignore_index=True)
        # 修正合計數值
        sum_cols = pd.DataFrame(body)[pd.DataFrame(body)['統計期間'] != '警備隊'][[c_wk, c_yt, c_ly, '目標值']].sum()
        df_final.iloc[0, 1:5] = [sum_cols[c_wk], sum_cols[c_yt], sum_cols[c_ly], sum_cols[c_yt]-sum_cols[c_ly]]
        df_final.iloc[0, 5] = sum_cols['目標值']
        df_final.iloc[0, 6] = f"{sum_cols[c_yt]/sum_cols['目標值']:.0%}" if sum_cols['目標值'] > 0 else "0%"

        # 底部說明
        y_v, m_v, d_v = int(e_yt[:3])+1911, int(e_yt[3:5]), int(e_yt[5:])
        prog = ((date(y_v, m_v, d_v) - date(y_v, 1, 1)).days + 1) / (366 if calendar.isleap(y_v) else 365)
        f_text = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt[:3]}年{e_yt[3:5]}月{e_yt[5:]}日 (入案日期)應達成率為{prog:.1%}"
        
        st.success("✅ 解析完成")
        st.write(df_final.to_html(escape=False, index=False), unsafe_allow_html=True)
        
        # 網頁說明文字標紅顯示
        f_rich = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 :red[{e_yt[:3]}]年:red[{e_yt[3:5]}]月:red[{e_yt[5:]}]日 (入案日期)應達成率為:red[{prog:.1%}]"
        st.markdown(f"#### {f_rich}")

        if st.button("🚀 同步至雲端 (標題日期精準標紅)", type="primary"):
            with st.status("正在發送富文本指令...") as s:
                if sync_to_google_sheets(df_final, f_text):
                    st.write("✅ 同步成功！試算表內僅日期數字與符號為紅色。")
                    st.balloons()
                s.update(label="同步結束", state="complete")
    except Exception as e: st.error(f"錯誤：{e}")
