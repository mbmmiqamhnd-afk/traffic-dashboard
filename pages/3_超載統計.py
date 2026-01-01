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
st.title("🚛 超載自動統計 (v39 末列僅百分比標紅版)")

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 
TARGETS = {'聖亭所': 24, '龍潭所': 32, '中興所': 24, '石門所': 19, '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30}
UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'}
UNIT_DATA_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. 富文本格式化核心 (Google Sheets API)
# ==========================================
def apply_footer_red_format(ws, row_idx, col_idx, text):
    """
    專門處理末列：僅將百分比部分標紅 (例如 99.5%)
    """
    runs = []
    # 使用正規表達式找出百分比的位置 (數字+點+數字+%)
    match = re.search(r'(\d+\.?\d*%)', text)
    
    if match:
        start, end = match.start(), match.end()
        # 0 ~ start 為黑色
        runs.append({"startIndex": 0, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}})
        # start ~ end 為紅色粗體
        runs.append({"startIndex": start, "format": {"foregroundColor": {"red": 1.0, "green": 0, "blue": 0}, "bold": True}})
        # end ~ 結束 為黑色
        if end < len(text):
            runs.append({"startIndex": end, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}})
    
    return {
        "updateCells": {
            "rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}],
            "fields": "userEnteredValue,textFormatRuns",
            "range": {"sheetId": ws.id, "startRowIndex": row_idx-1, "endRowIndex": row_idx, "startColumnIndex": col_idx-1, "endColumnIndex": col_idx}
        }
    }

def apply_header_red_format(ws, row_idx, col_idx, text):
    """
    處理標題列：數字符號紅，中文黑
    """
    red_chars = set("0123456789~().%")
    runs = []
    last_is_red = None
    for i, char in enumerate(text):
        is_red = char in red_chars
        if is_red != last_is_red:
            format_run = {"startIndex": i}
            color = {"red": 1.0, "green": 0, "blue": 0} if is_red else {"red": 0, "green": 0, "blue": 0}
            format_run["format"] = {"foregroundColor": color, "bold": is_red}
            runs.append(format_run)
            last_is_red = is_red
    return {
        "updateCells": {
            "rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}],
            "fields": "userEnteredValue,textFormatRuns",
            "range": {"sheetId": ws.id, "startRowIndex": row_idx-1, "endRowIndex": row_idx, "startColumnIndex": col_idx-1, "endColumnIndex": col_idx}
        }
    }

def sync_to_google_sheets(df, footer_text):
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        ws = sh.get_worksheet(1)
        
        # 1. 寫入數據
        clean_cols = [re.sub(r'<[^>]+>', '', c) for c in df.columns]
        ws.update(range_name='A2', values=[clean_cols] + df.values.tolist())
        
        # 2. 構造批次格式請求
        requests = []
        for i, col_txt in enumerate(clean_cols[1:4], start=2): # 標題日期欄 B, C, D
            requests.append(apply_header_red_format(ws, 2, i, col_txt))
        
        footer_idx = 2 + len(df) + 1
        ws.update_cell(footer_idx, 1, footer_text)
        requests.append(apply_footer_red_format(ws, footer_idx, 1, footer_text))
        
        sh.batch_update({"requests": requests})
        return True
    except Exception as e:
        st.error(f"❌ 同步失敗: {e}")
        return False

# ==========================================
# 2. 解析與介面
# ==========================================
def parse_report(f):
    if not f: return {}, "0000000", "0000000"
    counts, s, e = {}, "0000000", "0000000"
    try:
        f.seek(0)
        top = pd.read_excel(f, header=None, nrows=15).to_string()
        m = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', top)
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
    except: return {}, "0000000", "0000000"

def get_html_header_rich(text):
    red_chars = "0123456789~().%"
    res = ""
    for c in text:
        res += f"<span style='color:red; font-weight:bold;'>{c}</span>" if c in red_chars else c
    return res

def get_html_footer_rich(text):
    # 僅標紅百分比部分
    match = re.search(r'(\d+\.?\d*%)', text)
    if match:
        target = match.group(1)
        return text.replace(target, f"<span style='color:red; font-weight:bold;'>{target}</span>")
    return text

files = st.file_uploader("上傳 3 個 stoneCnt 報表", accept_multiple_files=True, type=['xlsx', 'xls'])

if files and len(files) >= 3:
    try:
        f_wk, f_yt, f_ly = None, None, None
        for f in files:
            if "(1)" in f.name: f_yt = f
            elif "(2)" in f.name: f_ly = f
            else: f_wk = f
        
        d_wk, s_wk, e_wk = parse_report(f_wk)
        d_yt, s_yt, e_yt = parse_report(f_yt)
        d_ly, s_ly, e_ly = parse_report(f_ly)

        raw_wk = f"本期 ({s_wk[-4:]}~{e_wk[-4:]})"
        raw_yt = f"本年累計 ({s_yt}~{e_yt})"
        raw_ly = f"去年累計 ({s_ly}~{e_ly})"

        h_wk, h_yt, h_ly = map(get_html_header_rich, [raw_wk, raw_yt, raw_ly])

        body = []
        for u in UNIT_DATA_ORDER:
            yv, tv = d_yt.get(u, 0), TARGETS.get(u, 0)
            body.append({'統計期間': u, h_wk: d_wk.get(u, 0), h_yt: yv, h_ly: d_ly.get(u, 0), '本年與去年同期比較': yv - d_ly.get(u, 0), '目標值': tv, '達成率': f"{yv/tv:.0%}" if tv > 0 else "—"})
        
        df_body = pd.DataFrame(body)
        sum_cols = df_body[df_body['統計期間'] != '警備隊'][[h_wk, h_yt, h_ly, '目標值']].sum()
        total_row = pd.DataFrame([{'統計期間': '合計', h_wk: sum_cols[h_wk], h_yt: sum_cols[h_yt], h_ly: sum_cols[h_ly], '本年與去年同期比較': sum_cols[h_yt] - sum_cols[h_ly], '目標值': sum_cols['目標值'], '達成率': f"{sum_cols[h_yt]/sum_cols['目標值']:.0%}" if sum_cols['目標值'] > 0 else "0%"}])
        df_final = pd.concat([total_row, df_body], ignore_index=True)

        y, m, d = int(e_yt[:3])+1911, int(e_yt[3:5]), int(e_yt[5:])
        prog = ((date(y, m, d) - date(y, 1, 1)).days + 1) / (366 if calendar.isleap(y) else 365)
        f_plain = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt[:3]}年{e_yt[3:5]}月{e_yt[5:]}日 (入案日期)應達成率為{prog:.1%}"
        f_html = get_html_footer_rich(f_plain)

        st.success("✅ 解析成功")
        st.write(df_final.to_html(escape=False, index=False), unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        st.write(f"#### {f_html}", unsafe_allow_html=True)

        if st.button("🚀 同步雲端 (精準格式版)", type="primary"):
            with st.status("正在同步...") as s:
                df_sync = df_final.copy()
                df_sync.columns = ['統計期間', raw_wk, raw_yt, raw_ly, '本年與去年同期比較', '目標值', '達成率']
                if sync_to_google_sheets(df_sync, f_plain):
                    st.write("✅ 同步完成！末列僅百分比標紅。")
                    st.balloons()
    except Exception as e: st.error(f"錯誤：{e}")
