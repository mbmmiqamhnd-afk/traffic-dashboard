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
st.title("🚛 超載自動統計 (v38 精準標紅版)")

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
def apply_precise_red_format(ws, row_idx, col_idx, text):
    """
    建立 Google Sheets 富文本請求：
    數字、~、( )、.、% 為紅色粗體
    其餘 (包含中文字 年月日) 為黑色正常
    """
    # 定義標紅的符號集 (不含中文字)
    red_chars = set("0123456789~().%")
    
    runs = []
    last_is_red = None
    
    for i, char in enumerate(text):
        is_red = char in red_chars
        if is_red != last_is_red:
            format_run = {"startIndex": i}
            if is_red:
                format_run["format"] = {"foregroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}, "bold": True}
            else:
                format_run["format"] = {"foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}, "bold": False}
            runs.append(format_run)
            last_is_red = is_red

    return {
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
                "startRowIndex": row_idx - 1, "endRowIndex": row_idx,
                "startColumnIndex": col_idx - 1, "endColumnIndex": col_idx
            }
        }
    }

def sync_to_google_sheets(df, footer_text):
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        ws = sh.get_worksheet(1)
        
        # 1. 寫入基本數據
        clean_cols = [re.sub(r'<[^>]+>', '', c) for c in df.columns]
        payload = [clean_cols] + df.values.tolist()
        ws.update(range_name='A2', values=payload)
        
        # 2. 構造批次格式請求
        requests = []
        # 標題日期欄位標紅 (B2, C2, D2)
        for i, col_txt in enumerate(clean_cols[1:4], start=2):
            requests.append(apply_precise_red_format(ws, 2, i, col_txt))
        
        # 末端說明列標紅 (交通分隊下兩列)
        footer_idx = 2 + len(df) + 1
        ws.update_cell(footer_idx, 1, footer_text)
        requests.append(apply_precise_red_format(ws, footer_idx, 1, footer_text))
        
        sh.batch_update({"requests": requests})
        return True
    except Exception as e:
        st.error(f"❌ 同步失敗: {e}")
        return False

# ==========================================
# 2. 解析邏輯
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

# ==========================================
# 3. 網頁呈現
# ==========================================
def get_html_rich_text(text):
    """將數字與符號用 HTML 標紅"""
    red_chars = "0123456789~().%"
    new_text = ""
    for char in text:
        if char in red_chars:
            new_text += f"<span style='color:red; font-weight:bold;'>{char}</span>"
        else:
            new_text += char
    return new_text

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

        # 欄位名稱
        raw_wk = f"本期 ({s_wk[-4:]}~{e_wk[-4:]})"
        raw_yt = f"本年累計 ({s_yt}~{e_yt})"
        raw_ly = f"去年累計 ({s_ly}~{e_ly})"

        # HTML 版名稱 (預覽用)
        html_wk, html_yt, html_ly = map(get_html_rich_text, [raw_wk, raw_yt, raw_ly])

        # 數據計算
        body = []
        for u in UNIT_DATA_ORDER:
            yv, tv = d_yt.get(u, 0), TARGETS.get(u, 0)
            body.append({
                '統計期間': u, html_wk: d_wk.get(u, 0), html_yt: yv, html_ly: d_ly.get(u, 0),
                '本年與去年同期比較': yv - d_ly.get(u, 0), '目標值': tv, '達成率': f"{yv/tv:.0%}" if tv > 0 else "—"
            })
        
        df_body = pd.DataFrame(body)
        sum_cols = df_body[df_body['統計期間'] != '警備隊'][[html_wk, html_yt, html_ly, '目標值']].sum()
        total_row = pd.DataFrame([{'統計期間': '合計', html_wk: sum_cols[html_wk], html_yt: sum_cols[html_yt], html_ly: sum_cols[html_ly], '本年與去年同期比較': sum_cols[html_yt] - sum_cols[html_ly], '目標值': sum_cols['目標值'], '達成率': f"{sum_cols[html_yt]/sum_cols['目標值']:.0%}" if sum_cols['目標值'] > 0 else "0%"}])
        df_final = pd.concat([total_row, df_body], ignore_index=True)

        # 說明文字
        y, m, d = int(e_yt[:3])+1911, int(e_yt[3:5]), int(e_yt[5:])
        prog = ((date(y, m, d) - date(y, 1, 1)).days + 1) / (366 if calendar.isleap(y) else 365)
        footer_plain = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt[:3]}年{e_yt[3:5]}月{e_yt[5:]}日 (入案日期)應達成率為{prog:.1%}"
        footer_html = get_html_rich_text(footer_plain)

        st.success("✅ 解析成功")
        st.write(df_final.to_html(escape=False, index=False), unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        st.write(f"#### {footer_html}", unsafe_allow_html=True)

        if st.button("🚀 同步雲端 (精準標紅)", type="primary"):
            with st.status("正在同步精密格式...") as s:
                # 重新映射回純淨名稱給寫入用
                df_sync = df_final.copy()
                df_sync.columns = ['統計期間', raw_wk, raw_yt, raw_ly, '本年與去年同期比較', '目標值', '達成率']
                if sync_to_google_sheets(df_sync, footer_plain):
                    st.write("✅ 同步成功！中文字維持黑色，數字符號已標紅。")
                    st.balloons()
    except Exception as e: st.error(f"錯誤：{e}")
