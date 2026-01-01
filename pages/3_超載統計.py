import streamlit as st
import pandas as pd
import re
import io
import smtplib
import gspread
from datetime import date
import calendar
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

st.set_page_config(page_title="超載統計", layout="wide", page_icon="🚛")
st.title("🚛 超載自動統計 (v35 標題日期標紅版)")

# --- 核心重置按鈕 ---
if st.button("🧹 徹底重置環境", type="primary"):
    st.cache_data.clear()
    st.cache_resource.clear()
    for key in st.session_state.keys():
        del st.session_state[key]
    st.success("✅ 已清空！請重新整理頁面 (F5)。")
    st.stop()

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 

TARGETS = {'聖亭所': 24, '龍潭所': 32, '中興所': 24, '石門所': 19, '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30}
UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'}
UNIT_DATA_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. 核心函數
# ==========================================
def update_sheet_final(df, footer_text, sheet_url):
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(sheet_url)
        ws = sh.get_worksheet(1) 
        # 寫入純淨文字
        clean_cols = [re.sub(r'<[^>]+>', '', c) for c in df.columns]
        payload = [clean_cols] + df.values.tolist() + [[footer_text] + [""]*(len(df.columns)-1)]
        ws.update(range_name='A2', values=payload)
        return True
    except: return False

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

# ==========================================
# 2. 執行邏輯
# ==========================================
files = st.file_uploader("上傳 3 個 stoneCnt 報表檔案", accept_multiple_files=True, type=['xlsx', 'xls'])

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

        # 定義「標紅日期」的 HTML 標頭
        red_span = "<span style='color:red; font-weight:bold;'>"
        end_span = "</span>"
        
        col_wk = f"本期 {red_span}({e_wk[-4:2]}~{e_wk[-4:]}){end_span}" # 修正截取邏輯
        col_wk = f"本期 {red_span}({s_wk[-4:]}~{e_wk[-4:]}){end_span}"
        col_yt = f"本年累計 {red_span}({s_yt}~{e_yt}){end_span}"
        col_ly = f"去年累計 {red_span}({s_ly}~{e_ly}){end_span}"

        # 建立數據
        body = []
        for u in UNIT_DATA_ORDER:
            yv, tv = d_yt.get(u, 0), TARGETS.get(u, 0)
            body.append({
                '統計期間': u, col_wk: d_wk.get(u, 0), col_yt: yv, col_ly: d_ly.get(u, 0),
                '本年與去年同期比較': yv - d_ly.get(u, 0), '目標值': tv, '達成率': f"{yv/tv:.0%}" if tv > 0 else "—"
            })
        
        df_t = pd.DataFrame(body)
        sum_d = df_t[df_t['統計期間'] != '警備隊'][[col_wk, col_yt, col_ly, '目標值']].sum()
        total_row = pd.DataFrame([{'統計期間': '合計', col_wk: sum_d[col_wk], col_yt: sum_d[col_yt], col_ly: sum_d[col_ly], '本年與去年同期比較': sum_d[col_yt] - sum_d[col_ly], '目標值': sum_d['目標值'], '達成率': f"{sum_d[col_yt]/sum_d['目標值']:.0%}" if sum_d['目標值'] > 0 else "0%"}])
        df_final = pd.concat([total_row, df_t], ignore_index=True)

        # 3. 說明文字標紅邏輯
        try:
            y_val, m_val, d_val = int(e_yt[:3])+1911, int(e_yt[3:5]), int(e_yt[5:])
            prog = ((date(y_val, m_val, d_val) - date(y_val, 1, 1)).days + 1) / (366 if calendar.isleap(y_val) else 365)
            f_plain = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt[:3]}年{e_yt[3:5]}月{e_yt[5:]}日 (入案日期)應達成率為{prog:.1%}"
            f_rich = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 :red[{e_yt[:3]}]年:red[{e_yt[3:5]}]月:red[{e_yt[5:]}]日 (入案日期)應達成率為:red[{prog:.1%}]"
        except: f_plain = "日期錯誤"; f_rich = f_plain

        st.success("✅ 數據解析完成")
        
        # 使用 HTML 渲染表格以達成標題局部標紅
        st.write("### 📋 報表預覽 (標題日期已標紅)")
        st.write(df_final.to_html(escape=False, index=False), unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"#### {f_rich}")

        # 4. 同步與下載
        st.markdown("---")
        if st.button("🚀 同步試算表並產出報表", type="primary"):
            with st.status("執行中...") as s:
                if update_sheet_final(df_final, f_plain, GOOGLE_SHEET_URL):
                    st.write("✅ 試算表同步成功 (已自動過濾 HTML 標籤)")
                    st.balloons()
                s.update(label="完成", state="complete")

        # Excel 下載需去除標籤
        out_excel = io.BytesIO()
        df_excel = df_final.copy()
        df_excel.columns = [re.sub(r'<[^>]+>', '', c) for c in df_excel.columns]
        df_excel.to_excel(out_excel, index=False)
        st.download_button("📥 下載 Excel 報表", out_excel.getvalue(), f"Report_{e_yt}.xlsx")

    except Exception as e:
        st.error(f"錯誤：{e}")
