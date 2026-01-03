import streamlit as st
import pandas as pd
import re
import io
import smtplib
import gspread
import calendar
import traceback
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header

# --- 初始化 ---
st.set_page_config(page_title="超載統計", layout="wide", page_icon="🚛")
st.title("🚛 超載自動統計 (v48 標題格式化版)")

# ==========================================
# 0. 核心設定 (請確認 Secrets 已填寫)
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit" 
TARGETS = {'聖亭所': 24, '龍潭所': 32, '中興所': 24, '石門所': 19, '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30}
UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'}
UNIT_DATA_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

# ==========================================
# 1. 核心格式指令 (Google Sheets API)
# ==========================================
def get_footer_precise_red_req(ws_id, row_idx, col_idx, text):
    runs = [{"startIndex": 0, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}}]
    anchor = "應達成率為"
    idx = text.find(anchor)
    if idx != -1:
        search_part = text[idx + len(anchor):]
        match = re.search(r'(\d+\.?\d*%)', search_part)
        if match:
            start_pos = idx + len(anchor) + match.start()
            end_pos = idx + len(anchor) + match.end()
            runs.append({"startIndex": start_pos, "format": {"foregroundColor": {"red": 1.0, "green": 0, "blue": 0}, "bold": True}})
            if end_pos < len(text):
                runs.append({"startIndex": end_pos, "format": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False}})
    return {"updateCells": {"rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}], "fields": "userEnteredValue,textFormatRuns", "range": {"sheetId": ws_id, "startRowIndex": row_idx-1, "endRowIndex": row_idx, "startColumnIndex": col_idx-1, "endColumnIndex": col_idx}}}

def get_header_num_red_req(ws_id, row_idx, col_idx, text):
    red_chars = set("0123456789~().%")
    runs = []
    last_is_red = None
    for i, char in enumerate(text):
        is_red = char in red_chars
        if is_red != last_is_red:
            color = {"red": 1.0, "green": 0, "blue": 0} if is_red else {"red": 0, "green": 0, "blue": 0}
            runs.append({"startIndex": i, "format": {"foregroundColor": color, "bold": is_red}})
            last_is_red = is_red
    return {"updateCells": {"rows": [{"values": [{"userEnteredValue": {"stringValue": text}, "textFormatRuns": runs}]}], "fields": "userEnteredValue,textFormatRuns", "range": {"sheetId": ws_id, "startRowIndex": row_idx-1, "endRowIndex": row_idx, "startColumnIndex": col_idx-1, "endColumnIndex": col_idx}}}

# ==========================================
# 2. 寄信功能與偵錯
# ==========================================
def send_report_email(excel_bytes, subject):
    try:
        if "email" not in st.secrets:
            return "錯誤：未在 Secrets 中設定 [email] 資訊"
        
        user = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg['Subject'] = Header(subject, 'utf-8').encode()
        msg['From'] = user
        msg['To'] = user
        msg.attach(MIMEText("自動產生的超載報表已同步，請查閱附件。", 'plain'))
        
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(excel_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="Overload_Report.xlsx"')
        msg.attach(part)
        
        # 建立連線並寄信
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(user, pwd)
        server.send_message(msg)
        server.quit()
        return "成功"
    except Exception as e:
        return f"郵件錯誤詳細資訊：{str(e) if str(e) else repr(e)}"

# ==========================================
# 3. 解析邏輯
# ==========================================
def parse_report(f):
    if not f: return {}, "0000000", "0000000"
    counts, s, e = {}, "0000000", "0000000"
    try:
        f.seek(0)
        df_top = pd.read_excel(f, header=None, nrows=15)
        text_block = df_top.to_string()
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
                        short = UNIT_MAP.get(u, u)
                        if short in UNIT_DATA_ORDER: counts[short] = counts.get(short, 0) + int(nums[-1])
                        u = None
        return counts, s, e
    except Exception as ex:
        raise ValueError(f"解析失敗: {ex}")

# ==========================================
# 4. 主程式執行
# ==========================================
# 側邊欄工具
st.sidebar.markdown("### 🛠️ 郵件測試工具")
if st.sidebar.button("📧 寄送測試郵件 (不含附件)"):
    res = send_report_email(b"", "超載統計 - 寄信測試")
    if res == "成功": st.sidebar.success("✅ 郵件設定正確，測試信已送出")
    else: st.sidebar.error(f"❌ 測試失敗\n{res}")

files = st.file_uploader("請同時上傳 3 個 stoneCnt 報表", accept_multiple_files=True, type=['xlsx', 'xls'])

if files and len(files) >= 3:
    try:
        file_hash = "".join(sorted([f.name + str(f.size) for f in files]))
        
        f_wk, f_yt, f_ly = None, None, None
        for f in files:
            if "(1)" in f.name: f_yt = f
            elif "(2)" in f.name: f_ly = f
            else: f_wk = f
        
        if not all([f_wk, f_yt, f_ly]):
            st.error("❌ 檔案命名不符合規則，請確認是否有 (1) 與 (2)。")
            st.stop()

        # 解析數據
        d_wk, s_wk, e_wk = parse_report(f_wk)
        d_yt, s_yt, e_yt = parse_report(f_yt)
        d_ly, s_ly, e_ly = parse_report(f_ly)

        # 欄位標題與標紅 HTML
        raw_wk = f"本期 ({s_wk[-4:]}~{e_wk[-4:]})"
        raw_yt = f"本年累計 ({s_yt[-4:]}~{e_yt[-4:]})"
        raw_ly = f"去年累計 ({s_ly[-4:]}~{e_ly[-4:]})"

        def h_html(t): return "".join([f"<span style='color:red; font-weight:bold;'>{c}</span>" if c in "0123456789~().%" else c for c in t])
        h_wk, h_yt, h_ly = map(h_html, [raw_wk, raw_yt, raw_ly])

        # 組裝數據表格
        body = []
        for u in UNIT_DATA_ORDER:
            yv, tv = d_yt.get(u, 0), TARGETS.get(u, 0)
            body.append({'統計期間': u, h_wk: d_wk.get(u, 0), h_yt: yv, h_ly: d_ly.get(u, 0), '本年與去年同期比較': yv - d_ly.get(u, 0), '目標值': tv, '達成率': f"{yv/tv:.0%}" if tv > 0 else "—"})
        
        df_body = pd.DataFrame(body)
        sum_v = df_body[df_body['統計期間'] != '警備隊'][[h_wk, h_yt, h_ly, '目標值']].sum()
        total_row = pd.DataFrame([{'統計期間': '合計', h_wk: sum_v[h_wk], h_yt: sum_v[h_yt], h_ly: sum_v[h_ly], '本年與去年同期比較': sum_v[h_yt] - sum_v[h_ly], '目標值': sum_v['目標值'], '達成率': f"{sum_v[h_yt]/sum_v['目標值']:.0%}" if sum_v['目標值'] > 0 else "0%"}])
        df_final = pd.concat([total_row, df_body], ignore_index=True)

        # 底部說明文字與百分比標紅
        y, m, d = int(e_yt[:3])+1911, int(e_yt[3:5]), int(e_yt[5:])
        prog_str = f"{((date(y, m, d) - date(y, 1, 1)).days + 1) / (366 if calendar.isleap(y) else 365):.1%}"
        f_plain = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 {e_yt[:3]}年{e_yt[3:5]}月{e_yt[5:]}日 (入案日期)應達成率為{prog_str}"
        f_html = f_plain.replace(prog_str, f"<span style='color:red; font-weight:bold;'>{prog_str}</span>")

        # 預覽介面
        st.success("✅ 數據解析成功！")
        # 顯示時稍微美化一下標題 (純前端顯示)
        st.markdown(f"<h3 style='text-align: center; color: blue;'>取締超載違規件數統計表</h3>", unsafe_allow_html=True)
        st.write(df_final.to_html(escape=False, index=False), unsafe_allow_html=True)
        st.write(f"#### {f_html}", unsafe_allow_html=True)

        # --- 自動化流程 ---
        if st.session_state.get("processed_hash") != file_hash:
            with st.status("🚀 執行雲端同步與自動寄信...") as s:
                try:
                    # ==========================================
                    # 1. Google Sheets 同步 (修改版：新增標題列)
                    # ==========================================
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    ws = sh.get_worksheet(1) # 假設是第2個工作表
                    
                    clean_cols = ['統計期間', raw_wk, raw_yt, raw_ly, '本年與去年同期比較', '目標值', '達成率']
                    
                    # A. 寫入資料
                    # A1 寫入標題
                    ws.update(range_name='A1', values=[['取締超載違規件數統計表']])
                    # A2 開始寫入欄位名稱與資料
                    ws.update(range_name='A2', values=[clean_cols] + df_final.values.tolist())
                    
                    # B. 格式化請求 (標題 + 內文紅字)
                    reqs = []
                    
                    # (1) 標題格式：合併 A1:G1、藍色粗體、置中、字型加大
                    reqs.append({
                        "mergeCells": {
                            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 7},
                            "mergeType": "MERGE_ALL"
                        }
                    })
                    reqs.append({
                        "updateCells": {
                            "rows": [{"values": [{"userEnteredFormat": {
                                "horizontalAlignment": "CENTER",
                                "verticalAlignment": "MIDDLE",
                                "textFormat": {"foregroundColor": {"blue": 1.0}, "fontSize": 18, "bold": True}
                            }}]}],
                            "fields": "userEnteredFormat",
                            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1}
                        }
                    })

                    # (2) 內文紅字邏輯 (注意 row_idx 要 +1，因為多了一列標題)
                    # 欄位標題 (現在在第 2 列)
                    for i, t in enumerate(clean_cols[1:4], 2):
                        reqs.append(get_header_num_red_req(ws.id, 2, i, t))
                    
                    # 底部說明文字 (現在在 資料長度 + 2(標題列) + 1(緩衝) + 1(index修正) = len + 4)
                    footer_row_idx = 2 + len(df_final) + 1
                    reqs.append(get_footer_precise_red_req(ws.id, footer_row_idx, 1, f_plain))
                    
                    sh.batch_update({"requests": reqs})
                    st.write("✅ 試算表同步與格式化完成 (含標題)")

                    # ==========================================
                    # 2. 自動寄信 (修改版：Excel 增加標題)
                    # ==========================================
                    st.write("📧 正在準備郵件附件並寄信...")
                    df_sync = df_final.copy()
                    df_sync.columns = clean_cols
                    
                    df_excel_buffer = io.BytesIO()
                    
                    # 使用 ExcelWriter 引擎來製作漂亮的標題
                    with pd.ExcelWriter(df_excel_buffer, engine='xlsxwriter') as writer:
                        # 資料從第 2 列開始寫 (startrow=1，Excel index 從 0 開始算)
                        df_sync.to_excel(writer, index=False, startrow=1, sheet_name='Sheet1')
                        
                        workbook = writer.book
                        worksheet = writer.sheets['Sheet1']
                        
                        # 定義標題格式
                        title_format = workbook.add_format({
                            'bold': True,
                            'font_size': 18,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_color': 'blue'
                        })
                        
                        # 合併 A1:G1 並寫入標題
                        worksheet.merge_range('A1:G1', '取締超載違規件數統計表', title_format)
                        
                        # (選用) 調整欄寬讓它好看一點
                        worksheet.set_column('A:A', 15)
                        worksheet.set_column('B:G', 12)

                    mail_res = send_report_email(df_excel_buffer.getvalue(), f"🚛 超載報表 - {e_yt} ({prog_str})")
                    
                    if mail_res == "成功":
                        st.write("✅ 電子郵件自動寄送成功")
                    else:
                        st.error(f"❌ 郵件自動寄送失敗\n{mail_res}")

                    st.session_state["processed_hash"] = file_hash
                    st.balloons()
                    s.update(label="自動化流程處理完畢", state="complete")
                    
                except Exception as ex:
                    st.error(f"❌ 自動化流程中斷: {ex}")
                    st.write(traceback.format_exc())

    except Exception as e:
        st.error(f"⚠️ 嚴重錯誤: {e}")
