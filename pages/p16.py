import streamlit as st
import pandas as pd
import io
import re
import traceback
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime, timedelta

# ==========================================
# 0. 系統初始化與狀態管理
# ==========================================
st.set_page_config(page_title="勤務督導報告自動生成系統", page_icon="🚓", layout="wide")

if "unit_reports" not in st.session_state:
    st.session_state.unit_reports = {}

try:
    from menu import show_sidebar
    show_sidebar()
except:
    pass

st.markdown("""
    <style>
    @font-face { font-family: 'Kaiu'; src: url('kaiu.ttf'); }
    .stTextArea textarea {
        font-family: 'Kaiu', "標楷體", sans-serif !important;
        font-size: 19px !important;
        line-height: 1.7 !important;
        color: #1c1c1c !important;
    }
    .stTabs [data-baseweb="tab-list"] button { font-size: 18px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 1. 寄信功能
# ==========================================
def send_gmail(subject, body, receiver_email):
    try:
        sender_email = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        msg = MIMEText(body, 'plain', 'utf-8')
        msg['Subject'] = Header(subject, 'utf-8')
        msg['From'] = f"督導助手 <{sender_email}>"
        msg['To'] = receiver_email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        return True
    except Exception as e:
        st.error(f"寄信失敗：{e}")
        return False

# ==========================================
# 2. 高精準解析引擎 (修正直線分隔符 | 定位問題)
# ==========================================
def d_safe_int(val):
    try: return int(float(str(val).split('.')[0].replace(',', '')))
    except: return 0

def d_normalize_code(c):
    c_str = str(c).strip().upper()
    c_str = c_str.translate(str.maketrans('０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ', '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'))
    c_str = c_str.replace(".0", "")
    return str(int(c_str)) if c_str.isdigit() else c_str

def d_parse_time(val):
    val_str = str(val).strip().replace("\n", "").replace(" ", "")
    # 過濾包含年月日字眼的雜訊
    if any(x in val_str for x in ["年", "月", "日"]): return None, None
    
    # 🌟 核心修正：加入對 | (直線) 的支援，精準捕捉 06|07 格式
    m = re.search(r'(?<!\d)(\d{1,2})[:：\-\s~～/|]+(\d{1,2})(?!\d)', val_str)
    if m:
        sh, eh = int(m.group(1)), int(m.group(2))
        return sh, eh
        
    # 支援單一小時數字
    m_single = re.fullmatch(r'(\d{1,2})', val_str)
    if m_single:
        sh = int(m_single.group(1))
        return sh, sh + 1
        
    return None, None

def d_extract_duty(d_file, hour):
    res = {'v_name': '解析失敗', 'cadre_status': '無幹部資料', 'unit_name': '未偵測單位', 'term': '該所'}
    try:
        df = pd.read_excel(d_file, header=None, dtype=str).fillna("")
        
        # A. 偵測單位與類型
        unit_full = ""
        for r in range(5):
            rt = "".join([str(x) for x in df.iloc[r].values])
            m = re.search(r'([\u4e00-\u9fa5]+(分局|派出所|分隊|警備隊|隊))', rt)
            if m: 
                unit_full = m.group(1)
                res['unit_name'] = unit_full
                break
        
        is_guard = "警備隊" in unit_full or ("隊" in unit_full and "分隊" not in unit_full)
        res['term'] = "該隊" if is_guard or "分隊" in unit_full else "該所"
        loc_term = res['term'][1:]
        
        # B. 全表人員雷達
        f_map = {}
        all_text = " ".join(df.astype(str).values.flatten())
        p = r'([A-Z0-9]{1,2})\s*(所長|副所長|隊長|副隊長|分隊長|小隊長|巡官|巡佐|警員|實習)[\s\n]*([\u4e00-\u9fa5]{2,4})'
        matches = re.findall(p, all_text)
        for m in matches:
            f_map[d_normalize_code(m[0])] = f"{m[1]}{m[2]}"
            
        # C. 時間座標鎖定法
        t_cols, tr_idx = {}, -1
        for r in range(15): 
            tmp = {}
            for c in range(len(df.columns)):
                sh, eh = d_parse_time(df.iloc[r, c])
                if sh is not None: tmp[c] = (sh, eh)
            if len(tmp) > len(t_cols): 
                t_cols = tmp
                tr_idx = r
                
        # 對位抵達時間
        target_col = -1
        adj_h = hour if hour >= 6 else hour + 24
        for c, (sh, eh) in t_cols.items():
            s = sh if sh >= 6 else sh + 24
            e = eh if eh > sh else eh + 24
            if e < s: e += 24 # 跨夜修正
            if s <= adj_h < e: 
                target_col = c; break

        if target_col != -1 and tr_idx != -1:
            # 1. 值班偵測
            for r in range(tr_idx + 1, min(tr_idx + 25, len(df))):
                row_head = "".join(df.iloc[r, :target_col+1])
                if any(x in row_head for x in ["值", "班"]):
                    val = str(df.iloc[r, target_col])
                    m_code = re.search(r'[A-Z0-9]{1,2}', val)
                    if m_code:
                        code = d_normalize_code(m_code.group(0))
                        res['v_name'] = f_map.get(code, f"警員({code})")
                        break
            
            # 2. 幹部動態
            c_codes = ["01", "02", "A", "B", "C"] if is_guard else ["A", "B", "C"]
            default_t = {"01":"隊長","02":"副隊長","A":"所長","B":"副所長","C":"幹部"}
            c_notes = []
            
            for code in c_codes:
                if code not in f_map and code not in ["A", "B"]: continue
                fname = f_map.get(code, default_t.get(code, "幹部"))
                is_off = False
                for r in range(max(0, len(df)-15), len(df)):
                    row_str = "".join(df.iloc[r, :]).upper()
                    if code in row_str and any(k in "".join(df.iloc[r, :4]) for k in ["休","輪","假","補"]):
                        is_off = True; break
                
                d_names = set()
                for r in range(tr_idx + 1, len(df)):
                    cell_val = str(df.iloc[r, target_col])
                    cell_codes = [d_normalize_code(x) for x in re.findall(r'[A-Z0-9]{1,2}', cell_val)]
                    if code in cell_codes:
                        is_off = False
                        row_n = "".join(df.iloc[r, :5])
                        kw_map = {"巡":"巡邏", "守":"守望", "望":"守望", "臨":"臨檢", "交":"交整", "路":"路檢", "督":"督勤", "備":"備勤", "專":"專案", "辦":"偵辦刑案", "淨":"專案"}
                        for k, kn in kw_map.items():
                            if k in row_n or k in cell_val: d_names.add(kn)
                
                if is_off: c_notes.append(f"{fname}休假")
                else:
                    if d_names:
                        sh_s, eh_s = t_cols[target_col]
                        c_notes.append(f"{fname}在{loc_term}督勤，編排{sh_s:02d}-{eh_s if eh_s!=0 else 24:02d}時段{'、'.join(sorted(d_names))}勤務")
                    else: c_notes.append(f"{fname}在{loc_term}督勤")
            res['cadre_status'] = "；".join(c_notes) + "。"
        else:
            res['v_name'] = "時段對位失敗"
            res['cadre_status'] = f"無法定位 {hour:02d} 點的欄位。請確認班表第 1-15 列中是否有正確的時間標題列（例如 11-12、12-13 或 12|13）。"
    except Exception as e: 
        res['cadre_status'] = f"解析發生錯誤：{str(e)}"
    return res

# ==========================================
# 3. 裝備解析
# ==========================================
def d_extract_equip(e_file, hour):
    try:
        df = pd.read_excel(e_file, header=None).fillna("")
        df_s = df.astype(str)
        col_map = {"gun": 2, "bullet": 3, "radio": 6, "vest": 11}
        for r in range(min(10, len(df))):
            for c in range(len(df.columns)):
                v = str(df.iloc[r, c])
                if "手槍" in v: col_map["gun"] = c
                if "子彈" in v: col_map["bullet"] = c
        sub = df.iloc[:35]; sub_s = df_s.iloc[:35]
        def get_v(kw, k):
            rows = sub[sub_s.iloc[:, 1].str.contains(kw, na=False)]
            return d_safe_int(rows.iloc[-1, col_map[k]]) if not rows.empty else 0
        r_fix = sub[sub_s.iloc[:, 1].str.contains("送", na=False)]
        return {"gi":get_v("在","gun"), "go":get_v("出","gun"), "gf":d_safe_int(r_fix.iloc[-1, col_map["gun"]]) if not r_fix.empty else 0,
                "bi":get_v("在","bullet"), "bo":get_v("出","bullet"),
                "ri":get_v("在","radio"), "ro":get_v("出","radio"), "rf":d_safe_int(r_fix.iloc[-1, col_map["radio"]]) if not r_fix.empty else 0,
                "vi":get_v("在","vest"), "vo":get_v("出","vest")}
    except: return None

# ==========================================
# 4. 主介面 UI
# ==========================================
st.header("📋 勤務督導報告自動生成")

c_s1, c_s2 = st.columns(2)
with c_s1:
    insp_date = st.date_input("選擇督導日期", datetime.now(), key="insp_d")
    num_units = st.number_input("待督導單位數量", 1, 8, 3, key="num_u")
with c_s2:
    d_e = insp_date - timedelta(days=1)
    d_5, d_3 = insp_date - timedelta(days=5), insp_date - timedelta(days=3)
    st.info(f"📅 檢測區間：監錄({d_5.strftime('%m%d')}-{d_e.strftime('%m%d')}) / 勤教({d_3.strftime('%m%d')}-{d_e.strftime('%m%d')})")

u_tabs = st.tabs([f"🏢 單位 {i+1}" for i in range(num_units)] + ["📄 總匯整報告"])

for i in range(num_units):
    with u_tabs[i]:
        u_time = st.time_input("抵達時間", datetime.now().time(), key=f"ut_{i}")
        col_f1, col_f2 = st.columns(2)
        with col_f1: u_duty = st.file_uploader(f"單位 {i+1} 勤務表 (.xlsx)", type=['xlsx'], key=f"ud_{i}")
        with col_f2: u_eq = st.file_uploader(f"單位 {i+1} 交接簿 (.xlsx)", type=['xlsx'], key=f"ue_{i}")

        if u_duty and u_eq:
            dr = d_extract_duty(u_duty, u_time.hour)
            er = d_extract_equip(u_eq, u_time.hour)
            uname, t = dr['unit_name'], dr['term']
            
            lns = [
                f"{u_time.strftime('%H%M')}，{t}值班{dr['v_name']}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。",
                f"{t}駐地監錄設備及天羅地網系統均運作正常，無故障，{d_5.strftime('%m月%d日')}至{d_e.strftime('%m月%d日')}有逐日檢測2次以上紀錄。",
                f"{t}{d_3.strftime('%m月%d日')}至{d_e.strftime('%m月%d日')}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。",
                f"{t}環境內務擺設整齊清潔，符合規定。",
                f"{t}手槍出勤 {er['go'] if er else 0} 把、在所 {er['gi'] if er else 0} 把，子彈出勤 {er['bo'] if er else 0} 顆、在所 {er['bi'] if er else 0} 顆，無線電出勤 {er['ro'] if er else 0} 臺、在所 {er['ri'] if er else 0} 臺；防彈背心出勤 {er['vo'] if er else 0} 件、在所 {er['vi'] if er else 0} 件，幹部對械彈每日檢查管制良好，符合規定。",
                f"本日{dr['cadre_status']}",
                f"{t}酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。"
            ]
            final_text = "\n".join([f"{idx+1}、{line}" for idx, line in enumerate(lns)])
            st.session_state.unit_reports[i] = f"【{uname} 督導報告】\n{final_text}"
            
            if "失敗" in dr['v_name']:
                st.error(f"❌ {uname} 解析失敗：{dr['cadre_status']}")
            else:
                st.success(f"✅ {uname} 解析完成")
            st.text_area("預覽報告", final_text, height=350, key=f"preview_{i}")

with u_tabs[-1]:
    reports_list = [st.session_state.unit_reports[k] for k in sorted(st.session_state.unit_reports.keys()) if k < num_units]
    if reports_list:
        full_text = ("\n\n" + "─" * 40 + "\n\n").join(reports_list)
        st.subheader("📋 所有單位匯整結果")
        st.text_area("匯整文本", full_text, height=600)
        
        st.divider()
        st.subheader("📫 同步寄送報告至 Gmail")
        target_mail = st.text_input("收件信箱", "mbmmiqamhnd@gmail.com")
        if st.button("🚀 立即寄送郵件"):
            with st.spinner("郵件發送中..."):
                if send_gmail(f"勤務督導報告匯整_{insp_date.strftime('%Y%m%d')}", full_text, target_mail):
                    st.success(f"✅ 報告已成功寄送至 {target_mail}")
    else:
        st.warning("目前尚無解析完成的資料，請先至各單位分頁上傳檔案。")
