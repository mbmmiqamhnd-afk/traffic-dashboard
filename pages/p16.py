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
# 0. 系統初始化與選單配置
# ==========================================
st.set_page_config(page_title="交通業務與督導整合系統", page_icon="🚓", layout="wide")

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
# 1. 寄信功能函數
# ==========================================
def send_gmail(subject, body, receiver_email):
    try:
        sender_email = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        msg = MIMEText(body, 'plain', 'utf-8')
        msg['Subject'] = Header(subject, 'utf-8')
        msg['From'] = f"督導報告助手 <{sender_email}>"
        msg['To'] = receiver_email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        return True
    except Exception as e:
        st.error(f"寄信失敗：{e}")
        return False

# ==========================================
# 2. 解析核心引擎 (已整合積極偵測邏輯)
# ==========================================
def d_safe_int(val):
    try: return int(float(str(val).split('.')[0].replace(',', '')))
    except: return 0

def d_normalize_code(c):
    c_str = str(c).strip().replace(".0", "").upper()
    return str(int(c_str)) if c_str.isdigit() else c_str

def d_parse_time(val):
    val_str = str(val).strip().replace("\n", "").replace(" ", "").replace("|", "-")
    if val_str in ["", "nan", "NaN"]: return None, None
    m_time = re.search(r'(\d{1,2})[~~\-－—–_]+(\d{1,2})', val_str)
    if m_time: return int(m_time.group(1)), int(m_time.group(2))
    return None, None

def d_extract_equip(e_file, hour):
    try:
        df = pd.read_excel(e_file, header=None) if not e_file.name.endswith('csv') else pd.read_csv(e_file, header=None)
        df_s = df.astype(str)
        col_map = {"gun": None, "bullet": None, "radio": None, "vest": None}
        for r in range(min(12, len(df))):
            for c in range(len(df.columns)):
                v = str(df.iloc[r, c]).replace(" ", "").replace("\n", "")
                if col_map["gun"] is None and "手槍" in v: col_map["gun"] = c
                if col_map["bullet"] is None and "子彈" in v: col_map["bullet"] = c
                if col_map["radio"] is None and "無線電" in v: col_map["radio"] = c
                if col_map["vest"] is None and "背心" in v: col_map["vest"] = c
        col_map = {k: (v if v is not None else d) for k, v, d in zip(col_map.keys(), col_map.values(), [2, 3, 6, 11])}
        stop_r = len(df)
        for r in range(min(10, len(df)), len(df)):
            nums = re.findall(r'\d{1,2}', str(df.iloc[r, 0]))
            if nums and int(nums[0]) > hour and (int(nums[0]) - hour < 12):
                stop_r = r; break
        sub = df.iloc[:stop_r]; sub_s = df_s.iloc[:stop_r]
        def get_v(kw, k):
            rows = sub[sub_s.iloc[:, 1].str.contains(kw, na=False)]
            return d_safe_int(rows.iloc[-1, col_map[k]]) if not rows.empty else 0
        r_fix = sub[sub_s.iloc[:, 1].str.contains("送", na=False)]
        gf = d_safe_int(r_fix.iloc[-1, col_map["gun"]]) if not r_fix.empty else 0
        rf = d_safe_int(r_fix.iloc[-1, col_map["radio"]]) if not r_fix.empty else 0
        return {"gi":get_v("在","gun"), "go":get_v("出","gun"), "gf":gf, "bi":get_v("在","bullet"), "bo":get_v("出","bullet"), "ri":get_v("在","radio"), "ro":get_v("出","radio"), "rf":rf, "vi":get_v("在","vest"), "vo":get_v("出","vest")}
    except: return None

def d_extract_duty(d_file, hour):
    res = {'v_name': '解析失敗', 'cadre_status': '無幹部資料', 'unit_name': '未偵測單位', 'term': '該所'}
    try:
        df = pd.read_excel(d_file, header=None, dtype=str).fillna("")
        for r in range(5):
            rt = "".join([str(x) for x in df.iloc[r].values])
            m = re.search(r'([\u4e00-\u9fa5]+(分局|派出所|分隊|局))', rt)
            if m: res['unit_name'] = m.group(1); break
        is_traffic_unit = "分隊" in res['unit_name']
        res['term'] = "該分隊" if is_traffic_unit else "該所"
        
        full = " ".join([str(x).strip() for x in df.values.flatten() if x])
        p = r'([A-Z]|[0-9]{1,2})\s*(所長|副所長|分隊長|小隊長|巡官|巡佐|警員|警員兼副所長|實習)[\s\n]*([\u4e00-\u9fa5]{2,4})'
        matches = re.findall(p, full)
        n_map, f_map = {}, {}
        for m in matches:
            code_id, title, name = d_normalize_code(m[0]), m[1].strip(), m[2]
            if len(name) >= 2: n_map[code_id] = name; f_map[code_id] = f"{title}{name}"
            
        tr_idx, t_cols, t_col_idx = 2, {}, -1
        for r in range(6):
            tmp = {c: d_parse_time(df.iloc[r, c]) for c in range(len(df.columns)) if d_parse_time(df.iloc[r, c])[0] is not None}
            if len(tmp) > len(t_cols): tr_idx, t_cols = r, tmp
        for c, (sh, eh) in t_cols.items():
            ce = eh if eh > sh else eh + 24
            ch = hour if hour >= 6 or ce <= 24 else hour + 24
            if sh <= ch < ce: t_col_idx = c
            
        vr_idx = tr_idx + 1
        for r in range(tr_idx + 1, min(tr_idx + 8, len(df))):
            if "值" in "".join([str(x) for x in df.iloc[r, :4]]): vr_idx = r; break
        
        if t_col_idx != -1:
            raw = str(df.iloc[vr_idx, t_col_idx]).strip()
            mc = re.search(r'[A-Za-z0-9]{1,2}', raw)
            code_v = d_normalize_code(mc.group(0)) if mc else ""
            if code_v in f_map: res['v_name'] = f_map[code_v]
            else:
                for cid, nm in n_map.items():
                    if nm in raw: res['v_name'] = f_map[cid]; break
                    
        titles_dict = {"A": "分隊長" if is_traffic_unit else "所長", "B": "小隊長" if is_traffic_unit else "副所長", "C": "幹部"}
        c_notes = []
        for code_c in ["A", "B", "C"]:
            full_name = f_map.get(code_c, titles_dict[code_c])
            found, is_actually_off, d_names = False, False, set()
            for r in range(vr_idx, len(df)):
                cell_val = str(df.iloc[r, t_col_idx]) if t_col_idx != -1 else ""
                cell_codes = [d_normalize_code(x) for x in re.findall(r'[A-Za-z0-9]{1,2}', cell_val)]
                if code_c in cell_codes:
                    found = True
                    dt_area = "".join([str(x) for x in df.iloc[r, :2]])
                    row_all_text = "".join([str(x) for x in df.iloc[r, :]])
                    has_work_kw = any(kw in row_all_text for kw in ["巡","守","望","臨","交","路","督","勤","備","辦","內","專"])
                    if any(k in dt_area for k in ["休", "假", "輪", "補"]) and not has_work_kw:
                        is_actually_off = True
                    else:
                        is_actually_off = False
                    kw_map = {"巡":"巡邏", "守":"守望", "望":"守望", "臨":"臨檢", "交":"交整", "路":"路檢", "督":"督導", "備":"備勤", "辦":"辦公", "內":"內勤", "專":"專案"}
                    for k, kn in kw_map.items():
                        if k in row_all_text: d_names.add(kn)
            if not found or is_actually_off: c_notes.append(f"{full_name}休假")
            else:
                if d_names:
                    sh_s, eh_s = t_cols.get(t_col_idx, (0,0))
                    e_str = "24" if eh_s in (24, 0) else f"{eh_s % 24:02d}"
                    c_notes.append(f"{full_name}在所督勤，編排{sh_s:02d}至{e_str}時段{'、'.join(sorted(list(d_names)))}勤務")
                else: c_notes.append(f"{full_name}在所督勤")
        res['cadre_status'] = "；".join(c_notes) + "。"
    except: pass
    return res

# ==========================================
# 3. 主介面邏輯
# ==========================================
main_tabs = st.tabs(["📊 數據處理", "📋 勤務督導報告"])

with main_tabs[1]:
    st.header("📋 勤務督導報告自動生成")
    c_s1, c_s2 = st.columns(2)
    with c_s1:
        insp_date = st.date_input("督導日期", datetime.now(), key="insp_d")
        num_units = st.number_input("單位數量", 1, 8, 3, key="num_u")
    with c_s2:
        d_e = insp_date - timedelta(days=1)
        d_5, d_3 = insp_date - timedelta(days=5), insp_date - timedelta(days=3)
        st.info(f"📅 區間：監錄({d_5.strftime('%m%d')}-{d_e.strftime('%m%d')}) / 勤教({d_3.strftime('%m%d')}-{d_e.strftime('%m%d')})")

    u_tabs = st.tabs([f"🏢 單位 {i+1}" for i in range(num_units)] + ["📄 總匯整報告"])
    all_final_reports = []

    for i in range(num_units):
        with u_tabs[i]:
            u_time = st.time_input("抵達時間", datetime.now().time(), key=f"ut_{i}")
            col_f1, col_f2 = st.columns(2)
            with col_f1: u_duty = st.file_uploader("上傳勤務表", type=['xlsx'], key=f"ud_{i}")
            with col_f2: u_eq = st.file_uploader("上傳交接簿", type=['xlsx'], key=f"ue_{i}")

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
                final_text = "\n".join([f"{idx+1}、{line.replace('該所', t)}" for idx, line in enumerate(lns)])
                all_final_reports.append(f"【{uname} 督導報告】\n{final_text}")
                st.text_area("預覽", final_text, height=350, key=f"txt_{i}")

    with u_tabs[-1]:
        if all_final_reports:
            full_text = ("\n\n" + "─" * 40 + "\n\n").join(all_final_reports)
            st.text_area("📄 總匯整結果", full_text, height=600)
            target_mail = st.text_input("收件信箱", "mbmmiqamhnd@gmail.com")
            if st.button("🚀 寄送至 Gmail"):
                if send_gmail(f"督導報告_{insp_date.strftime('%Y%m%d')}", full_text, target_mail):
                    st.success("郵件寄送成功！")
