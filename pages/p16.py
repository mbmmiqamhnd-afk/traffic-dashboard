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
# 0. 系統初始化
# ==========================================
st.set_page_config(page_title="交通業務與督導整合系統", page_icon="🚓", layout="wide")

if "unit_reports" not in st.session_state:
    st.session_state.unit_reports = {}

try:
    from menu import show_sidebar
    show_sidebar()
except:
    pass

# ==========================================
# 1. 寄信功能
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
# 2. 進化版解析引擎 (專治警備隊與高平格式)
# ==========================================
def d_safe_int(val):
    try: return int(float(str(val).split('.')[0].replace(',', '')))
    except: return 0

def d_normalize_code(c):
    c_str = str(c).strip().replace(".0", "").upper()
    return str(int(c_str)) if c_str.isdigit() else c_str

def d_parse_time(val):
    val_str = str(val).strip().replace("\n", "").replace(" ", "")
    m = re.search(r'(\d{1,2})[~~\-－—–_]+(\d{1,2})', val_str)
    return (int(m.group(1)), int(m.group(2))) if m else (None, None)

def d_extract_duty(d_file, hour):
    res = {'v_name': '解析失敗', 'cadre_status': '無幹部資料', 'unit_name': '未偵測單位', 'term': '該所'}
    try:
        df = pd.read_excel(d_file, header=None, dtype=str).fillna("")
        
        # 1. 偵測單位與稱謂屬性
        is_guard_unit = False
        for r in range(5):
            rt = "".join([str(x) for x in df.iloc[r].values])
            m = re.search(r'([\u4e00-\u9fa5]+(分局|派出所|分隊|局|隊))', rt)
            if m: 
                res['unit_name'] = m.group(1)
                if "警備隊" in res['unit_name']: is_guard_unit = True
                break
        res['term'] = "該隊" if is_guard_unit or "分隊" in res['unit_name'] else "該所"
        
        # 2. 人員雷達 (支援 隊長/副隊長 職稱)
        all_cells = df.astype(str).values.flatten()
        p = r'([A-Z0-9]{1,2})\s*(所長|副所長|隊長|副隊長|分隊長|小隊長|巡官|巡佐|警員|警員兼副所長|實習)[\s\n]*([\u4e00-\u9fa5]{2,4})'
        f_map = {}
        for cell in all_cells:
            m = re.search(p, cell)
            if m: f_map[d_normalize_code(m.group(1))] = f"{m.group(2)}{m.group(3)}"
        
        # 3. 時間座標偵測
        t_cols = {}
        for r in range(10): 
            for c in range(len(df.columns)):
                sh, eh = d_parse_time(df.iloc[r, c])
                if sh is not None: t_cols[c] = (sh, eh)
            if len(t_cols) > 5: break
            
        target_col = -1
        adj_h = hour if hour >= 6 else hour + 24
        for c, (sh, eh) in t_cols.items():
            s_time, e_time = sh, (eh if eh > sh else eh + 24)
            if s_time <= adj_h < e_time:
                target_col = c; break

        if target_col != -1:
            # 偵測值班人員 (避開幹部代號)
            for r in range(min(20, len(df))):
                row_h = "".join(df.iloc[r, :5])
                if "值" in row_h:
                    val = str(df.iloc[r, target_col])
                    mc = re.search(r'[A-Z0-9]{1,2}', val)
                    if mc:
                        code = d_normalize_code(mc.group(0))
                        if code not in ["A", "B", "C"]: # 值班通常不是 A/B/C
                            res['v_name'] = f_map.get(code, f"警員{code}")
                            break
            
            # 4. 幹部動態 (高平與警備隊特化)
            if is_guard_unit:
                titles = {"A": "隊長", "B": "副隊長", "C": "幹部"}
            else:
                titles = {"A": "所長", "B": "副所長", "C": "幹部"}
            
            c_notes = []
            for code in ["A", "B", "C"]:
                if code not in f_map and code != "A": continue # 警備隊人數少，若沒定義則跳過
                fname = f_map.get(code, titles[code])
                is_off = False
                
                # 底部休假區掃描
                for r in range(max(0, len(df)-10), len(df)):
                    if code in str(df.iloc[r, :]).upper() and any(k in "".join(df.iloc[r, :3]) for k in ["休","輪","假"]):
                        is_off = True; break
                
                d_names = set()
                # 垂直勤務掃描 (修正高平副所長偵測)
                for r in range(len(df)):
                    cell_val = str(df.iloc[r, target_col])
                    if code in re.findall(r'[A-Z0-9]{1,2}', cell_val):
                        is_off = False 
                        row_n = "".join(df.iloc[r, :5])
                        kw_map = {"巡":"巡邏", "守":"守望", "臨":"臨檢", "交":"交整", "路":"路檢", "督":"督勤", "備":"備勤"}
                        for k, kn in kw_map.items():
                            if k in row_n or k in cell_val: d_names.add(kn)
                
                if is_off: c_notes.append(f"{fname}休假")
                else:
                    if d_names:
                        sh, eh = t_cols[target_col]
                        c_notes.append(f"{fname}在所督勤，編排{sh:02d}-{eh if eh!=0 else 24:02d}時段{'、'.join(sorted(d_names))}勤務")
                    else: c_notes.append(f"{fname}在所督勤")
            res['cadre_status'] = "；".join(c_notes) + "。"
    except: pass
    return res

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
# 3. 主 UI 介面
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
        st.info(f"📅 監錄({d_5.strftime('%m%d')}-{d_e.strftime('%m%d')}) / 勤教({d_3.strftime('%m%d')}-{d_e.strftime('%m%d')})")

    u_tabs = st.tabs([f"🏢 單位 {i+1}" for i in range(num_units)] + ["📄 總匯整報告"])

    for i in range(num_units):
        with u_tabs[i]:
            u_time = st.time_input("抵達時間", datetime.now().time(), key=f"ut_{i}")
            col_f1, col_f2 = st.columns(2)
            with col_f1: u_duty = st.file_uploader(f"單位 {i+1} 勤務表", type=['xlsx'], key=f"ud_{i}")
            with col_f2: u_eq = st.file_uploader(f"單位 {i+1} 交接簿", type=['xlsx'], key=f"ue_{i}")

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
                st.success(f"✅ {uname} 解析完成")
                st.text_area("預覽報告", final_text, height=350, key=f"preview_{i}")

    with u_tabs[-1]:
        reports_list = [st.session_state.unit_reports[k] for k in sorted(st.session_state.unit_reports.keys()) if k < num_units]
        if reports_list:
            full_text = ("\n\n" + "─" * 40 + "\n\n").join(reports_list)
            st.text_area("📄 總匯整結果", full_text, height=600)
            target_mail = st.text_input("收件信箱", "mbmmiqamhnd@gmail.com")
            if st.button("🚀 寄送至 Gmail"):
                if send_gmail(f"督導報告匯整_{insp_date.strftime('%Y%m%d')}", full_text, target_mail):
                    st.success("郵件寄送成功！")
