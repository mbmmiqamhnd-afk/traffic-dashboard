import streamlit as st
import pandas as pd
import io
import re
import gspread
import traceback
from datetime import datetime, timedelta

# ==========================================
# 0. 系統初始化與選單配置
# ==========================================
st.set_page_config(page_title="交通業務與督導整合引擎", page_icon="🚓", layout="wide")

try:
    from menu import show_sidebar
    show_sidebar()
except:
    pass

st.markdown("""
    <style>
    @font-face { font-family: 'Kaiu'; src: url('kaiu.ttf'); }
    .stTextArea textarea { font-family: 'Kaiu', "標楷體", sans-serif !important; font-size: 19px !important; line-height: 1.7 !important; color: #1c1c1c !important; }
    .stTabs [data-baseweb="tab-list"] button { font-size: 18px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 1. 督導功能解析引擎
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
        for r in range(min(10, len(df))):
            for c in range(len(df.columns)):
                v = str(df.iloc[r, c]).replace(" ", "").replace("　", "").replace("\n", "")
                if col_map["gun"] is None and ("手槍" in v or ("槍" in v and "手" in v)): col_map["gun"] = c
                if col_map["bullet"] is None and ("子彈" in v or ("彈" in v and "子" in v)): col_map["bullet"] = c
                if col_map["radio"] is None and "無線電" in v: col_map["radio"] = c
                if col_map["vest"] is None and ("背心" in v or "防彈衣" in v): col_map["vest"] = c
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
        return {"gi":get_v("在","gun"), "go":get_v("出","gun"), "gf":gf,
                "bi":get_v("在","bullet"), "bo":get_v("出","bullet"),
                "ri":get_v("在","radio"), "ro":get_v("出","radio"), "rf":rf,
                "vi":get_v("在","vest"), "vo":get_v("出","vest")}
    except: return None

def d_extract_duty(d_file, hour):
    res = {'v_name': '未偵測', 'cadre_status': '無幹部資料', 'unit_name': '未偵測單位', 'term': '該所'}
    try:
        df = pd.read_excel(d_file, header=None, dtype=str).fillna("")
        for r in range(5):
            rt = "".join([str(x) for x in df.iloc[r].values])
            m = re.search(r'([\u4e00-\u9fa5]+(分局|派出所|分隊|局))', rt)
            if m: res['unit_name'] = m.group(1); break
        
        # 🌟 核心稱謂偵測
        is_traffic_unit = "分隊" in res['unit_name']
        res['term'] = "該分隊" if is_traffic_unit else "該所"
        
        full = " ".join([str(x).strip() for x in df.values.flatten() if x])
        p = r'(?<![A-Za-z0-9])([A-Z]|[0-9]{1,2})\s*(所長|副所長|分隊長|小隊長|巡官|巡佐|警員|實習)\s*([\u4e00-\u9fa5]{2,4})'
        matches = re.findall(p, full)
        
        n_map, f_map = {}, {}
        for m in matches:
            code, title, name = d_normalize_code(m[0]), m[1].strip(), m[2]
            for t in ["所", "副", "巡", "警", "實", "員", "長", "隊"]:
                if name.endswith(t): name = name[:-1]
            if len(name) >= 2: 
                n_map[code] = name
                f_map[code] = f"{title}{name}"
                
        tr_idx, t_cols, t_col_idx = 2, {}, -1
        for r in range(5):
            tmp = {}
            for c in range(len(df.columns)):
                sh, eh = d_parse_time(df.iloc[r, c])
                if sh is not None: tmp[c] = (sh, eh)
            if len(tmp) > len(t_cols): tr_idx, t_cols = r, tmp
        for c, (sh, eh) in t_cols.items():
            ce, ch = (eh if eh > sh else eh + 24), (hour if hour >= 6 or (eh if eh > sh else eh + 24) <= 24 else hour + 24)
            if sh <= ch < ce: t_col_idx = c
            
        vr_idx = tr_idx + 1
        for r in range(tr_idx + 1, min(tr_idx + 4, len(df))):
            if "值" in str(df.iloc[r, 0]) + str(df.iloc[r, 1]): vr_idx = r; break
        if t_col_idx != -1:
            raw = str(df.iloc[vr_idx, t_col_idx]).strip()
            mc = re.search(r'[A-Za-z0-9]{1,2}', raw)
            if mc:
                code = d_normalize_code(mc.group(0))
                res['v_name'] = f_map.get(code, f"未建檔:{code}")
            else:
                for c, n in n_map.items():
                    if n in raw: res['v_name'] = f_map.get(c, n); break
                    
        c_notes = []
        default_titles = {"A": "分隊長" if is_traffic_unit else "所長", 
                          "B": "小隊長" if is_traffic_unit else "副所長", 
                          "C": "幹部"}

        for code in ["A", "B", "C"]:
            full_name = f_map.get(code, default_titles[code])
            found, is_off, p_slots, d_names = False, False, [], set()
            for r in range(vr_idx, len(df)):
                dt = str(df.iloc[r, 0]) + str(df.iloc[r, 1])
                is_l = any(k in dt for k in ["休", "假", "輪", "輸", "補", "外"])
                for c, (sh, eh) in t_cols.items():
                    if code in [d_normalize_code(x) for x in re.findall(r'[A-Za-z0-9]{1,2}', str(df.iloc[r, c]))]:
                        found = True
                        if is_l: is_off = True
                        else:
                            is_e = False
                            for k, kn in zip(["巡","守","臨","交","路"],["巡邏","守望","臨檢","交整","路檢"]):
                                if k in dt or k in str(df.iloc[r, c]): d_names.add(kn); is_e = True
                            if is_e: p_slots.append((sh, eh))
            if not found or is_off: 
                c_notes.append(f"{full_name}休假")
            else:
                if p_slots:
                    ms, me = min([s[0] for s in p_slots]), max([s[1] for s in p_slots])
                    e_str = "24" if me==24 or me==0 else f"{me%24:02d}"
                    c_notes.append(f"{full_name}在所督勤，編排{ms:02d}至{e_str}時段{'、'.join(sorted(list(d_names)))}勤務")
                else: 
                    c_notes.append(f"{full_name}在所督勤")
        res['cadre_status'] = "；".join(c_notes) + "。"
    except: res['v_name'] = "解析失敗"
    return res

# ==========================================
# 2. 原有的交通執法數據處理邏輯 (process_ 函數)
# ==========================================
# [請完整貼上您原本 process_tech_enforcement, process_overload... 等函數內容]

# ==========================================
# 3. 主介面分頁整合
# ==========================================
main_tabs = st.tabs(["📊 數據自動化處理", "📋 勤務督導報告"])

with main_tabs[0]:
    st.header("📈 執法數據全自動批次中心")
    # (此處接您原本的 file_uploader 與業務程式碼)

with main_tabs[1]:
    st.header("📋 勤務督導報告自動生成")
    c_s1, c_s2 = st.columns(2)
    with c_s1:
        insp_date = st.date_input("選擇督導日期", datetime.now(), key="insp_d")
        num_units = st.number_input("單位數量", 1, 8, 3, key="num_u")
    with c_s2:
        d_end = insp_date - timedelta(days=1)
        d_s5, d_s3 = insp_date - timedelta(days=5), insp_date - timedelta(days=3)
        d_end_s, d_s5_s, d_s3_s = d_end.strftime('%m月%d日'), d_s5.strftime('%m月%d日'), d_s3.strftime('%m月%d日')
        st.info(f"📅 監錄區間：{d_s5_s}至{d_end_s} / 勤教：{d_s3_s}至{d_end_s}")

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
                uname = dr['unit_name']
                t = dr['term'] # 🌟 取得「該分隊」或「該所」
                
                lns = []
                # 1. 值班 (修正：使用變數 t)
                lns.append(f"{u_time.strftime('%H%M')}，{t}值班{dr['v_name']}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。")
                # 2. 監錄 (修正：使用變數 t)
                lns.append(f"{t}駐地監錄設備及天羅地網系統均運作正常，無故障，{d_s5_s}至{d_end_s}有逐日檢測2次以上紀錄。")
                # 3. 勤教 (修正：使用變數 t)
                lns.append(f"{t}{d_s3_s}至{d_end_s}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。")
                # 4. 內務 (修正：使用變數 t)
                lns.append(f"{t}環境內務擺設整齊清潔，符合規定。")
                # 5. 裝備 (修正：使用變數 t)
                eq = er if er else {"gi":0,"go":0,"gf":0,"bi":0,"bo":0,"ri":0,"ro":0,"rf":0,"vi":0,"vo":0}
                fix = f"（另有槍枝 {eq['gf']} 把、無線電 {eq['rf']} 臺送修中）" if (eq['gf']+eq['rf']) > 0 else ""
                lns.append(f"{t}手槍出勤 {eq['go']} 把、在所 {eq['gi']} 把，子彈出勤 {eq['bo']} 顆、在所 {eq['bi']} 顆，無線電出勤 {eq['ro']} 臺、在所 {eq['ri']} 臺；防彈背心出勤 {eq['vo']} 件、在所 {eq['vi']} 件，幹部對械彈每日檢查管制良好，符合規定{fix}。")
                # 6. 幹部
                lns.append(f"本日{dr['cadre_status']}")
                # 7. 酒測 (修正：使用變數 t)
                lns.append(f"{t}酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。")

                final = "\n".join([f"{idx+1}、{line}" for idx, line in enumerate(lns)])
                all_final_reports.append(f"【{uname} 督導報告】\n{final}")
                st.success(f"✅ {uname} 報告已完成")
                st.text_area("預覽", final, height=350, key=f"txt_{i}")

    with u_tabs[-1]:
        if all_final_reports:
            st.text_area("📄 總匯整結果", "\n\n--------------------\n\n".join(all_final_reports), height=600)
