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
# 2. 超進化全天候解析引擎 (警備隊拘留室加強版)
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
    if any(x in val_str for x in ["年", "月", "日", "號"]): return None, None
    m = re.search(r'(?<!\d)(\d{1,2})[:：\-\s~～/|]+(\d{1,2})(?!\d)', val_str)
    if m:
        sh, eh = int(m.group(1)), int(m.group(2))
        if 0 <= sh <= 24 and 0 <= eh <= 30: return sh, eh
    return None, None

def d_extract_duty(d_file, hour):
    res = {
        'v_name': '解析失敗', 
        'detention_name': None, # 拘留室人員
        'cadre_status': '無幹部資料', 
        'unit_name': '未偵測單位', 
        'term': '該所', 
        'loc_term': '所', 
        'has_skyline': True
    }
    try:
        df = pd.read_excel(d_file, header=None, dtype=str).fillna("")

        # A. 偵測單位
        unit_full = ""
        for r in range(5):
            rt = "".join([str(x) for x in df.iloc[r].values]).replace(" ", "")
            m = re.search(r'([\u4e00-\u9fa5]+(派出所|分駐所|警備隊|分隊|中隊|大隊|隊))', rt)
            if m: unit_full = m.group(1); res['unit_name'] = unit_full; break

        is_guard_unit = "警備隊" in unit_full
        if "分隊" in unit_full:
            res['term'] = "該分隊"; res['has_skyline'] = False
        elif "隊" in unit_full:
            res['term'] = "該隊"; res['has_skyline'] = False
        else:
            res['term'] = "該所"; res['has_skyline'] = True
        res['loc_term'] = res['term'][1:]

        # B. 人員雷達
        f_map = {}
        all_text = " ".join(df.astype(str).values.flatten())
        p = r'([A-Z0-9]{1,2})\s*(所長|副所長|隊長|副隊長|分隊長|小隊長|警務佐|巡官|巡佐|警員|實習)[\s\n,]*([\u4e00-\u9fa5]{2,4})'
        for m in re.findall(p, all_text):
            name = m[2].strip()
            if any(x in name for x in ["姓名", "職稱", "代號", "人員"]): continue
            f_map[d_normalize_code(m[0])] = f"{m[1]}{name}"

        # C. 時間座標鎖定
        t_cols, tr_idx = {}, -1
        for r in range(12):
            tmp = {}
            for c in range(len(df.columns)):
                sh, eh = d_parse_time(df.iloc[r, c])
                if sh is not None: tmp[c] = (sh, eh)
            if len(tmp) > len(t_cols): t_cols = tmp; tr_idx = r

        adj_h = hour if hour >= 6 else hour + 24
        target_col = -1
        for c, (sh, eh) in t_cols.items():
            s, e = (sh if sh >= 6 else sh + 24), (eh if eh > sh else eh + 24)
            if s <= adj_h < e: target_col = c; break

        if target_col != -1 and tr_idx != -1:
            footer_idx = len(df)
            for r in range(len(df)-1, max(0, len(df)-40), -1):
                row_all = "".join(df.iloc[r, :]).replace(" ", "")
                if any(x in row_all for x in ["輪休", "主管簽章", "備註", "合計", "人數"]):
                    footer_idx = r; break

            # 1. 🌟 值班與拘留室偵測 (僅警備隊觸發拘留室搜尋)
            v_found = False
            for r in range(tr_idx + 1, min(footer_idx, len(df))):
                cell_a = str(df.iloc[r, 0]).strip()
                row_head = "".join(df.iloc[r, :target_col+1])
                
                # A. 偵測值班人員
                if "值班" in cell_a or any(x in row_head for x in ["值", "班"]):
                    if not v_found: # 確保只抓第一列
                        cell_val = str(df.iloc[r, target_col]).strip()
                        # 警備隊規則：僅看上列
                        t_val = cell_val.split('\n')[0] if is_guard_unit and '\n' in cell_val else cell_val
                        mc = re.search(r'[A-Z0-9]{1,2}', t_val)
                        if mc:
                            res['v_name'] = f_map.get(d_normalize_code(mc.group(0)), f"警員({mc.group(0)})")
                        else:
                            res['v_name'] = "該時段無值班人員"
                        v_found = True
                
                # B. 🌟 偵測拘留室人員 (僅限警備隊)
                if is_guard_unit and "拘留" in cell_a:
                    d_val = str(df.iloc[r, target_col]).strip()
                    md = re.search(r'[A-Z0-9]{1,2}', d_val)
                    if md:
                        res['detention_name'] = f_map.get(d_normalize_code(md.group(0)), f"警員({md.group(0)})")

            if not v_found: res['v_name'] = "該時段無值班人員"

            # 2. 幹部動態 (排序與合併)
            target_titles = ["所長", "副所長", "隊長", "副隊長", "分隊長", "小隊長", "警務佐"]
            def cadre_rank(code):
                title = f_map[code]
                if any(x in title for x in ["所長", "隊長", "分隊長"]) and "副" not in title: return 0
                if "副" in title: return 1
                return 2
            target_codes = sorted([c for c, i in f_map.items() if any(t in i for t in target_titles)], key=cadre_rank)

            c_notes = []
            for code in target_codes:
                fname = f_map.get(code, "幹部")
                all_duties = []
                is_off = False
                for c_idx, (sh, eh) in t_cols.items():
                    for r in range(tr_idx + 1, footer_idx):
                        if code in [d_normalize_code(x) for x in re.findall(r'[A-Z0-9]{1,2}', str(df.iloc[r, c_idx]))]:
                            duty_parts = []
                            for scan_c in range(0, c_idx):
                                txt = str(df.iloc[r, scan_c]).strip()
                                txt = re.sub(r'^[0-9一二三四五六七八九十、\s\.\、]+', '', txt)
                                if txt and len(txt) >= 2 and not any(x in txt for x in ["代號", "職稱", "姓名", "人員", "所長", "隊長", "警員"]):
                                    duty_parts.append(txt)
                            duty_name = duty_parts[0] if duty_parts else "勤務"
                            if "值班" in duty_name and len(duty_parts) > 1: duty_name = f"{duty_name}({duty_parts[1]})"
                            if any(k in duty_name for k in ["休", "輪", "假", "補", "外宿"]): is_off = True; continue
                            all_duties.append({'s': sh if sh >= 6 else sh + 24, 'e': eh if eh > sh else eh + 24, 'name': duty_name, 'sh_orig': sh, 'eh_orig': eh})

                if is_off: c_notes.append(f"{fname}休假")
                elif all_duties:
                    all_duties.sort(key=lambda x: x['s'])
                    merged = []
                    for d in all_duties:
                        if not merged: merged.append(d)
                        else:
                            last = merged[-1]
                            if d['s'] == last['e'] and d['name'] == last['name']:
                                last['e'] = d['e']; last['eh_orig'] = d['eh_orig']
                            else: merged.append(d)
                    summary = [f"{m['sh_orig']:02d}-{(24 if m['eh_orig'] in [0, 24] else m['eh_orig']):02d}{m['name']}" for m in merged]
                    c_notes.append(f"{fname}在{res['loc_term']}督勤，編排" + "、".join(summary) + "勤務")
                else: c_notes.append(f"{fname}在{res['loc_term']}督勤")
            res['cadre_status'] = "；".join(c_notes) + "。"
    except Exception as e: res['cadre_status'] = f"解析中斷：{str(e)}"
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
                if "手槍" in str(df.iloc[r, c]): col_map["gun"] = c
                if "子彈" in str(df.iloc[r, c]): col_map["bullet"] = c
        sub = df.iloc[:35]; sub_s = df_s.iloc[:35]
        def get_v(kw, k):
            rows = sub[sub_s.iloc[:, 1].str.contains(kw, na=False)]
            return d_safe_int(rows.iloc[-1, col_map[k]]) if not rows.empty else 0
        return {"gi": get_v("在", "gun"), "go": get_v("出", "gun"),
                "bi": get_v("在", "bullet"), "bo": get_v("出", "bullet"),
                "ri": get_v("在", "radio"), "ro": get_v("出", "radio"),
                "vi": get_v("在", "vest"), "vo": get_v("出", "vest")}
    except: return None

# ==========================================
# 4. 主 UI
# ==========================================
st.header("📋 勤務督導報告自動生成系統")
insp_date = st.date_input("選擇督導日期", datetime.now(), key="insp_d")
num_units = st.number_input("待督導單位數量", 1, 8, 3, key="num_u")
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
            t = dr['term']; loc_term = dr['loc_term']
            d_e = insp_date - timedelta(days=1)
            d_5, d_3 = (insp_date - timedelta(days=5)), (insp_date - timedelta(days=3))
            
            # 🌟 第 1 項：值班、拘留室綜合判定
            duty_desc = ""
            if dr['v_name'] == "該時段無值班人員":
                duty_desc = "該時段無值班人員"
            else:
                duty_desc = f"值班{dr['v_name']}"
            
            # 警備隊額外增加拘留室人員文字
            detention_text = ""
            if "警備隊" in dr['unit_name']:
                if dr['detention_name']:
                    detention_text = f"、拘留室值班人員{dr['detention_name']}"
                else:
                    detention_text = "、拘留室目前無人拘留"
            
            line_1 = f"{u_time.strftime('%H%M')}，{t}{duty_desc}{detention_text}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。"
            
            # 第 2 項：監錄系統判定
            system_desc = "駐地監錄設備及天羅地網系統" if dr['has_skyline'] else "駐地監錄設備"
            line_2 = f"{t}{system_desc}均運作正常，無故障，{d_5.strftime('%m月%d日')}至{d_e.strftime('%m月%d日')}有逐日檢測2次以上紀錄。"
            
            lns = [
                line_1,
                line_2,
                f"{t}{d_3.strftime('%m月%d日')}至{d_e.strftime('%m月%d日')}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。",
                f"{t}環境內務擺設整齊清潔，符合規定。",
                f"{t}手槍出勤 {er['go'] if er else 0} 把、在{loc_term} {er['gi'] if er else 0} 把，子彈出勤 {er['bo'] if er else 0} 顆、在{loc_term} {er['bi'] if er else 0} 顆，無線電出勤 {er['ro'] if er else 0} 臺、在{loc_term} {er['ri'] if er else 0} 臺；防彈背心出勤 {er['vo'] if er else 0} 件、在{loc_term} {er['vi'] if er else 0} 件，幹部對械彈每日檢查管制良好，符合規定。",
                f"本日{dr['cadre_status']}",
                f"{t}酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。"
            ]
            final_text = "\n".join([f"{idx+1}、{line}" for idx, line in enumerate(lns)])
            st.session_state.unit_reports[i] = f"【{dr['unit_name']} 督導報告】\n{final_text}"
            st.success(f"✅ {dr['unit_name']} 解析完成")
            st.text_area("預覽報告", final_text, height=350, key=f"preview_{i}")

with u_tabs[-1]:
    reports_list = [st.session_state.unit_reports[k] for k in sorted(st.session_state.unit_reports.keys()) if k < num_units]
    if reports_list:
        full_text = ("\n\n" + "─" * 40 + "\n\n").join(reports_list)
        st.subheader("📋 匯整結果")
        st.text_area("匯整文本", full_text, height=600)
        target_mail = st.text_input("收件信箱", "mbmmiqamhnd@gmail.com")
        if st.button("🚀 立即寄送郵件"):
            if send_gmail(f"勤務督導報告匯整_{insp_date.strftime('%Y%m%d')}", full_text, target_mail):
                st.success(f"✅ 郵件發送成功")
    else: st.warning("請先上傳檔案。")
