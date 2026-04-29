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
# 2. 核心工具函式 (v2)
# ==========================================
def safe_int(val):
    try:
        return int(float(str(val).split('.')[0].replace(',', '')))
    except:
        return 0

def d_normalize_code(c):
    c_str = str(c).strip().upper()
    c_str = c_str.translate(str.maketrans('０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ', '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'))
    c_str = c_str.replace(".0", "")
    return str(int(c_str)) if c_str.isdigit() else c_str

def parse_time_header(cell):
    """'06\n|\n07' → (6, 7)"""
    nums = re.findall(r'\d+', str(cell))
    if len(nums) >= 2:
        h1, h2 = int(nums[0]), int(nums[1])
        if 0 <= h1 <= 24 and 0 <= h2 <= 24:
            return h1, h2
    return None, None

def adj(h):
    """將小時換算為跨日比較用的連續數（夜班 0-5 → 24-29）"""
    return h if h >= 6 else h + 24

def build_fmap(df):
    """從勤務表底部對照區解析 代號 → 職稱姓名"""
    fmap = {}
    for r in range(len(df)):
        if '代號' in str(df.iloc[r, 0]) and '職稱' in str(df.iloc[r, 0]):
            for rr in range(r, min(r + 20, len(df))):
                if '代號' not in str(df.iloc[rr, 0]):
                    continue
                c = 1
                while c < len(df.columns) - 1:
                    code = str(df.iloc[rr, c]).strip()
                    name = str(df.iloc[rr, c + 1]).strip() if c + 1 < len(df.columns) else ''
                    name = re.sub(r'\s+\d{1,2}$', '', name).strip()
                    if (code and name and code not in ('', 'nan') and name not in ('', 'nan') and
                            re.match(r'^[A-Za-z0-9甲乙丙丁]{1,3}$', code)):
                        fmap[code] = name
                    c += 6
    return fmap

def find_target_col(df, hour):
    """在 Row2 時段列找到對應 hour 的欄索引"""
    TIME_ROW = 2
    t_cols = {}
    for c in range(13, len(df.columns)):
        h1, h2 = parse_time_header(df.iloc[TIME_ROW, c])
        if h1 is not None:
            t_cols[c] = (h1, h2)

    adj_h = adj(hour)
    for c, (sh, eh) in sorted(t_cols.items()):
        s = adj(sh)
        e = adj(eh) if eh != sh else adj(sh) + 1
        if eh == 0:
            e = 24
        if s <= adj_h < e:
            return c, t_cols
    return -1, t_cols

_SKIP_DUTY_NAMES = {'勤務\n人員\n代號\n職稱\n姓名', '代號', '職稱', '姓名', '勤務備註', '員警', '時段', '項目'}

def _clean_duty_name(raw):
    """🌟 清洗掉不應作為勤務名稱的字串，並移除括號說明與贅字"""
    raw = str(raw).strip()
    
    # 遇到標題殘影直接回傳 None 跳過
    if raw in _SKIP_DUTY_NAMES or ('代號' in raw and '職稱' in raw):
        return None
        
    # 斬斷括號(半形或全形)及其包含的所有內容，例如 "內部管理(監督員" -> "內部管理"
    raw = re.sub(r'[\(（].*$', '', raw).strip()
    
    # 若字尾帶有「勤務」則移除，避免最後組裝變成「勤務勤務」
    if raw.endswith('勤務') and len(raw) > 2:
        raw = raw[:-2].strip()
        
    if not raw:
        return '勤務'
        
    if len(raw) > 15:
        raw = raw[:15]
    return raw

# ==========================================
# 3. 勤務表解析 (v2)
# ==========================================
def extract_duty_v2(d_file, hour):
    res = {
        'v_name': '解析失敗', 'detention_name': None,
        'cadre_status': '無幹部資料', 'unit_name': '未偵測單位',
        'term': '該所', 'loc_term': '所', 'has_skyline': True, 'is_guard_unit': False
    }
    try:
        df = pd.read_excel(d_file, header=None, dtype=str).fillna('')

        # ── 1. 單位名稱 ──
        for r in range(3):
            rt = str(df.iloc[r, 0]).replace(' ', '')
            m = re.search(r'([\u4e00-\u9fa5]+(派出所|分駐所|警備隊|分隊|中隊|大隊))', rt)
            if m:
                unit_full = m.group(1)
                res['unit_name'] = unit_full
                res['is_guard_unit'] = '警備隊' in unit_full
                if '分隊' in unit_full:
                    res['term'] = '該分隊'; res['has_skyline'] = False
                elif '隊' in unit_full:
                    res['term'] = '該隊'; res['has_skyline'] = False
                else:
                    res['term'] = '該所'; res['has_skyline'] = True
                res['loc_term'] = res['term'][1:]
                break

        # ── 2. 人員對照表 ──
        fmap = build_fmap(df)

        # ── 3. 時段欄 ──
        target_col, t_cols = find_target_col(df, hour)
        if target_col == -1:
            res['v_name'] = '找不到對應時段欄'
            return res

        # ── 4. 值班與拘留室人員 ──
        v_found = False
        for r in range(3, 30):
            if r >= len(df): break
            col0 = str(df.iloc[r, 0]).strip()
            col1 = str(df.iloc[r, 1]).strip()
            
            if col0 == '值班':
                cell = str(df.iloc[r, target_col]).strip()
                if not cell or cell == 'nan' or len(cell) > 10: continue
                codes = re.findall(r'[A-Z甲乙丙丁][0-9]?|[0-9]{2}', cell)
                valid_codes = [c for c in codes if re.match(r'^[A-Z0-9甲乙丙丁]{1,3}$', c)]
                if valid_codes:
                    code = valid_codes[0]
                    res['v_name'] = fmap.get(code, f'警員({code})')
                    v_found = True
                    break
            
            if '拘留' in col1 and res['is_guard_unit']:
                cell = str(df.iloc[r, target_col]).strip()
                codes = re.findall(r'[A-Z甲乙丙丁][0-9]?|[0-9]{2}', cell)
                valid_codes = [c for c in codes if re.match(r'^[A-Z0-9甲乙丙丁]{1,3}$', c)]
                if valid_codes:
                    res['detention_name'] = fmap.get(valid_codes[0], f'警員({valid_codes[0]})')

        if not v_found:
            res['v_name'] = '該時段無值班人員'

        # ── 5. 幹部動態 ──
        target_titles = ['所長', '副所長', '隊長', '副隊長', '分隊長', '小隊長', '警務佐']
        def rank(code):
            t = fmap.get(code, '')
            if any(x in t for x in ['所長', '隊長', '分隊長']) and '副' not in t: return 0
            if '副' in t: return 1
            return 2
        cadre_codes = sorted([c for c in fmap if any(t in fmap[c] for t in target_titles)], key=rank)

        footer_row = len(df)
        for r in range(3, len(df)):
            col0 = str(df.iloc[r, 0]).strip()
            if any(x in col0 for x in ['請假人員', '重要', '勤務\n備註', '原定勤務', '員警']):
                footer_row = r; break

        c_notes = []
        for code in cadre_codes:
            fname_full = fmap[code]
            d_list = []
            is_off = False

            for r in range(3, footer_row):
                col0, col1 = str(df.iloc[r, 0]).strip(), str(df.iloc[r, 1]).strip()
                if not col0: continue

                for c, (sh, eh) in t_cols.items():
                    cell_val = str(df.iloc[r, c]).strip()
                    if not cell_val or cell_val == 'nan': continue
                    
                    raw_codes = re.findall(r'[A-Z甲乙丙丁][0-9]?|[0-9]{2}', cell_val)
                    if code not in [x for x in raw_codes if re.match(r'^[A-Z0-9甲乙丙丁]{1,3}$', x)]:
                        continue

                    # 直接傳入 col1 交由 cleaner 判斷與截斷
                    duty_name = col1 if col1 and len(col1) >= 2 else col0
                    
                    if any(k in duty_name for k in ['休', '輪', '假', '補', '外宿']):
                        is_off = True; continue

                    duty_name = _clean_duty_name(duty_name)
                    if duty_name is None: continue

                    s = adj(sh)
                    e = adj(eh) if eh != sh else adj(sh) + 1
                    if eh == 0: e = 24
                    d_list.append({'sh': sh, 'eh': eh, 's': s, 'e': e, 'n': duty_name})

            if not d_list:
                for r in range(3, len(df)):
                    col0, col1 = str(df.iloc[r, 0]).strip(), str(df.iloc[r, 1]).strip()
                    if any(k in col0 + col1 for k in ['輪休', '慰休', '公假', '補休', '事假', '病假']):
                        for c in range(13, len(df.columns)):
                            if code in re.findall(r'[A-Z甲乙丙丁][0-9]?|[0-9]{2}', str(df.iloc[r, c])):
                                is_off = True; break

            # 勤務優先判定：有勤務就絕不報休假
            if d_list:
                d_list.sort(key=lambda x: x['s'])
                merged = []
                for d in d_list:
                    if merged and d['s'] == merged[-1]['e'] and d['n'] == merged[-1]['n']:
                        merged[-1]['e'] = d['e']
                        merged[-1]['eh'] = d['eh']
                    else:
                        merged.append(dict(d))
                parts = [f"{m['sh']:02d}-{(24 if m['eh'] == 0 else m['eh']):02d}{m['n']}" for m in merged]
                c_notes.append(f"{fname_full}在{res['loc_term']}督勤，編排{'、'.join(parts)}勤務")
            elif is_off:
                c_notes.append(f'{fname_full}休假')
            else:
                c_notes.append(f'{fname_full}在{res["loc_term"]}督勤')

        res['cadre_status'] = '；'.join(c_notes) + '。' if c_notes else '無幹部資料。'

    except Exception as e:
        res['cadre_status'] = f'解析中斷：{e}'
    return res

# ==========================================
# 4. 交接簿解析 (v2)
# ==========================================
def extract_equip_v2(e_file):
    try:
        df = pd.read_excel(e_file, header=None).fillna('')
        header_row, col_map = 2, {}

        for c in range(len(df.columns)):
            v = str(df.iloc[header_row, c]).replace('\n', '').replace(' ', '')
            if v == '手槍': col_map['gun'] = c
            if '子彈' in v:
                prev = str(df.iloc[header_row, c - 1]).replace('\n', '') if c > 0 else ''
                if '手槍' in prev: col_map['bullet'] = c
            if '無線電' in v: col_map['radio'] = c
            if '防彈背心' in v: col_map['vest'] = c

        for c in range(2, len(df.columns)):
            v = str(df.iloc[header_row, c]).replace('\n', '')
            if 'gun' not in col_map and '手槍' in v and '子彈' not in v: col_map['gun'] = c
            if 'bullet' not in col_map and '子彈' in v and '手槍' in str(df.iloc[header_row, c - 1]).replace('\n', ''): col_map['bullet'] = c
            if 'radio' not in col_map and '無線電' in v: col_map['radio'] = c
            if 'vest' not in col_map and '防彈背心' in v: col_map['vest'] = c

        last_zi = last_zo = -1
        for r in range(3, len(df)):
            lbl = str(df.iloc[r, 1]).replace('\n', '').strip()
            if '在' in lbl and any(x in lbl for x in ['所', '隊']): last_zi = r
            if '出' in lbl and '勤' in lbl: last_zo = r

        def get(row, key):
            if row < 0 or key not in col_map: return 0
            return safe_int(df.iloc[row, col_map[key]])

        return {
            'gi': get(last_zi, 'gun'),    'go': get(last_zo, 'gun'),
            'bi': get(last_zi, 'bullet'), 'bo': get(last_zo, 'bullet'),
            'ri': get(last_zi, 'radio'),  'ro': get(last_zo, 'radio'),
            'vi': get(last_zi, 'vest'),   'vo': get(last_zo, 'vest'),
        }
    except Exception:
        return None

# ==========================================
# 5. 主介面 UI
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
            dr = extract_duty_v2(u_duty, u_time.hour)
            er = extract_equip_v2(u_eq)
            
            if not er:
                er = {'gi':0, 'go':0, 'bi':0, 'bo':0, 'ri':0, 'ro':0, 'vi':0, 'vo':0}
                
            t, loc = dr['term'], dr['loc_term']
            d_e = insp_date - timedelta(days=1)
            d_5, d_3 = (insp_date - timedelta(days=5)), (insp_date - timedelta(days=3))
            
            if dr['v_name'] == "該時段無值班人員":
                line_1 = f"{u_time.strftime('%H%M')}，{t}該時段無值班人員。"
            else:
                line_1 = f"{u_time.strftime('%H%M')}，{t}值班{dr['v_name']}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。"
            
            lns = [
                line_1,
                f"{t}{'駐地監錄設備及天羅地網系統' if dr['has_skyline'] else '駐地監錄設備'}均運作正常，無故障，{d_5.strftime('%m月%d日')}至{d_e.strftime('%m月%d日')}有逐日檢測2次以上紀錄。",
                f"{t}{d_3.strftime('%m月%d日')}至{d_e.strftime('%m月%d日')}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。",
                f"{t}環境內務擺設整齊清潔，符合規定。",
                f"{t}手槍出勤 {er['go']} 把、在{loc} {er['gi']} 把，子彈出勤 {er['bo']} 顆、在{loc} {er['bi']} 顆，無線電出勤 {er['ro']} 臺、在{loc} {er['ri']} 臺；防彈背心出勤 {er['vo']} 件、在{loc} {er['vi']} 件，幹部對械彈每日檢查管制良好，符合規定。",
                f"本日{dr['cadre_status']}",
                f"{t}酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。"
            ]
            
            if dr['is_guard_unit']:
                lns.append(f"拘留室值班警員{dr['detention_name']}，對人犯監控良好，無異常狀況發生。" if dr['detention_name'] else "拘留室目前無人犯。")
            
            final_text = "\n".join([f"{idx+1}、{line}" for idx, line in enumerate(lns)])
            st.session_state.unit_reports[i] = f"【{dr['unit_name']} 督導報告】\n{final_text}"
            
            if "中斷" in dr['cadre_status'] or "失敗" in dr['v_name']:
                st.error(f"⚠️ {dr['unit_name']} 解析可能不完全：{dr['cadre_status']}")
            else:
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
    else:
        st.warning("請先上傳檔案。")
