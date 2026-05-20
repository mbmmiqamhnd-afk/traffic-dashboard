import streamlit as st
import pandas as pd
import io
import sys
import os
import re
import json
import traceback
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime, timedelta
from pdf2image import convert_from_bytes

try:
    import google.generativeai as genai
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

# ==========================================
# 0. 系統初始化與路徑設定
# ==========================================
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        pass

st.set_page_config(page_title="勤務督導報告自動生成系統", page_icon="🚓", layout="wide")

if "unit_reports" not in st.session_state:
    st.session_state.unit_reports = {}

# ==========================================
# 1. Gemini API 初始化與設定
# ==========================================
model = None
try:
    api_key = st.secrets["api"]["GOOGLE_API_KEY"]
    if GENAI_AVAILABLE:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.5-flash')
except Exception as e:
    st.error(f"Gemini API 初始化失敗，請檢查 secrets 設定: {e}")

safety_settings = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
]

# ==========================================
# 2. 寄信功能
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
# 3. 核心工具函式
# ==========================================
def safe_int(val):
    try:
        return int(float(str(val).split('.')[0].replace(',', '')))
    except:
        return 0

def d_normalize_code(c):
    c_str = str(c).strip().upper()
    c_str = c_str.translate(str.maketrans(
        '０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ',
        '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'))
    c_str = c_str.replace(".0", "")
    return str(int(c_str)) if c_str.isdigit() else c_str

def parse_time_header(cell):
    nums = re.findall(r'\d+', str(cell))
    if len(nums) >= 2:
        h1, h2 = int(nums[0]), int(nums[1])
        if 0 <= h1 <= 24 and 0 <= h2 <= 24:
            return h1, h2
    return None, None

def adj(h):
    return h if h >= 6 else h + 24

def build_fmap(df):
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
    raw = str(raw).strip()
    if raw in _SKIP_DUTY_NAMES or ('代號' in raw and '職稱' in raw):
        return None
    raw = re.sub(r'[\(（].*$', '', raw).strip()
    if raw.endswith('勤務') and len(raw) > 2:
        raw = raw[:-2].strip()
    if not raw:
        return '勤務'
    if len(raw) > 15:
        raw = raw[:15]
    return raw

# ==========================================
# 4. 勤務表解析
# ==========================================
def extract_duty_v2(d_file, hour):
    res = {
        'v_name': '解析失敗', 'detention_name': None,
        'cadre_status': '無幹部資料', 'unit_name': '未偵測單位',
        'term': '該所', 'loc_term': '所', 'has_skyline': True,
        'is_guard_unit': False, 'roster': []
    }
    try:
        df = pd.read_excel(d_file, header=None, dtype=str).fillna('')

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

        fmap = build_fmap(df)
        res['roster'] = list(fmap.values())

        target_col, t_cols = find_target_col(df, hour)
        if target_col == -1:
            res['v_name'] = '找不到對應時段欄'
            return res

        v_found = False
        for r in range(3, 30):
            if r >= len(df):
                break
            col0 = str(df.iloc[r, 0]).strip()
            col1 = str(df.iloc[r, 1]).strip()

            if res['is_guard_unit'] and ('拘留' in col0 or '拘留' in col1):
                if not res['detention_name']:
                    cell = str(df.iloc[r, target_col]).strip()
                    codes = re.findall(r'[A-Z甲乙丙丁][0-9]?|[0-9]{2}', cell)
                    valid_codes = [c for c in codes if re.match(r'^[A-Z0-9甲乙丙丁]{1,3}$', c)]
                    if valid_codes:
                        res['detention_name'] = fmap.get(valid_codes[0], f'警員({valid_codes[0]})')

            if not v_found and ('值班' in col0 or '值班' in col1):
                if res['is_guard_unit']:
                    cell_raw = str(df.iloc[r, target_col])
                    cell = cell_raw.split('\n')[0].strip()
                    codes = re.findall(r'[A-Z甲乙丙丁][0-9]?|[0-9]{2}', cell)
                    valid_codes = [c for c in codes if re.match(r'^[A-Z0-9甲乙丙丁]{1,3}$', c)]
                    if valid_codes:
                        res['v_name'] = fmap.get(valid_codes[0], f'警員({valid_codes[0]})')
                    else:
                        res['v_name'] = '該時段無值班人員'
                    v_found = True
                else:
                    cell = str(df.iloc[r, target_col]).strip()
                    if cell and cell != 'nan' and len(cell) <= 10:
                        codes = re.findall(r'[A-Z甲乙丙丁][0-9]?|[0-9]{2}', cell)
                        valid_codes = [c for c in codes if re.match(r'^[A-Z0-9甲乙丙丁]{1,3}$', c)]
                        if valid_codes:
                            c = valid_codes[0]
                            base_name = fmap.get(c, f"({c})")
                            # 清理重複職稱或簡單格式化
                            if not any(title in base_name for title in ['警員', '巡佐', '隊長', '所長', '副所長']):
                                res['v_name'] = f"警員 {base_name}"
                            else:
                                res['v_name'] = base_name
                            v_found = True

        if not v_found:
            res['v_name'] = '該時段無值班人員'

        target_titles = ['所長', '副所長', '隊長', '副隊長', '分隊長', '小隊長', '警務佐']

        def rank(code):
            t = fmap.get(code, '')
            if any(x in t for x in ['所長', '隊長', '分隊長']) and '副' not in t:
                return 0
            if '副' in t:
                return 1
            return 2

        cadre_codes = sorted(
            [c for c in fmap if any(t in fmap[c] for t in target_titles)],
            key=rank
        )

        footer_row = len(df)
        for r in range(3, len(df)):
            col0 = str(df.iloc[r, 0]).strip()
            if any(x in col0 for x in ['請假人員', '重要', '勤務\n備註', '原定勤務', '員警']):
                footer_row = r
                break

        c_notes = []
        for code in cadre_codes:
            fname_full = fmap[code]
            d_list = []
            is_off = False

            for r in range(3, footer_row):
                col0, col1 = str(df.iloc[r, 0]).strip(), str(df.iloc[r, 1]).strip()
                if not col0:
                    continue

                for c, (sh, eh) in t_cols.items():
                    cell_val = str(df.iloc[r, c]).strip()
                    if not cell_val or cell_val == 'nan':
                        continue

                    raw_codes = re.findall(r'[A-Z甲乙丙丁][0-9]?|[0-9]{2}', cell_val)
                    if code not in [x for x in raw_codes if re.match(r'^[A-Z0-9甲乙丙丁]{1,3}$', x)]:
                        continue

                    duty_name = col1 if col1 and len(col1) >= 2 else col0
                    if any(k in duty_name for k in ['休', '輪', '假', '補', '外宿']):
                        is_off = True
                        continue

                    duty_name = _clean_duty_name(duty_name)
                    if duty_name is None:
                        continue

                    s = adj(sh)
                    e = adj(eh) if eh != sh else adj(sh) + 1
                    if eh == 0:
                        e = 24
                    d_list.append({'sh': sh, 'eh': eh, 's': s, 'e': e, 'n': duty_name})

            if not d_list:
                for r in range(3, len(df)):
                    col0, col1 = str(df.iloc[r, 0]).strip(), str(df.iloc[r, 1]).strip()
                    if any(k in col0 + col1 for k in ['輪休', '慰休', '公假', '補休', '事假', '病假']):
                        for c in range(13, len(df.columns)):
                            if code in re.findall(r'[A-Z甲乙丙丁][0-9]?|[0-9]{2}', str(df.iloc[r, c])):
                                is_off = True
                                break

            if d_list:
                grouped_duties = {}
                for d in d_list:
                    grouped_duties.setdefault(d['s'], []).append(d)

                filtered_d_list = []
                for s_time, items in grouped_duties.items():
                    if len(items) > 1:
                        non_internal = [x for x in items if '內部管理' not in x['n']]
                        if non_internal:
                            filtered_d_list.extend(non_internal)
                        else:
                            filtered_d_list.append(items[0])
                    else:
                        filtered_d_list.extend(items)

                filtered_d_list.sort(key=lambda x: x['s'])

                merged = []
                for d in filtered_d_list:
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
# 5. 交接簿解析
# ==========================================
def extract_equip_v2(e_file):
    try:
        df = pd.read_excel(e_file, header=None).fillna('')
        header_row, col_map = 2, {}

        for c in range(len(df.columns)):
            v = str(df.iloc[header_row, c]).replace('\n', '').replace(' ', '')
            if v == '手槍':
                col_map['gun'] = c
            if '子彈' in v:
                prev = str(df.iloc[header_row, c - 1]).replace('\n', '') if c > 0 else ''
                if '手槍' in prev:
                    col_map['bullet'] = c
            if '無線電' in v:
                col_map['radio'] = c
            if '防彈背心' in v:
                col_map['vest'] = c

        for c in range(2, len(df.columns)):
            v = str(df.iloc[header_row, c]).replace('\n', '')
            if 'gun' not in col_map and '手槍' in v and '子彈' not in v:
                col_map['gun'] = c
            if 'bullet' not in col_map and '子彈' in v and '手槍' in str(df.iloc[header_row, c - 1]).replace('\n', ''):
                col_map['bullet'] = c
            if 'radio' not in col_map and '無線電' in v:
                col_map['radio'] = c
            if 'vest' not in col_map and '防彈背心' in v:
                col_map['vest'] = c

        last_zi = last_zo = -1
        for r in range(3, len(df)):
            lbl = str(df.iloc[r, 1]).replace('\n', '').strip()
            if '在' in lbl and any(x in lbl for x in ['所', '隊']):
                last_zi = r
            if '出' in lbl and '勤' in lbl:
                last_zo = r

        def get(row, key):
            if row < 0 or key not in col_map:
                return 0
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
# 6. Gemini 2.5 Vision 刑案單辨識核心
# ==========================================
def parse_crime_pdf_gemini(pdf_file, roster: list, unit_idx: int) -> list:
    if model is None:
        st.error("Gemini 模型未初始化，無法辨識刑案單。")
        return []

    pdf_file.seek(0)
    images = convert_from_bytes(pdf_file.read(), dpi=150)
    results = []
    roster_str = "、".join(roster)

    prompt = "請提取：嫌疑人, 查獲時間, 查獲地點, 觸犯法條, 查獲員警(請完整提取「職稱+姓名」，例如「警員蕭漢祥」)。\n"
    prompt += "名冊供比對參考：" + roster_str + "。\n"
    prompt += "請嚴格回傳 JSON Array 格式，即使只有一筆資料也要放在陣列中，例如：\n"
    prompt += '[{"嫌疑人": "王大明", "查獲時間": "10時00分", "查獲地點": "某路段", "觸犯法條": "公共危險", "查獲員警": "警員李小華"}]'

    total_pages = len(images)
    for i, img in enumerate(images):
        try:
            st.info(f"單位 {unit_idx+1} 🚀 AI 正在辨識刑案單第 {i+1}/{total_pages} 頁...")
            response = model.generate_content([prompt, img], safety_settings=safety_settings)
            raw_text = response.text.strip()

            # 移除 markdown code fence
            lines = raw_text.splitlines()
            if lines and lines[0].startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            raw_text = "\n".join(lines).strip()

            if raw_text and raw_text not in ("[]", "{}"):
                parsed = json.loads(raw_text)
                if isinstance(parsed, list):
                    results.extend(parsed)
                elif isinstance(parsed, dict):
                    results.append(parsed)
        except json.JSONDecodeError:
            st.warning(f"第 {i+1} 頁 JSON 解析失敗，略過。原始回傳：{raw_text[:200]}")
        except Exception as e:
            st.warning(f"第 {i+1} 頁辨識失敗：{e}")

    return results

# ==========================================
# 7. 報告組合 (完全對齊您的目標格式)
# ==========================================
def build_report(duty_info: dict, equip: dict, crimes: list, time_str: str, sup_date: datetime) -> str:
    unit = duty_info.get('unit_name', '未知單位')
    term = duty_info.get('term', '該所')
    loc_term = duty_info.get('loc_term', '所')
    v_name = duty_info.get('v_name', '不明')
    cadre = duty_info.get('cadre_status', '無幹部資料。')
    has_skyline = duty_info.get('has_skyline', True)
    is_guard = duty_info.get('is_guard_unit', False)
    detention = duty_info.get('detention_name')

    # 時間格式處理
    d_e = (sup_date - timedelta(days=1)).strftime("%m月%d日")
    d_3 = (sup_date - timedelta(days=3)).strftime("%m月%d日")
    d_5 = (sup_date - timedelta(days=5)).strftime("%m月%d日")

    lines = []
    lines.append(f"【{unit} 督導報告】")
    
    idx = 1
    
    # 第 1 點：值班情形
    if "無值班人員" in v_name:
        lines.append(f"{idx}、{time_str}，{term}該時段無值班人員。")
    else:
        lines.append(f"{idx}、{time_str}，{term}值班{v_name}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。")
    idx += 1

    # 第 2 點：駐地監錄設備
    skyline_str = "及天羅地網系統" if has_skyline else ""
    lines.append(f"{idx}、{term}駐地監錄設備{skyline_str}均運作正常，無故障，{d_5}至{d_e}有逐日檢測2次以上紀錄。")
    idx += 1

    # 第 3 點：勤前教育
    lines.append(f"{idx}、{term}{d_3}至{d_e}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。")
    idx += 1

    # 第 4 點：環境內務
    lines.append(f"{idx}、{term}環境內務擺設整齊清潔，符合規定。")
    idx += 1

    # 第 5 點：警械裝備
    if not equip:
        equip = {'gi':0, 'go':0, 'bi':0, 'bo':0, 'ri':0, 'ro':0, 'vi':0, 'vo':0}
    
    gi, go = equip.get('gi', 0), equip.get('go', 0)
    bi, bo = equip.get('bi', 0), equip.get('bo', 0)
    ri, ro = equip.get('ri', 0), equip.get('ro', 0)
    vi, vo = equip.get('vi', 0), equip.get('vo', 0)
    
    lines.append(f"{idx}、{term}手槍出勤 {go} 把、在{loc_term} {gi} 把，子彈出勤 {bo} 顆、在{loc_term} {bi} 顆，無線電出勤 {ro} 臺、在{loc_term} {ri} 臺；防彈背心出勤 {vo} 件、在{loc_term} {vi} 件，幹部對械彈每日檢查管制良好，符合規定。")
    idx += 1

    # 第 6 點：幹部督勤
    lines.append(f"{idx}、本日{cadre}")
    idx += 1

    # 第 7 點：酒測聯單
    lines.append(f"{idx}、{term}酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。")
    idx += 1

    # 第 8 點：拘留室 (警備隊專用)
    if is_guard:
        if detention:
            lines.append(f"{idx}、拘留室值班{detention}，對人犯監控良好，無異常狀況發生。")
        else:
            lines.append(f"{idx}、拘留室目前無人犯。")
        idx += 1

    # 接續點：AI 刑案單優蹟紀錄
    if crimes:
        for c in crimes:
            suspect = c.get('嫌疑人', '不明')
            t = c.get('查獲時間', '不明')
            loc = c.get('查獲地點', '不明')
            law = c.get('觸犯法條', '不明')
            officer = c.get('查獲員警', '不明')
            if isinstance(officer, list):
                officer = "、".join(officer)
            else:
                # 若 AI 回傳的是字串，將其中的半形/全形逗號替換為頓號，格式更美觀
                officer = str(officer).replace(', ', '、').replace(',', '、').replace('，', '、')
            lines.append(f"{idx}、優蹟紀錄：{term}同仁 {officer} 於 {t} 在 {loc} 查獲 {suspect} 涉嫌 {law} 案。")
            idx += 1

    return "\n".join(lines)


# ==========================================
# 8. Streamlit UI
# ==========================================
show_sidebar()

st.title("🚓 勤務督導報告自動生成系統")
st.markdown("上傳各單位勤務表、交接簿（選填）、刑案單（選填），自動生成督導報告。")

# --- 基本資訊 ---
with st.expander("📋 督導基本資訊", expanded=True):
    sup_date = st.date_input("督導日期", value=datetime.today())

sup_date_str = sup_date.strftime("%Y年%m月%d日")

# --- 單位上傳區 ---
st.markdown("---")
st.subheader("📁 上傳單位資料")

num_units = st.number_input("本次督導單位數量", min_value=1, max_value=10, value=1, step=1)

unit_inputs = []
for i in range(num_units):
    with st.expander(f"第 {i+1} 個單位", expanded=(i == 0)):
        u_time = st.time_input(f"抵達時間 (單位 {i+1})", value=datetime.now().time(), key=f"time_{i}")
        c1, c2, c3 = st.columns(3)
        with c1:
            d_file = st.file_uploader(f"勤務表 (Excel)", type=["xlsx", "xls"], key=f"duty_{i}")
        with c2:
            e_file = st.file_uploader(f"交接簿 (Excel，選填)", type=["xlsx", "xls"], key=f"equip_{i}")
        with c3:
            p_file = st.file_uploader(f"刑案單 (PDF，選填)", type=["pdf"], key=f"crime_{i}")
        unit_inputs.append((u_time, d_file, e_file, p_file))

# --- 生成按鈕 ---
st.markdown("---")
if st.button("🚀 生成督導報告", type="primary"):
    all_ready = all(d is not None for u_time, d, e, p in unit_inputs)
    if not all_ready:
        st.error("每個單位都必須至少上傳【勤務表】。")
        st.stop()

    st.session_state.unit_reports = {}
    progress = st.progress(0)

    for i, (u_time, d_file, e_file, p_file) in enumerate(unit_inputs):
        with st.spinner(f"正在處理第 {i+1} 個單位..."):
            # 勤務表解析
            duty_info = extract_duty_v2(io.BytesIO(d_file.read()), u_time.hour)

            # 交接簿解析
            equip = None
            if e_file:
                equip = extract_equip_v2(io.BytesIO(e_file.read()))

            # 刑案單辨識
            crimes = []
            if p_file and model is not None:
                crimes = parse_crime_pdf_gemini(p_file, duty_info.get('roster', []), i)

            # 組合報告 (傳入格式化的抵達時間，例如 "1545")
            time_str = u_time.strftime("%H%M")
            report_text = build_report(duty_info, equip, crimes, time_str, sup_date)
            
            st.session_state.unit_reports[i] = {
                'unit_name': duty_info.get('unit_name', f'單位{i+1}'),
                'report': report_text,
                'duty_info': duty_info,
            }

        progress.progress((i + 1) / num_units)

    st.success(f"✅ 已完成 {num_units} 個單位的報告生成！")

# --- 報告預覽與下載 ---
if st.session_state.unit_reports:
    st.markdown("---")
    st.subheader("📄 報告預覽")

    tabs = st.tabs([v['unit_name'] for v in st.session_state.unit_reports.values()])
    for tab, (idx, data) in zip(tabs, st.session_state.unit_reports.items()):
        with tab:
            report_text = data['report']
            edited = st.text_area(
                "可直接在此編輯報告內容",
                value=report_text,
                height=350,
                key=f"edit_{idx}"
            )
            st.session_state.unit_reports[idx]['report'] = edited

            col_dl, col_mail = st.columns(2)
            with col_dl:
                st.download_button(
                    label="⬇️ 下載報告 (.txt)",
                    data=edited.encode('utf-8-sig'),
                    file_name=f"{data['unit_name']}_督導報告_{sup_date_str}.txt",
                    mime="text/plain",
                    key=f"dl_{idx}"
                )
            with col_mail:
                receiver = st.text_input("收件人 Email", key=f"mail_{idx}")
                if st.button("📧 寄送報告", key=f"send_{idx}"):
                    if receiver:
                        subject = f"【督導報告】{data['unit_name']} {sup_date_str}"
                        if send_gmail(subject, edited, receiver):
                            st.success("郵件已寄出！")
                    else:
                        st.warning("請填寫收件人 Email。")

    # 合併下載
    st.markdown("---")
    all_text = "\n\n────────────────────────────────────────\n\n".join([v['report'] for v in st.session_state.unit_reports.values()])
    st.download_button(
        label="⬇️ 下載全部報告（合併）",
        data=all_text.encode('utf-8-sig'),
        file_name=f"督導報告_全部_{sup_date_str}.txt",
        mime="text/plain"
    )
