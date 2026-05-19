import streamlit as st
import pandas as pd
import io
import sys
import os
import re
import json
import traceback
import smtplib
import google.generativeai as genai
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime, timedelta
from pdf2image import convert_from_bytes

# 自動將上層目錄加入路徑 (相容您的系統架構)
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        pass 

# ==========================================
# 0. 系統初始化與狀態管理
# ==========================================
if "unit_reports" not in st.session_state:
    st.session_state.unit_reports = {}

# 初始化 Gemini 2.5 Flash API
try:
    api_key = st.secrets["api"]["GOOGLE_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
except Exception as e:
    st.error(f"Gemini API 初始化失敗，請檢查 secrets 設定: {e}")

# 🛑 關鍵防護 1：關閉所有安全攔截，避免真實姓名(如葉煥堂)被誤判為洩漏個資而阻擋
safety_settings = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
]

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
# 2. 核心工具與解析函式
# ==========================================
def safe_int(val):
    try:
        return int(float(str(val).split('.')[0].replace(',', '')))
    except:
        return 0

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
        if eh == 0: e = 24
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

def extract_duty_v2(d_file, hour):
    res = {
        'v_name': '解析失敗', 'detention_name': None,
        'cadre_status': '無幹部資料', 'unit_name': '未偵測單位',
        'term': '該所', 'loc_term': '所', 'has_skyline': True, 'is_guard_unit': False, 'roster': []
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
            if r >= len(df): break
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
                            res['v_name'] = fmap.get(valid_codes[0], f'警員({valid_codes[0]})')
                            v_found = True

        if not v_found:
            res['v_name'] = '該時段無值班人員'

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

            if d_list:
                grouped_duties = {}
                for d in d_list:
                    grouped_duties.setdefault(d['s'], []).append(d)

                filtered_d_list = []
                for s_time, items in grouped_duties.items():
                    if len(items) > 1:
                        non_internal = [x for x in items if '內部管理' not in x['n']]
                        if non_internal: filtered_d_list.extend(non_internal)
                        else: filtered_d_list.append(items[0])
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
# 3. Gemini 2.5 Vision 刑案單強效辨識核心
# ==========================================
def parse_crime_pdf_gemini(pdf_file, roster: list, unit_idx: int) -> list:
    pdf_file.seek(0)
    images = convert_from_bytes(pdf_file.read(), dpi=150)
    results = []
    roster_str = "、".join(roster)
    
    prompt = (
        f"請提取：嫌疑人, 查獲時間, 查獲地點, 觸犯法條, 查獲員警(請完整提取「職稱+姓名」，例如「警員蕭漢祥」)。\n"
        f"名冊供比對參考：{roster_str}。\n"
        "請嚴格回傳 JSON Array (列表) 格式，即使只有一筆資料也要放在陣列中，例如：\n"
        "[\n"
        "  {\n"
        '    "嫌疑人": "王大明",\n'
        '    "查獲時間": "115年05月18日 10時00分",\n'
        '    "查獲地點": "桃園市龍潭區某路段",\n'
        '    "觸犯法條": "公共危險",\n'
        '    "查獲員警": "警員李小華、巡佐張大山"\n'
        "  }\n"
        "]"
    )
    
    total_pages = len(images)
    for i, img in enumerate(images):
        try:
            st.info(f"單位 {unit_idx+1} 🚀 AI 正在辨識刑案單第 {i+1}/{total_pages} 頁...")
            response = model.generate_content([prompt, img], safety_settings=safety_settings)
            raw_text = response.text.strip()
            
            if raw_text.startswith("```"):
                raw_text = re.sub(r'^
http://googleusercontent.com/immersive_entry_chip/0
