import pandas as pd
import re


# ──────────────────────────────────────────────────────────
# 工具函式
# ──────────────────────────────────────────────────────────
def safe_int(val):
    try:
        return int(float(str(val).split('.')[0].replace(',', '')))
    except:
        return 0


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
    """
    從勤務表底部對照區解析 代號 → 職稱姓名。
    對照區特徵：col[0] 同時含 '代號' 與 '職稱'；
    每行以 6 欄為一組：[代號][職稱姓名×5][代號][職稱姓名×5]…
    """
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
                    # 去掉姓名末尾的警勤區番號（如 "警員 林軒宇 10"）
                    name = re.sub(r'\s+\d{1,2}$', '', name).strip()
                    if (code and name and
                            code not in ('', 'nan') and name not in ('', 'nan') and
                            re.match(r'^[A-Za-z0-9甲乙丙丁]{1,3}$', code)):
                        fmap[code] = name
                    c += 6
    return fmap


def find_target_col(df, hour):
    """在 Row2 時段列找到對應 hour 的欄索引。回傳 (col_idx, t_cols_dict)"""
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
            e = 24  # 23-00 時段
        if s <= adj_h < e:
            return c, t_cols
    return -1, t_cols


# ──────────────────────────────────────────────────────────
# 勤務表解析（v2）
# ──────────────────────────────────────────────────────────
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
                res['loc_term'] = res['term'][1:]
                break

        # ── 2. 人員對照表 ──
        fmap = build_fmap(df)

        # ── 3. 時段欄 ──
        target_col, t_cols = find_target_col(df, hour)
        if target_col == -1:
            res['v_name'] = '找不到對應時段欄'
            return res

        # ── 4. 值班人員 ──
        # 結構：col[0]='值班'，col[1]=勤務方式，col[target_col]=代號
        # 龍潭：Row05 是一般值班（col[1]='值班'）
        # 高平：Row03 是一般值班（col[1]='值班'）
        # 警備隊：Row04 是勤務中心值勤（代號為班別，非個人）
        v_found = False
        for r in range(3, 30):
            if r >= len(df):
                break
            col0 = str(df.iloc[r, 0]).strip()
            col1 = str(df.iloc[r, 1]).strip()
            if col0 != '值班':
                continue
            # 跳過「值班副所長」說明列（col1 含副所長但 target_col 無代號）
            cell = str(df.iloc[r, target_col]).strip()
            if not cell or cell == 'nan':
                continue
            # 跳過純文字說明列（整行文字超長）
            if len(cell) > 10:
                continue
            # 提取第一個有效代號（字母或數字）
            codes = re.findall(r'[A-Z甲乙丙丁][0-9]?|[0-9]{2}', cell)
            valid_codes = [c for c in codes if re.match(r'^[A-Z0-9甲乙丙丁]{1,3}$', c)]
            if valid_codes:
                code = valid_codes[0]
                if code in fmap:
                    res['v_name'] = fmap[code]
                else:
                    # 如果是班別（甲乙丙丁），找對應人員
                    res['v_name'] = fmap.get(code, f'警員({code})')
                v_found = True
                break
            # 拘留室（警備隊專用）
            if '拘留' in col1 and res['is_guard_unit']:
                if valid_codes:
                    res['detention_name'] = fmap.get(valid_codes[0], f'警員({valid_codes[0]})')

        if not v_found:
            res['v_name'] = '該時段無值班人員'

        # ── 5. 幹部動態 ──
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

        # 找到資料區的結束列（出現「請假人員」或「重要紀事」）
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

            # 掃全部勤務列（Row3 ~ footer_row）
            for r in range(3, footer_row):
                col0 = str(df.iloc[r, 0]).strip()
                col1 = str(df.iloc[r, 1]).strip()

                # 跳過空行或結構列
                if not col0:
                    continue

                for c, (sh, eh) in t_cols.items():
                    cell_val = str(df.iloc[r, c]).strip()
                    if not cell_val or cell_val == 'nan':
                        continue
                    # 提取代號列表（支援多代號如 '丙1 丁'、'D 03'）
                    raw_codes = re.findall(r'[A-Z甲乙丙丁][0-9]?|[0-9]{2}', cell_val)
                    raw_codes = [x for x in raw_codes if re.match(r'^[A-Z0-9甲乙丙丁]{1,3}$', x)]
                    if code not in raw_codes:
                        continue

                    # 勤務名稱取 col1（限前8字避免過長）
                    duty_name = col1[:8] if col1 and len(col1) >= 2 else col0
                    if not duty_name:
                        duty_name = '勤務'
                    if any(k in duty_name for k in ['休', '輪', '假', '補', '外宿']):
                        is_off = True
                        continue

                    # 過濾掉標題殘影
                    duty_name = _clean_duty_name(duty_name)
                    if duty_name is None:
                        continue

                    s = adj(sh)
                    e = adj(eh) if eh != sh else adj(sh) + 1
                    if eh == 0:
                        e = 24
                    d_list.append({'sh': sh, 'eh': eh, 's': s, 'e': e, 'n': duty_name})

            # 若在休假/輪休列看到代號 → 休假
            for r in range(3, len(df)):
                col0 = str(df.iloc[r, 0]).strip()
                col1 = str(df.iloc[r, 1]).strip()
                if any(k in col0 + col1 for k in ['輪休', '慰休', '公假', '補休', '事假', '病假']):
                    for c in range(13, len(df.columns)):
                        cv = str(df.iloc[r, c])
                        if code in re.findall(r'[A-Z甲乙丙丁][0-9]?|[0-9]{2}', cv):
                            is_off = True
                            break

            if is_off:
                c_notes.append(f'{fname_full}休假')
            elif d_list:
                d_list.sort(key=lambda x: x['s'])
                merged = []
                for d in d_list:
                    if (merged and d['s'] == merged[-1]['e'] and d['n'] == merged[-1]['n']):
                        merged[-1]['e'] = d['e']
                        merged[-1]['eh'] = d['eh']
                    else:
                        merged.append(dict(d))
                parts = [
                    f"{m['sh']:02d}-{(24 if m['eh'] == 0 else m['eh']):02d}{m['n']}"
                    for m in merged
                ]
                c_notes.append(
                    f"{fname_full}在{res['loc_term']}督勤，編排{'、'.join(parts)}勤務"
                )
            else:
                c_notes.append(f'{fname_full}在{res["loc_term"]}督勤')

        res['cadre_status'] = '；'.join(c_notes) + '。' if c_notes else '無幹部資料。'

    except Exception as e:
        import traceback
        res['cadre_status'] = f'解析中斷：{e}'
        res['v_name'] = res.get('v_name', '解析失敗')

    return res


# ──────────────────────────────────────────────────────────
# 交接簿解析（v2）
# ──────────────────────────────────────────────────────────
def extract_equip_v2(e_file):
    """
    從 Row02 標題列動態偵測欄位，取最新一筆「在所」「出勤」的數值。
    龍潭/高平：col2=手槍, col3=子彈, col6=無線電, col8=防彈背心
    警備隊：   col2=手槍, col3=子彈, col4=無線電, col6=防彈背心
    """
    try:
        df = pd.read_excel(e_file, header=None).fillna('')
        header_row = 2
        col_map = {}

        for c in range(len(df.columns)):
            v = str(df.iloc[header_row, c]).replace('\n', '').replace(' ', '')
            # 手槍（不含子彈）
            if v == '手槍':
                col_map['gun'] = c
            # 子彈（上一欄是手槍 → 手槍子彈；否則 M16 子彈不要）
            if '子彈' in v:
                prev = str(df.iloc[header_row, c - 1]).replace('\n', '') if c > 0 else ''
                if '手槍' in prev:
                    col_map['bullet'] = c
            if '無線電' in v:
                col_map['radio'] = c
            if '防彈背心' in v:
                col_map['vest'] = c

        # fallback：按欄位實際標題逐欄搜尋
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

        # 找最新「在所」「出勤」列
        last_zi = last_zo = -1
        for r in range(3, len(df)):
            lbl = str(df.iloc[r, 1]).replace('\n', '').strip()
            if '在' in lbl and '所' in lbl:
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
    except Exception as e:
        return None

# 幹部勤務名稱清洗
_SKIP_DUTY_NAMES = {'勤務\n人員\n代號\n職稱\n姓名', '代號', '職稱', '姓名', '勤務備註', '員警', '時段', '項目'}

def _clean_duty_name(raw):
    """清洗掉不應作為勤務名稱的字串"""
    raw = raw.strip()
    if raw in _SKIP_DUTY_NAMES:
        return None
    if '代號' in raw and '職稱' in raw:
        return None
    if len(raw) > 15:
        raw = raw[:15]
    return raw or None
