import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# 分頁配置
st.set_page_config(page_title="督導報告 v7.0 - 時空對位版", layout="wide")

# 套用標楷體風格
st.markdown(f"""
    <style>
    @font-face {{
        font-family: 'Kaiu';
        src: url('kaiu.ttf');
    }}
    .stTextArea textarea {{
        font-family: 'Kaiu', "標楷體", sans-serif !important;
        font-size: 19px !important;
        line-height: 1.7 !important;
        color: #1c1c1c !important;
    }}
    </style>
    """, unsafe_allow_html=True)

st.title("📋 督導報告極速生成器 v7.0 (矩陣與時間對位版)")

# --- 側邊欄：檔案與時間設定 ---
with st.sidebar:
    st.header("📂 數據源匯入")
    duty_file = st.file_uploader("1. 上傳『勤務分配表』", type=['xlsx', 'csv'])
    equip_file = st.file_uploader("2. 上傳『值班裝備交接簿』", type=['xlsx', 'csv'])
    
    st.divider()
    target_time = st.time_input("督導時間 (自動對位時段與裝備)", datetime.now().time())
    time_str = target_time.strftime('%H%M')
    target_hour = target_time.hour
    
    # 日期推算
    today = datetime.now()
    d_m5, d_m1 = [(today - timedelta(days=i)).strftime('%m月%d日') for i in [5, 1]]
    d_m3 = (today - timedelta(days=3)).strftime('%m月%d日')

# --- 核心解析引擎 ---
def safe_int(val):
    try: return int(float(str(val).split('.')[0].replace(',', '')))
    except: return 0

def normalize_code(c):
    c_str = str(c).strip().replace(".0", "").upper()
    if c_str.isdigit(): return str(int(c_str))
    return c_str

def parse_time_cell(val):
    """破解 Excel 日期魔咒，提取時間區間"""
    val_str = str(val).strip().replace("\n", "").replace(" ", "").replace("|", "-")
    if val_str in ["", "nan", "NaN"]: return None, None
    
    m_date = re.search(r'20\d{2}-(\d{2})-(\d{2})', val_str)
    if m_date: return int(m_date.group(1)), int(m_date.group(2))
        
    m_time = re.search(r'(\d{1,2})[~~\-－—–_]+(\d{1,2})', val_str)
    if m_time: return int(m_time.group(1)), int(m_time.group(2))
        
    return None, None

def extract_full_logic(d_file, e_file, hour):
    res = {'v_name': '未偵測', 'cadre_status': '無幹部資料', 'eq': None, 'debug_info': {}}
    try:
        # 1. 讀取勤務分配表
        df = pd.read_excel(d_file, header=None, dtype=str).fillna("")
        
        # A. 取得完美番號對照表
        full_text = " ".join([str(x).strip() for x in df.values.flatten() if x])
        pattern = r'(?<![A-Za-z0-9])([A-Z]|[0-9]{1,2})\s*(所長|副所長|巡官|巡佐|警員|實習)\s*([\u4e00-\u9fa5]{2,4})'
        matches = re.findall(pattern, full_text)
        
        name_map = {}
        for m in matches:
            code = normalize_code(m[0])
            name = m[2]
            for t_char in ["所", "副", "巡", "警", "實", "員", "長"]:
                if name.endswith(t_char): name = name[:-1]
            if len(name) >= 2: name_map[code] = name
            
        res['debug_info']['1. 番號對照表'] = name_map

        # B. 定位 X 軸 (時間軸)
        time_row_idx = 2 
        time_cols = {}
        target_col_idx = -1
        
        for r in range(5):
            count = 0
            temp_cols = {}
            for c_idx in range(len(df.columns)):
                start_h, end_h = parse_time_cell(df.iloc[r, c_idx])
                if start_h is not None and end_h is not None:
                    temp_cols[c_idx] = (start_h, end_h)
                    count += 1
            if count > max(len(time_cols), 3):
                time_row_idx = r
                time_cols = temp_cols

        for c_idx, (start_h, end_h) in time_cols.items():
            calc_end = end_h if end_h > start_h else end_h + 24
            calc_hour = hour if hour >= 6 or calc_end <= 24 else hour + 24
            
            if start_h <= calc_hour < calc_end:
                target_col_idx = c_idx

        # C. 抓取值班人員
        v_row_idx = time_row_idx + 1 
        for r in range(time_row_idx + 1, min(time_row_idx + 4, len(df))):
            if "值" in str(df.iloc[r, 0]) + str(df.iloc[r, 1]):
                v_row_idx = r
                break
                
        if target_col_idx != -1:
            raw_cell = str(df.iloc[v_row_idx, target_col_idx]).strip()
            m_code = re.search(r'[A-Za-z0-9]{1,2}', raw_cell)
            if m_code:
                code = normalize_code(m_code.group(0))
                res['v_name'] = name_map.get(code, f"未建檔番號:{code}")
            else:
                for code, name in name_map.items():
                    if name in raw_cell: res['v_name'] = name; break

        # D. 抓取幹部動態
        cadre_notes = []
        for code in ["A", "B", "C"]:
            name = name_map.get(code, {"A":"所長", "B":"副所長", "C":"幹部"}.get(code))
            is_off = False
            patrol_slots = []
            duty_names = set() 
            found_in_matrix = False

            for r in range(v_row_idx, len(df)):
                duty_title = str(df.iloc[r, 0]) + str(df.iloc[r, 1])
                is_leave_row = any(k in duty_title for k in ["休", "假", "輪", "輸", "補"])

                for c_idx, (s_h, e_h) in time_cols.items():
                    cell_val = str(df.iloc[r, c_idx])
                    cell_codes = [normalize_code(x) for x in re.findall(r'[A-Za-z0-9]{1,2}', cell_val)]
                    
                    if code in cell_codes:
                        found_in_matrix = True
                        if is_leave_row:
                            is_off = True
                        else:
                            is_external = False
                            if "巡" in duty_title or "巡" in cell_val: duty_names.add("巡邏"); is_external = True
                            if "守" in duty_title or "守" in cell_val: duty_names.add("守望"); is_external = True
                            if "臨" in duty_title or "臨" in cell_val: duty_names.add("臨檢"); is_external = True
                            if "交" in duty_title or "交" in cell_val: duty_names.add("交整"); is_external = True
                            if "路" in duty_title or "路" in cell_val: duty_names.add("路檢"); is_external = True
                            
                            if is_external:
                                patrol_slots.append((s_h, e_h))

            if not found_in_matrix or is_off:
                cadre_notes.append(f"{name}休假")
            else:
                if patrol_slots:
                    min_s = min([s[0] for s in patrol_slots])
                    max_e = max([s[1] for s in patrol_slots])
                    e_str = "24" if max_e == 24 or max_e == 0 else f"{max_e%24:02d}"
                    d_str = "、".join(sorted(list(duty_names)))
                    cadre_notes.append(f"{name}在所督勤，編排{min_s:02d}至{e_str}時段{d_str}勤務")
                else:
                    cadre_notes.append(f"{name}在所督勤")
                    
        if cadre_notes:
            res['cadre_status'] = "；".join(cadre_notes) + "。"
            
    except Exception as e:
        res['v_name'] = "解析失敗"
        res['cadre_status'] = f"幹部解析錯誤: {e}"

    # 🌟 2. 解析裝備交接簿 (時間切片引擎)
    try:
        df_e = pd.read_csv(e_file, header=None) if e_file.name.endswith('csv') else pd.read_excel(e_file, header=None)
        df_e_s = df_e.astype(str)
        
        stop_idx = len(df_e)
        for r in range(len(df_e)):
            t_val = df_e_s.iloc[r, 0] # 假設時間在 A 欄
            nums = re.findall(r'\d{1,2}', t_val)
            if nums:
                row_hour = int(nums[0])
                # 如果交接簿的時間大於督導時間，就停止 (防呆：避免跨夜 02時 誤判 20時)
                if row_hour > hour and (row_hour - hour < 12):
                    stop_idx = r
                    break
                    
        # 切出督導時間之前的紀錄
        df_e_valid = df_e.iloc[:stop_idx]
        df_e_s_valid = df_e_s.iloc[:stop_idx]
        
        # 安全獲取最後一筆匹配紀錄
        def get_latest_row(df, df_s, keyword):
            rows = df[df_s.iloc[:, 1].str.contains(keyword, na=False)]
            return rows.iloc[-1] if not rows.empty else None
            
        r_in = get_latest_row(df_e_valid, df_e_s_valid, "在")
        r_out = get_latest_row(df_e_valid, df_e_s_valid, "出")
        r_fix = get_latest_row(df_e_valid, df_e_s_valid, "送")
        
        def get_val(row, idx):
            return safe_int(row.iloc[idx]) if row is not None and idx < len(row) else 0

        res['eq'] = {
            "gi": get_val(r_in, 2), "go": get_val(r_out, 2), "gf": get_val(r_fix, 2),
            "bi": get_val(r_in, 3), "bo": get_val(r_out, 3), "bf": get_val(r_fix, 3),
            "ri": get_val(r_in, 6), "ro": get_val(r_out, 6), "rf": get_val(r_fix, 6),
            "vi": get_val(r_in, 11), "vo": get_val(r_out, 11), "vf": get_val(r_fix, 11)
        }
        res['debug_info']['2. 裝備切片行數'] = stop_idx
    except Exception as e: 
        res['eq'] = None
        res['debug_info']['裝備解析錯誤'] = str(e)
        
    return res

# --- 主畫面執行 ---
if duty_file and equip_file:
    with st.spinner("啟動時空矩陣引擎，精準定位人員與裝備動態..."):
        data = extract_full_logic(duty_file, equip_file, target_hour)
    
    st.success("✅ 報告生成完畢！")
    
    c1, c2 = st.columns(2)
    with c1:
        check_mon = st.checkbox("駐地監錄/天羅地網正常", value=True)
        check_edu = st.checkbox("勤前教育宣導落實", value=True)
    with c2:
        check_env = st.checkbox("環境內務擺設整齊", value=True)
        check_alc = st.checkbox("酒測聯單無跳號", value=True)

    lines = []
    
    lines.append(f"{time_str}，該所值班警員{data['v_name']}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。")
    if check_mon: lines.append(f"該所駐地監錄設備及天羅地網系統均運作正常，無故障，{d_m5}至{d_m1}有逐日檢測2次以上紀錄。")
    if check_edu: lines.append(f"該所{d_m3}至{d_m1}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。")
    if check_env: lines.append(f"該所環境內務擺設整齊清潔，符合規定。")
        
    eq = data['eq'] if data['eq'] else {"gi":0, "go":0, "gf":0, "bi":0, "bo":0, "bf":0, "ri":0, "ro":0, "rf":0, "vi":0, "vo":0, "vf":0}
    fix_str = f"（另有槍枝 {eq['gf']} 把、無線電 {eq['rf']} 臺送修中）" if (eq['gf'] + eq['rf']) > 0 else ""
    lines.append(f"該所手槍出勤 {eq['go']} 把、在所 {eq['gi']} 把，子彈出勤 {eq['bo']} 顆、在所 {eq['bi']} 顆，無線電出勤 {eq['ro']} 臺、在所 {eq['ri']} 臺；防彈背心出勤 {eq['vo']} 件、在所 {eq['vi']} 件，幹部對械彈每日檢查管制良好，符合規定{fix_str}。")
    
    lines.append(f"本日{data['cadre_status']}")
    if check_alc: lines.append(f"該所酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。")

    final_report = "\n".join([f"{i+1}、{line}" for i, line in enumerate(lines)])
    
    st.markdown("---")
    st.subheader("📋 最終督導報告 (時空對位版)")
    st.text_area("複製回貼公務系統：", value=final_report, height=450)
    
    with st.expander("🛠️ 查看系統自動辨識結果 (除錯專區)"):
        st.write(f"✅ 值班員警：{data['v_name']}")
        st.write(f"✅ 幹部動態：{data['cadre_status']}")
        st.json(data['debug_info'])

else:
    st.info("👋 匯入兩份檔案後，系統將自動啟動矩陣雷達引擎！")
