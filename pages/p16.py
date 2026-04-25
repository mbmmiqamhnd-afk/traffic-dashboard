import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# 分頁配置
st.set_page_config(page_title="督導報告 v7.0 - XY矩陣定位版", layout="wide")

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

st.title("📋 督導報告極速生成器 v7.0 (矩陣對位引擎)")

# --- 側邊欄：檔案與時間設定 ---
with st.sidebar:
    st.header("📂 數據源匯入")
    duty_file = st.file_uploader("1. 上傳『勤務分配表』", type=['xlsx', 'csv'])
    equip_file = st.file_uploader("2. 上傳『值班裝備交接簿』", type=['xlsx', 'csv'])
    
    st.divider()
    target_time = st.time_input("督導時間 (自動對位時段)", datetime.now().time())
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
    """將番號標準化，過濾空白與小數點"""
    c_str = str(c).strip().replace(".0", "").upper()
    if c_str.isdigit(): return str(int(c_str))
    return c_str

def extract_matrix_logic(d_file, e_file, hour):
    res = {'v_name': '未偵測', 'cadre_status': '無幹部資料', 'eq': None, 'debug_info': {}}
    try:
        # 1. 讀取勤務分配表 (全轉為字串防止 datetime 誤判)
        df = pd.read_excel(d_file, header=None, dtype=str).fillna("")
        
        # 🌟 A. 取得完美番號對照表
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
            
        res['debug_info']['對照表'] = name_map

        # 🌟 B. 定位 X 軸 (時間軸)
        time_row_idx = 2 # 根據提示：時段在第3列 (Index 2)
        time_cols = {}   # 記錄 {欄位Index: (起始時, 結束時, 原始字串)}
        target_col_idx = -1
        
        for c_idx in range(len(df.columns)):
            val = str(df.iloc[time_row_idx, c_idx]).strip()
            
            # 處理各種時間格式：06-07, 06~07, 06－07
            clean_val = re.sub(r'[～－—–_~]', '-', val).replace(":00", "").replace("：00", "")
            
            # 若為 Pandas DateTime 格式 (06-07 變為 2026-06-07)
            m_dt = re.search(r'^\d{4}-(\d{2})-(\d{2})', clean_val)
            if m_dt:
                start_h, end_h = int(m_dt.group(1)), int(m_dt.group(2))
            else:
                m = re.search(r'(\d{1,2})\s*-\s*(\d{1,2})', clean_val)
                if m:
                    start_h, end_h = int(m.group(1)), int(m.group(2))
                else:
                    continue

            # 跨夜邏輯處理
            if end_h == 0 or end_h < start_h: end_h += 24
            calc_hour = hour
            if calc_hour < 6 and end_h > 24: calc_hour += 24

            time_cols[c_idx] = (start_h, end_h, val)

            # 檢查督導時間是否落在本格
            if start_h <= calc_hour < end_h:
                target_col_idx = c_idx

        res['debug_info']['時段對位'] = {"找到的欄位 Index": target_col_idx, "該欄時段": time_cols.get(target_col_idx)}

        # 🌟 C. 定位 Y 軸與交叉點 (抓取值班人員)
        # 根據提示：值班代號在第4列 (Index 3)，但我們加上保險掃描機制
        v_row_idx = 3 
        for r in range(10):
            if "值" in str(df.iloc[r, 0]) or "值" in str(df.iloc[r, 1]):
                v_row_idx = r
                break
                
        if target_col_idx != -1:
            raw_cell = str(df.iloc[v_row_idx, target_col_idx])
            res['debug_info']['值班格子內容'] = raw_cell
            
            # 萃取格子內的番號
            m_code = re.search(r'[A-Za-z0-9]{1,2}', raw_cell)
            if m_code:
                code = normalize_code(m_code.group(0))
                res['v_name'] = name_map.get(code, f"未建檔番號:{code}")
            else:
                res['v_name'] = f"格子無番號 ({raw_cell})"

        # 🌟 D. 矩陣雷達：抓取幹部動態 (A, B, C)
        cadre_notes = []
        for code in ["A", "B", "C"]:
            name = name_map.get(code, {"A":"所長", "B":"副所長", "C":"幹部"}.get(code))
            found_today = False
            is_off = False
            patrol_slots = []

            # 掃描整張矩陣表 (第4列以下，時段欄位以內)
            for r in range(3, len(df)):
                duty_title = str(df.iloc[r, 0]) + str(df.iloc[r, 1])
                for c_idx in time_cols.keys():
                    cell_val = str(df.iloc[r, c_idx])
                    
                    # 檢查該格子是否包含該幹部的番號
                    tokens = [normalize_code(x) for x in re.findall(r'[A-Za-z0-9]{1,2}', cell_val)]
                    if code in tokens:
                        found_today = True
                        if any(k in duty_title for k in ["休", "假", "補", "外"]):
                            is_off = True
                        elif "巡" in duty_title:
                            patrol_slots.append(time_cols[c_idx]) # 加入巡邏時間

            # 判定邏輯
            if not found_today or is_off:
                cadre_notes.append(f"{name}休假")
            else:
                if patrol_slots:
                    min_s = min([s[0] for s in patrol_slots])
                    max_e = max([s[1] for s in patrol_slots])
                    min_s_str = f"{min_s%24:02d}"
                    max_e_str = "24" if max_e == 24 else f"{max_e%24:02d}"
                    cadre_notes.append(f"{name}在所督勤，編排{min_s_str}至{max_e_str}時段巡邏勤務")
                else:
                    cadre_notes.append(f"{name}在所督勤")
                    
        if cadre_notes:
            res['cadre_status'] = "；".join(cadre_notes) + "。"
            
    except Exception as e:
        res['v_name'] = "解析失敗"
        res['cadre_status'] = f"解析錯誤: {e}"

    # 2. 解析裝備交接簿 (完美穩定維持)
    try:
        df_e = pd.read_csv(e_file, header=None) if e_file.name.endswith('csv') else pd.read_excel(e_file, header=None)
        df_e_s = df_e.astype(str)
        
        r_in = df_e[df_e_s.iloc[:, 1].str.contains("在", na=False)].iloc[-1]
        r_out = df_e[df_e_s.iloc[:, 1].str.contains("出", na=False)].iloc[-1]
        r_fix = df_e[df_e_s.iloc[:, 1].str.contains("送", na=False)].iloc[-1]
        
        res['eq'] = {
            "gi": safe_int(r_in.iloc[2]), "go": safe_int(r_out.iloc[2]), "gf": safe_int(r_fix.iloc[2]),
            "bi": safe_int(r_in.iloc[3]), "bo": safe_int(r_out.iloc[3]), "bf": safe_int(r_fix.iloc[3]),
            "ri": safe_int(r_in.iloc[6]), "ro": safe_int(r_out.iloc[6]), "rf": safe_int(r_fix.iloc[6]),
            "vi": safe_int(r_in.iloc[11]), "vo": safe_int(r_out.iloc[11]), "vf": safe_int(r_fix.iloc[11])
        }
    except: res['eq'] = None
    return res

# --- 主畫面執行 ---
if duty_file and equip_file:
    with st.spinner("啟動 XY 矩陣定位引擎，正在擷取人員動態..."):
        data = extract_matrix_logic(duty_file, equip_file, target_hour)
    
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
    st.subheader("📋 最終督導報告 (矩陣定位版)")
    st.text_area("複製回貼公務系統：", value=final_report, height=450)
    
    with st.expander("🛠️ 查看系統自動辨識結果 (完美除錯區)"):
        st.json(data['debug_info'])
        st.write(f"✅ 值班員警：{data['v_name']}")
        st.write(f"✅ 幹部動態：{data['cadre_status']}")

else:
    st.info("👋 匯入兩份檔案後，系統將自動啟動矩陣定位引擎！")
