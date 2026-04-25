import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# 分頁配置
st.set_page_config(page_title="督導報告 v7.0 - 坐標定位版", layout="wide")

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

st.title("📋 督導報告極速生成器 v7.0 (絕對坐標定位法)")

# --- 側邊欄：檔案與時間設定 ---
with st.sidebar:
    st.header("📂 數據源匯入")
    duty_file = st.file_uploader("1. 上傳『勤務分配表』", type=['xlsx', 'csv'])
    equip_file = st.file_uploader("2. 上傳『值班裝備交接簿』", type=['xlsx', 'csv'])
    
    st.divider()
    target_time = st.time_input("督導時間 (自動對位時段)", datetime.now().time())
    time_str = target_time.strftime('%H%M')
    target_hour = target_time.hour
    
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

def extract_full_logic(d_file, e_file, hour):
    res = {'v_name': '未偵測', 'cadre_status': '無幹部資料', 'eq': None, 'debug_map': {}, 'debug_cols': {}}
    try:
        # 1. 讀取勤務分配表
        df = pd.read_excel(d_file, header=None)
        df_ffill = df.ffill()
        
        # 🌟 A. 完美番號對照表
        full_text = " ".join([str(x).strip() for x in df.values.flatten() if pd.notna(x)])
        pattern = r'(?<![A-Za-z0-9])([A-Z]|[0-9]{1,2})\s*(所長|副所長|巡官|巡佐|警員|實習)\s*([\u4e00-\u9fa5]{2,4})'
        matches = re.findall(pattern, full_text)
        
        name_map = {}
        for m in matches:
            code, name = normalize_code(m[0]), m[2]
            for t_char in ["所", "副", "巡", "警", "實", "員", "長"]:
                if name.endswith(t_char): name = name[:-1]
            if len(name) >= 2: name_map[code] = name
        res['debug_map'] = name_map

        # 🌟 B. 動態掃描「時段列」(尋找 06-07, 12-14 這種特徵)
        time_row_idx = 2 # 預設第 3 列 (Index 2)
        time_cols = {} # 記錄 [行號] 對應的 [時段字串]
        target_col_idx = -1
        
        # 掃描前 10 列確保不漏接
        for r in range(10):
            row_str = " ".join([str(x) for x in df.iloc[r].values])
            if len(re.findall(r'\d{1,2}[-~]\d{1,2}', row_str)) >= 2:
                time_row_idx = r
                break
        
        # 解析該列的每一個時段格子
        time_row = df.iloc[time_row_idx]
        for c_idx in range(len(time_row)):
            cell_val = str(time_row.iloc[c_idx]).strip().replace("\n", "")
            m = re.search(r'(\d{1,2})\s*[-~]\s*(\d{1,2})', cell_val)
            if m:
                time_cols[c_idx] = cell_val
                start_h, end_h = int(m.group(1)), int(m.group(2))
                
                # 處理跨夜時段 (例如 22-02)
                if end_h == 0 or end_h < start_h: end_h += 24
                calc_hour = hour
                if calc_hour < 6 and end_h > 24: calc_hour += 24
                    
                # 判斷督導時間是否落在這個格子內
                if start_h <= calc_hour < end_h:
                    target_col_idx = c_idx

        res['debug_cols'] = {"時段列": time_row_idx, "鎖定目標欄位": target_col_idx, "解析到的時段": time_cols}

        # 🌟 C. 抓取值班人員 (絕對坐標：時段列的下一列)
        if target_col_idx != -1:
            # 根據您的提示：第3列是時段，第4列(下一列)是值班代號
            duty_code_raw = str(df_ffill.iloc[time_row_idx + 1, target_col_idx])
            m_code = re.search(r'[A-Za-z0-9]{1,2}', duty_code_raw)
            if m_code:
                duty_code = normalize_code(m_code.group(0))
                res['v_name'] = name_map.get(duty_code, f"未建檔番號:{duty_code}")
            else:
                # 若格子裡沒代號直接寫了名字
                for code, name in name_map.items():
                    if name in duty_code_raw:
                        res['v_name'] = name
                        break

        # 🌟 D. 抓取幹部動態 (A, B, C)
        cadre_notes = []
        for code in ["A", "B", "C"]:
            name = name_map.get(code, {"A":"所長", "B":"副所長", "C":"幹部"}.get(code))
            
            c_row_idx = -1
            for r in range(4, len(df_ffill)):
                for c in range(4): # 在前4欄尋找幹部代號
                    if normalize_code(df_ffill.iloc[r, c]) == code:
                        c_row_idx = r
                        break
                if c_row_idx != -1: break
            
            if c_row_idx != -1:
                duty_now = str(df_ffill.iloc[c_row_idx, target_col_idx]) if target_col_idx != -1 else ""
                if any(k in duty_now for k in ["休", "假", "補", "外"]):
                    cadre_notes.append(f"{name}休假")
                else:
                    # 搜尋所有被標記為巡邏的時段欄位
                    slots = []
                    for c_idx, t_str in time_cols.items():
                        if "巡" in str(df_ffill.iloc[c_row_idx, c_idx]):
                            slots.append(t_str)
                    
                    if slots:
                        # 擷取最前與最後的時間點
                        all_hours = []
                        for s in slots:
                            m = re.search(r'(\d{1,2})\s*[-~]\s*(\d{1,2})', s)
                            if m: all_hours.extend([int(m.group(1)), int(m.group(2))])
                        if all_hours:
                            s_t, e_t = min(all_hours), max(all_hours)
                            e_str = "24" if e_t == 24 else f"{e_t:02d}"
                            cadre_notes.append(f"{name}在所督勤，編排{s_t:02d}至{e_str}時段巡邏勤務")
                    else:
                        cadre_notes.append(f"{name}在所督勤")
            elif code in name_map:
                cadre_notes.append(f"{name}休假")
                
        if cadre_notes:
            res['cadre_status'] = "；".join(cadre_notes) + "。"
            
    except Exception as e:
        res['v_name'] = "解析失敗"
        res['cadre_status'] = f"解析錯誤: {e}"

    # 2. 解析裝備交接簿
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
    with st.spinner("啟動 XY 絕對坐標引擎，正在擷取人員動態..."):
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
    st.subheader("📋 最終督導報告 (坐標解密版)")
    st.text_area("複製回貼公務系統：", value=final_report, height=450)
    
    with st.expander("🛠️ 查看系統自動辨識結果 (除錯專區)"):
        st.json(data['debug_cols'])
        st.write(f"✅ 值班員警：{data['v_name']}")
        st.write(f"✅ 幹部動態：{data['cadre_status']}")

else:
    st.info("👋 匯入兩份檔案後，系統將自動啟動坐標解密引擎！")
