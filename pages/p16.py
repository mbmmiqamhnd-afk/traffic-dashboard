import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# 分頁配置
st.set_page_config(page_title="督導報告 v7.0 - 終極無敵版", layout="wide")

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

st.title("📋 督導報告極速生成器 v7.0 (全自動雷達版)")

# --- 側邊欄：檔案與時間設定 ---
with st.sidebar:
    st.header("📂 數據源匯入")
    duty_file = st.file_uploader("1. 上傳『勤務分配表』", type=['xlsx', 'csv'])
    equip_file = st.file_uploader("2. 上傳『值班裝備交接簿』", type=['xlsx', 'csv'])
    
    st.divider()
    target_time = st.time_input("督導時間 (自動對位時段)", datetime.now().time())
    time_str = target_time.strftime('%H%M')
    target_hour = target_time.hour
    
    # 日期推算 (格式為 04月20日)
    today = datetime.now()
    d_m5, d_m1 = [(today - timedelta(days=i)).strftime('%m月%d日') for i in [5, 1]]
    d_m3 = (today - timedelta(days=3)).strftime('%m月%d日')

# --- 核心解析引擎 ---
def safe_int(val):
    try:
        return int(float(str(val).split('.')[0].replace(',', '')))
    except:
        return 0

def normalize_code(c):
    """將番號標準化，例如 '07', ' 7 ', 'A' 統一轉為乾淨的 '7', 'A'"""
    c_str = str(c).strip().replace(".0", "").upper()
    if c_str.isdigit():
        return str(int(c_str))
    return c_str

def extract_full_logic(d_file, e_file, hour):
    res = {'v_name': '未偵測', 'cadre_status': '無幹部資料', 'eq': None, 'debug_map': {}}
    try:
        # 1. 讀取勤務分配表
        df = pd.read_excel(d_file, header=None)
        
        # 🌟 A. 全域平坦化雷達：徹底無視 Excel 合併儲存格
        # 將整張表轉為單一巨大字串，完全不怕擠在同一格
        full_text = " ".join([str(x).strip() for x in df.values.flatten() if pd.notna(x)])
        
        # 強力正則表達式：尋找 [番號] + [職稱] + [姓名]
        # (?<![A-Za-z0-9]) 確保番號前不是英文數字，避免抓錯
        pattern = r'(?<![A-Za-z0-9])([A-Z]|[0-9]{1,2})\s*(所長|副所長|巡官|巡佐|警員|實習)\s*([\u4e00-\u9fa5]{2,4})'
        matches = re.findall(pattern, full_text)
        
        name_map = {}
        for m in matches:
            code = normalize_code(m[0])
            name = m[2]
            # 專殺「劉憬霖警」這類黏字的 Bug：把結尾的職稱字眼切掉
            for t_char in ["所", "副", "巡", "警", "實", "員", "長"]:
                if name.endswith(t_char):
                    name = name[:-1]
            if len(name) >= 2:
                name_map[code] = name
                
        res['debug_map'] = name_map
        
        # 🌟 B. 定位主表時段
        df_main = df.head(45).ffill() 
        col_idx = 4 + (hour // 2)
        
        # 🌟 C. 抓取值班人員
        duty_rows = df_main[df_main.iloc[:, col_idx].astype(str).str.contains(r"值.*班", regex=True, na=False)]
        if not duty_rows.empty:
            found_name = False
            # 掃描該列前四格尋找番號
            for search_col in range(4):
                r_code = normalize_code(duty_rows.iloc[0, search_col])
                if r_code in name_map:
                    res['v_name'] = name_map[r_code]
                    found_name = True
                    break
            
            if not found_name:
                fallback_code = normalize_code(duty_rows.iloc[0, 1])
                if fallback_code == "" or fallback_code == "NAN":
                    fallback_code = normalize_code(duty_rows.iloc[0, 0])
                res['v_name'] = name_map.get(fallback_code, f"番號[{fallback_code}]未建檔")
        
        # 🌟 D. 抓取幹部動態 (鎖定 A, B, C)
        cadre_notes = []
        for code in ["A", "B", "C"]:
            # 若對照表沒抓到，預設給職稱
            c_name = name_map.get(code, {"A":"所長", "B":"副所長", "C":"幹部"}.get(code))
            
            # 找尋幹部番號所在列
            c_row = pd.DataFrame()
            for search_col in range(4):
                temp_row = df_main[df_main.iloc[:, search_col].apply(normalize_code) == code]
                if not temp_row.empty:
                    c_row = temp_row
                    break
                    
            if not c_row.empty:
                duty_now = str(c_row.iloc[0, col_idx])
                if any(k in duty_now for k in ["休", "假", "補"]):
                    cadre_notes.append(f"{c_name}休假")
                else:
                    slots = []
                    for c_col in range(4, 16):
                        if "巡邏" in str(c_row.iloc[0, c_col]):
                            h = (c_col - 4) * 2
                            slots.append(f"{h:02d}-{h+2:02d}")
                    if slots:
                        time_range = f"{slots[0][:2]}至{slots[-1][-2:]}"
                        cadre_notes.append(f"{c_name}在所督勤，編排{time_range}時段巡邏勤務")
                    else:
                        cadre_notes.append(f"{c_name}在所督勤")
        
        if cadre_notes:
            res['cadre_status'] = "；".join(cadre_notes) + "。"
            
    except Exception as e:
        res['v_name'] = "解析失敗"
        res['cadre_status'] = f"幹部解析錯誤: {e}"

    # 2. 解析裝備交接簿 (完美版維持不變)
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
    except:
        res['eq'] = None
        
    return res

# --- 主畫面執行 ---
if duty_file and equip_file:
    with st.spinner("啟動全域雷達掃描，正在比對勤務番號與裝備數據..."):
        data = extract_full_logic(duty_file, equip_file, target_hour)
    
    st.success("✅ 數據擷取與番號比對完成！")
    
    c1, c2 = st.columns(2)
    with c1:
        check_mon = st.checkbox("駐地監錄/天羅地網正常", value=True)
        check_edu = st.checkbox("勤前教育宣導落實", value=True)
    with c2:
        check_env = st.checkbox("環境內務擺設整齊", value=True)
        check_alc = st.checkbox("酒測聯單無跳號", value=True)

    lines = []
    
    # 1. 值班
    lines.append(f"{time_str}，該所值班警員{data['v_name']}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。")
    
    # 2. 監錄
    if check_mon:
        lines.append(f"該所駐地監錄設備及天羅地網系統均運作正常，無故障，{d_m5}至{d_m1}有逐日檢測2次以上紀錄。")
    
    # 3. 勤教
    if check_edu:
        lines.append(f"該所{d_m3}至{d_m1}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。")
    
    # 4. 內務
    if check_env:
        lines.append(f"該所環境內務擺設整齊清潔，符合規定。")
        
    # 5. 裝備
    eq = data['eq'] if data['eq'] else {"gi":0, "go":0, "gf":0, "bi":0, "bo":0, "bf":0, "ri":0, "ro":0, "rf":0, "vi":0, "vo":0, "vf":0}
    fix_str = f"（另有槍枝 {eq['gf']} 把、無線電 {eq['rf']} 臺送修中）" if (eq['gf'] + eq['rf']) > 0 else ""
    lines.append(f"該所手槍出勤 {eq['go']} 把、在所 {eq['gi']} 把，子彈出勤 {eq['bo']} 顆、在所 {eq['bi']} 顆，無線電出勤 {eq['ro']} 臺、在所 {eq['ri']} 臺；防彈背心出勤 {eq['vo']} 件、在所 {eq['vi']} 件，幹部對械彈每日檢查管制良好，符合規定{fix_str}。")
    
    # 6. 幹部動態
    lines.append(f"本日{data['cadre_status']}")
    
    # 7. 酒測
    if check_alc:
        lines.append(f"該所酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。")

    # 最終生成
    final_report = "\n".join([f"{i+1}、{line}" for i, line in enumerate(lines)])
    
    st.markdown("---")
    st.subheader("📋 最終督導報告 (自動對位版)")
    st.text_area("複製回貼公務系統：", value=final_report, height=450)
    
    with st.expander("🛠️ 查看系統自動辨識結果 (完美對位區)"):
        st.write(f"偵測時段欄位：Index {4 + (target_hour // 2)}")
        st.write(f"✅ 完美過濾的番號對照表：")
        st.json(data['debug_map'])
        st.write(f"值班員警：{data['v_name']}")
        st.write(f"幹部動態：{data['cadre_status']}")
        st.write(f"裝備數據：{eq}")

else:
    st.info("👋 匯入兩份檔案後，系統將自動啟動番號雷達掃描！")
