import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re
import traceback

# --- 1. 分頁基本配置 ---
st.set_page_config(page_title="督導報告 v7.0 - 職稱精準版", layout="wide")

# --- 2. 套用標楷體風格 ---
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

st.title("📋 督導報告極速生成器 v7.0 (職稱動態對位版)")

# --- 3. 核心輔助工具 ---
def safe_int(val):
    try: return int(float(str(val).split('.')[0].replace(',', '')))
    except: return 0

def normalize_code(c):
    c_str = str(c).strip().replace(".0", "").upper()
    if c_str.isdigit(): return str(int(c_str))
    return c_str

def parse_time_cell(val):
    val_str = str(val).strip().replace("\n", "").replace(" ", "").replace("|", "-")
    if val_str in ["", "nan", "NaN"]: return None, None
    m_date = re.search(r'20\d{2}-(\d{2})-(\d{2})', val_str)
    if m_date: return int(m_date.group(1)), int(m_date.group(2))
    m_time = re.search(r'(\d{1,2})[~~\-－—–_]+(\d{1,2})', val_str)
    if m_time: return int(m_time.group(1)), int(m_time.group(2))
    return None, None

# --- 4. 裝備座標解析引擎 ---
def extract_equip_dynamic(e_file, hour):
    try:
        df = pd.read_csv(e_file, header=None) if e_file.name.endswith('csv') else pd.read_excel(e_file, header=None)
        df_s = df.astype(str)
        
        col_map = {"gun": 2, "bullet": 3, "radio": 6, "vest": 11} 
        
        for r in range(min(10, len(df))):
            for c in range(len(df.columns)):
                val_c = str(df.iloc[r, c]).replace(" ", "").replace("　", "").replace("\n", "")
                if "手槍" in val_c or ("槍" in val_c and "手" in val_c): col_map["gun"] = c
                if "子彈" in val_c or ("彈" in val_c and "子" in val_c): col_map["bullet"] = c
                if "無線電" in val_c: col_map["radio"] = c
                if "防彈背心" in val_c or "防彈衣" in val_c: col_map["vest"] = c

        stop_row = len(df)
        for r_idx in range(min(10, len(df)), len(df)):
            t_val = str(df.iloc[r_idx, 0])
            nums = re.findall(r'\d{1,2}', t_val)
            if nums:
                row_h = int(nums[0])
                if row_h > hour and (row_h - hour < 12):
                    stop_row = r_idx
                    break
        
        df_sub = df.iloc[:stop_row]
        df_sub_s = df_s.iloc[:stop_row]
        
        def get_v(keyword, equip_key):
            rows = df_sub[df_sub_s.iloc[:, 1].str.contains(keyword, na=False)]
            if not rows.empty: return safe_int(rows.iloc[-1, col_map[equip_key]])
            return 0

        result = {
            "gi": get_v("在", "gun"), "go": get_v("出", "gun"), "gf": get_v("送", "gun"),
            "bi": get_v("在", "bullet"), "bo": get_v("出", "bullet"), "bf": get_v("送", "bullet"),
            "ri": get_v("在", "radio"), "ro": get_v("出", "radio"), "rf": get_v("送", "radio"),
            "vi": get_v("在", "vest"), "vo": get_v("出", "vest"), "vf": get_v("送", "vest")
        }
        result["debug_map"] = col_map 
        return result
    except Exception as e:
        return {"gi":0, "go":0, "gf":0, "bi":0, "bo":0, "bf":0, "ri":0, "ro":0, "rf":0, "vi":0, "vo":0, "vf":0, "debug_map": f"錯誤: {e}"}

# --- 5. 勤務邏輯解析 (職稱還原版) ---
def extract_duty_logic(d_file, hour):
    res = {'v_name': '未偵測', 'cadre_status': '無幹部資料', 'debug_info': {}}
    try:
        df = pd.read_excel(d_file, header=None, dtype=str).fillna("")
        
        full_text = " ".join([str(x).strip() for x in df.values.flatten() if x])
        pattern = r'(?<![A-Za-z0-9])([A-Z]|[0-9]{1,2})\s*(所長|副所長|巡官|巡佐|警員|實習)\s*([\u4e00-\u9fa5]{2,4})'
        matches = re.findall(pattern, full_text)
        
        name_map = {}
        full_title_map = {} # 🌟 新增：儲存包含職稱的完整姓名
        
        for m in matches:
            code = normalize_code(m[0])
            title = m[1].strip()
            name = m[2]
            for t_char in ["所", "副", "巡", "警", "實", "員", "長"]:
                if name.endswith(t_char): name = name[:-1]
            if len(name) >= 2: 
                name_map[code] = name
                full_title_map[code] = f"{title}{name}" # 例如：巡佐傅錫城
                
        res['debug_info']['對照表'] = full_title_map

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
                # 🌟 使用帶有職稱的名字
                res['v_name'] = full_title_map.get(code, f"未建檔:{code}")
            else:
                for code, name in name_map.items():
                    if name in raw_cell: res['v_name'] = full_title_map.get(code, name); break

        cadre_notes = []
        for code in ["A", "B", "C"]:
            # 幹部動態維持使用純姓名 (例如：鄭榮捷休假)，較符合公文語意
            name = name_map.get(code, {"A":"所長", "B":"副所長", "C":"幹部"}.get(code))
            is_off = False
            patrol_slots = []
            duty_names = set() 
            found_in_matrix = False

            for r in range(v_row_idx, len(df)):
                duty_title = str(df.iloc[r, 0]) + str(df.iloc[r, 1])
                is_leave_row = any(k in duty_title for k in ["休", "假", "輪", "輸", "補", "外"])

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
        
    return res

# --- 6. 側邊欄與畫面執行區 ---
with st.sidebar:
    st.header("📂 數據源匯入")
    duty_file = st.file_uploader("1. 上傳『勤務分配表』", type=['xlsx', 'csv'])
    equip_file = st.file_uploader("2. 上傳『值班裝備交接簿』", type=['xlsx', 'csv'])
    
    st.divider()
    target_time = st.time_input("督導時間 (自動對位)", datetime.now().time())
    time_str = target_time.strftime('%H%M')
    target_hour = target_time.hour
    
    today = datetime.now()
    d_m5, d_m1 = [(today - timedelta(days=i)).strftime('%m月%d日') for i in [5, 1]]
    d_m3 = (today - timedelta(days=3)).strftime('%m月%d日')

if duty_file and equip_file:
    try:
        with st.spinner("🚀 啟動導彈鎖定雷達，掃描裝備座標..."):
            duty_data = extract_duty_logic(duty_file, target_hour)
            eq_data = extract_equip_dynamic(equip_file, target_hour)
            
        st.success("✅ 報告與裝備數據自動對位完成！")
        
        st.write("💡 **勾選欲加入的報告內容 (順序自動編排)：**")
        c1, c2 = st.columns(2)
        with c1:
            check_mon = st.checkbox("駐地監錄/天羅地網正常", value=True)
            check_edu = st.checkbox("勤前教育宣導落實", value=True)
        with c2:
            check_env = st.checkbox("環境內務擺設整齊", value=True)
            check_alc = st.checkbox("酒測聯單無跳號", value=True)

        lines = []
        
        # 🌟 這裡改成動態帶入職稱，例如「該所值班警員陳秉贏」或「該所值班巡佐傅錫城」
        lines.append(f"{time_str}，該所值班{duty_data['v_name']}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。")
        
        if check_mon: lines.append(f"該所駐地監錄設備及天羅地網系統均運作正常，無故障，{d_m5}至{d_m1}有逐日檢測2次以上紀錄。")
        if check_edu: lines.append(f"該所{d_m3}至{d_m1}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。")
        if check_env: lines.append(f"該所環境內務擺設整齊清潔，符合規定。")
            
        eq = eq_data if eq_data else {"gi":0, "go":0, "gf":0, "bi":0, "bo":0, "bf":0, "ri":0, "ro":0, "rf":0, "vi":0, "vo":0, "vf":0}
        fix_str = f"（另有槍枝 {eq['gf']} 把、無線電 {eq['rf']} 臺送修中）" if (eq['gf'] + eq['rf']) > 0 else ""
        lines.append(f"該所手槍出勤 {eq['go']} 把、在所 {eq['gi']} 把，子彈出勤 {eq['bo']} 顆、在所 {eq['bi']} 顆，無線電出勤 {eq['ro']} 臺、在所 {eq['ri']} 臺；防彈背心出勤 {eq['vo']} 件、在所 {eq['vi']} 件，幹部對械彈每日檢查管制良好，符合規定{fix_str}。")
        
        lines.append(f"本日{duty_data['cadre_status']}")
        if check_alc: lines.append(f"該所酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。")

        final_report = "\n".join([f"{i+1}、{line}" for i, line in enumerate(lines)])
        
        st.markdown("---")
        st.subheader("📋 最終督導報告 (職稱精準版)")
        st.text_area("複製回貼公務系統：", value=final_report, height=450)
        
        with st.expander("🛠️ 查看系統自動辨識結果 (除錯專區)"):
            st.write(f"✅ 裝備欄位偵測結果 (Index號)：{eq_data.get('debug_map', '未取得')}")
            st.write(f"✅ 值班人員：{duty_data['v_name']}")
            st.write(f"✅ 幹部動態：{duty_data['cadre_status']}")
            
    except Exception as e:
        st.error("系統發生錯誤，請確認檔案格式。")
        st.write(traceback.format_exc())
else:
    st.info("👋 請於左側上傳今日 Excel 檔案，系統將自動提取所有督導數據。")
