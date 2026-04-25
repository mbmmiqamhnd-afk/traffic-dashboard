import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re
import traceback

# --- 1. 分頁基本配置 ---
st.set_page_config(page_title="督導報告 v8.3 - 督導日自動推算版", layout="wide")

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
    .stTabs [data-baseweb="tab-list"] button {{
        font-size: 18px;
        font-weight: bold;
    }}
    </style>
    """, unsafe_allow_html=True)

st.title("📋 督導報告極速生成器 v8.3 (督導日自動推算)")

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
        col_map = {"gun": None, "bullet": None, "radio": None, "vest": None}
        for r in range(min(10, len(df))):
            for c in range(len(df.columns)):
                val_c = str(df.iloc[r, c]).replace(" ", "").replace("　", "").replace("\n", "")
                if col_map["gun"] is None and ("手槍" in val_c or ("槍" in val_c and "手" in val_c)): col_map["gun"] = c
                if col_map["bullet"] is None and ("子彈" in val_c or ("彈" in val_c and "子" in val_c)): col_map["bullet"] = c
                if col_map["radio"] is None and "無線電" in val_c: col_map["radio"] = c
                if col_map["vest"] is None and ("防彈背心" in val_c or "防彈衣" in val_c): col_map["vest"] = c
        col_map["gun"] = col_map["gun"] if col_map["gun"] is not None else 2
        col_map["bullet"] = col_map["bullet"] if col_map["bullet"] is not None else 3
        col_map["radio"] = col_map["radio"] if col_map["radio"] is not None else 6
        col_map["vest"] = col_map["vest"] if col_map["vest"] is not None else 11
        stop_row = len(df)
        for r_idx in range(min(10, len(df)), len(df)):
            t_val = str(df.iloc[r_idx, 0])
            nums = re.findall(r'\d{1,2}', t_val)
            if nums:
                row_h = int(nums[0])
                if row_h > hour and (row_h - hour < 12):
                    stop_row = r_idx; break
        df_sub = df.iloc[:stop_row]
        df_sub_s = df_s.iloc[:stop_row]
        def get_v(keyword, equip_key):
            rows = df_sub[df_sub_s.iloc[:, 1].str.contains(keyword, na=False)]
            if not rows.empty: return safe_int(rows.iloc[-1, col_map[equip_key]])
            return 0
        return {
            "gi": get_v("在", "gun"), "go": get_v("出", "gun"), "gf": get_v("送", "gun"),
            "bi": get_v("在", "bullet"), "bo": get_v("出", "bullet"), "bf": get_v("送", "bullet"),
            "ri": get_v("在", "radio"), "ro": get_v("出", "radio"), "rf": get_v("送", "radio"),
            "vi": get_v("在", "vest"), "vo": get_v("出", "vest"), "vf": get_v("送", "vest")
        }
    except: return None

# --- 5. 勤務與單位名稱解析 ---
def extract_duty_logic(d_file, hour):
    res = {'v_name': '未偵測', 'cadre_status': '無幹部資料', 'unit_name': '未偵測單位'}
    try:
        df = pd.read_excel(d_file, header=None, dtype=str).fillna("")
        for r in range(5):
            row_text = "".join([str(x) for x in df.iloc[r].values])
            match = re.search(r'([\u4e00-\u9fa5]+(分局|派出所|分隊|局))', row_text)
            if match: res['unit_name'] = match.group(1); break
        full_text = " ".join([str(x).strip() for x in df.values.flatten() if x])
        pattern = r'(?<![A-Za-z0-9])([A-Z]|[0-9]{1,2})\s*(所長|副所長|巡官|巡佐|警員|實習)\s*([\u4e00-\u9fa5]{2,4})'
        matches = re.findall(pattern, full_text)
        name_map, full_title_map = {}, {}
        for m in matches:
            code, title, name = normalize_code(m[0]), m[1].strip(), m[2]
            for t_char in ["所", "副", "巡", "警", "實", "員", "長"]:
                if name.endswith(t_char): name = name[:-1]
            if len(name) >= 2: name_map[code] = name; full_title_map[code] = f"{title}{name}" 
        time_row_idx, time_cols, target_col_idx = 2, {}, -1
        for r in range(5):
            count, temp_cols = 0, {}
            for c_idx in range(len(df.columns)):
                start_h, end_h = parse_time_cell(df.iloc[r, c_idx])
                if start_h is not None and end_h is not None: temp_cols[c_idx] = (start_h, end_h); count += 1
            if count > max(len(time_cols), 3): time_row_idx = r; time_cols = temp_cols
        for c_idx, (start_h, end_h) in time_cols.items():
            calc_end = end_h if end_h > start_h else end_h + 24
            calc_hour = hour if hour >= 6 or calc_end <= 24 else hour + 24
            if start_h <= calc_hour < calc_end: target_col_idx = c_idx
        v_row_idx = time_row_idx + 1 
        for r in range(time_row_idx + 1, min(time_row_idx + 4, len(df))):
            if "值" in str(df.iloc[r, 0]) + str(df.iloc[r, 1]): v_row_idx = r; break
        if target_col_idx != -1:
            raw_cell = str(df.iloc[v_row_idx, target_col_idx]).strip()
            m_code = re.search(r'[A-Za-z0-9]{1,2}', raw_cell)
            if m_code:
                code = normalize_code(m_code.group(0)); res['v_name'] = full_title_map.get(code, f"未建檔:{code}")
            else:
                for code, name in name_map.items():
                    if name in raw_cell: res['v_name'] = full_title_map.get(code, name); break
        cadre_notes = []
        for code in ["A", "B", "C"]:
            name = name_map.get(code, {"A":"所長", "B":"副所長", "C":"幹部"}.get(code))
            is_off, patrol_slots, duty_names, found_in_matrix = False, [], set(), False
            for r in range(v_row_idx, len(df)):
                dt = str(df.iloc[r, 0]) + str(df.iloc[r, 1])
                is_leave = any(k in dt for k in ["休", "假", "輪", "輸", "補", "外"])
                for c_idx, (s_h, e_h) in time_cols.items():
                    cell_val = str(df.iloc[r, c_idx])
                    cell_codes = [normalize_code(x) for x in re.findall(r'[A-Za-z0-9]{1,2}', cell_val)]
                    if code in cell_codes:
                        found_in_matrix = True
                        if is_leave: is_off = True
                        else:
                            is_ext = False
                            if "巡" in dt or "巡" in cell_val: duty_names.add("巡邏"); is_ext = True
                            if "守" in dt or "守" in cell_val: duty_names.add("守望"); is_ext = True
                            if "臨" in dt or "臨" in cell_val: duty_names.add("臨檢"); is_ext = True
                            if "交" in dt or "交" in cell_val: duty_names.add("交整"); is_ext = True
                            if "路" in dt or "路" in cell_val: duty_names.add("路檢"); is_ext = True
                            if is_ext: patrol_slots.append((s_h, e_h))
            if not found_in_matrix or is_off: cadre_notes.append(f"{name}休假")
            else:
                if patrol_slots:
                    min_s, max_e = min([s[0] for s in patrol_slots]), max([s[1] for s in patrol_slots])
                    e_str = "24" if max_e == 24 or max_e == 0 else f"{max_e%24:02d}"
                    cadre_notes.append(f"{name}在所督勤，編排{min_s:02d}至{e_str}時段{'、'.join(sorted(list(duty_names)))}勤務")
                else: cadre_notes.append(f"{name}在所督勤")
        if cadre_notes: res['cadre_status'] = "；".join(cadre_notes) + "。"
    except Exception as e: res['v_name'] = "解析失敗"; res['cadre_status'] = f"錯誤: {e}"
    return res

# --- 6. 側邊欄設定 ---
with st.sidebar:
    st.header("⚙️ 督導任務設定")
    num_units = st.number_input("本次督導單位數量", min_value=1, max_value=8, value=3)
    st.divider()
    
    st.subheader("📅 督導日期設定")
    # 只需選擇督導日期 (起始日)
    insp_date = st.date_input("選擇督導日期 (起始日)", datetime.now())
    
    # --- 自動推算邏輯 ---
    # 監錄系統通常查到昨天為止 (若督導日是 26 號，查 21~25 號)
    date_end = insp_date - timedelta(days=1)
    date_start_5d = insp_date - timedelta(days=5)
    date_start_3d = insp_date - timedelta(days=3)
    
    # 格式化
    date_end_str = date_end.strftime('%m月%d日')
    date_start_5d_str = date_start_5d.strftime('%m月%d日')
    date_start_3d_str = date_start_3d.strftime('%m月%d日')
    
    st.write(f"系統將自動帶入：")
    st.write(f"🔹 監錄區間：{date_start_5d_str}-{date_end_str}")
    st.write(f"🔹 勤教區間：{date_start_3d_str}-{date_end_str}")

# --- 7. 動態分頁處理 ---
tab_names = [f"🏢 單位 {i+1}" for i in range(num_units)] + ["📄 總匯整報告"]
tabs = st.tabs(tab_names)
all_final_reports = [] 

for i in range(num_units):
    with tabs[i]:
        st.subheader(f"第 {i+1} 站督導數據匯入")
        target_time = st.time_input("抵達督導時間", datetime.now().time(), key=f"time_{i}")
        time_str, target_hour = target_time.strftime('%H%M'), target_time.hour
        
        col_f1, col_f2 = st.columns(2)
        with col_f1: duty_file = st.file_uploader(f"1. 上傳勤務表", type=['xlsx'], key=f"duty_{i}")
        with col_f2: equip_file = st.file_uploader(f"2. 上傳裝備交接簿", type=['xlsx'], key=f"eq_{i}")

        st.write("💡 該單位檢查狀況：")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            check_mon = st.checkbox("駐地監錄/天羅地網正常", value=True, key=f"c_mon_{i}")
            check_edu = st.checkbox("勤前教育宣導落實", value=True, key=f"c_edu_{i}")
        with col_c2:
            check_env = st.checkbox("環境內務擺設整齊", value=True, key=f"c_env_{i}")
            check_alc = st.checkbox("酒測聯單無跳號", value=True, key=f"c_alc_{i}")

        if duty_file and equip_file:
            duty_data = extract_duty_logic(duty_file, target_hour)
            eq_data = extract_equip_dynamic(equip_file, target_hour)
            u_name = duty_data['unit_name']
            
            lines = []
            lines.append(f"{time_str}，該所值班{duty_data['v_name']}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。")
            
            # 使用自動推算的日期
            if check_mon:
                lines.append(f"該所駐地監錄設備及天羅地網系統均運作正常，無故障，{date_start_5d_str}至{date_end_str}有逐日檢測2次以上紀錄。")
            if check_edu:
                lines.append(f"該所{date_start_3d_str}至{date_end_str}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。")
            
            if check_env: lines.append(f"該所環境內務擺設整齊清潔，符合規定。")
            eq = eq_data if eq_data else {"gi":0, "go":0, "gf":0, "bi":0, "bo":0, "bf":0, "ri":0, "ro":0, "rf":0, "vi":0, "vo":0, "vf":0}
            fix_str = f"（另有槍枝 {eq['gf']} 把、無線電 {eq['rf']} 臺送修中）" if (eq['gf'] + eq['rf']) > 0 else ""
            lines.append(f"該所手槍出勤 {eq['go']} 把、在所 {eq['gi']} 把，子彈出勤 {eq['bo']} 顆、在所 {eq['bi']} 顆，無線電出勤 {eq['ro']} 臺、在所 {eq['ri']} 臺；防彈背心出勤 {eq['vo']} 件、在所 {eq['vi']} 件，幹部對械彈每日檢查管制良好，符合規定{fix_str}。")
            lines.append(f"本日{duty_data['cadre_status']}")
            if check_alc: lines.append(f"該所酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。")

            final_report = "\n".join([f"{idx+1}、{line}" for idx, line in enumerate(lines)])
            all_final_reports.append(f"【{u_name} 督導報告】\n{final_report}")
            st.success(f"✅ {u_name} | 解析完畢！")
            st.text_area("預覽：", value=final_report, height=300, key=f"text_{i}")

# --- 8. 總匯整 ---
with tabs[-1]:
    st.subheader("📄 總匯整報告區")
    if all_final_reports:
        combined_text = "\n\n----------------------------------------\n\n".join(all_final_reports)
        st.text_area("點擊下方全選複製：", value=combined_text, height=600)
    else:
        st.info("⏳ 請先前往各單位分頁上傳檔案。")
