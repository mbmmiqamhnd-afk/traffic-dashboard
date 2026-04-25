import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# 分頁配置
st.set_page_config(page_title="督導報告 v7.0 - 座標精準版", layout="wide")

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

st.title("📋 督導報告極速生成器 v7.0 (裝備座標對焦版)")

# --- 側邊欄設定 ---
with st.sidebar:
    st.header("📂 數據源匯入")
    duty_file = st.file_uploader("1. 上傳『勤務分配表』", type=['xlsx', 'csv'])
    equip_file = st.file_uploader("2. 上傳『值班裝備交接簿』", type=['xlsx', 'csv'])
    
    st.divider()
    target_time = st.time_input("督導時間", datetime.now().time())
    time_str = target_time.strftime('%H%M')
    target_hour = target_time.hour
    
    today = datetime.now()
    d_m5, d_m1 = [(today - timedelta(days=i)).strftime('%m月%d日') for i in [5, 1]]
    d_m3 = (today - timedelta(days=3)).strftime('%m月%d日')

# --- 核心工具 ---
def safe_int(val):
    try: return int(float(str(val).split('.')[0].replace(',', '')))
    except: return 0

def normalize_code(c):
    c_str = str(c).strip().replace(".0", "").upper()
    if c_str.isdigit(): return str(int(c_str))
    return c_str

# --- 裝備座標解析引擎 ---
def extract_equip_dynamic(e_file, hour):
    try:
        df = pd.read_csv(e_file, header=None) if e_file.name.endswith('csv') else pd.read_excel(e_file, header=None)
        df_s = df.astype(str)
        
        # 1. 在第 3 列 (Index 2) 尋找各裝備的「欄位 Index」
        header_row = df_s.iloc[2]
        col_map = {
            "gun": 2,    # 預設
            "bullet": 3, # 預設
            "radio": 6,  # 預設
            "vest": 11   # 預設
        }
        for c_idx, val in enumerate(header_row):
            val_c = val.replace(" ", "").replace("\n", "")
            if "槍" in val_c and "手" in val_c: col_map["gun"] = c_idx
            if "彈" in val_c and "子" in val_c: col_map["bullet"] = c_idx
            if "無線電" in val_c: col_map["radio"] = c_idx
            if "背心" in val_c or "防彈衣" in val_c: col_map["vest"] = c_idx

        # 2. 在第 A 欄 (Index 0) 尋找符合時間的「截止列 Row」
        stop_row = len(df)
        for r_idx in range(len(df)):
            t_val = df_s.iloc[r_idx, 0]
            nums = re.findall(r'\d{1,2}', t_val)
            if nums:
                row_h = int(nums[0])
                if row_h > hour and (row_h - hour < 12):
                    stop_row = r_idx
                    break
        
        # 3. 擷取該時間區間內最後的 在所/出勤/送修 狀態
        df_sub = df.iloc[:stop_row]
        df_sub_s = df_s.iloc[:stop_row]
        
        def get_v(keyword, equip_key):
            rows = df_sub[df_sub_s.iloc[:, 1].str.contains(keyword, na=False)]
            if not rows.empty:
                return safe_int(rows.iloc[-1, col_map[equip_key]])
            return 0

        return {
            "gi": get_v("在", "gun"), "go": get_v("出", "gun"), "gf": get_v("送", "gun"),
            "bi": get_v("在", "bullet"), "bo": get_v("出", "bullet"), "bf": get_v("送", "bullet"),
            "ri": get_v("在", "radio"), "ro": get_v("出", "radio"), "rf": get_v("送", "radio"),
            "vi": get_v("在", "vest"), "vo": get_v("出", "vest"), "vf": get_v("送", "vest")
        }
    except: return None

# --- 勤務邏輯解析 (矩陣定位) ---
def extract_duty_logic(d_file, hour):
    # 此處保留您之前成功的矩陣解析邏輯 (省略重複代碼以節省篇幅)
    # ... 包含時段對位與幹部動態解析 ...
    pass

# --- 畫面執行與組合 ---
# (此處呼叫 extract_equip_dynamic 並產出報告)
