import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# 分頁配置
st.set_page_config(page_title="督導報告 v7.0 - 代號比對版", layout="wide")

# 套用標楷體風格 (引用您根目錄下的 kaiu.ttf)
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

st.title("📋 督導報告極速生成器 v7.0 (全自動代號比對)")

# --- 側邊欄：檔案與時間設定 ---
with st.sidebar:
    st.header("📂 數據源匯入")
    duty_file = st.file_uploader("1. 上傳『勤務分配表』", type=['xlsx'])
    equip_file = st.file_uploader("2. 上傳『值班裝備交接簿』", type=['xlsx', 'csv'])
    
    st.divider()
    target_time = st.time_input("督導時間 (自動對位時段)", datetime.now().time())
    time_str = target_time.strftime('%H%M')
    target_hour = target_time.hour
    
    # 日期推算 (用於報告語句)
    today = datetime.now()
    d_m5, d_m1 = [(today - timedelta(days=i)).strftime('%m月%d日') for i in [5, 1]]
    d_m3 = (today - timedelta(days=3)).strftime('%m月%d日')

# --- 核心解析引擎 ---
def safe_int(val):
    try:
        return int(float(str(val).split('.')[0].replace(',', '')))
    except:
        return 0

def extract_full_logic(d_file, e_file, hour):
    res = {}
    try:
        # 1. 讀取勤務分配表
        df = pd.read_excel(d_file, header=None)
        
        # 🌟 A. 建立人員對照表 (代號 -> 姓名)
        name_map = {}
        for i, row in df.iterrows():
            row_str = "".join([str(x) for x in row if pd.notna(x)])
            if "代號" in row_str and "姓名" in row_str:
                # 往後掃描對照區
                for j in range(i + 1, min(i + 30, len(df))):
                    sub_r = df.iloc[j]
                    # 警用格式通常是 [代號, 職稱, 姓名] 循環
                    for k in range(0, len(sub_r) - 2, 3):
                        code = str(sub_r[k]).strip().replace(".0", "")
                        name = str(sub_r[k+2]).strip()
                        if len(name) >= 2 and len(name) <= 4:
                            name_map[code] = name
        
        # 🌟 B. 定位主表時段
        df_main = df.head(45).ffill() 
        col_idx = 4 + (hour // 2) # E欄=00-02, F欄=02-04...
        
        # 🌟 C. 抓取值班人員 (代號反查)
        duty_rows = df_main[df_main.iloc[:, col_idx].astype(str).str.contains("值班", na=False)]
        v_name = "未偵測"
        if not duty_rows.empty:
            # 抓取該列的代號 (通常在 B 欄 / Index 1)
            row_code = str(duty_rows.iloc[0, 1]).strip().replace(".0", "")
            v_name = name_map.get(row_code, "未對應姓名")
        res['v_name'] = v_name
        
        # 🌟 D. 抓取幹部動態 (代號 A, B, C)
        cadre_notes = []
        for code in ["A", "B", "C"]:
            name = name_map.get(code)
            if name:
                # 找主表中該代號的列
                c_row = df_main[df_main.iloc[:, 1].astype(str).str.replace(".0", "") == code]
                if not c_row.empty:
                    duty_now = str(c_row.iloc[0, col_idx])
                    if any(k in duty_now for k in ["休", "假", "補"]):
                        cadre_notes.append(f"{name}休假")
                    else:
                        # 掃描全天巡邏班表
                        slots = []
                        for c in range(4, 16):
                            if "巡邏" in str(c_row.iloc[0, c]):
                                h = (c - 4) * 2
                                slots.append(f"{h:02d}-{h+2:02d}")
                        if slots:
                            time_range = f"{slots[0][:2]}至{slots[-1][-2:]}"
                            cadre_notes.append(f"{name}在所督勤，編排{time_range}時段巡邏勤務")
                        else:
                            cadre_notes.append(f"{name}在所督勤")
        res['cadre_status'] = "；".join(cadre_notes) + "。"
    except Exception as e:
        res['v_name'], res['cadre_status'] = "解析失敗", f"幹部解析錯誤: {e}"

    # 2. 解析裝備交接簿 (依照您給的數據座標)
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
    with st.spinner("正在執行代號比對與自動解析..."):
        data = extract_full_logic(duty_file, equip_file, target_hour)
    
    st.success("✅ 全自動解析完成！")
    
    # 勾選事項
    c1, c2 = st.columns(2)
    with c1:
        check_mon = st.checkbox("駐地監錄/天羅地網正常", value=True)
        check_edu = st.checkbox("勤前教育宣導落實", value=True)
    with c2:
        check_env = st.checkbox("環境內務擺設整潔", value=True)
        check_alc = st.checkbox("酒測聯單無跳號", value=True)

    # 組合文字 (完全還原您的範例格式)
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
    st.subheader("📋 最終督導報告 (標楷體)")
    st.text_area("直接複製到公務系統：", value=final_report, height=450)
    
    with st.expander("🛠️ 查看後台自動辨識細節 (Debug)"):
        st.write(f"偵測時段欄位：Index {4 + (target_hour // 2)}")
        st.write(f"值班員警姓名：{data['v_name']}")
        st.write(f"幹部動態內容：{data['cadre_status']}")
        st.write(f"裝備數據：{eq}")

else:
    st.info("👋 請上傳檔案，系統會自動在下方尋找人名對照表進行比對。")
