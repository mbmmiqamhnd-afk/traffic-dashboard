import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# 分頁配置
st.set_page_config(page_title="督導報告 v7.0 - 邏輯修正版", layout="wide")

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
    }}
    </style>
    """, unsafe_allow_html=True)

st.title("📋 督導報告極速生成器 v7.0 (對照表比對版)")

# --- 側邊欄設定 ---
with st.sidebar:
    st.header("📂 數據源匯入")
    duty_file = st.file_uploader("1. 上傳『勤務分配表』", type=['xlsx', 'csv'])
    equip_file = st.file_uploader("2. 上傳『值班裝備交接簿』", type=['xlsx', 'csv'])
    
    st.divider()
    target_time = st.time_input("督導時間 (自動對位班表)", datetime.now().time())
    time_str = target_time.strftime('%H%M')
    target_hour = target_time.hour
    
    # 日期推算
    today = datetime.now()
    d_m5, d_m1 = [(today - timedelta(days=i)).strftime('%m月%d日') for i in [5, 1]]
    d_m3 = (today - timedelta(days=3)).strftime('%m月%d日')

# --- 核心解析引擎 (重點：對照表比對) ---
def extract_duty_logic(d_file, hour):
    try:
        # 1. 讀取原始資料
        df = pd.read_csv(d_file, header=None) if d_file.name.endswith('csv') else pd.read_excel(d_file, header=None)
        
        # 🌟 步驟 A：建立人名對照表 (代號 -> 姓名)
        # 搜尋包含「代號」與「姓名」的列，通常在班表下方
        name_map = {}
        for idx, row in df.iterrows():
            row_vals = [str(x).replace(" ", "").strip() for x in row.values if pd.notna(x)]
            row_str = "".join(row_vals)
            # 偵測到對照表區域
            if "代號" in row_str and "姓名" in row_str:
                # 往後掃描 20 列來抓取對照資料
                for sub_idx in range(idx + 1, min(idx + 25, len(df))):
                    sub_row = df.iloc[sub_idx].values
                    # 警用格式通常是 [代號, 職稱, 姓名] 這樣三格一組，重複出現
                    for i in range(0, len(sub_row) - 2, 3):
                        code = str(sub_row[i]).strip().replace(".0", "")
                        name = str(sub_row[i+2]).strip()
                        if code.isdigit() and len(name) >= 2 and len(name) <= 4:
                            name_map[code] = name
        
        # 🌟 步驟 B：處理主表 (填補合併儲存格)
        df_main = df.head(40).ffill() # 只處理上方主表區
        col_idx = 4 + (hour // 2) # 定位時段欄位 (E欄=00-02)
        
        # 🌟 步驟 C：尋找值班人員
        # 先找出在該時段欄位填有「值班」的「列」，再對應回第一欄的「代號」
        v_name = "未偵測"
        duty_rows = df_main[df_main.iloc[:, col_idx].astype(str).str.contains("值班", na=False)]
        for _, r in duty_rows.iterrows():
            # 抓取該列的代號 (通常在 A 欄或 B 欄)
            row_code = str(r.iloc[1]).strip().replace(".0", "") if not pd.isna(r.iloc[1]) else str(r.iloc[0]).strip().replace(".0", "")
            if row_code in name_map:
                v_name = name_map[row_code]
                break
        
        # 🌟 步驟 D：幹部動態比對
        cadre_notes = []
        for c_name in ["鄭榮捷", "邱品淳", "曹培翔"]:
            # 先找幹部的代號
            c_code = next((code for code, name in name_map.items() if c_name in name), None)
            if c_code:
                # 到主表找該代號的列
                c_row = df_main[df_main.iloc[:, 1].astype(str).str.replace(".0","") == c_code]
                if not c_row.empty:
                    duty_now = str(c_row.iloc[0, col_idx])
                    if any(k in duty_now for k in ["休", "假", "補"]):
                        cadre_notes.append(f"{c_name}休假")
                    else:
                        slots = []
                        for c in range(4, 16):
                            if "巡邏" in str(c_row.iloc[0, c]):
                                h = (c-4)*2
                                slots.append(f"{h:02d}-{h+2:02d}")
                        if slots:
                            cadre_notes.append(f"{c_name}在所督勤，編排{slots[0][:2]}至{slots[-1][-2:]}時段巡邏勤務")
                        else:
                            cadre_notes.append(f"{c_name}在所督勤")
        
        return v_name, "；".join(cadre_notes) + "。"
    except Exception as e:
        return f"解析失敗", f"錯誤: {e}"

# --- 裝備解析引擎 (維持 100% 正確的完美邏輯) ---
def extract_equip_data(e_file):
    try:
        df = pd.read_csv(e_file, header=None) if e_file.name.endswith('csv') else pd.read_excel(e_file, header=None)
        df_str = df.astype(str)
        def get_v(row, idx): return int(float(str(row.iloc[idx]).split('.')[0].replace(',', '')))
        
        r_in = df[df_str.iloc[:, 1].str.contains("在", na=False)].iloc[-1]
        r_out = df[df_str.iloc[:, 1].str.contains("出", na=False)].iloc[-1]
        r_fix = df[df_str.iloc[:, 1].str.contains("送", na=False)].iloc[-1]
        
        return {
            "gi": get_v(r_in, 2), "go": get_v(r_out, 2), "gf": get_v(r_fix, 2),
            "bi": get_v(r_in, 3), "bo": get_v(r_out, 3), "bf": get_v(r_fix, 3),
            "ri": get_v(r_in, 6), "ro": get_v(r_out, 6), "rf": get_v(r_fix, 6),
            "vi": get_v(r_in, 11), "vo": get_v(r_out, 11), "vf": get_v(r_fix, 11)
        }
    except: return None

# --- 畫面生成 ---
if duty_file and equip_file:
    v_person, cadre_text = extract_duty_logic(duty_file, target_hour)
    eq = extract_equip_data(equip_file)
    
    if eq:
        st.success("✅ 數據擷取完成 (已完成代號比對)")
        
        # 報告勾選項目
        c1, c2 = st.columns(2)
        with c1:
            check_m = st.checkbox("監錄/天羅地網正常", value=True)
            check_e = st.checkbox("勤前教育宣導落實", value=True)
        with c2:
            check_v = st.checkbox("內務擺設整齊", value=True)
            check_a = st.checkbox("酒測聯單無跳號", value=True)

        # 組合文字
        lines = [
            f"{time_str}，該所值班警員{v_person}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。",
            f"該所駐地監錄設備及天羅地網系統均運作正常，無故障，{d_m5}至{d_m1}有逐日檢測2次以上紀錄。" if check_m else None,
            f"該所{d_m3}至{d_m1}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。" if check_e else None,
            f"該所環境內務擺設整齊清潔，符合規定。" if check_v else None,
            f"該所手槍出勤 {eq['go']} 把、在所 {eq['gi']} 把，子彈出勤 {eq['bo']} 顆、在所 {eq['bi']} 顆，無線電出勤 {eq['ro']} 臺、在所 {eq['ri']} 臺；防彈背心出勤 {eq['vo']} 件、在所 {eq['vi']} 件，幹部對械彈每日檢查管制良好，符合規定。" + (f"（另有槍枝 {eq['gf']} 把送修中）" if eq['gf']>0 else ""),
            f"本日{cadre_text}",
            f"該所酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。" if check_a else None
        ]
        
        # 清除 None 並加編號
        final_list = [l for l in lines if l is not None]
        final_report = "\n".join([f"{i+1}、{txt}" for i, txt in enumerate(final_list)])
        
        st.markdown("---")
        st.text_area("複製回貼公務系統：", value=final_report, height=450)
    else:
        st.error("裝備交接簿解析失敗。")
else:
    st.info("請上傳 Excel 檔案。系統會自動前往「下方」尋找人名對照表進行比對。")
