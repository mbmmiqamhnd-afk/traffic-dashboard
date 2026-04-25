import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# 分頁配置
st.set_page_config(page_title="督導報告 v7.0 - 深度自動化", layout="wide")

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

st.title("📋 督導報告極速生成器 v7.0 (全自動版)")

# --- 側邊欄設定 ---
with st.sidebar:
    st.header("📂 數據源匯入")
    duty_file = st.file_uploader("1. 上傳『勤務分配表』", type=['xlsx', 'csv'])
    equip_file = st.file_uploader("2. 上傳『值班裝備交接簿』", type=['xlsx', 'csv'])
    
    st.divider()
    target_time = st.time_input("督導時間 (自動對位班表)", datetime.now().time())
    time_str = target_time.strftime('%H%M')
    target_hour = target_time.hour
    
    # 日期推算 (用於報告語句)
    today = datetime.now()
    d_m5, d_m3, d_m1 = [(today - timedelta(days=i)).strftime('%m月%d日') for i in [5, 3, 1]]

# --- 核心解析引擎 ---
def safe_parse_int(val):
    try:
        return int(float(str(val).split('.')[0].replace(',', '')))
    except:
        return 0

def extract_all_data(d_file, e_file, hour):
    results = {}
    
    # 1. 解析【勤務分配表】
    try:
        # 讀取並填補合併儲存格
        df_d = pd.read_csv(d_file, header=None) if d_file.name.endswith('csv') else pd.read_excel(d_file, header=None)
        df_d = df_d.ffill()
        
        # 定位時段欄位：聖亭所格式通常 E 欄 (Index 4) 是 00-02，以此類推
        # 每 2 小時一格，所以 (hour // 2) 決定位移
        col_idx = 4 + (hour // 2)
        
        # 搜尋「值班」人員
        duty_row = df_d[df_d.iloc[:, col_idx].astype(str).str.contains("值班", na=False)]
        results['v_person'] = duty_row.iloc[0, 1] if not duty_row.empty else "未偵測"
        
        # 搜尋幹部動態 (鄭榮捷、邱品淳、曹培翔)
        cadre_notes = []
        for name in ["鄭榮捷", "邱品淳", "曹培翔"]:
            row = df_d[df_d.iloc[:, 1].astype(str).str.contains(name, na=False)]
            if not row.empty:
                # 檢查當前時段勤務
                curr_duty = str(row.iloc[0, col_idx])
                if "休" in curr_duty:
                    cadre_notes.append(f"{name}休假")
                else:
                    # 搜尋全天巡邏時段
                    patrol_slots = []
                    for c in range(4, 16):
                        if "巡邏" in str(row.iloc[0, c]):
                            h = (c-4)*2
                            patrol_slots.append(f"{h:02d}-{h+2:02d}")
                    if patrol_slots:
                        # 合併連續時段 (簡化處理)
                        time_range = f"{patrol_slots[0][:2]}至{patrol_slots[-1][-2:]}"
                        cadre_notes.append(f"{name}在所督勤，編排{time_range}時段巡邏勤務")
                    else:
                        cadre_notes.append(f"{name}在所督勤({curr_duty})")
        results['cadre_status'] = "；".join(cadre_notes) + "。"
    except Exception as e:
        results['v_person'], results['cadre_status'] = "解析失敗", "幹部資料解析失敗。"

    # 2. 解析【裝備交接簿】
    try:
        df_e = pd.read_csv(e_file, header=None) if e_file.name.endswith('csv') else pd.read_excel(e_file, header=None)
        df_e_str = df_e.astype(str)
        
        # 定位最後一筆 在所/出勤/送修
        row_in = df_e[df_e_str.iloc[:, 1].str.contains("在", na=False)].iloc[-1]
        row_out = df_e[df_e_str.iloc[:, 1].str.contains("出", na=False)].iloc[-1]
        row_fix = df_e[df_e_str.iloc[:, 1].str.contains("送", na=False)].iloc[-1]
        
        results['eq'] = {
            "gun_in": safe_parse_int(row_in.iloc[2]), "gun_out": safe_parse_int(row_out.iloc[2]), "gun_fix": safe_parse_int(row_fix.iloc[2]),
            "bul_in": safe_parse_int(row_in.iloc[3]), "bul_out": safe_parse_int(row_out.iloc[3]), "bul_fix": safe_parse_int(row_fix.iloc[3]),
            "rad_in": safe_parse_int(row_in.iloc[6]), "rad_out": safe_parse_int(row_out.iloc[6]), "rad_fix": safe_parse_int(row_fix.iloc[6]),
            "vest_in": safe_parse_int(row_in.iloc[11]), "vest_out": safe_parse_int(row_out.iloc[11]), "vest_fix": safe_parse_int(row_fix.iloc[11])
        }
    except:
        results['eq'] = None

    return results

# --- 主畫面執行 ---
if duty_file and equip_file:
    with st.spinner("正在讀取 Excel 數據並自動對位..."):
        data = extract_all_data(duty_file, equip_file, target_hour)
    
    st.success("✅ 數據已自動擷取完成！")
    
    # 建立勾選清單 (決定報告內容)
    col1, col2 = st.columns(2)
    with col1:
        check_monitor = st.checkbox("駐地監錄/天羅地網正常", value=True)
        check_edu = st.checkbox("勤前教育宣導落實", value=True)
    with col2:
        check_env = st.checkbox("環境內務擺設整齊", value=True)
        check_alcohol = st.checkbox("酒測聯單無跳號", value=True)

    # 組合報告語句
    lines = []
    
    # 1. 值班
    lines.append(f"{time_str}，該所值班警員{data['v_person']}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。")
    
    # 2. 監錄
    if check_monitor:
        lines.append(f"該所駐地監錄設備及天羅地網系統均運作正常，無故障，{d_m5}至{d_m1}有逐日檢測2次以上紀錄。")
    
    # 3. 勤教
    if check_edu:
        lines.append(f"該所{d_m3}至{d_m1}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。")
    
    # 4. 內務
    if check_env:
        lines.append(f"該所環境內務擺設整齊清潔，符合規定。")
        
    # 5. 裝備 (自動帶入數據)
    eq = data['eq'] if data['eq'] else {"gun_in": 0, "gun_out": 0, "gun_fix": 0, "bul_in": 0, "bul_out": 0, "bul_fix": 0, "rad_in": 0, "rad_out": 0, "rad_fix": 0, "vest_in": 0, "vest_out": 0, "vest_fix": 0}
    
    fix_msg = f"（另有槍枝 {eq['gun_fix']} 把、無線電 {eq['rad_fix']} 臺送修中）" if (eq['gun_fix'] + eq['rad_fix']) > 0 else ""
    
    lines.append(f"該所手槍出勤 {eq['gun_out']} 把、在所 {eq['gun_in']} 把，子彈出勤 {eq['bul_out']} 顆、在所 {eq['bul_in']} 顆，無線電出勤 {eq['rad_out']} 臺、在所 {eq['rad_in']} 臺；防彈背心出勤 {eq['vest_out']} 件、在所 {eq['vest_in']} 件，幹部對械彈每日檢查管制良好，符合規定{fix_msg}。")
    
    # 6. 幹部動態
    lines.append(f"本日{data['cadre_status']}")
    
    # 7. 酒測
    if check_alcohol:
        lines.append(f"該所酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。")

    # 最終文字生成
    final_report = "\n".join([f"{i+1}、{line}" for i, line in enumerate(lines)])
    
    st.markdown("---")
    st.subheader("📋 產出之督導報告 (標楷體)")
    st.text_area("直接複製到公務系統：", value=final_report, height=450)
    
    # 預覽區 (Debug 資訊，若抓不到人名可看這)
    with st.expander("🛠️ 查看系統自動辨識結果 (Debug)"):
        st.write(f"偵測到時段欄位：Index {4 + (target_hour // 2)}")
        st.write(f"值班員警：{data['v_person']}")
        st.write(f"幹部動態：{data['cadre_status']}")
        st.write(f"械彈數據：{eq}")

else:
    st.info("👋 請上傳『勤務分配表』與『值班裝備交接簿』。系統會根據您的督導時間自動對位資料。")
