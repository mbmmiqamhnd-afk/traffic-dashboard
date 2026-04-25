import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# 分頁配置
st.set_page_config(page_title="督導報告 v7.0 - 自動擷取版", layout="wide")

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

st.title("📋 警用督導報告極速生成器 v7.0")
st.info("💡 系統已連結 Excel 解析引擎：自動擷取值班人員、幹部動態及械彈數量。")

# --- 側邊欄：檔案與時間設定 ---
with st.sidebar:
    st.header("📂 數據源匯入")
    duty_file = st.file_uploader("1. 上傳『勤務分配表』", type=['xlsx'])
    equip_file = st.file_uploader("2. 上傳『值班裝備交接簿』", type=['xlsx', 'csv'])
    
    st.divider()
    target_time = st.time_input("督導時間 (自動對位時段)", datetime.now().time())
    time_str = target_time.strftime('%H%M')
    target_hour = target_time.hour
    
    # 日期推算
    today = datetime.now()
    d_minus_5 = (today - timedelta(days=5)).strftime('%m月%d日')
    d_minus_3 = (today - timedelta(days=3)).strftime('%m月%d日')
    d_minus_1 = (today - timedelta(days=1)).strftime('%m月%d日')

# --- 核心解析引擎 ---
def extract_duty_info(d_file, hour):
    try:
        df = pd.read_excel(d_file, header=None).ffill()
        # 尋找時段欄位 (通常在第 4 或 5 列)
        # 定位現在小時所屬的欄位 (例如 11 點屬於 10-12 時段)
        col_idx = 4 + (hour // 2)  # 聖亭所班表通常每 2 小時一格，從第 4 欄開始
        
        # 1. 抓取值班人員
        v_person = df[df.iloc[:, col_idx].astype(str).str.contains("值班", na=False)].iloc[0, 1]
        
        # 2. 抓取幹部動態
        cadre_notes = []
        cadres = ["鄭榮捷", "邱品淳", "曹培翔"]
        for name in cadres:
            row = df[df.iloc[:, 1].astype(str).str.contains(name, na=False)]
            if not row.empty:
                duty_type = row.iloc[0, col_idx]
                if "巡邏" in str(duty_type):
                    # 搜尋該員當天所有的巡邏時段
                    patrol_hours = []
                    for c in range(4, 16):
                        if "巡邏" in str(row.iloc[0, c]):
                            h_start = (c - 4) * 2
                            patrol_hours.append(f"{h_start:02d}-{h_start+2:02d}")
                    p_str = "、".join(patrol_hours)
                    cadre_notes.append(f"{name}在所督勤，編排{p_str}時段巡邏勤務")
                elif "休" in str(duty_type):
                    cadre_notes.append(f"{name}休假")
                else:
                    cadre_notes.append(f"{name}{duty_type}")
        
        cadre_final = "；".join(cadre_notes) + "。"
        return v_person, cadre_final
    except:
        return "人員解析失敗", "幹部動態解析失敗。"

def extract_equip_info(e_file):
    try:
        df = pd.read_csv(e_file, header=None) if e_file.name.endswith('csv') else pd.read_excel(e_file, header=None)
        df_str = df.astype(str)
        
        row_in = df[df_str.iloc[:, 1].str.contains("在\n所|在所", na=False)].iloc[-1]
        row_out = df[df_str.iloc[:, 1].str.contains("出\n勤|出勤", na=False)].iloc[-1]
        row_fix = df[df_str.iloc[:, 1].str.contains("送\n修|送修", na=False)].iloc[-1]
        
        def get_v(row, idx): return int(float(str(row.iloc[idx]).split('.')[0]))
        
        return {
            "gun_in": get_v(row_in, 2), "gun_out": get_v(row_out, 2), "gun_fix": get_v(row_fix, 2),
            "bul_in": get_v(row_in, 3), "bul_out": get_v(row_out, 3), "bul_fix": get_v(row_fix, 3),
            "rad_in": get_v(row_in, 6), "rad_out": get_v(row_out, 6), "rad_fix": get_v(row_fix, 6),
            "vest_in": get_v(row_in, 11), "vest_out": get_v(row_out, 11), "vest_fix": get_v(row_fix, 11)
        }
    except:
        return None

# --- 畫面生成 ---
if duty_file and equip_file:
    v_name, c_status = extract_duty_info(duty_file, target_hour)
    eq = extract_equip_info(equip_file)
    
    if eq:
        st.write("💡 **自動勾選督導事項 (順序將自動編排)：**")
        c1, c2 = st.columns(2)
        with c1:
            check_monitor = st.checkbox("✅ 監錄設備/天羅地網正常", value=True)
            check_edu = st.checkbox("✅ 勤教宣導(酒駕/優先權)落實", value=True)
        with c2:
            check_env = st.checkbox("✅ 環境內務整潔", value=True)
            check_alcohol = st.checkbox("✅ 酒測聯單無跳號", value=True)

        # 組合標準格式
        lines = []
        
        # 1. 值班
        lines.append(f"{time_str}，該所值班警員{v_name}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。")
        
        # 2. 監錄
        if check_monitor:
            lines.append(f"該所駐地監錄設備及天羅地網系統均運作正常，無故障，{d_minus_5}至{d_minus_1}有逐日檢測2次以上紀錄。")
        
        # 3. 勤教
        if check_edu:
            lines.append(f"該所{d_minus_3}至{d_minus_1}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。")
        
        # 4. 內務
        if check_env:
            lines.append(f"該所環境內務擺設整齊清潔，符合規定。")
            
        # 5. 裝備 (全自動擷取)
        fix_str = ""
        if eq['gun_fix'] > 0 or eq['rad_fix'] > 0:
            fix_str = f"（另有槍枝 {eq['gun_fix']} 把、無線電 {eq['rad_fix']} 臺送修中）"
            
        lines.append(f"該所手槍出勤 {eq['gun_out']} 把、在所 {eq['gun_in']} 把，子彈出勤 {eq['bul_out']} 顆、在所 {eq['bul_in']} 顆，無線電出勤 {eq['rad_out']} 臺、在所 {eq['rad_in']} 臺；防彈背心出勤 {eq['vest_out']} 件、在所 {eq['vest_in']} 件，幹部對械彈每日檢查管制良好，符合規定{fix_str}。")
        
        # 6. 幹部動態 (全自動解析)
        lines.append(f"本日{c_status}")
        
        # 7. 酒測
        if check_alcohol:
            lines.append(f"該所酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。")

        # 最終文本
        final_text = "\n".join([f"{i+1}、{line}" for i, line in enumerate(lines)])
        
        st.markdown("---")
        st.subheader("📋 最終督導報告 (標楷體)")
        st.text_area("直接複製到公務系統：", value=final_text, height=450)
        st.success("✨ 報告已依照聖亭所最新資料自動產出！")
    else:
        st.error("❌ 裝備交接簿解析失敗，請確認檔案內容。")
else:
    st.warning("👋 請先在左側上傳今日 Excel 檔案，系統將自動提取所有督導數據。")
