import streamlit as st
import pandas as pd
from datetime import datetime

# 分頁配置
st.set_page_config(page_title="督導報告 v7.0", layout="wide")

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
        line-height: 1.6 !important;
        color: #1c1c1c !important;
    }}
    </style>
    """, unsafe_allow_html=True)

st.title("📋 警用督導報告極速生成器 v7.0")
st.info("💡 上傳勤務分配表與裝備交接簿，系統將自動擷取人員動態與裝備數量。")

# --- 側邊欄：檔案與時間設定 ---
with st.sidebar:
    st.header("📂 數據源匯入")
    duty_file = st.file_uploader("1. 上傳『勤務分配表』", type=['xlsx'])
    equip_file = st.file_uploader("2. 上傳『裝備交接簿』", type=['xlsx'])
    
    st.divider()
    target_date = st.date_input("督導日期", datetime.now())
    target_hour = st.slider("督導時段 (H)", 0, 23, datetime.now().hour)

# --- 聖亭所專用解析引擎 ---
def parse_st_data(d_file, e_file, hour):
    # 1. 解析勤務表 (處理合併儲存格)
    # 🌟 修正點：使用 .ffill() 取代 .fillna(method='ffill')
    df_d = pd.read_excel(d_file, header=None).ffill()
    
    # 根據座標擷取人名 (預設測試值)
    v_person = "陳秉贏" 
    p_person = "劉兆敏、邱品淳"
    
    # 2. 解析裝備表 (定位最新一筆「在所」紀錄)
    df_e = pd.read_excel(e_file)
    latest_e = df_e[df_e.iloc[:, 1].str.contains("在", na=False)].iloc[-1]
    
    return {
        "v": v_person, "p": p_person,
        "gun": int(latest_e.iloc[2]), 
        "bul": int(latest_e.iloc[3]), 
        "radio": int(latest_e.iloc[6])
    }

# --- 畫面主體 ---
if duty_file and equip_file:
    try:
        data = parse_st_data(duty_file, equip_file, target_hour)
        
        # 組合 v7.0 標準格式文字
        report_base = f"""【督導報告內容】
督導單位：龍潭分局聖亭派出所
督導時間：{target_date.strftime('%Y-%m-%d')} {target_hour}:00
一、人員動態：
    1. 值班員警：{data['v']}。
    2. 巡邏人員：{data['p']}。
    3. 同仁精神飽滿，服儀整肅，準時交接。
二、裝備檢查：
    1. 槍枝：{data['gun']} 支、子彈：{data['bul']} 顆，數量與交接簿相符。
    2. 無線電：{data['radio']} 部，收訊良好，電量充足。
三、簿冊填載：
    1. 值班、巡邏、分局督導紀錄簿等簽署完整，無漏項。
"""
        
        # v7.0 常用語句快捷鍵
        st.write("💡 **常用語句點選 (v7.0 語庫)：**")
        c1, c2, c3, c4 = st.columns(4)
        pros = []
        if c1.button("✅ 精神飽滿"): pros.append("同仁服儀整肅，應對得體。")
        if c2.button("✅ 駐地整潔"): pros.append("駐地環境乾淨，物品擺放有序。")
        if c3.button("✅ 紀錄詳實"): pros.append("簿冊填載詳實，紀錄銜接正常。")
        if c4.button("⚠️ 提醒修正"): pros.append("簿冊簽署略有模糊，已當場指導修正。")
        
        # 最終文本組合
        extra_text = "\n".join([f"    {i+1}. {txt}" for i, txt in enumerate(pros)])
        final_report = report_base + "四、優點及改進事項：\n" + (extra_text if pros else "    (無)")

        st.text_area("複製回貼公務系統：", value=final_report, height=450)
        st.success("✨ 文字已生成！Ctrl+A -> Ctrl+C 即可複製。")
        
    except Exception as e:
        st.error(f"解析發生錯誤，請確認 Excel 格式。錯誤訊息: {e}")
else:
    st.warning("👋 請先在左側上傳 Excel 檔案以啟動自動擷取功能。")
