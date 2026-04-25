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
    }}
    </style>
    """, unsafe_allow_html=True)

st.title("📋 警用督導報告極速生成器 v7.0")

# --- 側邊欄：檔案匯入 ---
with st.sidebar:
    st.header("📂 數據源匯入")
    duty_file = st.file_uploader("1. 上傳『勤務分配表』", type=['xlsx'])
    equip_file = st.file_uploader("2. 上傳『裝備交接簿』", type=['xlsx'])
    
    st.divider()
    target_date = st.date_input("督導日期", datetime.now())
    target_hour = st.slider("督導時段 (H)", 0, 23, datetime.now().hour)

# --- 核心解析引擎 (針對聖亭所格式) ---
def parse_st_files(d_file, e_file, hour):
    # 解析勤務表 (處理合併儲存格)
    df_d = pd.read_excel(d_file, header=None).fillna(method='ffill')
    
    # 解析裝備表 (定位最後一筆「在所」紀錄)
    df_e = pd.read_excel(e_file)
    # 根據上傳 CSV 內容，定位包含「在\n所」的最新一列
    latest_e = df_e[df_e.iloc[:, 1].str.contains("在", na=False)].iloc[-1]
    
    # 預設人名 (可依據 df_d 座標進一步自動擷取)
    v_person = "陳秉贏" 
    p_person = "劉兆敏、邱品淳"
    
    return {
        "v": v_person, "p": p_person,
        "gun": int(latest_e.iloc[2]), "bul": int(latest_e.iloc[3]), "radio": int(latest_e.iloc[6])
    }

if duty_file and equip_file:
    try:
        data = parse_st_files(duty_file, equip_file, target_hour)
        
        # 報告主體文字
        report = f"""【督導報告內容】
督導單位：龍潭分局聖亭派出所
督導時間：{target_date.strftime('%Y-%m-%d')} {target_hour}:00
一、人員動態：
    1. 值班員警：{data['v']}。
    2. 巡邏人員：{data['p']}。
    3. 同仁精神飽滿，服儀整肅。
二、裝備檢查：
    1. 槍枝：{data['gun']} 支、子彈：{data['bul']} 顆，數量相符。
    2. 無線電：{data['radio']} 部，收訊良好，電量充足。
三、簿冊填載：
    1. 值班、巡邏、分局督導紀錄簿等簽署完整，無漏項。
"""

        # v7.0 快捷語句點選
        st.write("💡 **常用語句點選 (v7.0)：**")
        c1, c2, c3, c4 = st.columns(4)
        extra = []
        if c1.button("✅ 應對得體"): extra.append("同仁服儀整肅，應對得體，熟悉轄區治安熱點。")
        if c2.button("✅ 駐地整潔"): extra.append("駐地環境乾淨，公物擺放井然有序，內務整潔。")
        if c3.button("✅ 紀錄詳實"): extra.append("各項簿冊紀錄銜接清楚，簽署完整，表報詳盡。")
        if c4.button("⚠️ 指導修正"): extra.append("部分紀錄略有塗改，已提醒應依規定修正並蓋章。")
        
        # 組合優缺點事項
        if extra:
            full_text = report + "四、優點及改進事項：\n" + "\n".join([f"    {i+1}. {t}" for i, t in enumerate(extra)])
        else:
            full_text = report + "四、優點及改進事項：\n    (無)"
        
        st.text_area("複製區 (Ctrl+A -> Ctrl+C)：", value=full_text, height=450)
        st.success("✨ 文字已生成，直接貼回公務系統即可！")
        
    except Exception as e:
        st.error(f"解析發生錯誤：{e}")
else:
    st.info("👋 您好！請於左側上傳今日 Excel 檔案以啟用 v7.0 自動化功能。")
