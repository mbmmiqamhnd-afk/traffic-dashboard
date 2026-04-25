import streamlit as st

def show_sidebar():
    with st.sidebar:
        # 1. 系統大標題
        st.title("🚓 交通執法自動化分析引擎")
        st.info("本系統專為處理繁雜交通執法報表設計，支援自動辨識與雲端同步。")
        
        st.divider() # 分隔線
        
        # 2. 核心數據處理
        st.subheader("📊 數據與分析")
        st.page_link("app.py", label="全自動批次處理中心", icon="⚙️")
        
        # 3. 勤務與專案規劃
        st.subheader("📅 勤務與專案規劃")
        st.page_link("pages/p09.py", label="聯合稽查勤務規劃", icon="🚓") 
        st.page_link("pages/p10.py", label="防制危險駕車", icon="🚓")
        st.page_link("pages/p11.py", label="防制危險駕車 (月份版)", icon="📅")
        st.page_link("pages/p12.py", label="行人及護老交通安全", icon="🚶")
        st.page_link("pages/p13.py", label="取締砂石車", icon="🚛")
        st.page_link("pages/p14.py", label="二階段勤務規劃", icon="🚓")
        st.page_link("pages/p15.py", label="三合一勤務規劃系統", icon="📋")
        
        # 4. 輔助工具
        st.subheader("🛠️ 輔助工具")
        st.page_link("pages/p05.py", label="商標頁碼工具", icon="🔖")
        st.page_link("pages/p06.py", label="PDF 轉 PPTX 工具", icon="📂")
        # 👇 這是我們新增的 v7 工具連結 👇
        st.page_link("pages/p16.py", label="督導報告極速生成器 v7.0", icon="📋")
