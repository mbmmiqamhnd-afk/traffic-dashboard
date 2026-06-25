import streamlit as st

def show_sidebar():
    with st.sidebar:
        # 1. 系統大標題
        st.title("🚓 交通執法自動化分析引擎")
        st.info("本系統專為處理繁雜交通執法報表設計，支援自動辨識與雲端同步。")

        st.divider()

        # 2. 核心數據處理
        st.subheader("📊 數據與分析")
        st.page_link("app.py", label="回首頁 / 批次處理中心", icon="⚙️")
        st.page_link("pages/p17.py", label="交通疏導時數彙整", icon="⏱️")
        st.page_link("pages/p18.py", label="獎勵金點數統計表", icon="💰")

        st.divider()

        # 3. 勤務與專案規劃
        st.subheader("📅 勤務與專案規劃")
        st.page_link("pages/p09.py", label="聯合稽查勤務規劃", icon="🚓")          
        st.page_link("pages/p20.py", label="聯合稽查(二階段)勤務規劃", icon="🚓") 
        st.page_link("pages/p10.py", label="防制危險駕車", icon="🚓")
        st.page_link("pages/p11.py", label="防制危險駕車 (月份版)", icon="📅")
        st.page_link("pages/p12.py", label="行人及護老交通安全", icon="🚶")
        st.page_link("pages/p13.py", label="取締砂石車", icon="🚛")
        st.page_link("pages/p14.py", label="二階段勤務規劃", icon="🚓")
        st.page_link("pages/p15.py", label="三合一勤務規劃系統", icon="📋") # 原有功能
        st.page_link("pages/p19.py", label="二合一勤務規劃系統", icon="📋")
        st.page_link("pages/p21.py", label="三階段專案勤務規劃系統", icon="🚀") # 🆕 全新功能

        st.divider()

        # 4. 輔助工具
        st.subheader("🛠️ 輔助工具")
        st.page_link("pages/p05.py", label="商標頁碼工具", icon="🔖")
        st.page_link("pages/p06.py", label="PDF 轉 PPTX 工具", icon="📂")
        st.page_link("pages/p16.py", label="督導報告極速生成器 v7.0", icon="📋")

def main():
    show_sidebar()

    st.title("🚓 歡迎使用交通執法自動化分析引擎")
    st.markdown("""
    請從左側選單選擇您要使用的功能。

    ✅ **最新系統更新：**
    * 已於「勤務與專案規劃」專區新增 **[三階段專案勤務規劃系統]**，支援第一階段「機動攔檢」、第二階段「場所臨檢」與第三階段「定點路檢」的獨立表單規劃。
    * 已於「勤務與專案規劃」專區新增 **[聯合稽查(二階段)勤務規劃]**，完美融合環保局臨時檢驗站開設功能與二階段定點路檢排班機制。
    * 已於「數據與分析」專區新增 **[獎勵金點數統計表]** 功能，可自動比對匯入交通事故與疏導時數。
    """)

if __name__ == "__main__":
    main()
