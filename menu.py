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
        # ✅ 新增績效結算系統在這裡 (假設檔名為 p27.py)
        st.page_link("pages/p27.py", label="舉發績效結算系統", icon="👮")
        st.page_link("pages/p17.py", label="交通疏導時數彙整", icon="⏱️")
        st.page_link("pages/p18.py", label="獎勵金點數統計表", icon="💰")
        st.page_link("pages/p24.py", label="噪音改裝車輛嘉獎統計", icon="🏍️")
        st.page_link("pages/p25.py", label="連假專案督勤表生成", icon="🗓️") 
        st.page_link("pages/p26.py", label="行人及護老專案督勤表生成", icon="👵")

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
        st.page_link("pages/p15.py", label="三合一勤務規劃系統", icon="📋") 
        st.page_link("pages/p19.py", label="二合一勤務規劃系統", icon="📋")
        st.page_link("pages/p21.py", label="三階段專案勤務規劃系統", icon="🚀") 
        st.page_link("pages/p22.py", label="綜合勤務規劃(新)", icon="📋")
        st.page_link("pages/p23.py", label="純巡邏勤務規劃", icon="🚓")

        st.divider()

        # 4. 輔助工具
        st.subheader("🛠️ 輔助工具")
        # ✅ 將合併後的工具指向 p06.py，並從選單中移除舊的 p05.py
        st.page_link("pages/p06.py", label="綜合檔案加工與轉檔中心", icon="🗂️")
        st.page_link("pages/p16.py", label="督導報告極速生成器 v7.0", icon="📋")

def main():
    show_sidebar()

    st.title("🚓 歡迎使用交通執法自動化分析引擎")
    st.markdown("""
    請從左側選單選擇您要使用的功能。

    ✅ **最新系統更新：**
    * 已於「數據與分析」專區新增 **[舉發績效結算系統]**，支援上傳配分表與員警個人績效表，自動查表結算並同步至 Google Sheets 或匯出 Excel。
    * 已於「輔助工具」專區將商標頁碼工具與 PDF 轉檔工具完美合併為 **[綜合檔案加工與轉檔中心]**，單一介面即可處理「商標頁碼遮蓋」、「PDF轉PPTX/圖片」及「多圖合併PDF」等需求。
    * 已於「數據與分析」專區新增 **[連假專案督勤表生成]**，可自動化匯出符合格式的連續假期與週末督勤日期表。
    * 已於「數據與分析」專區新增 **[噪音改裝車輛嘉獎統計]**，支援上、下半年自動偵測模式，並提供一鍵比對與商數計算匯出功能。
    * 已於「勤務與專案規劃」專區新增 **[純巡邏勤務規劃]**，提供專注於機動巡邏與交通稽查的獨立表單介面。
    * 已於「勤務與專案規劃」專區新增 **[三階段專案勤務規劃系統]**，支援第一階段「機動攔檢」、第二階段「場所臨檢」與第三階段「定點路檢」的獨立表單規劃。
    * 已於「勤務與專案規劃」專區新增 **[聯合稽查(二階段)勤務規劃]**，完美融合環保局臨時檢驗站開設功能與二階段定點路檢排班機制。
    * 已於「數據與分析」專區新增 **[獎勵金點數統計表]** 功能，可自動比對匯入交通事故與疏導時數。
    """)

if __name__ == "__main__":
    main()
