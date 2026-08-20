import streamlit as st
import pandas as pd
import numpy as np
import io

# 載入您自訂的側邊欄模組
import menu

def process_traffic_data(file):
    """讀取並清洗自選匯出 Excel 資料"""
    try:
        # 讀取「案件明細」工作表，跳過前兩列表頭
        df = pd.read_excel(file, sheet_name="案件明細", header=2)
        
        # 將第一列設為正確的欄位名稱
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
        
        # 確保所有必要欄位都存在
        cols_to_keep = ['單號', '簡式車種名稱', '違規法條1', '違規事實1', '入案日', '舉發員警1']
        
        missing_cols = [col for col in cols_to_keep if col not in df.columns]
        if missing_cols:
            st.error(f"上傳的檔案缺少以下必要欄位：{', '.join(missing_cols)}，請確認是否為正確的『自選匯出.xlsx』格式。")
            return None
            
        df = df[cols_to_keep]
        return df
    
    except Exception as e:
        st.error(f"檔案解析失敗，錯誤訊息：{str(e)}")
        return None

def calculate_merits_for_officer(group):
    """計算單一員警的預估嘉獎次數 (依入案日排序以處理加倍機制)"""
    group = group.sort_values(by='入案日')
    
    total_merits = 0
    other_plate_cases = 0  
    fake_plate_cases = 0   
    
    for idx, row in group.iterrows():
        violation = str(row['違規事實1'])
        vehicle = str(row['簡式車種名稱'])
        
        multiplier = 2 if total_merits >= 9 else 1
        
        if '偽造' in violation or '變造' in violation:
            fake_plate_cases += 1
            if '汽車' in vehicle:
                total_merits += 2 * multiplier
            else:
                total_merits += 1 * multiplier
                
        elif '他車' in violation:
            other_plate_cases += 1
            if other_plate_cases % 2 == 0:
                total_merits += 1 * multiplier
                
    return pd.Series({
        '偽變造件數 (件)': fake_plate_cases,
        '他車牌照件數 (件)': other_plate_cases,
        '總合件數 (件)': fake_plate_cases + other_plate_cases,
        '預估嘉獎次數 (次)': total_merits
    })

def main():
    st.set_page_config(page_title="偽變造車牌專案 敘獎統計", layout="wide")
    
    # === 呼叫您的自訂側邊欄 ===
    try:
        menu.show_sidebar()
    except Exception as e:
        st.sidebar.error("無法載入側邊欄，請確認根目錄下有 menu.py")
    # ==========================
        
    st.title("🔎 偽變造車牌專案 - 自動敘獎統計系統")
    st.markdown("本模組專為計算「加強取締查緝偽(變)造車牌及違法(規)權利車」專案期間出力人員敘獎所設計。")
    st.divider()
    
    uploaded_file = st.file_uploader("請上傳『自選匯出.xlsx』(資料來源需包含簡式車種名稱與違規事實)", type=["xlsx"])
    
    if uploaded_file is not None:
        with st.spinner("資料處理中，請稍候..."):
            df = process_traffic_data(uploaded_file)
            
        if df is not None:
            with st.expander("📄 檢視原始案件明細", expanded=False):
                st.dataframe(df, use_container_width=True)
            
            st.subheader("📊 員警專案敘獎統計表")
            
            merit_stats = df.groupby('舉發員警1').apply(calculate_merits_for_officer).reset_index()
            merit_stats = merit_stats.sort_values(by=['預估嘉獎次數 (次)', '總合件數 (件)'], ascending=[False, False]).reset_index(drop=True)
            
            styled_df = merit_stats.style.background_gradient(subset=['預估嘉獎次數 (次)'], cmap='Blues')
            st.dataframe(styled_df, use_container_width=True)
            
            csv = merit_stats.to_csv(index=False).encode('utf-8-sig')
            
            col1, col2 = st.columns([1, 4])
            with col1:
                st.download_button(
                    label="📥 匯出敘獎統計名冊 (CSV)",
                    data=csv,
                    file_name='偽變造車牌專案敘獎統計表.csv',
                    mime='text/csv',
                    type="primary"
                )
            
            with col2:
                st.info("💡 **系統計算標準：**\n"
                        "1. **偽變造車牌**：汽車記嘉獎2次，機車記嘉獎1次。\n"
                        "2. **懸掛他車號牌**：每 2 件記嘉獎1次。\n"
                        "3. **獎勵加倍**：累計達 9 次嘉獎門檻後，後續入案之件數獎勵自動加倍計算。")

if __name__ == "__main__":
    main()
