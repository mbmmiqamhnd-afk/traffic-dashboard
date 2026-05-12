import streamlit as st
import pandas as pd
import io
import re

# 1. 頁面配置
st.set_page_config(page_title="交通疏導時數彙整", page_icon="⏱️", layout="wide")

def run_app():
    st.title("⏱️ 交通疏導勤務時數彙整 (時段精確比對版)")
    st.markdown("---")

    # --- 側邊欄：手動設定尖峰時段 (避免配置表格式不一抓不到) ---
    st.sidebar.header("⚙️ 尖峰時段設定")
    st.sidebar.info("請設定明細表中對應的尖峰時段欄位名稱")
    
    am_peak = st.sidebar.text_input("上午尖峰時段 (用逗號隔開)", value="07-08,08-09")
    pm_peak = st.sidebar.text_input("下午尖峰時段 (用逗號隔開)", value="17-18,18-19")
    
    # 轉換成清單並去除空格
    peak_intervals = [i.strip() for i in (am_peak + "," + pm_peak).split(',') if i.strip()]

    # --- 主介面 ---
    st.subheader("1️⃣ 上傳每日勤務明細")
    st.info(f"當前計算時段：{', '.join(peak_intervals)}")
    
    uploaded_files = st.file_uploader("請選取當月勤務明細檔 (CSV/Excel)", accept_multiple_files=True, type=['csv', 'xlsx'])

    if uploaded_files:
        all_records = []
        
        for file in uploaded_files:
            try:
                # 讀取檔案 (跳過前兩行標題列)
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, skiprows=2, encoding='utf-8-sig')
                    except:
                        df = pd.read_csv(file, skiprows=2, encoding='cp950')
                else:
                    df = pd.read_excel(file, skiprows=2)

                # 清洗欄位名稱 (移除換行與空格)
                df.columns = [re.sub(r'[\s\n\r]', '', str(c)) for c in df.columns]
                
                # 找出存在的尖峰時段欄位
                available_peak_cols = [c for c in df.columns if c in peak_intervals]
                
                # 單位名稱
                unit_name = re.split(r'\d+', file.name)[0]

                for _, row in df.iterrows():
                    name = str(row.get('姓名', '')).replace('\n', '').strip()
                    if not name or name in ['nan', 'None', '', '合計', '總計']: continue
                    
                    watch_hours = 0
                    # 僅掃描「尖峰時段」的欄位
                    for col in available_peak_cols:
                        content = str(row.get(col, '')).replace('\n', '')
                        if "守望" in content:
                            watch_hours += 1
                    
                    if watch_hours > 0:
                        all_records.append({
                            "單位": unit_name,
                            "姓名": name,
                            "守望時數": watch_hours,
                            "來源日期": file.name
                        })
            except Exception as e:
                st.error(f"解析 {file.name} 失敗：{e}")

        # --- 顯示結果 ---
        if all_records:
            full_df = pd.DataFrame(all_records)
            summary = full_df.groupby(['單位', '姓名'])['守望時數'].sum().reset_index()
            summary = summary.sort_values(['單位', '守望時數'], ascending=[True, False])
            
            tab1, tab2 = st.tabs(["🏆 月總計結果", "🔍 明細對帳"])
            with tab1:
                st.subheader("📊 彙整統計結果")
                st.dataframe(summary, use_container_width=True, hide_index=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    summary.to_excel(writer, index=False, sheet_name='月彙整')
                    full_df.to_excel(writer, index=False, sheet_name='對帳明細')
                st.download_button("📥 下載統計報表", output.getvalue(), "尖峰時數統計.xlsx")
                
            with tab2:
                st.subheader("📋 每日詳細紀錄")
                st.dataframe(full_df, use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 讀取成功，但在設定的尖峰時段內未發現『守望』紀錄。")
            st.write("目前偵測到的欄位標題：", df.columns.tolist())

if __name__ == "__main__":
    run_app()
