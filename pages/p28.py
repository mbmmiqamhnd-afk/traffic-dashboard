import streamlit as st
import pandas as pd
from io import BytesIO

# ==========================================
# 頁面設定
# ==========================================
st.set_page_config(page_title="危險駕車取締獎勵統計", layout="wide", page_icon="🚨")
st.title("🚨 危險駕車與改裝車輛取締 - 獎勵統計儀表板")
st.markdown("""
**💡 獎勵計算標準：**
* **【Group A：每 2 輛嘉獎一次】**：違反道交條例第13條第1款、第18條第1項及第43條第3項。
* **【Group B：每 4 輛嘉獎一次】**：違反道交條例第16條第1項第1、2款（限定排氣管及消音器設備）、第43條第1項第1、3、4、5款。
* *註：累積達一大功 (9次嘉獎) 後，以倍數計算並以此類推。*
""")
st.divider()

# ==========================================
# 檔案上傳區塊
# ==========================================
uploaded_file = st.file_uploader("📂 請上傳舉發違規資料檔 (支援格式：xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        with st.spinner('資料處理中，請稍候...'):
            # 讀取資料
            df = pd.read_excel(uploaded_file, sheet_name="list2")
            
            # 確保欄位為字串型態，避免比對出錯及 NaN 問題
            df['條款1'] = df['條款1'].astype(str).str.strip()
            df['違規事實1'] = df['違規事實1'].astype(str).fillna('')
            df['舉發員警'] = df['舉發員警'].astype(str).str.strip()
            
            # ==========================================
            # 條件篩選邏輯區
            # ==========================================
            # Group A (2件 = 1嘉獎): 第13-1, 18-1, 43-3 條
            mask_A = df['條款1'].str.startswith(('131', '181', '433'))
            
            # Group B (4件 = 1嘉獎): 
            # 1. 第16條 且 違規事實包含「排氣管」或「消音器」
            mask_16 = df['條款1'].str.startswith('16') & df['違規事實1'].str.contains('排氣管|消音器', na=False)
            
            # 2. 第43-1-1, 43-1-3, 43-1-4, 43-1-5 (直接比對前五碼)
            mask_43_1 = df['條款1'].str.startswith(('43101', '43103', '43104', '43105'))
            
            mask_B = mask_16 | mask_43_1
            
            # 切分出符合條件的 DataFrame
            df_A = df[mask_A]
            df_B = df[mask_B]
            
            # ==========================================
            # 獎勵核算區
            # ==========================================
            # 計算各員警件數
            counts_A = df_A['舉發員警'].value_counts().rename("GroupA_件數")
            counts_B = df_B['舉發員警'].value_counts().rename("GroupB_件數")
            
            # 以外部合併 (Outer Join) 確保所有有件數的員警都在名單上
            reward_df = pd.concat([counts_A, counts_B], axis=1).fillna(0).astype(int)
            
            # 排除可能因為資料空白導致的空白員警名稱 (例如：'nan')
            if 'nan' in reward_df.index:
                reward_df = reward_df.drop('nan')
                
            # 計算嘉獎數 (無條件捨去取整數)
            reward_df['嘉獎次數'] = (reward_df['GroupA_件數'] // 2) + (reward_df['GroupB_件數'] // 4)
            
            # 計算大功數 (9次嘉獎 = 1大功) 與 剩餘嘉獎
            reward_df['大功數'] = reward_df['嘉獎次數'] // 9
            reward_df['剩餘嘉獎'] = reward_df['嘉獎次數'] % 9
            
            # 整理報表格式並排序
            reward_df = reward_df.sort_values(by=["嘉獎次數", "GroupA_件數"], ascending=[False, False]).reset_index()
            
            # 動態判斷並重新命名警員欄位 (解決 pandas reset_index 的命名衝突)
            if '舉發員警' in reward_df.columns:
                reward_df = reward_df.rename(columns={"舉發員警": "舉發員警名稱"})
            elif 'index' in reward_df.columns:
                reward_df = reward_df.rename(columns={"index": "舉發員警名稱"})
            
            # 過濾出至少有一件符合的員警
            reward_df = reward_df[(reward_df['GroupA_件數'] > 0) | (reward_df['GroupB_件數'] > 0)]
            
            # 重新排列欄位順序
            reward_df = reward_df[['舉發員警名稱', 'GroupA_件數', 'GroupB_件數', '嘉獎次數', '大功數', '剩餘嘉獎']]

        # ==========================================
        # 畫面呈現區
        # ==========================================
        st.success("✅ 資料處理完成！")
        
        # 數據概況指標
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("適用【2件1嘉獎】", f"{len(df_A)} 件")
        col2.metric("適用【4件1嘉獎】", f"{len(df_B)} 件")
        col3.metric("核發嘉獎總數", f"{reward_df['嘉獎次數'].sum()} 次")
        col4.metric("達成大功總數", f"{reward_df['大功數'].sum()} 次")
        
        st.divider()
        
        # 顯示核算結果表
        st.subheader("🏆 員警獎勵核算總表")
        st.dataframe(reward_df, use_container_width=True)
        
        # 匯出 Excel 功能
        if not reward_df.empty:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                reward_df.to_excel(writer, index=False, sheet_name='獎勵結算表')
                df_A.to_excel(writer, index=False, sheet_name='GroupA_明細')
                df_B.to_excel(writer, index=False, sheet_name='GroupB_明細')
            output.seek(0)
            
            st.download_button(
                label="📥 下載完整統計報表 (Excel)",
                data=output,
                file_name="危險駕車取締獎勵結算表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        st.divider()
        
        # 原始資料驗證區塊 (使用 Tabs 進行分類)
        st.subheader("🔍 案件明細檢核區")
        tab1, tab2 = st.tabs(["📌 Group A：2件1嘉獎 (第13-1, 18-1, 43-3條)", "📌 Group B：4件1嘉獎 (第16條改管, 43-1條)"])
        
        display_columns = ['單號', '車牌', '違規日期', '條款1', '違規事實1', '舉發員警', '舉發單位']
        
        with tab1:
            if not df_A.empty:
                st.dataframe(df_A[display_columns], use_container_width=True)
            else:
                st.info("查無符合 Group A 條件之案件。")
                
        with tab2:
            if not df_B.empty:
                st.dataframe(df_B[display_columns], use_container_width=True)
            else:
                st.info("查無符合 Group B 條件之案件。")

    except Exception as e:
        st.error(f"❌ 檔案處理發生錯誤，請確認欄位格式是否與標準資料庫一致。\n\n詳細錯誤訊息：{e}")
