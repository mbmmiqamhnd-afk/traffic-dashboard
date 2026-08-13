import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

st.set_page_config(page_title="無照駕駛移置保管敘獎統計", page_icon="🚗", layout="wide")

st.title("🚗 無照駕駛車輛移置保管敘獎統計系統")
st.markdown("""
依據最新公文 (桃警交字第1140031633號) 邏輯設計：
* **未成年**：每 **5** 件嘉獎一次，每季每人上限二次。
* **成年人 (小型車/汽車)**：每 **3** 件嘉獎一次，每季每人上限二次。
* **跨季累計**：當季未達標件數可保留至下一季 (不得跨年)。
""")

st.info("💡 請上傳系統產出的「自選匯出.xlsx」(需包含『案件明細』工作表)")

# 解析年齡字串提取數字
def parse_age(age_str):
    try:
        if pd.isna(age_str):
            return np.nan
        return int(re.sub(r'\D', '', str(age_str)))
    except:
        return np.nan

# 判斷適用案件類別
def categorize_case(row):
    fact = str(row['違規事實1'])
    law = str(row['違規法條1'])
    age = row['age_num']
    
    if pd.isna(age):
        return '不採計'
        
    # 未成年 (未滿18歲) 適用任何無照
    if age < 18:
        return '未成年'
        
    # 成年人 (18歲含以上) 僅適用小型車 (或法條內之汽車)
    if age >= 18:
        # 檢查是否為小型車或法條文字中包含汽車駕駛人
        is_small_car = '小型車' in fact or '汽車駕駛人' in fact
        # 排除明確標示機車的案件
        is_motorcycle = '機車' in fact
        
        if is_small_car and not is_motorcycle:
            # 確認是否為 21-1-1, 21-1-4, 21-1-5 或 21-2
            # 實務上法條代碼可能為 21101xx, 21104xx, 21105xx, 212xxxx
            if law.startswith('21101') or law.startswith('21104') or law.startswith('21105') or law.startswith('212'):
                return '成年'
            
    return '不採計'

uploaded_file = st.file_uploader("上傳自選匯出 Excel", type=["xlsx"])

if uploaded_file:
    with st.spinner("處理中..."):
        try:
            # 讀取 Excel 中的「案件明細」工作表
            xls = pd.ExcelFile(uploaded_file)
            if '案件明細' not in xls.sheet_names:
                st.error("❌ 找不到『案件明細』工作表，請確認您上傳的是正確的自選匯出報表。")
                st.stop()
                
            df_raw = pd.read_excel(xls, sheet_name='案件明細', header=None)
            
            # 定位欄位名稱列 (通常在第3或第4列，前幾列是報表標頭)
            header_idx = df_raw[df_raw.iloc[:, 1] == '違規法條1'].index[0]
            
            # 擷取真實資料並設定欄位名稱
            df = df_raw.iloc[header_idx+1:].copy()
            df.columns = ['單號', '違規法條1', '違規事實1', '入案日', '舉發員警1', '違規人年齡']
            df = df.dropna(subset=['單號', '舉發員警1'])
            
            # 處理年齡與分類
            df['age_num'] = df['違規人年齡'].apply(parse_age)
            df['案件類別'] = df.apply(categorize_case, axis=1)
            
            # 過濾掉不採計的資料
            df_valid = df[df['案件類別'] != '不採計'].copy()
            
            # 處理日期與季別 (入案日格式通常為民國年 1150106)
            def get_quarter(date_str):
                try:
                    month = int(str(date_str)[3:5])
                    if month <= 3: return 1
                    elif month <= 6: return 2
                    elif month <= 9: return 3
                    else: return 4
                except:
                    return 1
            df_valid['季別'] = df_valid['入案日'].apply(get_quarter)
            
            # 建立統計樞紐表
            summary = df_valid.pivot_table(
                index='舉發員警1', 
                columns=['案件類別', '季別'], 
                aggfunc='size', 
                fill_value=0
            )
            
            # 結算敘獎結果
            results = []
            for officer, row in summary.iterrows():
                # Q1
                juv_q1 = row.get(('未成年', 1), 0)
                adult_q1 = row.get(('成年', 1), 0)
                
                juv_merit_q1 = min(juv_q1 // 5, 2)
                juv_carry_to_q2 = juv_q1 - (juv_merit_q1 * 5)
                
                adult_merit_q1 = min(adult_q1 // 3, 2)
                adult_carry_to_q2 = adult_q1 - (adult_merit_q1 * 3)
                
                # Q2
                juv_q2_raw = row.get(('未成年', 2), 0)
                adult_q2_raw = row.get(('成年', 2), 0)
                
                juv_total_q2 = juv_q2_raw + juv_carry_to_q2
                juv_merit_q2 = min(juv_total_q2 // 5, 2)
                
                adult_total_q2 = adult_q2_raw + adult_carry_to_q2
                adult_merit_q2 = min(adult_total_q2 // 3, 2)
                
                # Q3 & Q4 可以以此類推擴充，目前針對您資料中的 Q1, Q2 展示
                
                total_merit = juv_merit_q1 + adult_merit_q1 + juv_merit_q2 + adult_merit_q2
                
                results.append({
                    '舉發員警': officer,
                    'Q1 未成年(件)': juv_q1,
                    'Q1 成年(件)': adult_q1,
                    'Q1 核算嘉獎數': juv_merit_q1 + adult_merit_q1,
                    
                    'Q2 未成年(含Q1保留)': juv_total_q2,
                    'Q2 成年(含Q1保留)': adult_total_q2,
                    'Q2 核算嘉獎數': juv_merit_q2 + adult_merit_q2,
                    
                    '總嘉獎數': total_merit
                })
                
            res_df = pd.DataFrame(results)
            
            if res_df.empty:
                st.warning("⚠️ 沒有計算出任何符合敘獎資格的資料，請確認入案日與案件是否吻合條件。")
            else:
                st.success("✅ 計算完成！")
                
                # 顯示過濾後的有效案件明細
                with st.expander("🔍 檢視有效案件明細 (供除錯與核對用)"):
                    st.dataframe(df_valid[['單號', '入案日', '舉發員警1', '違規人年齡', '案件類別', '違規事實1', '違規法條1']])
                
                # 顯示結算表
                st.subheader("🏆 敘獎結算結果")
                st.dataframe(res_df.style.highlight_max(subset=['總嘉獎數'], color='lightgreen'))
                
                # 下載功能
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    res_df.to_excel(writer, sheet_name='敘獎結算表', index=False)
                    df_valid.to_excel(writer, sheet_name='採計案件明細', index=False)
                output.seek(0)
                
                st.download_button(
                    label="📥 下載完整統計結果 (Excel)",
                    data=output,
                    file_name="無照駕駛敘獎統計表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"解析發生錯誤：{str(e)}")
            st.write("請確認您的自選匯出檔案格式沒有大幅變更。")
