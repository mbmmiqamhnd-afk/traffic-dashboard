import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font, Alignment, Border, Side

# 1. 頁面設定
st.set_page_config(page_title="交通事故統計 (A1/A2)", layout="wide", page_icon="🚑")
st.title("🚑 交通事故統計自動化系統")

st.markdown("""
### 📝 使用說明
1. 請上傳 **3 個原始報表檔案** (本週、今年累計、去年累計)。
2. 系統會**自動掃描檔案內容**分辨日期。
3. 自動產出 **A1/A2 統計報表** (Excel 格式，含標楷體與紅字標示)。
""")

# 2. 檔案上傳
uploaded_files = st.file_uploader("請上傳 3 個事故報表檔案", accept_multiple_files=True, key="acc_uploader")

# 3. 主邏輯
if uploaded_files and st.button("🚀 開始分析", key="btn_acc"):
    with st.spinner("正在智慧辨識檔案與計算中..."):
        try:
            # --- 函數定義區 ---
            def parse_raw(file_obj):
                """讀取 CSV 或 Excel"""
                try: 
                    return pd.read_csv(file_obj, header=None)
                except: 
                    file_obj.seek(0)
                    return pd.read_excel(file_obj, header=None)

            def clean_data(df_raw):
                """清理數據，標準化欄位"""
                # 強制轉字串避免讀取錯誤
                df_raw[0] = df_raw[0].astype(str)
                # 篩選含有 '總計' 或 '派出所' 或 '合計' 的列
                df_data = df_raw[df_raw[0].str.contains("總計|派出所|合計", na=False)].copy()
                df_data = df_data.reset_index(drop=True)
                
                # 欄位對應
                columns_map = {
                    0: "Station", 1: "Total_Cases", 2: "Total_Deaths", 3: "Total_Injuries",
                    4: "A1_Cases", 5: "A1_Deaths", 6: "A1_Injuries",
                    7: "A2_Cases", 8: "A2_Deaths", 9: "A2_Injuries", 10: "A3_Cases"
                }
                
                # 補足缺失欄位
                for i in range(11):
                    if i not in df_data.columns: df_data[i] = 0
                
                df_data = df_data.rename(columns=columns_map)
                df_data = df_data[list(columns_map.values())]
                
                # 轉數字
                for col in list(columns_map.values())[1:]:
                    df_data[col] = pd.to_numeric(df_data[col].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
                
                # 單位名稱簡化
                df_data['Station_Short'] = df_data['Station'].astype(str).str.replace('派出所', '所').str.replace('總計', '合計')
                return df_data

            # --- 智慧辨識檔案日期 ---
            file_data_map = {}
            debug_info = []

            for uploaded_file in uploaded_files:
                uploaded_file.seek(0)
                df = parse_raw(uploaded_file)
                
                found_dates = []
                date_str_found = "未找到日期"
                
                # 掃描前 5 列、前 3 欄
                for r in range(min(5, len(df))):
                    for c in range(min(3, len(df.columns))):
                        val = str(df.iloc[r, c])
                        # 尋找 113/01/01 或 113.1.1
                        dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', val)
                        if len(dates) >= 2:
                            found_dates = dates
                            date_str_found = val
                            break
                    if found_dates: break

                if found_dates:
                    start_y, start_m, start_d = map(int, found_dates[0])
                    end_y, end_m, end_d = map(int, found_dates[1])
                    
                    month_diff = (end_y - start_y) * 12 + (end_m - start_m)
                    days_diff = end_d - start_d
                    
                    # 判斷邏輯: 同一個月且天數差距小 -> 本期
                    if month_diff == 0 and days_diff < 20:
                        category = 'weekly'
                    else:
                        category = 'cumulative'
                        
                    file_data_map[uploaded_file.name] = {
                        'df': df, 
                        'category': category, 
                        'year': start_y, 
                        'raw_date': f"{start_y}/{start_m:02d}/{start_d:02d}-{end_y}/{end_m:02d}/{end_d:02d}"
                    }
                    debug_info.append(f"✅ {uploaded_file.name}: [{category}] ({found_dates[0]}~{found_dates[1]})")
                else:
                    debug_info.append(f"❌ {uploaded_file.name}: 無法識別日期")

            # --- 分配檔案 ---
            df_wk = None; df_cur = None; df_lst = None
            h_wk = ""; h_cur = ""; h_lst = ""

            for fname, data in file_data_map.items():
                if data['category'] == 'weekly':
                    df_wk = clean_data(data['df']); h_wk = data['raw_date']

            cumu_files = [d for d in file_data_map.values() if d['category'] == 'cumulative']
            if len(cumu_files) >= 2:
                cumu_files.sort(key=lambda x: x['year'], reverse=True)
                df_cur = clean_data(cumu_files[0]['df']); h_cur = cumu_files[0]['raw_date']
                df_lst = clean_data(cumu_files[1]['df']); h_lst = cumu_files[1]['raw_date']

            if df_wk is None or df_cur is None or df_lst is None:
                st.error("❌ 無法識別完整的 3 份檔案 (需包含：本期週報、今年累計、去年累計)。")
                with st.expander("🕵️‍♂️ 查看偵測細節"):
                    for info in debug_info: st.write(info)
                st.stop()

            # --- 計算 A1 ---
            a1_wk = df_wk[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'wk'})
            a1_cur = df_cur[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'cur'})
            a1_lst = df_lst[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'last'})
            m_a1 = pd.merge(a1_wk, a1_cur, on='Station_Short', how='outer')
            m_a1 = pd.merge(m_a1, a1_lst, on='Station_Short', how='outer').fillna(0)
            m_a1['Diff'] = m_a1['cur'] - m_a1['last']

            # --- 計算 A2 ---
            a2_wk = df_wk[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'wk'})
            a2_cur = df_cur[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'cur'})
            a2_lst = df_lst[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'last'})
            m_a2 = pd.merge(a2_wk, a2_cur, on='Station_Short', how='outer')
            m_a2 = pd.merge(m_a2, a2_lst, on='Station_Short', how='outer').fillna(0)
            m_a2['Diff'] = m_a2['cur'] - m_a2['last']
            # 計算增減率
            m_a2['Pct_Str'] = m_a2.apply(lambda x: f"{(x['Diff']/x['last']):.2%}" if x['last']!=0 else "-", axis=1)
            m_a2['Prev'] = "-" # 佔位符

            # --- 排序 ---
            target_order = ['合計', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
            for m in [m_a1, m_a2]:
                m['Station_Short'] = pd.Categorical(m['Station_Short'], categories=target_order, ordered=True)
                m.sort_values('Station_Short', inplace=True)

            # --- 整理最終表格 ---
            a1_final = m_a1[['Station_Short', 'wk', 'cur', 'last', 'Diff']].copy()
            a1_final.columns = ['單位', f'本期({h_wk})', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較']
            
            a2_final = m_a2[['Station_Short', 'wk', 'Prev', 'cur', 'last', 'Diff', 'Pct_Str']].copy()
            a2_final.columns = ['單位', f'本期({h_wk})', '前期', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較', '本年較去年增減比例']

            # --- 顯示結果 ---
            st.subheader("📊 A1 死亡人數統計")
            st.dataframe(a1_final, use_container_width=True, hide_index=True)
            
            st.subheader("📊 A2 受傷人數統計")
            st.dataframe(a2_final, use_container_width=True, hide_index=True)

            # --- 產出 Excel (含格式) ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # 1. 寫入資料
                a1_final.to_excel(writer, index=False, sheet_name='A1死亡人數')
                a2_final.to_excel(writer, index=False, sheet_name='A2受傷人數')
                
                # 2. 定義樣式
                font_normal = Font(name='標楷體', size=12)
                font_red_bold = Font(name='標楷體', size=12, bold=True, color="FF0000") # 紅色粗體
                font_bold = Font(name='標楷體', size=12, bold=True)
                align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
                border_style = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                      top=Side(style='thin'), bottom=Side(style='thin'))
                
                # 3. 針對每個分頁進行格式化
                for sheet_name in ['A1死亡人數', 'A2受傷人數']:
                    ws = writer.book[sheet_name]
                    
                    # 調整欄寬
                    for col in ws.columns:
                        col_letter = col[0].column_letter
                        ws.column_dimensions[col_letter].width = 20
                    
                    # 處理標題列 (第一列)
                    for cell in ws[1]:
                        cell.alignment = align_center
                        cell.border = border_style
                        # 判斷標題是否含日期關鍵字 -> 轉紅字
                        if any(x in str(cell.value) for x in ["本期", "累計", "/"]):
                            cell.font = font_red_bold
                        else:
                            cell.font = font_bold
                            
                    # 處理數據內容 (從第二列開始)
                    for row in ws.iter_rows(min_row=2):
                        for cell in row:
                            cell.alignment = align_center
                            cell.border = border_style
                            cell.font = font_normal

            st.download_button(
                label="📥 下載 Excel 報表 (含格式)", 
                data=output.getvalue(), 
                file_name=f'交通事故統計表_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx', 
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        except Exception as e:
            st.error(f"發生系統錯誤：{e}")
            st.exception(e) # 顯示詳細錯誤以便除錯
