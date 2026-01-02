import streamlit as st
import pandas as pd
import io
import re
from datetime import date
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(page_title="交通事故統計 (A1/A2)", layout="wide", page_icon="🚑")
st.title("🚑 交通事故統計自動化系統 (跨年修正版)")

st.markdown("""
### 📝 使用說明
1. 請上傳 **3 個原始報表檔案**。
2. 系統邏輯升級：
   - 自動以 **結束年份** 分組 (解決跨年週報歸屬問題)。
   - 優先以 **1月1日開始** 識別累計表 (解決年初累計天數少於週報的問題)。
""")

uploaded_files = st.file_uploader("請上傳 3 個事故報表檔案", accept_multiple_files=True, key="acc_uploader")

if uploaded_files and st.button("🚀 開始分析", key="btn_acc"):
    with st.spinner("正在進行邏輯辨識與計算..."):
        try:
            # --- 1. 基礎函數 ---
            def parse_raw(file_obj):
                try: return pd.read_csv(file_obj, header=None)
                except: 
                    file_obj.seek(0)
                    return pd.read_excel(file_obj, header=None)

            def clean_data(df_raw):
                df_raw[0] = df_raw[0].astype(str)
                df_data = df_raw[df_raw[0].str.contains("總計|派出所|合計", na=False)].copy()
                df_data = df_data.reset_index(drop=True)
                
                columns_map = {
                    0: "Station", 1: "Total_Cases", 2: "Total_Deaths", 3: "Total_Injuries",
                    4: "A1_Cases", 5: "A1_Deaths", 6: "A1_Injuries",
                    7: "A2_Cases", 8: "A2_Deaths", 9: "A2_Injuries", 10: "A3_Cases"
                }
                
                for i in range(11):
                    if i not in df_data.columns: df_data[i] = 0
                
                df_data = df_data.rename(columns=columns_map)
                df_data = df_data[list(columns_map.values())]
                
                for col in list(columns_map.values())[1:]:
                    df_data[col] = pd.to_numeric(df_data[col].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
                
                df_data['Station_Short'] = df_data['Station'].astype(str).str.replace('派出所', '所').str.replace('總計', '合計')
                return df_data

            # --- 2. 檔案掃描與資訊提取 ---
            files_meta = [] 
            debug_info = []

            for uploaded_file in uploaded_files:
                uploaded_file.seek(0)
                df = parse_raw(uploaded_file)
                
                found_dates = []
                for r in range(min(5, len(df))):
                    for c in range(min(3, len(df.columns))):
                        val = str(df.iloc[r, c])
                        dates = re.findall(r'(\d{3})[./](\d{1,2})[./](\d{1,2})', val)
                        if len(dates) >= 2:
                            found_dates = dates
                            break
                    if found_dates: break

                if found_dates:
                    try:
                        start_y, start_m, start_d = map(int, found_dates[0])
                        end_y, end_m, end_d = map(int, found_dates[1])
                        
                        d_start = date(start_y + 1911, start_m, start_d)
                        d_end = date(end_y + 1911, end_m, end_d)
                        duration_days = (d_end - d_start).days
                        
                        raw_date_str = f"{start_y}/{start_m:02d}/{start_d:02d}-{end_y}/{end_m:02d}/{end_d:02d}"
                        
                        files_meta.append({
                            'file': uploaded_file,
                            'df': df,
                            'start_tuple': (start_y, start_m, start_d), # 用於判斷 01/01
                            'end_year': end_y,    # 用結束年份來分組 (關鍵修正)
                            'duration': duration_days,
                            'raw_date': raw_date_str
                        })
                        debug_info.append(f"✅ {uploaded_file.name}: 結束年={end_y}, 開始={start_m}/{start_d}, 天數={duration_days}")
                    except:
                        debug_info.append(f"⚠️ {uploaded_file.name}: 日期解析失敗")
                else:
                    debug_info.append(f"❌ {uploaded_file.name}: 找不到日期")

            # --- 3. 智慧分配 (邏輯核心) ---
            # 依照「結束年份」排序，最大的為今年
            files_meta.sort(key=lambda x: x['end_year'], reverse=True)
            
            df_wk = None; df_cur = None; df_lst = None
            h_wk = ""; h_cur = ""; h_lst = ""
            
            if len(files_meta) >= 3:
                current_year_end = files_meta[0]['end_year']
                
                # 分組：今年結束的檔案 vs 以前年份結束的檔案
                current_files = [f for f in files_meta if f['end_year'] == current_year_end]
                past_files = [f for f in files_meta if f['end_year'] < current_year_end]
                
                # 1. 抓去年累計 (過去年份中，年份最大的)
                if past_files:
                    past_files.sort(key=lambda x: x['end_year'], reverse=True)
                    target = past_files[0]
                    df_lst = clean_data(target['df'])
                    h_lst = target['raw_date']
                
                # 2. 抓今年 (本期 vs 累計)
                if len(current_files) >= 2:
                    # 邏輯 A: 看誰是 01/01 開始 -> 那個就是累計
                    cumu_candidate = None
                    wk_candidate = None
                    
                    # 先找有沒有 01月01日 開始的檔案
                    starts_on_jan1 = [f for f in current_files if f['start_tuple'][1] == 1 and f['start_tuple'][2] == 1]
                    
                    if len(starts_on_jan1) == 1:
                        # 只有一個檔案是 01/01 開始 -> 它就是累計
                        cumu_candidate = starts_on_jan1[0]
                        # 另一個就是週報 (排除掉累計那個)
                        remaining = [f for f in current_files if f != cumu_candidate]
                        if remaining: wk_candidate = remaining[0]
                    else:
                        # 如果都沒有，或都有 (極端狀況)，退回到用「天數長短」判斷
                        # 天數長 = 累計, 天數短 = 週報
                        current_files.sort(key=lambda x: x['duration'])
                        wk_candidate = current_files[0]
                        cumu_candidate = current_files[-1]

                    if cumu_candidate:
                        df_cur = clean_data(cumu_candidate['df']); h_cur = cumu_candidate['raw_date']
                    if wk_candidate:
                        df_wk = clean_data(wk_candidate['df']); h_wk = wk_candidate['raw_date']

            # --- 4. 檢核 ---
            if df_wk is None or df_cur is None or df_lst is None:
                st.error("❌ 邏輯判斷失敗，請確認上傳檔案是否包含：一份去年、一份今年累計、一份今年週報。")
                with st.expander("🕵️‍♂️ 偵測與分組細節"):
                    for info in debug_info: st.write(info)
                st.stop()

            # --- 5. 計算 (A1/A2) ---
            # A1
            a1_wk = df_wk[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'wk'})
            a1_cur = df_cur[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'cur'})
            a1_lst = df_lst[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'last'})
            m_a1 = pd.merge(a1_wk, a1_cur, on='Station_Short', how='outer')
            m_a1 = pd.merge(m_a1, a1_lst, on='Station_Short', how='outer').fillna(0)
            m_a1['Diff'] = m_a1['cur'] - m_a1['last']

            # A2
            a2_wk = df_wk[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'wk'})
            a2_cur = df_cur[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'cur'})
            a2_lst = df_lst[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'last'})
            m_a2 = pd.merge(a2_wk, a2_cur, on='Station_Short', how='outer')
            m_a2 = pd.merge(m_a2, a2_lst, on='Station_Short', how='outer').fillna(0)
            m_a2['Diff'] = m_a2['cur'] - m_a2['last']
            m_a2['Pct_Str'] = m_a2.apply(lambda x: f"{(x['Diff']/x['last']):.2%}" if x['last']!=0 else "-", axis=1)
            m_a2['Prev'] = "-"

            # 排序
            target_order = ['合計', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
            for m in [m_a1, m_a2]:
                m['Station_Short'] = pd.Categorical(m['Station_Short'], categories=target_order, ordered=True)
                m.sort_values('Station_Short', inplace=True)

            # 顯示
            a1_final = m_a1[['Station_Short', 'wk', 'cur', 'last', 'Diff']].copy()
            a1_final.columns = ['單位', f'本期({h_wk})', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較']
            
            a2_final = m_a2[['Station_Short', 'wk', 'Prev', 'cur', 'last', 'Diff', 'Pct_Str']].copy()
            a2_final.columns = ['單位', f'本期({h_wk})', '前期', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較', '本年較去年增減比例']

            st.subheader("📊 A1 死亡人數統計"); st.dataframe(a1_final, use_container_width=True, hide_index=True)
            st.subheader("📊 A2 受傷人數統計"); st.dataframe(a2_final, use_container_width=True, hide_index=True)

            # --- 6. Excel 產出 (維持美化格式) ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                a1_final.to_excel(writer, index=False, sheet_name='A1死亡人數')
                a2_final.to_excel(writer, index=False, sheet_name='A2受傷人數')
                
                font_normal = Font(name='標楷體', size=12)
                font_red_bold = Font(name='標楷體', size=12, bold=True, color="FF0000")
                font_bold = Font(name='標楷體', size=12, bold=True)
                align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
                border_style = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                
                for sheet_name in ['A1死亡人數', 'A2受傷人數']:
                    ws = writer.book[sheet_name]
                    for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 20
                    for cell in ws[1]:
                        cell.alignment = align_center
                        cell.border = border_style
                        if any(x in str(cell.value) for x in ["本期", "累計", "/"]): cell.font = font_red_bold
                        else: cell.font = font_bold
                    for row in ws.iter_rows(min_row=2):
                        for cell in row:
                            cell.alignment = align_center
                            cell.border = border_style
                            cell.font = font_normal

            st.download_button(label="📥 下載 Excel 報表", data=output.getvalue(), file_name=f'交通事故統計表_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e:
            st.error(f"錯誤：{e}")
            st.exception(e)
