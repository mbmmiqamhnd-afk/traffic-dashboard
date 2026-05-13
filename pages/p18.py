import streamlit as st
import pandas as pd
import io
import sys
import os
import re

# 自動將上層目錄加入路徑
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
try:
    from app import show_sidebar
except ImportError:
    def show_sidebar():
        pass 

def p18_page():
    show_sidebar()

    st.title("💰 龍潭分局 - 獎勵金點數統計表產生器")
    st.info("系統將自動辨識報表年月，並【強制產生】總表工作表於檔案首頁。")

    # 1. 參數設定
    with st.expander("⚙️ 點數權重設定", expanded=False):
        col1, col2, col3 = st.columns(3)
        p_a2 = col1.number_input("A2 點數/件", value=3.0, step=0.5)
        p_a3 = col2.number_input("A3 點數/件", value=1.0, step=0.5)
        p_traf = col3.number_input("交整點數/小時", value=1.0, step=0.5)

    # 2. 檔案上傳
    st.subheader("📂 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】(內含各單位數據)", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)", type=['xlsx'], accept_multiple_files=True)

    # 3. 執行運算
    if st.button("🚀 開始彙整並產生報表", type="primary"):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種資料皆已上傳！")
            return

        with st.spinner("正在辨識年月並強制產生總表..."):
            try:
                # --- 處理事故資料 ---
                df_acc = pd.read_excel(file_acc, header=4)
                df_acc['姓名'] = df_acc['姓名'].astype(str).str.strip()
                df_acc['A2類'] = pd.to_numeric(df_acc['A2類'], errors='coerce').fillna(0)
                df_acc['A3類'] = pd.to_numeric(df_acc['A3類'], errors='coerce').fillna(0)
                df_acc = df_acc[~df_acc['姓名'].isin(['nan', 'None', '', 'NaN'])]
                dict_acc = df_acc.groupby('姓名')[['A2類', 'A3類']].sum().to_dict(orient='index')

                # --- 處理交整資料 ---
                df_traf_all = pd.concat([pd.read_excel(f, sheet_name='月彙整總表') for f in file_traf_list])
                df_traf_all['姓名'] = df_traf_all['姓名'].astype(str).str.strip()
                df_traf_all['總計尖峰時數'] = pd.to_numeric(df_traf_all['總計尖峰時數'], errors='coerce').fillna(0)
                df_traf_all = df_traf_all[~df_traf_all['姓名'].isin(['nan', 'None', '', 'NaN'])]
                dict_traf = df_traf_all.groupby('姓名')['總計尖峰時數'].sum().to_dict()

                # --- 讀取基底檔案 ---
                dfs = pd.read_excel(file_template, sheet_name=None, header=None)
                
                # 🌟 尋找並移除舊的總表 (避免干擾)
                summary_sheet_key = None
                for k in list(dfs.keys()):
                    if '總表' in str(k) or '總計' in str(k):
                        summary_sheet_key = k
                        break
                if summary_sheet_key:
                    del dfs[summary_sheet_key]

                # 初始化統計
                unit_summary_data = [] 
                g_cite, g_acc, g_traf, g_all = 0, 0, 0, 0

                # --- 遍歷所有派出所分頁 ---
                for sheet_name in list(dfs.keys()):
                    df = dfs[sheet_name].astype(object) # 解除型別鎖定
                    header_row = None
                    cols = {}
                    
                    # 尋找欄位座標
                    for idx, row in df.iterrows():
                        vals = [str(x).strip() for x in row.values]
                        if '員警姓名' in vals:
                            header_row = idx
                            f_list = ['員警姓名', 'A2件數', 'A3件數', '事故點數', '交整時數', '交整點數', '個人總點數', '取締點數']
                            for f in f_list:
                                if f in vals: cols[f] = vals.index(f)
                            break
                    
                    if header_row is not None:
                        s_cite, s_acc, s_traf, s_all = 0, 0, 0, 0
                        
                        for r in range(header_row + 1, len(df)):
                            name = str(df.iloc[r, cols['員警姓名']]).strip()
                            
                            # 遇到小計或總計列
                            if name in ['小計', '總計']:
                                c_sub = pd.to_numeric(df.iloc[r, cols['取締點數']], errors='coerce')
                                c_sub = c_sub if not pd.isna(c_sub) else 0
                                
                                df.iloc[r, cols['事故點數']] = s_acc if s_acc > 0 else ""
                                df.iloc[r, cols['交整點數']] = s_traf if s_traf > 0 else ""
                                df.iloc[r, cols['個人總點數']] = c_sub + s_acc + s_traf
                                
                                if name == '小計':
                                    unit_summary_data.append({
                                        '單位名稱': sheet_name, '取締點數': c_sub, '事故點數': s_acc, 
                                        '交整點數': s_traf, '個人總點數': c_sub + s_acc + s_traf
                                    })
                                    g_cite += c_sub; g_acc += s_acc; g_traf += s_traf; g_all += (c_sub + s_acc + s_traf)
                                break
                            
                            # 一般員警名單填值
                            elif name not in ['nan', 'None', '', 'NaN']:
                                a2 = dict_acc.get(name, {}).get('A2類', 0)
                                a3 = dict_acc.get(name, {}).get('A3類', 0)
                                t_h = dict_traf.get(name, 0)
                                
                                ap = a2 * p_a2 + a3 * p_a3
                                tp = t_h * p_traf
                                cp = pd.to_numeric(df.iloc[r, cols['取締點數']], errors='coerce')
                                cp = cp if not pd.isna(cp) else 0
                                
                                df.iloc[r, cols['A2件數']] = a2 if a2 > 0 else ""
                                df.iloc[r, cols['A3件數']] = a3 if a3 > 0 else ""
                                df.iloc[r, cols['事故點數']] = ap if ap > 0 else ""
                                df.iloc[r, cols['交整時數']] = t_h if t_h > 0 else ""
                                df.iloc[r, cols['交整點數']] = tp if tp > 0 else ""
                                df.iloc[r, cols['個人總點數']] = cp + ap + tp
                                
                                s_acc += ap; s_traf += tp
                    dfs[sheet_name] = df # 存回記憶體

                # --- 🌟 強制生成置頂的全新總表 ---
                summary_rows = [['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']]
                for item in unit_summary_data:
                    summary_rows.append([item['單位名稱'], item['取締點數'], item['事故點數'], item['交整點數'], item['個人總點數']])
                summary_rows.append(['合計', g_cite, g_acc, g_traf, g_all])
                
                df_summary = pd.DataFrame(summary_rows)
                
                # 重新組合所有分頁 (將總表放在第一頁)
                final_dfs = {'總表': df_summary}
                for k, v in dfs.items():
                    final_dfs[k] = v

                # --- 🌟 智慧判讀年月邏輯 ---
                ext_year, ext_month = "115", "4" # 預設值
                found_date = False
                for s_name, df_temp in final_dfs.items():
                    for r in range(min(15, len(df_temp))):
                        for c in range(min(10, len(df_temp.columns))):
                            val = str(df_temp.iloc[r, c])
                            m = re.search(r'開單日期[：:\s]*(\d{3})(\d{2})', val)
                            if m:
                                ext_year = m.group(1)
                                ext_month = str(int(m.group(2)))
                                found_date = True; break
                        if found_date: break
                    if found_date: break
                    
                # 若內文沒找到，從檔名找
                if not found_date:
                    m_filename = re.search(r'(\d{2,4})[年\.\-/](\d{1,2})月?', file_template.name)
                    if m_filename:
                        ext_year, ext_month = m_filename.group(1), m_filename.group(2)

                # --- 網頁即時預覽總表 ---
                st.success(f"✅ 彙整完成！偵測到資料年月為：【{ext_year}年 {ext_month}月】")
                st.subheader("📊 【總表】預覽結果 (請確認是否出現數據)")
                
                # 建立一個供網頁預覽的漂亮表格
                preview_df = pd.DataFrame(unit_summary_data)
                if not preview_df.empty:
                    preview_df.loc[len(preview_df)] = ['合計', g_cite, g_acc, g_traf, g_all]
                    st.dataframe(preview_df)

                # --- 寫入 Excel 產出 ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for s_name, df_final in final_dfs.items():
                        df_final.to_excel(writer, sheet_name=s_name, header=False, index=False)
                
                # 🌟 完美產出指定檔名
                final_filename = f"桃園市政府警察局龍潭分局{ext_year}年{ext_month}月份處理道路交通安全人員獎勵金點數統計表.xlsx"
                
                st.download_button(
                    label=f"📥 點此下載：{final_filename}", 
                    data=output.getvalue(), 
                    file_name=final_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")

if __name__ == "__main__":
    p18_page()
