import streamlit as st
import pandas as pd
import io
import sys
import os
import re

# 自動將上層目錄加入路徑，以便能順利 import app.py 中的 show_sidebar
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
try:
    from app import show_sidebar
except ImportError:
    def show_sidebar():
        pass 

def p18_page():
    show_sidebar()

    st.title("💰 龍潭分局 - 獎勵金點數統計表產生器")
    st.info("系統會自動從檔案中判讀年月，產生標準檔名，並完美保留您的「總表」排版格式！")

    # 1. 參數設定
    with st.expander("⚙️ 點數權重設定", expanded=False):
        col1, col2, col3 = st.columns(3)
        p_a2 = col1.number_input("A2 點數/件", value=3.0, step=0.5)
        p_a3 = col2.number_input("A3 點數/件", value=1.0, step=0.5)
        p_traf = col3.number_input("交整點數/小時", value=1.0, step=0.5)

    # 2. 檔案上傳
    st.subheader("📂 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】(含各單位名單與總表)", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)", type=['xlsx'], accept_multiple_files=True)

    # 3. 執行運算
    if st.button("🚀 開始彙整並產生總表", type="primary"):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種當月資料皆已上傳！")
            return

        with st.spinner("正在進行深度彙整、計算總表並偵測年月..."):
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
                
                # 🌟 智慧判讀年月邏輯
                ext_year, ext_month = "115", "4" # 預設值
                
                # 策略 1：嘗試從檔名尋找 (例如 115年4月)
                m_filename = re.search(r'(\d{2,4})[年\.\-/](\d{1,2})月?', file_template.name)
                if m_filename:
                    ext_year, ext_month = m_filename.group(1), m_filename.group(2)
                else:
                    # 策略 2：從表格內容尋找「開單日期：1150401」
                    found_date = False
                    for sheet_name, df in dfs.items():
                        for r in range(min(15, len(df))): # 掃描前 15 行
                            for c in range(min(10, len(df.columns))):
                                val = str(df.iloc[r, c]).strip()
                                m_date = re.search(r'開單日期[：:\s]*(\d{3})(\d{2})', val)
                                if m_date:
                                    ext_year = m_date.group(1)
                                    ext_month = str(int(m_date.group(2))) # 轉成整數去掉 0，例如 '04' -> '4'
                                    found_date = True
                                    break
                            if found_date: break
                        if found_date: break

                # 初始化總表容器
                unit_summary = [] 
                g_cite, g_acc, g_traf, g_all = 0, 0, 0, 0

                # --- 遍歷所有分頁進行填值 ---
                for sheet_name, df in dfs.items():
                    df = df.astype(object) # 解除型別鎖定
                    dfs[sheet_name] = df
                    
                    if '總表' in sheet_name:
                        continue # 總表留到最後處理
                    
                    header_row = None
                    cols = {}
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
                            
                            # 處理小計列
                            if name in ['小計', '總計']:
                                c_sub = pd.to_numeric(df.iloc[r, cols['取締點數']], errors='coerce')
                                c_sub = c_sub if not pd.isna(c_sub) else 0
                                df.iloc[r, cols['事故點數']] = s_acc if s_acc > 0 else ""
                                df.iloc[r, cols['交整點數']] = s_traf if s_traf > 0 else ""
                                df.iloc[r, cols['個人總點數']] = c_sub + s_acc + s_traf
                                
                                if name == '小計':
                                    unit_summary.append({
                                        '單位': sheet_name, '取締': c_sub, '事故': s_acc, '交整': s_traf, '總計': c_sub + s_acc + s_traf
                                    })
                                    g_cite += c_sub; g_acc += s_acc; g_traf += s_traf; g_all += (c_sub + s_acc + s_traf)
                                break
                            
                            # 處理一般員警
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
                                
                                s_acc += ap; s_traf += tp; s_cite += cp

                # --- 處理「總表」分頁 (原地精準覆寫) ---
                summary_sheet_key = next((k for k in dfs.keys() if '總表' in str(k)), None)
                
                if summary_sheet_key:
                    df_sum = dfs[summary_sheet_key]
                    h_row = None
                    s_cols = {}
                    
                    # 尋找總表的標題列
                    for idx, row in df_sum.iterrows():
                        vals = [str(x).strip() for x in row.values]
                        if '單位名稱' in vals:
                            h_row = idx
                            for f in ['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']:
                                if f in vals: s_cols[f] = vals.index(f)
                            break
                    
                    if h_row is not None:
                        curr_r = h_row + 1
                        # 填入各派出所數據
                        for item in unit_summary:
                            if curr_r < len(df_sum):
                                df_sum.iloc[curr_r, s_cols['單位名稱']] = item['單位']
                                df_sum.iloc[curr_r, s_cols['取締點數']] = item['取締'] if item['取締'] > 0 else ""
                                df_sum.iloc[curr_r, s_cols['事故點數']] = item['事故'] if item['事故'] > 0 else ""
                                df_sum.iloc[curr_r, s_cols['交整點數']] = item['交整'] if item['交整'] > 0 else ""
                                df_sum.iloc[curr_r, s_cols['個人總點數']] = item['總計'] if item['總計'] > 0 else ""
                                curr_r += 1
                        
                        # 尋找並填寫「合計」列
                        for r_idx in range(curr_r, len(df_sum)):
                            if '合計' in str(df_sum.iloc[r_idx, s_cols.get('單位名稱', 0)]).strip():
                                df_sum.iloc[r_idx, s_cols['取締點數']] = g_cite
                                df_sum.iloc[r_idx, s_cols['事故點數']] = g_acc
                                df_sum.iloc[r_idx, s_cols['交整點數']] = g_traf
                                df_sum.iloc[r_idx, s_cols['個人總點數']] = g_all
                                break

                # --- 下載產出 ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for s_name, df_final in dfs.items():
                        df_final.to_excel(writer, sheet_name=s_name, header=False, index=False)

                st.success(f"✅ 彙整完成！已成功抓取資料年月：【{ext_year} 年 {ext_month} 月】")
                st.balloons()
                
                # 🌟 完美的自動化檔名
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
