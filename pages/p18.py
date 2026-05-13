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
    st.info("系統將自動偵測報表年月，並【強制產生】包含完整數據的總表。")

    # 1. 點數設定
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

    # 3. 執行彙整
    if st.button("🚀 開始彙整並產生報表", type="primary"):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種資料皆已上傳！")
            return

        with st.spinner("資料比對與總表彙整中..."):
            try:
                # --- A. 事故資料處理 ---
                df_acc = pd.read_excel(file_acc, header=4)
                df_acc['姓名'] = df_acc['姓名'].astype(str).str.strip()
                dict_acc = df_acc.groupby('姓名')[['A2類', 'A3類']].sum().to_dict(orient='index')

                # --- B. 交整資料處理 ---
                df_traf_all = pd.concat([pd.read_excel(f, sheet_name='月彙整總表') for f in file_traf_list])
                df_traf_all['姓名'] = df_traf_all['姓名'].astype(str).str.strip()
                dict_traf = df_traf_all.groupby('姓名')['總計尖峰時數'].sum().to_dict()

                # --- C. 讀取基底檔案 ---
                dfs = pd.read_excel(file_template, sheet_name=None, header=None)
                
                # 移除舊總表
                for k in list(dfs.keys()):
                    if '總表' in str(k): del dfs[k]

                # 初始化統計容器
                unit_summary_list = []
                g_cite, g_acc, g_traf, g_all = 0, 0, 0, 0

                # --- D. 遍歷分頁填值 ---
                for sheet_name in list(dfs.keys()):
                    df = dfs[sheet_name].astype(object)
                    header_row = None
                    cols = {}
                    
                    # 搜尋欄位位置
                    for idx, row in df.iterrows():
                        vals = [str(x).strip() for x in row.values]
                        if '員警姓名' in vals:
                            header_row = idx
                            targets = ['員警姓名', 'A2件數', 'A3件數', '事故點數', '交整時數', '交整點數', '個人總點數', '取締點數']
                            for t in targets:
                                if t in vals: cols[t] = vals.index(t)
                            break
                    
                    if header_row is not None:
                        s_cite, s_acc, s_traf = 0, 0, 0
                        
                        for r in range(header_row + 1, len(df)):
                            name_val = str(df.iloc[r, cols['員警姓名']]).strip()
                            
                            # 偵測小計行 (包含「小計」字樣)
                            if '小計' in name_val or '總計' in name_val:
                                c_sub = pd.to_numeric(df.iloc[r, cols['取締點數']], errors='coerce') or 0
                                df.iloc[r, cols['事故點數']] = s_acc if s_acc > 0 else ""
                                df.iloc[r, cols['交整點數']] = s_traf if s_traf > 0 else ""
                                df.iloc[r, cols['個人總點數']] = c_sub + s_acc + s_traf
                                
                                if '小計' in name_val:
                                    unit_summary_list.append({
                                        '單位名稱': sheet_name, '取締點數': c_sub, '事故點數': s_acc, 
                                        '交整點數': s_traf, '個人總點數': c_sub + s_acc + s_traf
                                    })
                                    g_cite += c_sub; g_acc += s_acc; g_traf += s_traf; g_all += (c_sub + s_acc + s_traf)
                                break
                            
                            elif name_val not in ['nan', 'None', '', 'NaN']:
                                a2 = dict_acc.get(name_val, {}).get('A2類', 0)
                                a3 = dict_acc.get(name_val, {}).get('A3類', 0)
                                t_h = dict_traf.get(name_val, 0)
                                
                                ap = a2 * p_a2 + a3 * p_a3
                                tp = t_h * p_traf
                                cp = pd.to_numeric(df.iloc[r, cols['取締點數']], errors='coerce') or 0
                                
                                df.iloc[r, cols['A2件數']] = a2 if a2 > 0 else ""
                                df.iloc[r, cols['A3件數']] = a3 if a3 > 0 else ""
                                df.iloc[r, cols['事故點數']] = ap if ap > 0 else ""
                                df.iloc[r, cols['交整時數']] = t_h if t_h > 0 else ""
                                df.iloc[r, cols['交整點數']] = tp if tp > 0 else ""
                                df.iloc[r, cols['個人總點數']] = cp + ap + tp
                                
                                s_acc += ap; s_traf += tp
                        dfs[sheet_name] = df

                # --- E. 產生總表數據 ---
                sum_rows = [['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']]
                for item in unit_summary_list:
                    sum_rows.append([item['單位名稱'], item['取締點數'], item['事故點數'], item['交整點數'], item['個人總點數']])
                sum_rows.append(['合計', g_cite, g_acc, g_traf, g_all])
                df_summary_final = pd.DataFrame(sum_rows)

                # --- F. 判斷年月 ---
                ext_year, ext_month = "115", "4"
                for s_name, df_scan in dfs.items():
                    content_str = df_scan.astype(str).values.flatten()
                    match = re.search(r'開單日期[：:\s]*(\d{3})(\d{2})', "".join(content_str))
                    if match:
                        ext_year, ext_month = match.group(1), str(int(match.group(2)))
                        break

                # --- G. 畫面顯示與預覽 ---
                st.success(f"✅ 運算完成！已偵測報表日期：{ext_year}年{ext_month}月")
                
                st.subheader("📊 總表數據預覽")
                if unit_summary_list:
                    preview_df = pd.DataFrame(unit_summary_list)
                    preview_df.loc[len(preview_df)] = ['合計', g_cite, g_acc, g_traf, g_all]
                    st.table(preview_df) # 使用 table 強制完整顯示
                else:
                    st.warning("⚠️ 偵測不到派出所分頁中的『小計』列，請確認範本中員警姓名下方是否有小計字樣。")

                # --- H. 封裝 Excel ---
                final_filename = f"桃園市政府警察局龍潭分局{ext_year}年{ext_month}月份處理道路交通安全人員獎勵金點數統計表.xlsx"
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # 強制將總表放在第一個 Sheet
                    df_summary_final.to_excel(writer, sheet_name='總表', header=False, index=False)
                    for s_name, df_final in dfs.items():
                        df_final.to_excel(writer, sheet_name=s_name, header=False, index=False)
                
                st.download_button(
                    label=f"📥 點此下載：{final_filename}", 
                    data=output.getvalue(), 
                    file_name=final_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ 處理過程中發生錯誤：{str(e)}")

if __name__ == "__main__":
    p18_page()
