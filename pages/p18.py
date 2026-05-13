import streamlit as st
import pandas as pd
import io
import sys
import os

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
    st.info("本工具會以您上傳的「當月統計表」為基底，自動填入事故與交整數據，並【強制產生】符合格式的總表。")

    # 1. 參數設定
    with st.expander("⚙️ 點數權重設定", expanded=False):
        col1, col2, col3 = st.columns(3)
        p_a2 = col1.number_input("A2 點數/件", value=3.0, step=0.5)
        p_a3 = col2.number_input("A3 點數/件", value=1.0, step=0.5)
        p_traf = col3.number_input("交整點數/小時", value=1.0, step=0.5)

    # 2. 檔案上傳
    st.subheader("📂 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】(含各單位名單與取締數據)", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)", type=['xlsx'], accept_multiple_files=True)

    # 3. 執行運算
    if st.button("🚀 開始彙整並產生總表", type="primary"):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種當月資料皆已上傳！")
            return

        with st.spinner("正在進行深度彙整與總表計算..."):
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
                
                # 初始化總表容器
                summary_data = [] 
                g_cite, g_acc, g_traf, g_all = 0, 0, 0, 0

                # 遍歷所有分頁進行填值
                for sheet_name, df in dfs.items():
                    # 強制轉為 object 以免寫入數字報錯
                    df = df.astype(object)
                    dfs[sheet_name] = df
                    
                    if '總表' in sheet_name:
                        continue # 總表留到最後處理
                    
                    header_row = None
                    cols = {}
                    # 尋找標題列
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
                        
                        # 逐人填值
                        for r in range(header_row + 1, len(df)):
                            name = str(df.iloc[r, cols['員警姓名']]).strip()
                            
                            # 處理小計列
                            if name in ['小計', '總計']:
                                # 讀取該單位原本的取締點數小計
                                c_sub = pd.to_numeric(df.iloc[r, cols['取締點數']], errors='coerce') or 0
                                df.iloc[r, cols['事故點數']] = s_acc if s_acc > 0 else ""
                                df.iloc[r, cols['交整點數']] = s_traf if s_traf > 0 else ""
                                df.iloc[r, cols['個人總點數']] = c_sub + s_acc + s_traf
                                
                                if name == '小計':
                                    summary_data.append([sheet_name, c_sub, s_acc, s_traf, c_sub + s_acc + s_traf])
                                    g_cite += c_sub; g_acc += s_acc; g_traf += s_traf; g_all += (c_sub + s_acc + s_traf)
                                break
                            
                            # 填寫員警個人數據
                            elif name not in ['nan', 'None', '', 'nan']:
                                a2 = dict_acc.get(name, {}).get('A2類', 0)
                                a3 = dict_acc.get(name, {}).get('A3類', 0)
                                t_h = dict_traf.get(name, 0)
                                
                                ap = a2 * p_a2 + a3 * p_a3
                                tp = t_h * p_traf
                                cp = pd.to_numeric(df.iloc[r, cols['取締點數']], errors='coerce') or 0
                                
                                df.iloc[r, cols['A2件數']] = a2 if a2 > 0 else ""
                                df.iloc[r, cols['A3件數']] = a3 if a3 > 0 else ""
                                df.iloc[r, cols['事故點數']] = ap if ap > 0 else ""
                                df.iloc[r, cols['交整時數']] = t_h if t_h > 0 else ""
                                df.iloc[r, cols['交整點數']] = tp if tp > 0 else ""
                                df.iloc[r, cols['個人總點數']] = cp + ap + tp
                                
                                s_acc += ap; s_traf += tp; s_cite += cp; s_all += (cp + ap + tp)

                # --- 處理「總表」分頁 ---
                summary_sheet_name = next((k for k in dfs.keys() if '總表' in k), '總表')
                
                # 如果範本中沒有總表，我們建立一個
                if summary_sheet_name not in dfs:
                    df_sum = pd.DataFrame(columns=['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數'])
                else:
                    df_sum = dfs[summary_sheet_name]

                # 建立全新的總表內容
                new_summary = [['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']]
                new_summary.extend(summary_data)
                new_summary.append(['合計', g_cite, g_acc, g_traf, g_all])
                
                # 轉成 DataFrame 並填回
                df_final_sum = pd.DataFrame(new_summary)
                dfs[summary_sheet_name] = df_final_sum

                # --- 下載產出 ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for s_name, df_final in dfs.items():
                        # 保持不寫入 Index 與 Header，完全依照原格式
                        df_final.to_excel(writer, sheet_name=s_name, header=False, index=False)

                st.success("✅ 彙整完成！已自動產出並更新「總表」內容。")
                st.balloons()
                
                # 🌟 關鍵：直接使用上傳檔案的名稱作為下載名稱
                final_filename = file_template.name
                
                st.download_button(
                    label=f"📥 下載：{final_filename}", 
                    data=output.getvalue(), 
                    file_name=final_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")

if __name__ == "__main__":
    p18_page()
