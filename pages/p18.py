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
    st.info("本工具會以當月上傳的統計表為基底，自動比對姓名並填入外部數據，最後自動彙整生成「總表」。")

    # 1. 參數設定
    with st.expander("⚙️ 點數權重設定 (預設已設定完畢，點擊可展開修改)", expanded=False):
        col1, col2, col3 = st.columns(3)
        p_a2 = col1.number_input("A2 點數/件", value=3.0, step=0.5)
        p_a3 = col2.number_input("A3 點數/件", value=1.0, step=0.5)
        p_traf = col3.number_input("交整點數/小時", value=1.0, step=0.5)

    # 2. 檔案上傳
    st.subheader("📂 當月原始資料上傳")
    c1, c2 = st.columns(2)
    # 更改了說明文字，符合您實際的操作邏輯
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】\n(內含當月名單與取締數據)", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】\n(內含 A1/A2/A3 數據)", type=['xls', 'xlsx'])
    
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】\n(可一次框選多個派出所檔案)", type=['xlsx'], accept_multiple_files=True)

    # 3. 執行運算
    if st.button("🚀 開始彙整並產生總表", type="primary"):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種當月檔案皆已完成上傳！")
            return

        with st.spinner("正在跨表彙整數據並產生總表..."):
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

                # --- 讀取當月基底檔案 ---
                dfs = pd.read_excel(file_template, sheet_name=None, header=None)
                
                # 初始化統計容器
                unit_summary = [] 
                grand_total = {"cite": 0, "acc": 0, "traf": 0, "all": 0} 

                # 解除型別鎖定
                for k in dfs.keys():
                    dfs[k] = dfs[k].astype(object)

                # 第一輪：處理各單位分頁
                for sheet_name, df in dfs.items():
                    if '總表' in sheet_name: continue
                    
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
                            
                            if name in ['小計', '總計']:
                                df.iloc[r, cols['事故點數']] = s_acc if s_acc > 0 else ""
                                df.iloc[r, cols['交整點數']] = s_traf if s_traf > 0 else ""
                                c_p = pd.to_numeric(df.iloc[r, cols['取締點數']], errors='coerce') or 0
                                df.iloc[r, cols['個人總點數']] = c_p + s_acc + s_traf
                                
                                if name == '小計':
                                    unit_summary.append({
                                        "單位": sheet_name, "取締": c_p, "事故": s_acc, "交整": s_traf, "總計": c_p + s_acc + s_traf
                                    })
                                    grand_total["cite"] += c_p
                                    grand_total["acc"] += s_acc
                                    grand_total["traf"] += s_traf
                                    grand_total["all"] += (c_p + s_acc + s_traf)
                                break
                            
                            elif name not in ['nan', 'None', '']:
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

                # 第二輪：精準填寫「總表」
                if '總表' in dfs:
                    df_sum = dfs['總表']
                    h_row = None
                    s_cols = {}
                    for idx, row in df_sum.iterrows():
                        vals = [str(x).strip() for x in row.values]
                        if '單位名稱' in vals:
                            h_row = idx
                            for f in ['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']:
                                if f in vals: s_cols[f] = vals.index(f)
                            break
                    
                    if h_row is not None:
                        curr_r = h_row + 1
                        for item in unit_summary:
                            if curr_r < len(df_sum):
                                df_sum.iloc[curr_r, s_cols['單位名稱']] = item['單位']
                                df_sum.iloc[curr_r, s_cols['取締點數']] = item['取締']
                                df_sum.iloc[curr_r, s_cols['事故點數']] = item['事故']
                                df_sum.iloc[curr_r, s_cols['交整點數']] = item['交整']
                                df_sum.iloc[curr_r, s_cols['個人總點數']] = item['總計']
                                curr_r += 1
                        
                        for r_idx in range(curr_r, len(df_sum)):
                            if str(df_sum.iloc[r_idx, s_cols['單位名稱']]).strip() == '合計':
                                df_sum.iloc[r_idx, s_cols['取締點數']] = grand_total["cite"]
                                df_sum.iloc[r_idx, s_cols['事故點數']] = grand_total["acc"]
                                df_sum.iloc[r_idx, s_cols['交整點數']] = grand_total["traf"]
                                df_sum.iloc[r_idx, s_cols['個人總點數']] = grand_total["all"]
                                break

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for s_name, df_final in dfs.items():
                        df_final.to_excel(writer, sheet_name=s_name, header=False, index=False)

                st.success("✅ 當月統計表彙整完畢！")
                
                # 🌟 極致自動化：直接沿用上傳檔案的名稱！
                final_filename = file_template.name
                
                st.download_button(
                    label=f"📥 下載彙整完成之統計表", 
                    data=output.getvalue(), 
                    file_name=final_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")

if __name__ == "__main__":
    p18_page()
