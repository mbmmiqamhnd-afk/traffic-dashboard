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
    st.warning("【數據精確修正】系統將徹底無視原始檔中的錯誤小計，並於名單末端重新產生正確加總。")

    # 1. 點數權重設定
    with st.expander("⚙️ 點數權重設定", expanded=False):
        col1, col2, col3 = st.columns(3)
        p_a2 = col1.number_input("A2 點數/件", value=3.0, step=0.5)
        p_a3 = col2.number_input("A3 點數/件", value=1.0, step=0.5)
        p_traf = col3.number_input("交整點數/小時", value=1.0, step=0.5)

    # 2. 檔案上傳
    st.subheader("📂 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)", type=['xlsx'], accept_multiple_files=True)

    if st.button("🚀 執行彙整與強制更新小計", type="primary"):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種資料皆已完成上傳！")
            return

        with st.spinner("正在清除舊有數據並重新計算正確小計..."):
            try:
                # --- A. 外部數據預處理 ---
                df_acc = pd.read_excel(file_acc, header=4)
                df_acc['姓名'] = df_acc['姓名'].astype(str).str.strip()
                dict_acc = df_acc.groupby('姓名')[['A2類', 'A3類']].sum().to_dict(orient='index')

                df_traf_all = pd.concat([pd.read_excel(f, sheet_name='月彙整總表') for f in file_traf_list])
                df_traf_all['姓名'] = df_traf_all['姓名'].astype(str).str.strip()
                dict_traf = df_traf_all.groupby('姓名')['總計尖峰時數'].sum().to_dict()

                # --- B. 讀取範本並安全偵測日期 ---
                dfs_raw = pd.read_excel(file_template, sheet_name=None, header=None)
                ext_year, ext_month = "115", "4"
                found_date = False
                for _, df_scan in dfs_raw.items():
                    for r in range(min(20, len(df_scan))):
                        for c in range(min(15, len(df_scan.columns))):
                            cell_val = str(df_scan.iloc[r, c])
                            m = re.search(r'開單日期[：:\s]*(\d{3})(\d{2})', cell_val)
                            if m:
                                ext_year, ext_month = m.group(1), str(int(m.group(2)))
                                found_date = True; break
                        if found_date: break
                    if found_date: break

                # --- C. 處理分頁與小計強制覆寫 ---
                final_sheets = {}
                summary_rows = []
                g_cite, g_acc, g_traf, g_all = 0, 0, 0, 0

                for sheet_name, df in dfs_raw.items():
                    if '總表' in sheet_name: continue
                    
                    # 尋找「員警姓名」標題座標
                    start_r, start_c = None, None
                    for r_idx, row in df.iterrows():
                        row_str = [str(x).strip() for x in row.values]
                        if '員警姓名' in row_str:
                            start_r, start_c = r_idx, row_str.index('員警姓名')
                            break
                    
                    if start_r is not None:
                        # 裁剪：移除標題與左側空欄
                        df_clean = df.iloc[start_r:, start_c:].copy()
                        df_clean.reset_index(drop=True, inplace=True)
                        df_clean.columns = [str(c).strip() for c in df_clean.iloc[0]]
                        df_clean = df_clean.drop(0).astype(object)
                        
                        # 🌟 強制修正：過濾掉原始資料中「可能錯誤」的小計/總計行
                        df_clean['姓名檢核'] = df_clean['員警姓名'].astype(str).str.strip()
                        df_clean = df_clean[~df_clean['姓名檢核'].str.contains('小計|總計|合計', na=False)]
                        df_clean = df_clean[df_clean['姓名檢核'] != 'nan']
                        df_clean = df_clean[df_clean['姓名檢核'] != '']
                        
                        df_clean = df_clean.drop(columns=['姓名檢核'])
                        
                        col_map = {c: i for i, c in enumerate(df_clean.columns)}
                        s_cite, s_acc, s_traf = 0, 0, 0
                        
                        # 重新計算每位員警數值
                        for r in range(len(df_clean)):
                            name = str(df_clean.iloc[r, col_map['員警姓名']]).strip()
                            a2 = dict_acc.get(name, {}).get('A2類', 0)
                            a3 = dict_acc.get(name, {}).get('A3類', 0)
                            th = dict_traf.get(name, 0)
                            
                            ap, tp = a2 * p_a2 + a3 * p_a3, th * p_traf
                            cp = pd.to_numeric(df_clean.iloc[r, col_map['取締點數']], errors='coerce') or 0
                            
                            if 'A2件數' in col_map: df_clean.iloc[r, col_map['A2件數']] = a2 if a2 > 0 else ""
                            if 'A3件數' in col_map: df_clean.iloc[r, col_map['A3件數']] = a3 if a3 > 0 else ""
                            if '事故點數' in col_map: df_clean.iloc[r, col_map['事故點數']] = ap if ap > 0 else ""
                            if '交整時數' in col_map: df_clean.iloc[r, col_map['交整時數']] = th if th > 0 else ""
                            if '交整點數' in col_map: df_clean.iloc[r, col_map['交整點數']] = tp if tp > 0 else ""
                            if '個人總點數' in col_map: df_clean.iloc[r, col_map['個人總點數']] = cp + ap + tp
                            
                            s_acc += ap; s_traf += tp; s_cite += cp

                        # 🌟 核心修正：在最後一名員警後方插入程式計算出的「正確小計」
                        correct_sub_row = [""] * len(df_clean.columns)
                        correct_sub_row[col_map['員警姓名']] = '小計'
                        
                        for col_n, c_i in col_map.items():
                            if col_n in ['員警姓名', '蓋章']: continue
                            # 垂直加總目前畫面上所有員警的正確數值
                            v_sum = pd.to_numeric(df_clean.iloc[:, c_i], errors='coerce').sum()
                            correct_sub_row[c_i] = v_sum if v_sum > 0 else 0
                        
                        df_clean.loc[len(df_clean)] = correct_sub_row
                        
                        summary_rows.append([sheet_name, s_cite, s_acc, s_traf, s_cite + s_acc + s_traf])
                        g_cite += s_cite; g_acc += s_acc; g_traf += s_traf; g_all += (s_cite + s_acc + s_traf)
                        
                        # 移除蓋章欄
                        if '蓋章' in df_clean.columns: df_clean = df_clean.drop(columns=['蓋章'])
                        final_sheets[sheet_name] = df_clean

                # --- D. 總表與輸出 ---
                sum_header = ['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']
                df_summary = pd.DataFrame([sum_header] + summary_rows + [['合計', g_cite, g_acc, g_traf, g_all]])

                st.success(f"✅ 成功產出 {ext_year}年{ext_month}月 報表，交通組小計已修正為正確加總。")
                st.table(pd.DataFrame(summary_rows + [['合計', g_cite, g_acc, g_traf, g_all]], columns=sum_header))

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_summary.to_excel(writer, sheet_name='總表', header=False, index=False)
                    for sn, df_f in final_sheets.items():
                        df_f.to_excel(writer, sheet_name=sn, index=False)

                filename = f"桃園市政府警察局龍潭分局{ext_year}年{ext_month}月份處理道路交通安全人員獎勵金點數統計表.xlsx"
                st.download_button(label=f"📥 下載統計表 (小計數據強制修正版)", data=output.getvalue(), file_name=filename)

            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")

if __name__ == "__main__":
    p18_page()
