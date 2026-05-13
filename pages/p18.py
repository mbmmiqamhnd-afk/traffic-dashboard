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
    st.success("【權重已更新】A2=10, A3=5, 交整=5。系統將強制修正所有單位(含交通組)的小計數值。")

    # 1. 點數權重設定 (依照您的需求調整預設值)
    with st.expander("⚙️ 點數權重設定 (目前已設為最新要求)", expanded=False):
        col1, col2, col3 = st.columns(3)
        p_a2 = col1.number_input("A2 點數/件", value=10.0, step=1.0)
        p_a3 = col2.number_input("A3 點數/件", value=5.0, step=1.0)
        p_traf = col3.number_input("交整點數/小時", value=5.0, step=1.0)

    # 2. 檔案上傳
    st.subheader("📂 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)", type=['xlsx'], accept_multiple_files=True)

    if st.button("🚀 執行彙整並強制修正所有數據", type="primary"):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種資料皆已完成上傳！")
            return

        with st.spinner("正在以新權重重新計算交通組等單位的小計..."):
            try:
                # --- A. 外部數據預處理 ---
                df_acc = pd.read_excel(file_acc, header=4)
                df_acc['姓名'] = df_acc['姓名'].astype(str).str.strip()
                dict_acc = df_acc.groupby('姓名')[['A2類', 'A3類']].sum().to_dict(orient='index')

                df_traf_all = pd.concat([pd.read_excel(f, sheet_name='月彙整總表') for f in file_traf_list])
                df_traf_all['姓名'] = df_traf_all['姓名'].astype(str).str.strip()
                dict_traf = df_traf_all.groupby('姓名')['總計尖峰時數'].sum().to_dict()

                # --- B. 讀取範本並偵測日期 ---
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
                    
                    # 定位標題
                    start_r, start_c = None, None
                    for r_idx, row in df.iterrows():
                        row_str = [str(x).strip() for x in row.values]
                        if '員警姓名' in row_str:
                            start_r, start_c = r_idx, row_str.index('員警姓名')
                            break
                    
                    if start_r is not None:
                        # 裁剪並清理
                        df_work = df.iloc[start_r:, start_c:].copy()
                        df_work.reset_index(drop=True, inplace=True)
                        df_work.columns = [str(c).strip() for c in df_work.iloc[0]]
                        df_work = df_work.drop(0).astype(object)
                        
                        col_map = {c: i for i, c in enumerate(df_work.columns)}
                        s_cite, s_acc, s_traf = 0, 0, 0
                        
                        # 辨識有效員警與小計行
                        member_indices = []
                        subtotal_idx = None
                        
                        for r in range(len(df_work)):
                            name_cell = str(df_work.iloc[r, col_map['員警姓名']]).strip()
                            if '小計' in name_cell or '總計' in name_cell:
                                subtotal_idx = r
                                break 
                            if name_cell not in ['nan', 'None', '']:
                                member_indices.append(r)

                        # 填寫並重新計算數據
                        for r in member_indices:
                            name = str(df_work.iloc[r, col_map['員警姓名']]).strip()
                            a2 = dict_acc.get(name, {}).get('A2類', 0)
                            a3 = dict_acc.get(name, {}).get('A3類', 0)
                            th = dict_traf.get(name, 0)
                            
                            ap, tp = a2 * p_a2 + a3 * p_a3, th * p_traf
                            cp = pd.to_numeric(df_work.iloc[r, col_map['取締點數']], errors='coerce') or 0
                            
                            # 強制覆寫所有欄位
                            if 'A2件數' in col_map: df_work.iloc[r, col_map['A2件數']] = a2 if a2 > 0 else ""
                            if 'A3件數' in col_map: df_work.iloc[r, col_map['A3件數']] = a3 if a3 > 0 else ""
                            if '事故點數' in col_map: df_work.iloc[r, col_map['事故點數']] = ap if ap > 0 else ""
                            if '交整時數' in col_map: df_work.iloc[r, col_map['交整時數']] = th if th > 0 else ""
                            if '交整點數' in col_map: df_work.iloc[r, col_map['交整點數']] = tp if tp > 0 else ""
                            if '個人總點數' in col_map: df_work.iloc[r, col_map['個人總點數']] = cp + ap + tp
                            
                            s_acc += ap; s_traf += tp; s_cite += cp

                        # 🌟 強制產出正確的小計列
                        if subtotal_idx is None:
                            # 若沒找到小計，手動在名單後新增一行
                            new_idx = len(df_work)
                            df_work.loc[new_idx] = [""] * len(df_work.columns)
                            df_work.iloc[new_idx, col_map['員警姓名']] = '小計'
                            subtotal_idx = new_idx

                        # 垂直重新加總 (確保數據絕對正確)
                        for col_n, c_i in col_map.items():
                            if col_n in ['員警姓名', '蓋章']: continue
                            col_sum = pd.to_numeric(df_work.iloc[member_indices, c_i], errors='coerce').sum()
                            df_work.iloc[subtotal_idx, c_i] = col_sum if col_sum > 0 else 0

                        summary_rows.append([sheet_name, s_cite, s_acc, s_traf, s_cite + s_acc + s_traf])
                        g_cite += s_cite; g_acc += s_acc; g_traf += s_traf; g_all += (s_cite + s_acc + s_traf)
                        
                        if '蓋章' in df_work.columns: df_work = df_work.drop(columns=['蓋章'])
                        final_sheets[sheet_name] = df_work

                # --- D. 總表與輸出 ---
                sum_header = ['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']
                df_summary = pd.DataFrame([sum_header] + summary_rows + [['合計', g_cite, g_acc, g_traf, g_all]])

                st.success(f"✅ 彙整完畢！日期：{ext_year}年{ext_month}月。權重：A2(10), A3(5), 交整(5)")
                st.table(pd.DataFrame(summary_rows + [['合計', g_cite, g_acc, g_traf, g_all]], columns=sum_header))

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_summary.to_excel(writer, sheet_name='總表', header=False, index=False)
                    for sn, df_f in final_sheets.items():
                        df_f.to_excel(writer, sheet_name=sn, index=False)

                filename = f"桃園市政府警察局龍潭分局{ext_year}年{ext_month}月份處理道路交通安全人員獎勵金點數統計表.xlsx"
                st.download_button(label=f"📥 下載統計表 (最新權重版)", data=output.getvalue(), file_name=filename)

            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")

if __name__ == "__main__":
    p18_page()
