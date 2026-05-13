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
    st.info("已升級「總表自動彙整」功能！系統會自動比對外部數據、填入各派出所分頁，最後自動生成完美的總表頁面。")

    # 1. 參數設定
    with st.expander("⚙️ 點數權重設定 (點擊可展開/收合)", expanded=True):
        col1, col2, col3 = st.columns(3)
        p_a2 = col1.number_input("A2 點數/件", value=3.0, step=0.5)
        p_a3 = col2.number_input("A3 點數/件", value=1.0, step=0.5)
        p_traf = col3.number_input("交整點數/小時", value=1.0, step=0.5)

    # 2. 檔案上傳
    st.subheader("📂 原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳【F龍潭分局_統計表】\n(需包含各單位與『總表』分頁)", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳【處理交通事故案件統計表】\n(內含 A1/A2/A3 數據)", type=['xls', 'xlsx'])
    
    file_traf_list = st.file_uploader("3. 上傳【各單位_交通疏導統計】\n(可一次框選多個派出所檔案)", type=['xlsx'], accept_multiple_files=True)

    # 3. 執行運算
    if st.button("🚀 開始彙整運算", type="primary"):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種檔案皆已完成上傳！")
            return

        with st.spinner("資料比對與運算中，請稍候..."):
            try:
                # ==========================================
                # A. 處理【交通事故資料】
                # ==========================================
                df_acc = pd.read_excel(file_acc, header=4)
                df_acc['姓名'] = df_acc['姓名'].astype(str).str.strip()
                df_acc['A2類'] = pd.to_numeric(df_acc['A2類'], errors='coerce').fillna(0)
                df_acc['A3類'] = pd.to_numeric(df_acc['A3類'], errors='coerce').fillna(0)
                
                df_acc = df_acc[~df_acc['姓名'].isin(['nan', 'None', '', 'NaN'])]
                dict_acc = df_acc.groupby('姓名')[['A2類', 'A3類']].sum().to_dict(orient='index')

                # ==========================================
                # B. 處理【交通疏導資料】
                # ==========================================
                df_traf_list = []
                for f in file_traf_list:
                    df_t = pd.read_excel(f, sheet_name='月彙整總表')
                    df_traf_list.append(df_t)
                
                df_traf_all = pd.concat(df_traf_list)
                df_traf_all['姓名'] = df_traf_all['姓名'].astype(str).str.strip()
                df_traf_all['總計尖峰時數'] = pd.to_numeric(df_traf_all['總計尖峰時數'], errors='coerce').fillna(0)
                
                df_traf_all = df_traf_all[~df_traf_all['姓名'].isin(['nan', 'None', '', 'NaN'])]
                dict_traf = df_traf_all.groupby('姓名')['總計尖峰時數'].sum().to_dict()

                # ==========================================
                # C. 讀取並填寫【基底統計表】
                # ==========================================
                dfs = pd.read_excel(file_template, sheet_name=None, header=None)
                
                # 解除 Pandas 型別鎖定，允許文字與數字混合
                for k in dfs.keys():
                    dfs[k] = dfs[k].astype(object)

                # --- 核心：三階段彙整機制 ---
                g_a2, g_a3, g_acc_p, g_t_h, g_t_p, g_cite_p, g_total_p = 0, 0, 0, 0, 0, 0, 0
                summary_stats = {}
                total_rows_to_fill = []

                # 【階段一】處理各派出所名單，計算每個人的點數並累加
                for sheet_name, df in dfs.items():
                    if '總表' in sheet_name:
                        continue
                        
                    header_row = None
                    cols = {}
                    for idx, row in df.iterrows():
                        vals = [str(x).strip() for x in row.values]
                        if '員警姓名' in vals and '取締件數' in vals:
                            header_row = idx
                            target_fields = ['員警姓名', 'A2件數', 'A3件數', '事故點數', '交整時數', '交整點數', '個人總點數', '取締點數']
                            for f in target_fields:
                                if f in vals: cols[f] = vals.index(f)
                            break
                    
                    if header_row is not None:
                        s_a2, s_a3, s_acc_p, s_t_h, s_t_p, s_cite_p, s_total_p = 0, 0, 0, 0, 0, 0, 0
                        
                        for r in range(header_row + 1, len(df)):
                            name = str(df.iloc[r, cols['員警姓名']]).strip()
                            
                            if name == '小計':
                                total_rows_to_fill.append({'sheet': sheet_name, 'row': r, 'type': '小計', 'cols': cols})
                            elif name == '總計':
                                total_rows_to_fill.append({'sheet': sheet_name, 'row': r, 'type': '總計', 'cols': cols})
                            elif name not in ['nan', 'None', '', 'NaN']:
                                a2 = dict_acc.get(name, {}).get('A2類', 0)
                                a3 = dict_acc.get(name, {}).get('A3類', 0)
                                t_h = dict_traf.get(name, 0)
                                
                                acc_p = a2 * p_a2 + a3 * p_a3
                                traf_p = t_h * p_traf
                                cite_p = pd.to_numeric(df.iloc[r, cols['取締點數']], errors='coerce')
                                cite_p = cite_p if not pd.isna(cite_p) else 0
                                total_p = cite_p + acc_p + traf_p
                                
                                df.iloc[r, cols['A2件數']] = a2 if a2 > 0 else ""
                                df.iloc[r, cols['A3件數']] = a3 if a3 > 0 else ""
                                df.iloc[r, cols['事故點數']] = acc_p if acc_p > 0 else ""
                                df.iloc[r, cols['交整時數']] = t_h if t_h > 0 else ""
                                df.iloc[r, cols['交整點數']] = traf_p if traf_p > 0 else ""
                                df.iloc[r, cols['個人總點數']] = total_p if total_p > 0 else ""
                                
                                s_a2 += a2; s_a3 += a3; s_acc_p += acc_p; s_t_h += t_h; s_t_p += traf_p
                                s_cite_p += cite_p; s_total_p += total_p
                                
                        # 儲存單位小計與全域總計
                        for item in reversed(total_rows_to_fill):
                            if item['sheet'] == sheet_name and item['type'] == '小計':
                                item['s_data'] = {'A2件數': s_a2, 'A3件數': s_a3, '事故點數': s_acc_p, '交整時數': s_t_h, '交整點數': s_t_p, '取締點數': s_cite_p, '個人總點數': s_total_p}
                                break
                                
                        summary_stats[sheet_name] = {
                            '取締點數': s_cite_p, '事故點數': s_acc_p, '交整點數': s_t_p, '個人總點數': s_total_p
                        }
                        
                        g_a2 += s_a2; g_a3 += s_a3; g_acc_p += s_acc_p; g_t_h += s_t_h; g_t_p += s_t_p
                        g_cite_p += s_cite_p; g_total_p += s_total_p

                # 【階段二】將計算好的「小計」與「總計」填回各單位分頁中
                for item in total_rows_to_fill:
                    df = dfs[item['sheet']]
                    r = item['row']
                    cols = item['cols']
                    if item['type'] == '小計':
                        sd = item.get('s_data', {})
                        df.iloc[r, cols['A2件數']] = sd.get('A2件數', 0) if sd.get('A2件數', 0) > 0 else ""
                        df.iloc[r, cols['A3件數']] = sd.get('A3件數', 0) if sd.get('A3件數', 0) > 0 else ""
                        df.iloc[r, cols['事故點數']] = sd.get('事故點數', 0) if sd.get('事故點數', 0) > 0 else ""
                        df.iloc[r, cols['交整時數']] = sd.get('交整時數', 0) if sd.get('交整時數', 0) > 0 else ""
                        df.iloc[r, cols['交整點數']] = sd.get('交整點數', 0) if sd.get('交整點數', 0) > 0 else ""
                        df.iloc[r, cols['取締點數']] = sd.get('取締點數', 0) if sd.get('取締點數', 0) > 0 else ""
                        df.iloc[r, cols['個人總點數']] = sd.get('個人總點數', 0) if sd.get('個人總點數', 0) > 0 else ""
                    elif item['type'] == '總計':
                        df.iloc[r, cols['A2件數']] = g_a2 if g_a2 > 0 else ""
                        df.iloc[r, cols['A3件數']] = g_a3 if g_a3 > 0 else ""
                        df.iloc[r, cols['事故點數']] = g_acc_p if g_acc_p > 0 else ""
                        df.iloc[r, cols['交整時數']] = g_t_h if g_t_h > 0 else ""
                        df.iloc[r, cols['交整點數']] = g_t_p if g_t_p > 0 else ""
                        df.iloc[r, cols['取締點數']] = g_cite_p if g_cite_p > 0 else ""
                        df.iloc[r, cols['個人總點數']] = g_total_p if g_total_p > 0 else ""

                # 【階段三】精準覆寫「總表」分頁
                summary_sheet_key = next((k for k in dfs.keys() if '總表' in k), None)
                if summary_sheet_key:
                    df_sum = dfs[summary_sheet_key]
                    header_row = None
                    cols = {}
                    for idx, row in df_sum.iterrows():
                        vals = [str(x).strip() for x in row.values]
                        if '單位名稱' in vals:
                            header_row = idx
                            for f in ['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']:
                                if f in vals: cols[f] = vals.index(f)
                            break
                            
                    if header_row is not None:
                        curr_r = header_row + 1
                        for unit_name, stats in summary_stats.items():
                            if curr_r >= len(df_sum):
                                df_sum.loc[curr_r] = [""] * len(df_sum.columns)
                            
                            df_sum.iloc[curr_r, cols.get('單位名稱', 0)] = unit_name
                            if '取締點數' in cols: df_sum.iloc[curr_r, cols['取締點數']] = stats['取締點數']
                            if '事故點數' in cols: df_sum.iloc[curr_r, cols['事故點數']] = stats['事故點數']
                            if '交整點數' in cols: df_sum.iloc[curr_r, cols['交整點數']] = stats['交整點數']
                            if '個人總點數' in cols: df_sum.iloc[curr_r, cols['個人總點數']] = stats['個人總點數']
                            curr_r += 1
                            
                        # 寫入最後的「合計」
                        if curr_r >= len(df_sum):
                            df_sum.loc[curr_r] = [""] * len(df_sum.columns)
                        df_sum.iloc[curr_r, cols.get('單位名稱', 0)] = '合計'
                        if '取締點數' in cols: df_sum.iloc[curr_r, cols['取締點數']] = g_cite_p
                        if '事故點數' in cols: df_sum.iloc[curr_r, cols['事故點數']] = g_acc_p
                        if '交整點數' in cols: df_sum.iloc[curr_r, cols['交整點數']] = g_traf_p
                        if '個人總點數' in cols: df_sum.iloc[curr_r, cols['個人總點數']] = g_total_p
                        curr_r += 1
                        
                        # 清除多餘的空白列，確保版面乾淨
                        while curr_r < len(df_sum):
                            for c in cols.values():
                                df_sum.iloc[curr_r, c] = ""
                            curr_r += 1

                # 4. 產出 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for sheet_name, df in dfs.items():
                        df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)

                st.success("✅ 報表與【總表】皆彙整成功！排版格式已完美還原。")
                st.balloons()
                
                st.download_button(
                    label="📥 下載已彙整之 龍潭分局獎勵金統計表.xlsx",
                    data=output.getvalue(),
                    file_name="F龍潭分局_統計表_含總表彙整.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ 處理過程中發生錯誤：{str(e)}")
                st.info("請確認上傳的檔案是否為原本的格式。")

if __name__ == "__main__":
    p18_page()
