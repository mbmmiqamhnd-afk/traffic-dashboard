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
        pass # 若讀取不到側邊欄，維持不報錯繼續執行

def p18_page():
    show_sidebar()

    st.title("💰 龍潭分局 - 獎勵金點數統計表產生器")
    st.info("本工具會自動比對姓名，將外部的「交通事故」與「交通疏導」數據填入您的「F龍潭分局_統計表」中，並自動加總小計與總計。")

    # 1. 參數設定
    with st.expander("⚙️ 點數權重設定 (點擊可展開/收合)", expanded=True):
        col1, col2, col3 = st.columns(3)
        p_a2 = col1.number_input("A2 點數/件", value=3.0, step=0.5)
        p_a3 = col2.number_input("A3 點數/件", value=1.0, step=0.5)
        p_traf = col3.number_input("交整點數/小時", value=1.0, step=0.5)

    # 2. 檔案上傳
    st.subheader("📂 原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳【F龍潭分局_統計表】\n(作為基底格式，內含取締數據)", type=['xlsx'])
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
                
                # 過濾空白姓名並加總
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
                
                # 過濾空白姓名並加總
                df_traf_all = df_traf_all[~df_traf_all['姓名'].isin(['nan', 'None', '', 'NaN'])]
                dict_traf = df_traf_all.groupby('姓名')['總計尖峰時數'].sum().to_dict()

                # ==========================================
                # C. 讀取並填寫【基底統計表】
                # ==========================================
                dfs = pd.read_excel(file_template, sheet_name=None, header=None)
                
                # 🌟 修復核心：解除 Pandas 3.0 的嚴格型別鎖定
                # 將所有分頁的資料格式轉為 object，允許在同一欄內混寫中文字與數字
                for k in dfs.keys():
                    dfs[k] = dfs[k].astype(object)

                output = io.BytesIO()
                g_a2, g_a3, g_acc_p, g_t_h, g_t_p = 0, 0, 0, 0, 0

                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for sheet_name, df in dfs.items():
                        header_row = None
                        cols = {}
                        
                        for idx, row in df.iterrows():
                            vals = [str(x).strip() for x in row.values]
                            if '員警姓名' in vals and '取締件數' in vals:
                                header_row = idx
                                target_fields = ['員警姓名', 'A2件數', 'A3件數', '事故點數', '交整時數', '交整點數', '個人總點數', '取締點數']
                                for f in target_fields:
                                    if f in vals: 
                                        cols[f] = vals.index(f)
                                break
                        
                        if header_row is not None:
                            s_a2, s_a3, s_acc_p, s_t_h, s_t_p = 0, 0, 0, 0, 0

                            for r in range(header_row + 1, len(df)):
                                name = str(df.iloc[r, cols['員警姓名']]).strip()
                                
                                if name == '小計':
                                    df.iloc[r, cols['A2件數']] = s_a2 if s_a2 > 0 else ""
                                    df.iloc[r, cols['A3件數']] = s_a3 if s_a3 > 0 else ""
                                    df.iloc[r, cols['事故點數']] = s_acc_p if s_acc_p > 0 else ""
                                    df.iloc[r, cols['交整時數']] = s_t_h if s_t_h > 0 else ""
                                    df.iloc[r, cols['交整點數']] = s_t_p if s_t_p > 0 else ""
                                    
                                    cite_p = pd.to_numeric(df.iloc[r, cols['取締點數']], errors='coerce')
                                    cite_p = cite_p if not pd.isna(cite_p) else 0
                                    df.iloc[r, cols['個人總點數']] = cite_p + s_acc_p + s_t_p
                                    
                                    g_a2 += s_a2; g_a3 += s_a3; g_acc_p += s_acc_p; g_t_h += s_t_h; g_t_p += s_t_p

                                elif name == '總計':
                                    df.iloc[r, cols['A2件數']] = g_a2 if g_a2 > 0 else ""
                                    df.iloc[r, cols['A3件數']] = g_a3 if g_a3 > 0 else ""
                                    df.iloc[r, cols['事故點數']] = g_acc_p if g_acc_p > 0 else ""
                                    df.iloc[r, cols['交整時數']] = g_t_h if g_t_h > 0 else ""
                                    df.iloc[r, cols['交整點數']] = g_t_p if g_t_p > 0 else ""
                                    
                                    cite_p = pd.to_numeric(df.iloc[r, cols['取締點數']], errors='coerce')
                                    cite_p = cite_p if not pd.isna(cite_p) else 0
                                    df.iloc[r, cols['個人總點數']] = cite_p + g_acc_p + g_t_p

                                elif name not in ['nan', 'None', '', 'NaN']:
                                    a2 = dict_acc.get(name, {}).get('A2類', 0)
                                    a3 = dict_acc.get(name, {}).get('A3類', 0)
                                    t_h = dict_traf.get(name, 0)
                                    
                                    acc_p = a2 * p_a2 + a3 * p_a3
                                    traf_p = t_h * p_traf
                                    cite_p = pd.to_numeric(df.iloc[r, cols['取締點數']], errors='coerce')
                                    cite_p = cite_p if not pd.isna(cite_p) else 0
                                    
                                    df.iloc[r, cols['A2件數']] = a2 if a2 > 0 else ""
                                    df.iloc[r, cols['A3件數']] = a3 if a3 > 0 else ""
                                    df.iloc[r, cols['事故點數']] = acc_p if acc_p > 0 else ""
                                    df.iloc[r, cols['交整時數']] = t_h if t_h > 0 else ""
                                    df.iloc[r, cols['交整點數']] = traf_p if traf_p > 0 else ""
                                    df.iloc[r, cols['個人總點數']] = cite_p + acc_p + traf_p
                                    
                                    s_a2 += a2; s_a3 += a3; s_acc_p += acc_p; s_t_h += t_h; s_t_p += traf_p

                        df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)

                st.success("✅ 報表彙整成功！排版格式已完美保留。")
                st.balloons()
                
                st.download_button(
                    label="📥 下載已彙整之 龍潭分局獎勵金統計表.xlsx",
                    data=output.getvalue(),
                    file_name="F龍潭分局_統計表_自動彙整完成.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ 處理過程中發生錯誤：{str(e)}")
                st.info("請確認上傳的檔案是否為原本的格式。")

if __name__ == "__main__":
    p18_page()
