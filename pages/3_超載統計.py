import streamlit as st
import pandas as pd
import numpy as np
import re

st.set_page_config(page_title="超載統計", layout="wide", page_icon="🚛")
st.title("🚛 超載 (stoneCnt) 自動統計")

st.markdown("""
### 📝 使用說明
1. 請上傳 **3 個** `stoneCnt` 系列的 Excel 檔案。
2. 系統自動辨識 `(1)`本年、`(2)`去年、無括號本期。
3. **自動排除**「警備隊」列入合計。
4. **自動帶入**各單位目標值。
""")

uploaded_files = st.file_uploader("請拖曳 3 個 stoneCnt 檔案至此", accept_multiple_files=True, type=['xlsx', 'xls'])

TARGETS = {'聖亭所': 24, '龍潭所': 32, '中興所': 24, '石門所': 19, '高平所': 16, '三和所': 9, '警備隊': 0, '交通分隊': 30}
UNIT_MAP = {'聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所', '警備隊': '警備隊', '龍潭交通分隊': '交通分隊'}
UNIT_ORDER = ['聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']

if uploaded_files and st.button("🚀 開始計算", key="btn_stone"):
    with st.spinner("正在分析超載數據..."):
        try:
            files_config = {"Week": None, "YTD": None, "Last_YTD": None}
            for f in uploaded_files:
                if "(1)" in f.name: files_config["YTD"] = f
                elif "(2)" in f.name: files_config["Last_YTD"] = f
                else: files_config["Week"] = f
            
            def parse_stone(f):
                if not f: return {}
                counts = {}
                xls = pd.ExcelFile(f)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet, header=None)
                    curr = None
                    for _, row in df.iterrows():
                        s = row.astype(str).str.cat(sep=' ')
                        if "舉發單位：" in s:
                            m = re.search(r"舉發單位：(\S+)", s)
                            if m: curr = m.group(1).strip()
                        if "總計" in s and curr:
                            nums = [float(x) for x in row if str(x).replace('.','',1).isdigit()]
                            if nums:
                                short = UNIT_MAP.get(curr, curr)
                                counts[short] = counts.get(short, 0) + int(nums[-1])
                                curr = None
                return counts

            d_wk = parse_stone(files_config["Week"])
            d_yt = parse_stone(files_config["YTD"])
            d_ly = parse_stone(files_config["Last_YTD"])

            rows = []
            for u in UNIT_ORDER:
                rows.append({
                    '單位': u, '本期': d_wk.get(u,0), '本年累計': d_yt.get(u,0), '去年累計': d_ly.get(u,0), '目標值': TARGETS.get(u,0)
                })
            
            df = pd.DataFrame(rows)
            df_calc = df.copy()
            df_calc.loc[df_calc['單位']=='警備隊', ['本期', '本年累計', '去年累計', '目標值']] = 0
            
            total = df_calc[['本期', '本年累計', '去年累計', '目標值']].sum().to_dict()
            total['單位'] = '合計'
            
            df_final = pd.concat([pd.DataFrame([total]), df], ignore_index=True)
            df_final['本年與去年同期比較'] = df_final['本年累計'] - df_final['去年累計']
            df_final['達成率'] = df_final.apply(lambda x: f"{x['本年累計']/x['目標值']:.2%}" if x['目標值']>0 else "—", axis=1)
            
            # 警備隊特殊顯示
            df_final.loc[df_final['單位']=='警備隊', ['本年與去年同期比較', '目標值', '達成率']] = "—"
            
            cols = ['單位', '本期', '本年累計', '去年累計', '本年與去年同期比較', '目標值', '達成率']
            df_final = df_final[cols]
            
            st.success("✅ 分析完成！")
            st.dataframe(df_final, use_container_width=True, hide_index=True)
            
            csv = df_final.to_csv(index=False).encode('utf-8-sig')
            st.download_button(label="📥 下載 CSV", data=csv, file_name='超載統計表.csv', mime='text/csv')

        except Exception as e: st.error(f"錯誤：{e}")
