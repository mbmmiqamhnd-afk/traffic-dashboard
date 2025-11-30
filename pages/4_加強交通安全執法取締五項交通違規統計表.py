import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="五項交通違規統計", layout="wide", page_icon="🚦")
st.title("🚦 加強交通安全執法取締統計表")

st.markdown("""
### 📝 操作說明
1. 請上傳 **6 個檔案** (本期/本年/去年 的「自選匯出」與「footman」)。
2. 自動執行：排除警備隊、交通組更名、整合行人數據、計算比較值。
""")

uploaded_files = st.file_uploader("請將 6 個檔案拖曳至此", accept_multiple_files=True)

if uploaded_files and st.button("🚀 開始分析", key="btn_five"):
    with st.spinner('正在分析數據中...'):
        try:
            file_map = {}
            for f in uploaded_files:
                name = f.name
                is_foot = 'footman' in name.lower()
                period = 'last' if '(2)' in name else ('curr' if '(1)' in name else 'week')
                file_map[f"{period}_{'foot' if is_foot else 'gen'}"] = {'file': f, 'name': name}

            def smart_read(fobj, fname):
                try:
                    fobj.seek(0)
                    if fname.endswith(('.xls', '.xlsx')): df = pd.read_excel(fobj, header=None, nrows=20)
                    else: df = pd.read_csv(fobj, header=None, nrows=20, encoding='utf-8')
                    idx = -1
                    for i, r in df.iterrows():
                        if '單位' in r.astype(str).values: idx = i; break
                    if idx == -1: idx = 3
                    fobj.seek(0)
                    if fname.endswith(('.xls', '.xlsx')): df = pd.read_excel(fobj, header=idx)
                    else: df = pd.read_csv(fobj, header=idx)
                    df.columns = [str(c).strip() for c in df.columns]
                    if '單位' not in df.columns:
                        match = [c for c in df.columns if '單位' in c]
                        if match: df.rename(columns={match[0]: '單位'}, inplace=True)
                        else: return pd.DataFrame(columns=['單位'])
                    return df
                except: return pd.DataFrame(columns=['單位'])

            def process(key_gen, key_foot, suffix):
                if key_gen not in file_map: return pd.DataFrame(columns=['單位'])
                df = smart_read(file_map[key_gen]['file'], file_map[key_gen]['name'])
                df = df[~df['單位'].isin(['合計', '總計', '小計', 'nan'])].dropna(subset=['單位']).copy()
                df['單位'] = df['單位'].astype(str).str.strip()
                for c in df.columns:
                    if c!='單位' and df[c].dtype=='object':
                        df[c] = pd.to_numeric(df[c].astype(str).str.replace(',','').str.replace('nan','0'), errors='coerce').fillna(0)
                
                cols = df.columns
                dui = [c for c in cols if str(c).startswith('35條')] + [c for c in ['73條2項','73條3項'] if c in cols]
                red = [c for c in cols if str(c).startswith('53條')]
                spd = [c for c in cols if str(c).startswith('43條')]
                yld = [c for c in cols if str(c).startswith('44條') or str(c).startswith('48條')]
                
                res = pd.DataFrame()
                res['單位'] = df['單位']
                res[f'酒駕_{suffix}'] = df[dui].sum(axis=1); res[f'闖紅燈_{suffix}'] = df[red].sum(axis=1)
                res[f'嚴重超速_{suffix}'] = df[spd].sum(axis=1); res[f'車不讓人_{suffix}'] = df[yld].sum(axis=1)
                
                if key_foot in file_map:
                    foot = smart_read(file_map[key_foot]['file'], file_map[key_foot]['name'])
                    ped_col = next((c for c in foot.columns if '78' in str(c)), None)
                    if ped_col:
                        if foot[ped_col].dtype=='object': foot[ped_col] = pd.to_numeric(foot[ped_col].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                        foot['單位'] = foot['單位'].astype(str).str.strip()
                        res = res.merge(foot[['單位', ped_col]], on='單位', how='left')
                        res.rename(columns={ped_col: f'行人違規_{suffix}'}, inplace=True)
                
                if f'行人違規_{suffix}' not in res.columns: res[f'行人違規_{suffix}'] = 0
                res[f'行人違規_{suffix}'] = res[f'行人違規_{suffix}'].fillna(0)
                return res

            df_w = process('week_gen', 'week_foot', '本期')
            df_c = process('curr_gen', 'curr_foot', '本年')
            df_l = process('last_gen', 'last_foot', '去年')

            full = df_c.merge(df_l, on='單位', how='outer').merge(df_w, on='單位', how='left').fillna(0)
            u_map = {'龍潭交通分隊': '交通分隊', '交通組': '科技執法', '聖亭派出所': '聖亭所', '龍潭派出所': '龍潭所', '中興派出所': '中興所', '石門派出所': '石門所', '高平派出所': '高平所', '三和派出所': '三和所'}
            full['Target_Unit'] = full['單位'].map(u_map)
            final = full[full['Target_Unit'].notna()].copy()

            if final.empty: st.error("❌ 資料錯誤：找不到對應單位"); st.stop()

            cats = ['酒駕', '闖紅燈', '嚴重超速', '車不讓人', '行人違規']
            for c in cats: final[f'{c}_比較'] = final[f'{c}_本年'] - final[f'{c}_去年']

            num_cols = final.columns.drop(['單位', 'Target_Unit'])
            total_row = final[num_cols].sum().to_frame().T; total_row['Target_Unit'] = '合計'
            result = pd.concat([total_row, final], ignore_index=True)

            order = ['合計', '科技執法', '交通分隊', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
            result['Target_Unit'] = pd.Categorical(result['Target_Unit'], categories=order, ordered=True)
            result.sort_values('Target_Unit', inplace=True)

            cols_out = ['Target_Unit']
            for p in ['本期', '本年', '去年', '比較']:
                for c in cats: cols_out.append(f'{c}_{p}')
            
            final_table = result[cols_out].copy()
            final_table.rename(columns={'Target_Unit': '單位'}, inplace=True)
            try: final_table.iloc[:, 1:] = final_table.iloc[:, 1:].astype(int)
            except: pass

            st.success("✅ 分析完成！")
            st.dataframe(final_table, use_container_width=True)
            csv = final_table.to_csv(index=False).encode('utf-8-sig')
            st.download_button(label="📥 下載 CSV", data=csv, file_name='交通違規統計.csv', mime='text/csv')

        except Exception as e: st.error(f"發生錯誤：{e}")
