import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# 設定頁面配置
st.set_page_config(page_title="交通事故統計自動化", page_icon="🚓", layout="wide")

def main():
    st.title("🚓 交通事故統計自動化工具")
    st.markdown("請上傳 **本週**、**今年累計**、**去年累計** 三份報表（支援 `.csv` 或 `.xlsx`），系統將自動辨識並產出報表。")

    # 1. 檔案上傳區 (取代 google.colab.files)
    uploaded_files = st.file_uploader("請一次選取三個檔案", accept_multiple_files=True, type=['csv', 'xlsx'])

    if len(uploaded_files) == 3:
        if st.button("開始分析"):
            with st.spinner('正在解析檔案與計算數據...'):
                try:
                    process_files(uploaded_files)
                except Exception as e:
                    st.error(f"發生錯誤：{e}")
    elif len(uploaded_files) > 0 and len(uploaded_files) != 3:
        st.warning(f"目前已上傳 {len(uploaded_files)} 個檔案，請確保剛好上傳 3 個檔案。")

def parse_police_stats_raw(file_obj):
    """讀取檔案並回傳 DataFrame"""
    try:
        # Streamlit 的 UploadedFile 可以直接讀取
        df_raw = pd.read_csv(file_obj, header=None)
    except:
        file_obj.seek(0)
        df_raw = pd.read_excel(file_obj, header=None)
    return df_raw

def process_files(uploaded_files):
    # --- 2. 智慧辨識檔案身分 ---
    file_data_map = []
    
    for file_obj in uploaded_files:
        # 重置指針以防讀取錯誤
        file_obj.seek(0)
        df = parse_police_stats_raw(file_obj)
        
        try:
            # 抓取日期字串
            date_str = df.iloc[1, 0].replace("統計日期：", "").strip()
            dates = re.findall(r'(\d{3})/(\d{2})/(\d{2})', date_str)
            
            if not dates:
                st.warning(f"無法識別日期：{file_obj.name}")
                continue
                
            start_y, start_m, start_d = map(int, dates[0])
            end_y, end_m, end_d = map(int, dates[1])
            
            # 判斷邏輯
            month_diff = (end_y - start_y) * 12 + (end_m - start_m)
            
            if month_diff == 0 and (end_d - start_d) < 20:
                category = 'weekly'
            else:
                category = f'cumulative_{start_y}'
            
            file_data_map.append({
                'df': df,
                'date_str': date_str,
                'category': category,
                'year': start_y
            })
        except Exception as e:
            st.error(f"檔案解析失敗 {file_obj.name}: {e}")
            return

    # 分配角色
    df_wk, df_cur, df_lst = None, None, None
    d_wk, d_cur, d_lst = "", "", ""

    # 找出 Weekly
    for data in file_data_map:
        if data['category'] == 'weekly':
            df_wk = data['df']
            d_wk = data['date_str']
            break
            
    # 找出 Current 和 Last (比較年份)
    cumulative_files = [d for d in file_data_map if 'cumulative' in d['category']]
    if len(cumulative_files) >= 2:
        cumulative_files.sort(key=lambda x: x['year'], reverse=True)
        df_cur, d_cur = cumulative_files[0]['df'], cumulative_files[0]['date_str']
        df_lst, d_lst = cumulative_files[1]['df'], cumulative_files[1]['date_str']
    
    if df_wk is None or df_cur is None or df_lst is None:
        st.error("❌ 自動辨識失敗，無法區分本週、今年與去年檔案，請檢查檔案內容。")
        return

    st.success(f"✅ 成功辨識：\n- **本期**: {d_wk}\n- **今年**: {d_cur}\n- **去年**: {d_lst}")

    # --- 3. 資料清理與計算 ---
    df_wk_clean = process_data(df_wk)
    df_cur_clean = process_data(df_cur)
    df_lst_clean = process_data(df_lst)

    # 準備標題日期
    h_wk = format_date(d_wk)
    h_cur = format_date(d_cur)
    h_lst = format_date(d_lst)

    # --- 合併 A1 ---
    a1_wk = df_wk_clean[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'wk'})
    a1_cur = df_cur_clean[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'cur'})
    a1_lst = df_lst_clean[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'last'})
    
    m_a1 = pd.merge(a1_wk, a1_cur, on='Station_Short', how='outer')
    m_a1 = pd.merge(m_a1, a1_lst, on='Station_Short', how='outer').fillna(0)
    m_a1['Diff'] = m_a1['cur'] - m_a1['last']

    # --- 合併 A2 ---
    a2_wk = df_wk_clean[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'wk'})
    a2_cur = df_cur_clean[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'cur'})
    a2_lst = df_lst_clean[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'last'})
    
    m_a2 = pd.merge(a2_wk, a2_cur, on='Station_Short', how='outer')
    m_a2 = pd.merge(m_a2, a2_lst, on='Station_Short', how='outer').fillna(0)
    m_a2['Diff'] = m_a2['cur'] - m_a2['last']
    m_a2['Pct'] = m_a2.apply(lambda x: (x['Diff']/x['last']) if x['last']!=0 else 0, axis=1)
    m_a2['Pct_Str'] = m_a2['Pct'].apply(lambda x: f"{x:.2%}")
    m_a2['Prev'] = "-"

    # 排序
    m_a1 = sort_stations(m_a1)
    m_a2 = sort_stations(m_a2)

    # 整理最終表格
    a1_final = m_a1[['Station_Short', 'wk', 'cur', 'last', 'Diff']].copy()
    a1_final.columns = ['單位', f'本期({h_wk})', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較']
    
    # 顯示用的 A2 表 (含 % 字串)
    a2_display = m_a2[['Station_Short', 'wk', 'Prev', 'cur', 'last', 'Diff', 'Pct_Str']].copy()
    a2_display.columns = ['單位', f'本期({h_wk})', '前期', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較', '本年較去年增減比例']

    # 下載用的 A2 表 (含 % 數值，方便 Excel 格式化)
    a2_download = m_a2[['Station_Short', 'wk', 'Prev', 'cur', 'last', 'Diff', 'Pct']].copy()
    a2_download.columns = ['單位', f'本期({h_wk})', '前期', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較', '本年較去年增減比例']

    # --- 4. 顯示結果與下載按鈕 ---
    st.markdown("### 📊 統計結果")
    
    st.subheader("1. A1 類交通事故死亡人數")
    st.dataframe(a1_final, use_container_width=True)
    
    st.subheader("2. A2 類交通事故受傷人數")
    st.dataframe(a2_display, use_container_width=True)

    # 產出 Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        a1_final.to_excel(writer, sheet_name='A1死亡人數', index=False)
        a2_download.to_excel(writer, sheet_name='A2受傷人數', index=False)
        
        # 設定 A2 百分比格式
        workbook  = writer.book
        worksheet = writer.sheets['A2受傷人數']
        percent_fmt = workbook.add_format({'num_format': '0.00%'})
        worksheet.set_column(6, 6, None, percent_fmt)
        
    output.seek(0)
    
    filename = f'交通事故統計表_{datetime.now().strftime("%Y%m%d")}.xlsx'
    
    st.download_button(
        label="📥 下載整理好的 Excel 報表",
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def process_data(df_raw):
    """資料清理核心邏輯"""
    df_data = df_raw[df_raw[0].notna()].copy()
    df_data = df_data[df_data[0].str.contains("總計|派出所")].copy()
    df_data = df_data.reset_index(drop=True)
    
    columns_map = {
        0: "Station", 1: "Total_Cases", 2: "Total_Deaths", 3: "Total_Injuries",
        4: "A1_Cases", 5: "A1_Deaths", 6: "A1_Injuries",
        7: "A2_Cases", 8: "A2_Deaths", 9: "A2_Injuries", 10: "A3_Cases"
    }
    df_data = df_data.rename(columns=columns_map)
    
    for c in list(columns_map.values()):
        if c not in df_data.columns: df_data[c] = 0
    df_data = df_data[list(columns_map.values())]
    
    for col in list(columns_map.values())[1:]:
        df_data[col] = pd.to_numeric(df_data[col].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
        
    df_data['Station_Short'] = df_data['Station'].str.replace('派出所', '所').str.replace('總計', '合計')
    
    # 重新計算合計
    df_stations = df_data[~df_data['Station_Short'].str.contains("合計")].copy()
    numeric_cols = df_data.columns[1:-1]
    total_row = df_stations[numeric_cols].sum()
    total_row['Station_Short'] = '合計'
    df_total = pd.DataFrame([total_row])
    
    return pd.concat([df_total, df_stations], ignore_index=True)

def sort_stations(df):
    target_order = ['合計', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
    order_map = {name: i for i, name in enumerate(target_order)}
    df['order'] = df['Station_Short'].map(order_map).fillna(99)
    return df.sort_values('order').drop(columns=['order'])

def format_date(s):
    m = re.findall(r'/(\d{2})/(\d{2})', s)
    return f"{m[0][0]}{m[0][1]}~{m[1][0]}{m[1][1]}" if len(m)>=2 else s

if __name__ == "__main__":
    main()
