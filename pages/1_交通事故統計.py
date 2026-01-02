import pandas as pd
import io
import re
from google.colab import files
from google.colab import auth
import gspread
from google.auth import default
from datetime import datetime

def analyze_traffic_stats_gsheet_smart():
    print("🚀 請上傳三個檔案（本週、今年累計、去年累計），順序與檔名不拘...")
    uploaded = files.upload()
    
    if len(uploaded) < 3:
        print("⚠️ 警告：檔案數量不足 3 個，可能會導致計算錯誤。")

    # --- 1. 定義解析函式 (先讀取所有檔案) ---
    def parse_police_stats_raw(file_obj):
        try:
            df_raw = pd.read_csv(file_obj, header=None)
        except:
            file_obj.seek(0)
            df_raw = pd.read_excel(file_obj, header=None)
        return df_raw

    # --- 2. 智慧辨識檔案身分 ---
    file_data_map = {} # 暫存解析後的資料
    
    print("🔍 正在分析檔案內容以自動分類...")
    
    for filename, content in uploaded.items():
        file_obj = io.BytesIO(content)
        df = parse_police_stats_raw(file_obj)
        
        # 抓取日期字串，例如 "統計日期：114/11/21 至 114/11/27"
        try:
            date_str = df.iloc[1, 0].replace("統計日期：", "").strip()
            # 簡單正則抓取年份與月份
            dates = re.findall(r'(\d{3})/(\d{2})/(\d{2})', date_str)
            # dates 結構: [('114', '11', '21'), ('114', '11', '27')]
            
            if not dates:
                print(f"⚠️ 無法識別日期：{filename}")
                continue
                
            start_y, start_m, start_d = map(int, dates[0])
            end_y, end_m, end_d = map(int, dates[1])
            
            # 判斷邏輯
            # 1. 期間很短 (小於 30 天) -> 本期 (Weekly)
            # 2. 期間很長 + 年份較大 -> 今年累計 (Current)
            # 3. 期間很長 + 年份較小 -> 去年累計 (Last)
            
            # 這裡用簡易判斷：若起始月是 1月 且 結束月大於 1月，通常是累計
            is_cumulative = (start_m == 1 and end_m >= 1)
            
            # 或是計算天數差異 (略過複雜 datetime，直接看月份跨度)
            month_diff = (end_y - start_y) * 12 + (end_m - start_m)
            
            if month_diff == 0 and (end_d - start_d) < 20:
                category = 'weekly'
            else:
                # 累計檔，比較年份
                # 這裡先存起來，等等比較哪一個年份大
                category = f'cumulative_{start_y}'
            
            file_data_map[filename] = {
                'df': df,
                'date_str': date_str,
                'category': category,
                'year': start_y
            }
        except Exception as e:
            print(f"⚠️ 檔案解析失敗 {filename}: {e}")

    # 分配角色
    df_wk = None
    df_cur = None
    df_lst = None
    d_wk = ""
    d_cur = ""
    d_lst = ""

    # 找出 Weekly
    for fname, data in file_data_map.items():
        if data['category'] == 'weekly':
            df_wk = data['df']
            d_wk = data['date_str']
            break
            
    # 找出 Current 和 Last (比較年份)
    cumulative_files = [d for d in file_data_map.values() if 'cumulative' in d['category']]
    if len(cumulative_files) >= 2:
        # 排序：年份大的在前
        cumulative_files.sort(key=lambda x: x['year'], reverse=True)
        
        df_cur = cumulative_files[0]['df']
        d_cur = cumulative_files[0]['date_str']
        
        df_lst = cumulative_files[1]['df']
        d_lst = cumulative_files[1]['date_str']
    
    # 檢查是否都找到了
    if df_wk is None or df_cur is None or df_lst is None:
        print("❌ 自動辨識失敗，請檢查檔案內容日期是否正確。")
        print(f"辨識結果: {[d['category'] for d in file_data_map.values()]}")
        return

    print(f"✅ 成功辨識：\n   本期: {d_wk}\n   今年: {d_cur}\n   去年: {d_lst}")

    # --- 3. 資料清理與計算函式 ---
    def process_data(df_raw):
        # 抓取資料列
        df_data = df_raw[df_raw[0].notna()].copy()
        df_data = df_data[df_data[0].str.contains("總計|派出所")].copy()
        df_data = df_data.reset_index(drop=True)
        
        columns_map = {
            0: "Station",
            1: "Total_Cases", 2: "Total_Deaths", 3: "Total_Injuries",
            4: "A1_Cases", 5: "A1_Deaths", 6: "A1_Injuries",
            7: "A2_Cases", 8: "A2_Deaths", 9: "A2_Injuries",
            10: "A3_Cases"
        }
        df_data = df_data.rename(columns=columns_map)
        
        for c in list(columns_map.values()):
            if c not in df_data.columns: df_data[c] = 0
        df_data = df_data[list(columns_map.values())]
        
        for col in list(columns_map.values())[1:]:
            df_data[col] = pd.to_numeric(df_data[col].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
            
        df_data['Station_Short'] = df_data['Station'].str.replace('派出所', '所').str.replace('總計', '合計')
        
        # 重新計算合計 (確保資料正確)
        df_stations = df_data[~df_data['Station_Short'].str.contains("合計")].copy()
        numeric_cols = df_data.columns[1:-1]
        total_row = df_stations[numeric_cols].sum()
        total_row['Station_Short'] = '合計'
        df_total = pd.DataFrame([total_row])
        
        return pd.concat([df_total, df_stations], ignore_index=True)

    df_wk_clean = process_data(df_wk)
    df_cur_clean = process_data(df_cur)
    df_lst_clean = process_data(df_lst)

    # 4. 準備標題日期
    def format_date(s):
        m = re.findall(r'/(\d{2})/(\d{2})', s)
        return f"{m[0][0]}{m[0][1]}~{m[1][0]}{m[1][1]}" if len(m)>=2 else s

    h_wk = format_date(d_wk)
    h_cur = format_date(d_cur)
    h_lst = format_date(d_lst)

    # 5. 合併與計算 (強制正確對應)
    
    # --- A1 ---
    # 確保欄位名稱唯一
    a1_wk = df_wk_clean[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'wk'})
    a1_cur = df_cur_clean[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'cur'})
    a1_lst = df_lst_clean[['Station_Short', 'A1_Deaths']].rename(columns={'A1_Deaths': 'last'})
    
    m_a1 = pd.merge(a1_wk, a1_cur, on='Station_Short', how='outer')
    m_a1 = pd.merge(m_a1, a1_lst, on='Station_Short', how='outer').fillna(0)
    
    m_a1['Diff'] = m_a1['cur'] - m_a1['last'] # 今年 - 去年

    # --- A2 ---
    a2_wk = df_wk_clean[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'wk'})
    a2_cur = df_cur_clean[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'cur'})
    a2_lst = df_lst_clean[['Station_Short', 'A2_Injuries']].rename(columns={'A2_Injuries': 'last'})
    
    m_a2 = pd.merge(a2_wk, a2_cur, on='Station_Short', how='outer')
    m_a2 = pd.merge(m_a2, a2_lst, on='Station_Short', how='outer').fillna(0)
    
    m_a2['Diff'] = m_a2['cur'] - m_a2['last'] # 今年 - 去年
    m_a2['Pct'] = m_a2.apply(lambda x: (x['Diff']/x['last']) if x['last']!=0 else 0, axis=1)
    m_a2['Pct_Str'] = m_a2['Pct'].apply(lambda x: f"{x:.2%}") # 轉成 % 字串
    m_a2['Prev'] = "-"

    # 6. 排序
    target_order = ['合計', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所']
    order_map = {name: i for i, name in enumerate(target_order)}
    
    m_a1['order'] = m_a1['Station_Short'].map(order_map).fillna(99)
    m_a1 = m_a1.sort_values('order').drop(columns=['order'])
    
    m_a2['order'] = m_a2['Station_Short'].map(order_map).fillna(99)
    m_a2 = m_a2.sort_values('order').drop(columns=['order'])

    # 7. 整理最終表格
    a1_final = m_a1[['Station_Short', 'wk', 'cur', 'last', 'Diff']].copy()
    a1_final.columns = ['單位', f'本期({h_wk})', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較']
    
    a2_final = m_a2[['Station_Short', 'wk', 'Prev', 'cur', 'last', 'Diff', 'Pct_Str']].copy()
    a2_final.columns = ['單位', f'本期({h_wk})', '前期', f'本年累計({h_cur})', f'去年累計({h_lst})', '本年與去年同期比較', '本年較去年增減比例']

    # --- GOOGLE SHEETS 串接 ---
    print("🔐 正在驗證 Google 帳號權限 (請在跳出的視窗點選『允許』)...")
    try:
        auth.authenticate_user()
        creds, _ = default()
        gc = gspread.authorize(creds)
    except Exception as e:
        print(f"❌ 驗證失敗：{e}")
        return

    # 建立新試算表
    sheet_name = f'交通事故統計表_整理結果_{datetime.now().strftime("%Y%m%d_%H%M%S")}'
    try:
        sh = gc.create(sheet_name)
    except Exception as e:
        print(f"❌ 建立試算表失敗：{e}")
        return

    # 寫入 A1 分頁
    try:
        ws1 = sh.sheet1
        ws1.update_title("A1死亡人數")
        ws1.update([a1_final.columns.values.tolist()] + a1_final.fillna(0).values.tolist())
    except Exception as e:
        print(f"⚠️ 寫入 A1 分頁時發生小問題：{e}")

    # 寫入 A2 分頁
    try:
        ws2 = sh.add_worksheet(title="A2受傷人數", rows=20, cols=10)
        ws2.update([a2_final.columns.values.tolist()] + a2_final.fillna(0).values.tolist())
    except Exception as e:
        print(f"⚠️ 寫入 A2 分頁時發生小問題：{e}")

    print("\n" + "="*40)
    print(f"✅ 成功！已自動辨識檔案並產生報表：")
    print(f"🔗 連結：{sh.url}")
    print("="*40)

if __name__ == "__main__":
    analyze_traffic_stats_gsheet_smart()
