import streamlit as st
import pandas as pd
import re
import io
from datetime import date

# --- 1. 定義識別與目標 ---
# 加強單位識別，防止「龍潭所」與「交通分隊」混淆
def get_standard_unit(raw_name):
    name = str(raw_name).strip()
    if '分隊' in name: return '交通分隊'
    if '科技' in name or '交通組' in name: return '科技執法'
    if '警備' in name: return '警備隊'
    if '聖亭' in name: return '聖亭所'
    if '龍潭' in name: return '龍潭所'
    if '中興' in name: return '中興所'
    if '石門' in name: return '石門所'
    if '高平' in name: return '高平所'
    if '三和' in name: return '三和所'
    return None

UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
TARGETS = {
    '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479,
    '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006
}

# --- 2. 核心解析函數 ---
def parse_report_precision(uploaded_file, sheet_keyword, target_col_idx):
    """
    target_col_idx: 
    - 抓本期/本年總計就傳 17 (R欄)
    - 抓去年同期總計就傳 20 (U欄)
    """
    try:
        content = uploaded_file.getvalue()
        xl = pd.ExcelFile(io.BytesIO(content))
        
        # 自動尋找包含關鍵字的工作表
        target_sheet = xl.sheet_names[0]
        for s in xl.sheet_names:
            if sheet_keyword in s:
                target_sheet = s
                break
        
        df = pd.read_excel(xl, sheet_name=target_sheet, header=None)
        
        # 提取日期以利後續自動辨識檔案
        info_text = "".join(df.iloc[:5].astype(str).values.flatten())
        match = re.search(r'(\d{3,7}).*至\s*(\d{3,7})', info_text)
        start_date = match.group(1) if match else "0000000"
        end_date = match.group(2) if match else "0000000"
        
        unit_data = {}
        # 遍歷每一列，動態捕捉單位
        for _, row in df.iterrows():
            unit_name = get_standard_unit(row.iloc[0])
            if unit_name and "合計" not in str(row.iloc[0]):
                # 清洗數值
                val_raw = str(row.iloc[target_col_idx]).replace(',', '').strip()
                val = float(val_raw) if val_raw not in ['', 'nan', 'None', '-'] else 0.0
                unit_data[unit_name] = unit_data.get(unit_name, 0) + val
                
        return {'data': unit_data, 'start': start_date, 'end': end_date}
    except Exception as e:
        st.error(f"解析失敗: {e}")
        return None

# --- 3. 主程式介面 ---
st.markdown("## 🚔 交通違規統計 (來源檔案精準對齊版)")

files = st.file_uploader("📂 請上傳 2 個檔案 (本期報表 & 累計報表)", accept_multiple_files=True)

if files and len(files) == 2:
    # A. 識別哪個是累計檔（天數較長者）
    meta = []
    for f in files:
        # 暫以 R 欄 (17) 讀取來測天數
        res = parse_report_precision(f, "重點違規統計表", 17)
        if res:
            try:
                s, e = res['start'], res['end']
                d = (date(int(e[:3])+1911, int(e[3:5]), int(e[5:])) - 
                     date(int(s[:3])+1911, int(s[3:5]), int(s[5:]))).days
                res['duration'] = d
            except: res['duration'] = 0
            res['file_obj'] = f
            meta.append(res)
            
    if len(meta) == 2:
        meta.sort(key=lambda x: x['duration'], reverse=True)
        f_long, f_short = meta[0]['file_obj'], meta[1]['file_obj']
        
        # B. 依照需求抓取指定欄位
        # 1. 本期：短檔 -> R欄 (Index 17)
        d_week = parse_report_precision(f_short, "重點違規統計表", 17)['data']
        # 2. 本年：長檔 -> 工作表(1) -> R欄 (Index 17)
        d_year = parse_report_precision(f_long, "(1)", 17)['data']
        # 3. 去年：長檔 -> 工作表(1) -> U欄 (Index 20)
        d_last = parse_report_precision(f_long, "(1)", 20)['data']
        
        # C. 組合報表
        rows = []
        for u in UNIT_ORDER:
            w, y, l = d_week.get(u,0), d_year.get(u,0), d_last.get(u,0)
            tgt = TARGETS.get(u, 0)
            diff = y - l
            rate = f"{(y/tgt):.1%}" if tgt > 0 else "0%"
            rows.append([u, int(w), int(y), int(l), int(diff), tgt, rate])
            
        df_res = pd.DataFrame(rows, columns=['單位', '本期數值(P-R)', '本年累計(P-R)', '去年同期(S-U)', '增減比較', '目標值', '達成率'])
        st.success("✅ 報表統計完成！")
        st.dataframe(df_res, use_container_width=True)
