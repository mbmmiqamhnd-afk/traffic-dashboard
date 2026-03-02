import streamlit as st
import pandas as pd
import gspread
import io
from datetime import datetime

# ==========================================
# 0. 基本設定
# ==========================================
st.set_page_config(page_title="龍潭分局交通戰情室", page_icon="🚓", layout="wide")

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"

# 單位與目標值設定 (依據上傳格式)
# 類別順序：酒駕, 闖紅燈, 嚴重超速, 車不讓人, 行人違規, 大型車違規
TARGET_CONFIG = {
    '聖亭所': [5, 115, 5, 16, 7, 10],
    '龍潭所': [6, 145, 7, 20, 9, 12],
    '中興所': [5, 115, 5, 16, 7, 10],
    '石門所': [3, 80, 4, 11, 5, 7],
    '高平所': [3, 80, 4, 11, 5, 7],
    '三和所': [2, 40, 2, 6, 2, 5],
    '交通分隊': [5, 115, 4, 16, 6, 8]
}

CATS = ["酒後駕車", "闖紅燈", "嚴重超速", "車不讓人", "行人違規", "大型車違規"]

# 法條歸類定義
LAW_MAP = {
    "酒後駕車": ["35條", "73條2項", "73條3項"],
    "闖紅燈": ["53條"],
    "嚴重超速": ["43條"],
    "車不讓人": ["44條", "48條"],
    "行人違規": ["78條"],
    "大型車違規": ["29條", "30條", "18之1條"]
}

# ==========================================
# 1. 主畫面
# ==========================================
st.title("🚓 龍潭分局交通數據戰情室")
st.header("🚦 (一) 加強交通安全執法取締五項交通違規統計表")
st.info("請在此執行原本的五項違規分析流程...")
# [保留您原本的功能代碼...]

st.markdown("<br><hr style='border:2px solid #ddd'><br>", unsafe_allow_html=True)

# ==========================================
# 2. 強化專案功能 (垂直整合在下方)
# ==========================================
st.header(f"📈 (二) {PROJECT_NAME}")

uploaded_file = st.file_uploader("上傳『法條件數統計報表』(CSV/XLSX)", type=["csv", "xlsx"], key="proj_sync")

if uploaded_file:
    # 讀取數據 (自動跳過前3行)
    df_raw = pd.read_csv(uploaded_file, skiprows=3) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, skiprows=3)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    
    # 資料處理
    results = []
    for unit, targets in TARGET_CONFIG.items():
        # 尋找報表中對應的單位行 (模糊匹配)
        row = df_raw[df_raw['單位'].str.contains(unit.replace('所',''), na=False)]
        
        unit_data = [unit]
        for i, cat in enumerate(CATS):
            # 根據 LAW_MAP 加總該類別的所有法條件數
            keywords = LAW_MAP[cat]
            matched_cols = [c for c in df_raw.columns if any(k in c for k in keywords)]
            count = row[matched_cols].sum(axis=1).values[0] if not row.empty else 0
            
            target = targets[i]
            ratio = f"{(count/target*100):.0f}%" if target > 0 else "0%"
            unit_data.extend([int(count), target, ratio])
        results.append(unit_data)

    # 建立 DataFrame
    cols = ["單位"]
    for cat in CATS:
        cols.extend([f"{cat}_取締", f"{cat}_目標", f"{cat}_達成率"])
    
    df_final = pd.DataFrame(results, columns=cols)
    
    # 計算合計列
    totals = ["合計"]
    for i in range(1, len(cols), 3):
        cnt_sum = df_final.iloc[:, i].sum()
        tgt_sum = df_final.iloc[:, i+1].sum()
        ratio_sum = f"{(cnt_sum/tgt_sum*100):.0f}%" if tgt_sum > 0 else "0%"
        totals.extend([int(cnt_sum), int(tgt_sum), ratio_sum])
    
    df_final = pd.concat([pd.DataFrame([totals], columns=cols), df_final]).reset_index(drop=True)

    # 預覽表格
    st.subheader("📋 專案執法達成率預覽")
    st.dataframe(df_final, use_container_width=True)

    # 同步按鈕
    if st.button("🚀 同步至雲端試算表 (格式化)", use_container_width=True):
        try:
            gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
            sh = gc.open_by_url(GOOGLE_SHEET_URL)
            try:
                ws = sh.worksheet(PROJECT_NAME)
            except:
                ws = sh.add_worksheet(title=PROJECT_NAME, rows=50, cols=20)
            
            # 依照您要求的檔案格式進行寫入
            now = datetime.now().strftime("%Y-%m-%d")
            header1 = [f"{PROJECT_NAME} ({now})"] + [""] * 18
            header2 = [""] + [c for c in CATS for _ in range(3)]
            header3 = ["單位"] + ["取締件數", "目標值", "達成率"] * 6
            data = df_final.values.tolist()
            
            ws.clear()
            ws.update(values=[header1, header2, header3] + data)
            
            # 合併單元格與格式化
            requests = [
                {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 19}, "mergeType": "MERGE_ALL"}},
            ]
            for j in range(1, 19, 3):
                requests.append({"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": j, "endColumnIndex": j+3}, "mergeType": "MERGE_ALL"}})
            
            sh.batch_update({"requests": requests})
            st.success(f"✅ 已成功同步至雲端，分頁：{PROJECT_NAME}")
            st.balloons()
        except Exception as e:
            st.error(f"同步失敗：{e}")
