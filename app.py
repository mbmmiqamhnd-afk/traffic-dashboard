import streamlit as st
import pandas as pd
import io
import re
import gspread
import shutil
import smtplib
import calendar
from datetime import datetime, timedelta, date
from pdf2image import convert_from_bytes
from pptx import Presentation
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header

# ==========================================
# 0. 系統初始化與全局設定
# ==========================================
st.set_page_config(page_title="龍潭分局交通智慧戰情室", page_icon="🚓", layout="wide")

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
TO_EMAIL = "mbmmiqamhnd@gmail.com"

# 目標值設定區
TARGETS_MAJOR = {'科技執法': 6006, '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479, '高平所': 1294, '三和所': 339, '交通分隊': 2526}
TARGETS_OVERLOAD = {'聖亭所': 20, '龍潭所': 27, '中興所': 20, '石門所': 16, '高平所': 14, '三和所': 8, '警備隊': 0, '交通分隊': 22}

# ==========================================
# 🛠️ 通用工具箱 (Helper Functions)
# ==========================================

def get_std_unit(n):
    n = str(n).strip()
    if '分隊' in n: return '交通分隊'
    if '科技' in n or '交通組' in n: return '科技執法'
    if '警備' in n: return '警備隊'
    for k in ['聖亭', '龍潭', '中興', '石門', '高平', '三和']:
        if k in n: return k + '所'
    return None

def sync_gsheet_batch(ws, title, df_data, font_size=16):
    """通用雲端同步：含藍紅標題格式化"""
    ws.clear()
    data = [ [title] ] + [df_data.columns.tolist()] + df_data.values.tolist()
    ws.update(values=data)
    
    blue = {"red": 0, "green": 0, "blue": 1}
    red = {"red": 1, "green": 0, "blue": 0}
    split_idx = title.find("(") if "(" in title else len(title)
    
    reqs = [{
        "updateCells": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1},
            "rows": [{"values": [{"userEnteredValue": {"stringValue": title},
                "textFormatRuns": [
                    {"startIndex": 0, "format": {"foregroundColor": blue, "bold": True, "fontSize": font_size}},
                    {"startIndex": split_idx, "format": {"foregroundColor": red, "bold": True, "fontSize": font_size}}
                ]}]}],
            "fields": "userEnteredValue,textFormatRuns"
        }
    }]
    ws.spreadsheet.batch_update({"requests": reqs})

# ==========================================
# 🏰 導覽選單
# ==========================================
with st.sidebar:
    st.title("🚓 龍潭分局戰情室")
    app_mode = st.selectbox("功能模組", ["🏠 智慧上傳中心", "📂 PDF 轉 PPTX 工具"])
    st.divider()
    st.info("💡 10秒流程：首頁直接拖入報表即可分析。")

# ==========================================
# 🏠 核心：智慧上傳中心
# ==========================================
if app_mode == "🏠 智慧上傳中心":
    st.header("📈 交通數據智慧分析中心")
    uploads = st.file_uploader("📂 拖入隨身碟中的報表檔案", type=["xlsx", "csv", "xls"], accept_multiple_files=True)

    if uploads:
        num = len(uploads)
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)

        # --- [1份檔案]：科技執法 ---
        if num == 1:
            f = uploads[0]
            if "list" in f.name.lower() or "地點" in f.name:
                st.success(f"📸 識別為「科技執法」：{f.name}")
                # 執行科技執法 logic ... (略)
                # 提示：此處可呼叫 sync_gsheet_batch(sh.get_worksheet(4), title, df, font_size=24)

        # --- [2份檔案]：重大交通違規 ---
        elif num == 2:
            st.success("✅ 識別為「重大交通違規」統計")
            # 執行重大違規 logic ... (略)
            # 提示：此處同步至 sh.get_worksheet(0)

        # --- [3份檔案]：強化專案 或 超載統計 ---
        elif num == 3:
            if any("stone" in f.name.lower() for f in uploads):
                st.success("🚛 識別為「超載違規」自動統計")
                # 執行超載統計 logic (含寄信) ... (略)
                # 提示：同步至 sh.get_worksheet(1)
            else:
                st.success("🔥 識別為「強化交通安全專案」統計")
                # 執行強化專案 logic ... (略)
                # 提示：同步至 sh.get_worksheet(5)

        # --- [4份檔案]：交通事故 A1/A2 ---
        elif num == 4:
            st.success("🚑 識別為「交通事故 A1/A2」統計")
            # 執行交通事故 logic ... (略)
            # 提示：同步至 sh.get_worksheet(2) & (3)

        else:
            st.warning(f"目前收到 {num} 份檔案，請確認數量。")

# ==========================================
# 📂 模式二：PDF 轉 PPTX
# ==========================================
elif app_mode == "📂 PDF 轉 PPTX 工具":
    st.header("📂 PDF 行政文書轉檔")
    # ... 原本轉檔代碼 ...
