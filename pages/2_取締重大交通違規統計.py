import streamlit as st
import pandas as pd
import re
import io
import gspread
from gspread_formatting import *

# ==========================================
# 0. 設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
UNIT_ORDER = ['科技執法', '聖亭所', '龍潭所', '中興所', '石門所', '高平所', '三和所', '警備隊', '交通分隊']
TARGETS = {'聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479, '高平所': 1294, '三和所': 339, '交通分隊': 2526, '警備隊': 0, '科技執法': 6006}
FOOTNOTE_TEXT = "重大交通違規指：「酒駕」、「闖紅燈」、「嚴重超速」、「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」及「不暫停讓行人」"

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

# --- 2. 雲端同步功能 (從 A2 開始，不更改 A1 格式) ---
def sync_to_specified_sheet(df):
    try:
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        ws = sh.get_worksheet(0)
        
        # 1. 準備資料列 (不包含總標題，僅包含兩層 Header 與數據)
        col_tuples = df.columns.tolist()
        top_row = [t[0] for t in col_tuples]
        bottom_row = [t[1] for t in col_tuples]
        data_list = [top_row, bottom_row] + df.values.tolist()
        
        # 2. 【核心修改】從 A2 開始寫入，不使用 ws.clear()
        ws.update(range_name='A2', values=data_list)
        
        # 計算相關 Index
        # 資料從第 2 列開始，所以 Row Index 起點為 1
        header_top_idx = 1
        header_bottom_idx = 2
        data_start_idx = 3
        footnote_row_idx = len(data_list) + 1 # 加上原本跳過的第 1 列
        
        red_color = {"red": 1.0, "green": 0.0, "blue": 0.0}
        black_color = {"red": 0.0, "green": 0.0, "blue": 0.0}
        
        # 3. 格式化請求 (僅針對第 2 列以後)
        requests = [
            # 統計期間標題合併 (Row Index 1 & 2)
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 0, "endColumnIndex": 1}, "mergeType": "MERGE_ALL"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 1, "endColumnIndex": 3}, "mergeType": "MERGE_ALL"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 3, "endColumnIndex": 5}, "mergeType": "MERGE_ALL"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 5, "endColumnIndex": 7}, "mergeType": "MERGE_ALL"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 7, "endColumnIndex": 8}, "mergeType": "MERGE_ALL"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 8, "endColumnIndex": 9}, "mergeType": "MERGE_ALL"}},
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 3, "startColumnIndex": 9, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}},
            
            # 備註列合併
            {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": footnote_row_idx - 1, "endRowIndex": footnote_row_idx, "startColumnIndex": 0, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}},
        ]
        
        # 4. 標題雙色邏輯 (套用在 Row Index 1)
        for i, text in enumerate(top_row):
            if "(" in text:
                paren_start = text.find("(")
                requests.append({
                    "updateCells": {
                        "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": i, "endColumnIndex": i+1},
                        "rows": [{
                            "values": [{
                                "textFormatRuns": [
                                    {"startIndex": 0, "format": {"foregroundColor": black_color}},
                                    {"startIndex": paren_start, "format": {"foregroundColor": red_color}}
                                ],
                                "userEnteredValue": {"stringValue": text}
                            }]
                        }],
                        "fields": "userEnteredValue,textFormatRuns"
                    }
                })

        # 5. 負值紅字規則 (資料區從 Row Index 3 開始)
        requests.extend([
            {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": ws.id, "startRowIndex": 3, "endRowIndex": footnote_row_idx - 1, "startColumnIndex": 7, "endColumnIndex": 8}],
                        "booleanRule": {
                            "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                            "format": {"textFormat": {"foregroundColor": red_color}}
                        }
                    }, "index": 0
                }
            }
        ])
        
        sh.batch_update({"requests": requests})
        return True
    except Exception as e:
        st.error(f"雲端同步失敗: {e}")
        return False

# --- 4. 解析與主介面邏輯 (略，與前版相同) ---
# ... (parse_excel_with_date_extraction 函數與 Streamlit UI 代碼)
