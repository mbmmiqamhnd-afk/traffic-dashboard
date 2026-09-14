import io
import os
import re
import time
import traceback
from datetime import datetime, timedelta

import gspread
import numpy as np
import pandas as pd
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ==========================================
# 0. 系統初始化與格式套件
# ==========================================
st.set_page_config(
    page_title="交通執法自動化分析引擎",
    page_icon="🚓",
    layout="wide",
)

try:
    from gspread_formatting import *
    HAS_FORMATTING = True
except ImportError:
    HAS_FORMATTING = False

# ==========================================
# 1. 全局常數與設定區
# ==========================================
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

try:
    GCP_CREDS = dict(st.secrets.get("gcp_service_account", {}))
except Exception:
    GCP_CREDS = None

# 業務常數設定
MAJOR_UNIT_ORDER = [
    "科技執法",
    "聖亭所",
    "龍潭所",
    "中興所",
    "石門所",
    "高平所",
    "三和所",
    "警備隊",
    "交通分隊",
]

MAJOR_TARGETS = {
    "聖亭所": 1941,
    "龍潭所": 2588,
    "中興所": 1941,
    "石門所": 1479,
    "高平所": 1294,
    "三和所": 339,
    "交通分隊": 2526,
    "警備隊": 0,
    "科技執法": 6006,
}

MAJOR_FOOTNOTE = (
    "重大交通違規指：「酒駕」、「闖紅燈」、「嚴重超速」、「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」及「不暫停讓行人」"
)

OVERLOAD_TARGETS = {
    "聖亭所": 20,
    "龍潭所": 27,
    "中興所": 20,
    "石門所": 16,
    "高平所": 14,
    "三和所": 8,
    "警備隊": 0,
    "交通分隊": 22,
}

OVERLOAD_UNIT_MAP = {
    "聖亭派出所": "聖亭所",
    "龍潭派出所": "龍潭所",
    "中興派出所": "中興所",
    "石門派出所": "石門所",
    "高平派出所": "高平所",
    "三和派出所": "三和所",
    "警備隊": "警備隊",
    "龍潭交通分隊": "交通分隊",
}

OVERLOAD_UNIT_ORDER = [
    "聖亭所",
    "龍潭所",
    "中興所",
    "石門所",
    "高平所",
    "三和所",
    "警備隊",
    "交通分隊",
]

PROJECT_NAME = "強化交通安全執法專案勤務取締件數統計表"
PROJECT_TARGETS = {
    "聖亭所": [5, 115, 5, 16, 7, 10],
    "龍潭所": [6, 145, 7, 20, 9, 12],
    "中興所": [5, 115, 5, 16, 7, 10],
    "石門所": [3, 80, 4, 11, 5, 7],
    "高平所": [3, 80, 4, 11, 5, 7],
    "三和所": [2, 40, 2, 6, 2, 5],
    "交通分隊": [5, 115, 4, 16, 6, 8],
    "交通組": [0, 0, 0, 0, 0, 0],
    "警備隊": [0, 0, 0, 0, 0, 0],
}

PROJECT_CATS = [
    "酒後駕車",
    "闖紅燈",
    "嚴重超速",
    "車不讓人",
    "行人違規",
    "大型車違規",
]

PROJECT_LAW_MAP = {
    "酒後駕車": ["35條", "73條2項", "73條3項"],
    "闖紅燈": ["53條"],
    "嚴重超速": ["43條", "40條"],
    "車不讓人": ["44條", "48條"],
    "行人違規": ["78條"],
}

THREE_MAJOR_CATS = ["闖紅燈", "逆向行駛", "不停讓行人"]
THREE_MAJOR_START_ROC = 1150901  # 統計起日

# ==========================================
# 2. Google Sheets & Drive 連線層
# ==========================================
def _gsheet_call_with_retry(fn, *args, max_retries=4, base_delay=5, **kwargs):
    for attempt in range(max_retries):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            if "429" in str(e) and attempt < max_retries - 1:
                wait = base_delay * (2 ** attempt)
                st.warning(f"⏳ Google Sheets API 限速 (429)，等待 {wait} 秒後重試... (第 {attempt+1} 次)")
                time.sleep(wait)
            else:
                raise

@st.cache_resource
def get_gsheet_connection():
    if GCP_CREDS:
        try:
            gc = gspread.service_account_from_dict(GCP_CREDS)
            sh = gc.open_by_url(GOOGLE_SHEET_URL)
            sh._cached_worksheets = _gsheet_call_with_retry(sh.worksheets)
            return sh
        except Exception as e:
            st.error(f"⚠️ Google Sheets 連線失敗: {e}")
    return None

class DriveVirtualFile(io.BytesIO):
    """具備檔案 ID 與虛擬檔案特性之記憶體物件"""
    def __init__(self, name, content_bytes, file_id=None, parents=None):
        super().__init__(content_bytes)
        self.name = name
        self.size = len(content_bytes)
        self.id = file_id
        self.parents = parents or []

@st.cache_resource
def get_drive_service():
    if not GCP_CREDS:
        return None
    creds = service_account.Credentials.from_service_account_info(
        GCP_CREDS, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)

def fetch_files_from_drive(folder_id):
    service = get_drive_service()
    if not service:
        st.error("❌ 無法初始化 Drive 服務，請確認 secrets.toml 設定")
        return []

    folder_id = str(folder_id).strip().replace('"', "").replace("'", "")
    try:
        results = (
            service.files()
            .list(
                q=f"'{folder_id}' in parents and trashed = false",
                fields="files(id, name, size, mimeType, parents)",
                pageSize=100,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )
        items = results.get("files", [])
    except Exception as e:
        if "404" in str(e) or "File not found" in str(e):
            st.error(
                f"❌ 服務帳號對資料夾 ID `{folder_id}` 無存取權。\n"
                "請確認該資料夾已將服務帳號加為協作者（編輯者）。"
            )
        else:
            st.error(f"❌ 查詢雲端硬碟檔案失敗：{e}")
        return []

    if items:
        with st.expander("🔎 此資料夾內服務帳號可辨識的檔案清單（點擊展開）"):
            for f in items:
                st.write(f"- **{f['name']}** (`{f.get('mimeType')}`)")
    else:
        st.error(f"❌ 服務帳號在資料夾 ID `{folder_id}` 底下找不到任何檔案！")
        return []

    valid_items = [
        f for f in items
        if any(f["name"].lower().endswith(ext) for ext in [".xlsx", ".xls", ".csv"])
    ]

    if not valid_items:
        st.warning(f"⚠️ 資料夾內找到 {len(items)} 個檔案，但都不是 .xlsx / .xls / .csv 報表。")
        return []

    st.caption(f"🔍 成功篩選出 {len(valid_items)} 個有效報表！")
    downloaded_files = []
    for item in valid_items:
        try:
            req = service.files().get_media(fileId=item["id"], supportsAllDrives=True)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = downloader.next_chunk()

            vfile = DriveVirtualFile(
                item["name"],
                fh.getvalue(),
                file_id=item["id"],
                parents=item.get("parents", []),
            )
            downloaded_files.append(vfile)
        except Exception as e:
            st.warning(f"檔案 {item['name']} 下載失敗: {e}")

    return downloaded_files

def _ws_update(ws, range_name, values):
    _gsheet_call_with_retry(ws.update, range_name=range_name, values=values)

def _ws_clear(ws):
    _gsheet_call_with_retry(ws.clear)

def _sh_batch_update(sh, body):
    _gsheet_call_with_retry(sh.batch_update, body)

def get_or_create_ws(sh, ws_name, rows=100, cols=20):
    cached = getattr(sh, "_cached_worksheets", [])
    ws = next((s for s in cached if s.title == ws_name), None)
    if not ws:
        ws = _gsheet_call_with_retry(
            sh.add_worksheet, title=ws_name, rows=str(rows), cols=str(cols)
        )
        sh._cached_worksheets.append(ws)
    return ws

def ensure_ws_capacity(ws, min_rows, min_cols):
    """確保工作表擁有足夠行列，避免超出邊界引發 APIError: 400"""
    try:
        cur_r = ws.row_count
        cur_c = ws.col_count
        target_r = max(cur_r, min_rows)
        target_c = max(cur_c, min_cols)
        if target_r > cur_r or target_c > cur_c:
            _gsheet_call_with_retry(ws.resize, rows=target_r, cols=target_c)
    except Exception:
        pass

# ==========================================
# 3. 輔助工具區
# ==========================================
def clean_unit_name(raw):
    if pd.isna(raw):
        return None
    n = str(raw).strip()
    if "分隊" in n:
        return "交通分隊"
    if any(k in n for k in ["科技", "交通組"]):
        return "科技執法"
    if "警備" in n:
        return "警備隊"
    for k in ["聖亭", "龍潭", "中興", "石門", "高平", "三和"]:
        if k in n:
            return k + "所"
    return None

def read_tabular_file(f):
    """通用讀取 Excel 或 CSV 檔案轉為 DataFrame"""
    f.seek(0)
    name = f.name.lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(f, header=None)
        except Exception:
            f.seek(0)
            return pd.read_csv(f, encoding="cp950", header=None)
    else:
        return pd.read_excel(f, header=None)

# ==========================================
# 4. 業務邏輯處理區
# ==========================================

# 1. 科技執法
def process_tech_enforcement(files, sh):
    st.markdown("### 📸 科技執法取締成效統計")
    all_dfs = []
    for f in files:
        df_raw = read_tabular_file(f)
        h_idx = -1
        for i in range(min(15, len(df_raw))):
            row_vals = [str(x).strip() for x in df_raw.iloc[i].values if pd.notna(x)]
            if any("地點" in x or "路段" in x or "取締項目" in x for x in row_vals):
                h_idx = i
                break
        f.seek(0)
        if h_idx != -1:
            df = pd.read_csv(f, skiprows=h_idx) if f.name.lower().endswith(".csv") else pd.read_excel(f, skiprows=h_idx)
            df.columns = [str(c).strip() for c in df.columns]
            all_dfs.append(df)
        else:
            all_dfs.append(df_raw)

    if not all_dfs:
        st.warning("⚠️ 科技執法無有效數據。")
        return

    df_combined = pd.concat(all_dfs, ignore_index=True)
    st.dataframe(df_combined.head(30), use_container_width=True)

    if sh:
        try:
            ws = get_or_create_ws(sh, "科技執法", rows=len(df_combined) + 10, cols=len(df_combined.columns) + 5)
            ensure_ws_capacity(ws, len(df_combined) + 10, len(df_combined.columns) + 5)
            _ws_clear(ws)
            out_grid = [df_combined.columns.tolist()] + df_combined.fillna("").astype(str).values.tolist()
            _ws_update(ws, "A1", out_grid)
            st.success("✅ 科技執法數據已同步至 Google Sheets！")
        except Exception as e:
            st.error(f"科技執法同步失敗: {e}")

# 2. 超載統計
def process_overload(files, sh):
    st.markdown("### 🚛 取締大貨車及砂石車超載統計")
    counts = {u: 0 for u in OVERLOAD_UNIT_ORDER}

    for f in files:
        df_raw = read_tabular_file(f)
        for r_idx in range(len(df_raw)):
            row = df_raw.iloc[r_idx]
            first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            matched_u = OVERLOAD_UNIT_MAP.get(first_val, clean_unit_name(first_val))
            if matched_u and matched_u in counts:
                nums = [pd.to_numeric(x, errors="coerce") for x in row.values if pd.notna(x)]
                valid_nums = [int(n) for n in nums if pd.notna(n) and n >= 0]
                if valid_nums:
                    counts[matched_u] += valid_nums[-1]

    rows = []
    for u in OVERLOAD_UNIT_ORDER:
        act = counts.get(u, 0)
        tgt = OVERLOAD_TARGETS.get(u, 0)
        rate = f"{(act / tgt * 100):.1f}%" if tgt > 0 else "-"
        diff = act - tgt if tgt > 0 else "-"
        rows.append({"單位": u, "目標件數": tgt, "取締件數": act, "達成率": rate, "增減件數": diff})

    df_res = pd.DataFrame(rows)
    tot_tgt = sum(OVERLOAD_TARGETS.values())
    tot_act = sum(counts.values())
    tot_rate = f"{(tot_act / tot_tgt * 100):.1f}%" if tot_tgt > 0 else "-"
    tot_row = pd.DataFrame([{"單位": "合計", "目標件數": tot_tgt, "取締件數": tot_act, "達成率": tot_rate, "增減件數": tot_act - tot_tgt}])
    df_final = pd.concat([tot_row, df_res], ignore_index=True)

    st.dataframe(df_final, hide_index=True, use_container_width=True)

    if sh:
        try:
            ws = get_or_create_ws(sh, "取締超載統計", rows=30, cols=10)
            ensure_ws_capacity(ws, len(df_final) + 5, len(df_final.columns) + 2)
            _ws_clear(ws)
            grid = [df_final.columns.tolist()] + df_final.values.tolist()
            _ws_update(ws, "A1", grid)
            st.success("✅ 超載統計數據已同步至 Google Sheets！")
        except Exception as e:
            st.error(f"超載統計同步失敗: {e}")

# 3. 重大交通違規
def process_major(files, sh):
    st.markdown("### 🚨 重大交通違規 取締成效統計")
    st.caption(f"📌 {MAJOR_FOOTNOTE}")
    counts = {u: 0 for u in MAJOR_UNIT_ORDER}

    for f in files:
        df_raw = read_tabular_file(f)
        h_idx = -1
        for i in range(min(20, len(df_raw))):
            row_strs = [str(x).strip() for x in df_raw.iloc[i].values if pd.notna(x)]
            if any("單位" in x or "所別" in x for x in row_strs):
                h_idx = i
                break

        if h_idx != -1:
            f.seek(0)
            df = pd.read_csv(f, skiprows=h_idx, encoding="cp950" if f.name.lower().endswith(".csv") else None) if f.name.lower().endswith(".csv") else pd.read_excel(f, skiprows=h_idx)
            df.columns = [str(c).strip() for c in df.columns]
            unit_col = next((c for c in df.columns if any(k in c for k in ["單位", "所別"])), None)
            total_col = next((c for c in df.columns if any(k in c for k in ["合計", "總計", "取締數", "件數"])), None)
            is_detail = any(k in "".join(df.columns) for k in ["條款", "法條", "違規事實", "單號"])

            if is_detail and unit_col:
                for _, r in df.iterrows():
                    u = clean_unit_name(r[unit_col])
                    if u in counts:
                        counts[u] += 1
            elif unit_col and total_col:
                for _, r in df.iterrows():
                    u = clean_unit_name(r[unit_col])
                    if u in counts:
                        val = pd.to_numeric(r[total_col], errors="coerce")
                        counts[u] += int(val) if pd.notna(val) else 0
        else:
            for r_idx in range(len(df_raw)):
                row = df_raw.iloc[r_idx]
                u = clean_unit_name(row.iloc[0])
                if u in counts:
                    nums = [pd.to_numeric(x, errors="coerce") for x in row.values if pd.notna(x)]
                    valid_nums = [int(n) for n in nums if pd.notna(n) and n > 0]
                    if valid_nums:
                        counts[u] += valid_nums[-1]

    data_list = []
    for u in MAJOR_UNIT_ORDER:
        actual = counts.get(u, 0)
        target = MAJOR_TARGETS.get(u, 0)
        rate = f"{(actual / target * 100):.1f}%" if target > 0 else "-"
        diff = actual - target if target > 0 else "-"
        data_list.append({"單位": u, "目標件數": target, "取締件數": actual, "達成率": rate, "增減件數": diff})

    df_major = pd.DataFrame(data_list)
    tot_target = sum(MAJOR_TARGETS.values())
    tot_actual = sum(counts.values())
    tot_rate = f"{(tot_actual / tot_target * 100):.1f}%" if tot_target > 0 else "-"
    tot_row = pd.DataFrame([{"單位": "合計", "目標件數": tot_target, "取締件數": tot_actual, "達成率": tot_rate, "增減件數": tot_actual - tot_target}])
    df_result = pd.concat([tot_row, df_major], ignore_index=True)

    st.dataframe(df_result, hide_index=True, use_container_width=True)

    if sh:
        try:
            ws = get_or_create_ws(sh, "重大交通違規", rows=30, cols=10)
            ensure_ws_capacity(ws, len(df_result) + 5, len(df_result.columns) + 2)
            _ws_clear(ws)
            grid = [df_result.columns.tolist()] + df_result.values.tolist()
            _ws_update(ws, "A1", grid)
            st.success("✅ 重大交通違規數據已成功同步至 Google 試算表！")
        except Exception as e:
            st.error(f"重大交通違規同步 Google Sheets 出錯：{e}")

# 4. 強化專案
def process_project(files, sh):
    st.markdown("### 🔥 強化交通安全執法專案勤務取締件數統計")
    units = list(PROJECT_TARGETS.keys())
    res_matrix = {u: {cat: 0 for cat in PROJECT_CATS} for u in units}

    for f in files:
        df_raw = read_tabular_file(f)
        h_idx = -1
        for i in range(min(20, len(df_raw))):
            row_strs = [str(x).strip() for x in df_raw.iloc[i].values if pd.notna(x)]
            if any("單位" in x or "所別" in x for x in row_strs):
                h_idx = i
                break
        f.seek(0)
        df = pd.read_csv(f, skiprows=h_idx, encoding="cp950" if f.name.lower().endswith(".csv") else None) if (h_idx != -1 and f.name.lower().endswith(".csv")) else (pd.read_excel(f, skiprows=h_idx) if h_idx != -1 else df_raw)
        df.columns = [str(c).strip() for c in df.columns]

        unit_col = next((c for c in df.columns if any(k in c for k in ["單位", "所別"])), None)
        law_col = next((c for c in df.columns if any(k in c for k in ["條款", "法條", "法規"])), None)
        fact_col = next((c for c in df.columns if any(k in c for k in ["違規事實", "事實", "違規項目"])), None)

        if unit_col:
            for _, r in df.iterrows():
                u = clean_unit_name(r[unit_col])
                if u not in res_matrix:
                    continue
                txt = f"{r.get(law_col, '')} {r.get(fact_col, '')}"
                if any(k in txt for k in ["35條", "酒駕", "酒後"]):
                    res_matrix[u]["酒後駕車"] += 1
                elif any(k in txt for k in ["53條", "闖紅燈"]):
                    res_matrix[u]["闖紅燈"] += 1
                elif any(k in txt for k in ["43條", "嚴重超速"]):
                    res_matrix[u]["嚴重超速"] += 1
                elif any(k in txt for k in ["44條", "48條", "車不讓人", "停讓行人"]):
                    res_matrix[u]["車不讓人"] += 1
                elif any(k in txt for k in ["78條", "行人違規", "行人穿越"]):
                    res_matrix[u]["行人違規"] += 1
                elif any(k in txt for k in ["大型車", "大貨車", "聯結車", "砂石車"]):
                    res_matrix[u]["大型車違規"] += 1

    rows = []
    for u in units:
        r_dict = {"單位": u}
        tot_u = 0
        for cat in PROJECT_CATS:
            c = res_matrix[u][cat]
            r_dict[cat] = c
            tot_u += c
        r_dict["合計"] = tot_u
        rows.append(r_dict)

    df_proj = pd.DataFrame(rows)
    st.dataframe(df_proj, hide_index=True, use_container_width=True)

    if sh:
        try:
            ws = get_or_create_ws(sh, "強化專案統計", rows=len(df_proj) + 5, cols=len(df_proj.columns) + 2)
            ensure_ws_capacity(ws, len(df_proj) + 5, len(df_proj.columns) + 2)
            _ws_clear(ws)
            grid = [df_proj.columns.tolist()] + df_proj.values.tolist()
            _ws_update(ws, "A1", grid)
            st.success("✅ 強化專案統計已同步至 Google Sheets！")
        except Exception as e:
            st.error(f"強化專案同步失敗: {e}")

# 5. 交通事故
def process_accident(files, sh):
    st.markdown("### 🚑 交通事故 (A1/A2/A3) 分析統計")
    all_dfs = []
    for f in files:
        df_raw = read_tabular_file(f)
        all_dfs.append(df_raw)
    if all_dfs:
        df_merged = pd.concat(all_dfs, ignore_index=True)
        st.dataframe(df_merged.head(25), use_container_width=True)
        if sh:
            try:
                ws = get_or_create_ws(sh, "交通事故分析", rows=len(df_merged) + 5, cols=len(df_merged.columns) + 2)
                ensure_ws_capacity(ws, len(df_merged) + 5, len(df_merged.columns) + 2)
                _ws_clear(ws)
                grid = [df_merged.columns.tolist()] + df_merged.fillna("").astype(str).values.tolist()
                _ws_update(ws, "A1", grid)
                st.success("✅ 交通事故統計已同步至 Google Sheets！")
            except Exception as e:
                st.error(f"交通事故同步失敗: {e}")

# 6. 靜桃計畫
def process_jing_tao(files, sh):
    st.markdown("### 🤫 靜桃計畫（改裝噪音車輛取締）統計")
    all_dfs = []
    for f in files:
        df_raw = read_tabular_file(f)
        all_dfs.append(df_raw)
    if all_dfs:
        df_merged = pd.concat(all_dfs, ignore_index=True)
        st.dataframe(df_merged.head(25), use_container_width=True)
        if sh:
            try:
                ws = get_or_create_ws(sh, "靜桃計畫統計", rows=len(df_merged) + 5, cols=len(df_merged.columns) + 2)
                ensure_ws_capacity(ws, len(df_merged) + 5, len(df_merged.columns) + 2)
                _ws_clear(ws)
                grid = [df_merged.columns.tolist()] + df_merged.fillna("").astype(str).values.tolist()
                _ws_update(ws, "A1", grid)
                st.success("✅ 靜桃計畫數據已同步至 Google Sheets！")
            except Exception as e:
                st.error(f"靜桃計畫同步失敗: {e}")

# 7. 取締三項重點違規（各項最新單日新增 vs 115年9月1日起累計）
def process_three_major_daily(files, sh):
    """
    統計各單位 115 年 9 月 1 日起之三項重點違規（闖紅燈、逆向行駛、不停讓行人）：
    - 各單項之「最後一日新增件數」
    - 各單項自「115年9月1日起累計件數」
    - 三項總計之最後一日新增與累計
    """
    if not files:
        st.warning("⚠️ 未偵測到可供統計三項重點違規之報表檔案。")
        return

    daily_records = []

    def parse_date_code_label(s):
        if not s:
            return None, None
        s = str(s).strip()
        m_roc = re.search(r"(1\d{2})[./\-_]?(\d{2})[./\-_]?(\d{2})", s)
        if m_roc:
            y, m, d = int(m_roc.group(1)), int(m_roc.group(2)), int(m_roc.group(3))
            return y * 10000 + m * 100 + d, f"{m:02d}/{d:02d}"
        m_ce = re.search(r"(20\d{2})[./\-_]?(\d{2})[./\-_]?(\d{2})", s)
        if m_ce:
            y, m, d = int(m_ce.group(1)) - 1911, int(m_ce.group(2)), int(m_ce.group(3))
            return y * 10000 + m * 100 + d, f"{m:02d}/{d:02d}"
        return None, None

    for f in files:
        df_raw = read_tabular_file(f)

        # 模式 A：明細清冊型
        is_detail = False
        detail_header_idx = -1
        for i in range(min(20, len(df_raw))):
            row_strs = [str(x).strip() for x in df_raw.iloc[i].values if pd.notna(x)]
            if any("單位" in x or "所別" in x for x in row_strs) and any("日" in x for x in row_strs):
                if any(k in "".join(row_strs) for k in ["條", "法條", "違規", "單號", "事實"]):
                    detail_header_idx = i
                    is_detail = True
                    break

        if is_detail:
            f.seek(0)
            is_csv = f.name.lower().endswith(".csv")
            df_det = (
                pd.read_csv(f, skiprows=detail_header_idx, encoding="cp950" if is_csv else None)
                if is_csv
                else pd.read_excel(f, skiprows=detail_header_idx)
            )
            df_det.columns = [str(c).strip() for c in df_det.columns]

            unit_col = next((c for c in df_det.columns if any(k in c for k in ["單位", "所別"])), None)
            date_col = next((c for c in df_det.columns if any(k in c for k in ["入案", "違規日", "舉發日", "單據日", "日期"])), None)
            law_col = next((c for c in df_det.columns if any(k in c for k in ["條款", "法條", "法規", "條"])), None)
            fact_col = next((c for c in df_det.columns if any(k in c for k in ["違規事實", "事實", "違規項目", "取締項目"])), None)

            if unit_col and date_col:
                for _, r in df_det.iterrows():
                    u = clean_unit_name(r[unit_col])
                    if not u:
                        continue
                    d_code, d_label = parse_date_code_label(r[date_col])
                    if not d_code or d_code < THREE_MAJOR_START_ROC:
                        continue

                    text_check = f"{r.get(law_col, '')} {r.get(fact_col, '')}"
                    target_cat = None
                    if "53" in str(r.get(law_col, "")) or "闖紅燈" in text_check:
                        target_cat = "闖紅燈"
                    elif any(k in str(r.get(law_col, "")) for k in ["45101", "45103", "45條"]) or any(k in text_check for k in ["逆向", "來車道", "不按遵行"]):
                        target_cat = "逆向行駛"
                    elif any(k in str(r.get(law_col, "")) for k in ["442", "443", "444", "44條", "482", "48條"]) or any(k in text_check for k in ["停讓", "行人", "車不讓人"]):
                        target_cat = "不停讓行人"

                    if target_cat:
                        daily_records.append({
                            "date_code": d_code,
                            "date_label": d_label,
                            "unit": u,
                            "cat": target_cat,
                            "count": 1,
                        })
            continue

        # 模式 B：彙總型報表
        header_idx = -1
        for i in range(min(15, len(df_raw))):
            row_strs = [str(x) for x in df_raw.iloc[i].values if pd.notna(x)]
            if any("闖紅燈" in v for v in row_strs) and any("逆向" in v for v in row_strs):
                header_idx = i
                break

        if header_idx != -1:
            text_top = df_raw.iloc[:header_idx, :5].to_string() + " " + f.name
            m_range = re.search(r"115(\d{4})\s*[至\-\~]\s*115(\d{4})", text_top)
            file_d_code, file_d_label = None, None

            if m_range:
                s_d, e_d = m_range.group(1), m_range.group(2)
                end_code = 1150000 + int(e_d)
                if end_code >= THREE_MAJOR_START_ROC:
                    file_d_code = end_code
                    file_d_label = f"{e_d[:2]}/{e_d[2:]}"
            else:
                d_c, d_l = parse_date_code_label(text_top)
                if d_c and d_c >= THREE_MAJOR_START_ROC:
                    file_d_code, file_d_label = d_c, d_l

            if not file_d_code:
                continue

            headers = df_raw.iloc[header_idx].values
            cat_cols = {"闖紅燈": [], "逆向行駛": [], "不停讓行人": []}
            curr_c = None
            for c_idx in range(len(headers)):
                h = str(headers[c_idx]).strip() if pd.notna(headers[c_idx]) else ""
                if "闖紅燈" in h:
                    curr_c = "闖紅燈"
                elif "逆向" in h:
                    curr_c = "逆向行駛"
                elif any(k in h for k in ["不暫停讓行人", "不停讓行人", "車不讓人"]):
                    curr_c = "不停讓行人"
                elif h in ["酒駕", "嚴重超速", "轉彎未依規定", "蛇行惡意逼車", "本年度", "去年度"]:
                    curr_c = None

                if curr_c:
                    cat_cols[curr_c].append(c_idx)

            for r_idx in range(header_idx + 1, len(df_raw)):
                row = df_raw.iloc[r_idx]
                u = clean_unit_name(row.iloc[0])
                if u and "合計" not in str(row.iloc[0]):
                    for cat, cols in cat_cols.items():
                        cnt = sum([int(pd.to_numeric(row.iloc[c], errors="coerce") or 0) for c in cols if c < len(row)])
                        daily_records.append({
                            "date_code": file_d_code,
                            "date_label": file_d_label,
                            "unit": u,
                            "cat": cat,
                            "count": cnt,
                        })

    if not daily_records:
        st.error("❌ 未能自所選檔案中解析出 115 年 9 月 1 日起之三項重點違規數據！")
        return

    df_all = pd.DataFrame(daily_records)
    unique_dates = df_all[["date_code", "date_label"]].drop_duplicates().sort_values("date_code")
    date_cols = unique_dates["date_label"].tolist()
    start_label = date_cols[0]
    latest_day = date_cols[-1]

    # 計算各項目細部統計
    cat_tables = {}
    for cat in THREE_MAJOR_CATS:
        df_c = df_all[df_all["cat"] == cat]
        if not df_c.empty:
            p_c = df_c.pivot_table(index="unit", columns="date_label", values="count", aggfunc="sum", fill_value=0)
            p_c = p_c.reindex(columns=date_cols, fill_value=0).astype(int)
            p_c = p_c.reindex(MAJOR_UNIT_ORDER, fill_value=0)
            p_c["累計"] = p_c.sum(axis=1)
            cat_tables[cat] = p_c
        else:
            cat_tables[cat] = pd.DataFrame(0, index=MAJOR_UNIT_ORDER, columns=date_cols + ["累計"])

    # 組合各項「最後一日新增」與「115年9月1日起累計」統計總表
    overview_rows = []
    for u in MAJOR_UNIT_ORDER:
        red_latest = int(cat_tables["闖紅燈"].loc[u, latest_day]) if latest_day in cat_tables["闖紅燈"].columns else 0
        red_total = int(cat_tables["闖紅燈"].loc[u, "累計"])

        rev_latest = int(cat_tables["逆向行駛"].loc[u, latest_day]) if latest_day in cat_tables["逆向行駛"].columns else 0
        rev_total = int(cat_tables["逆向行駛"].loc[u, "累計"])

        ped_latest = int(cat_tables["不停讓行人"].loc[u, latest_day]) if latest_day in cat_tables["不停讓行人"].columns else 0
        ped_total = int(cat_tables["不停讓行人"].loc[u, "累計"])

        overview_rows.append({
            "單位": u,
            f"闖紅燈({latest_day})": red_latest,
            "闖紅燈(9/1起累計)": red_total,
            f"逆向行駛({latest_day})": rev_latest,
            "逆向行駛(9/1起累計)": rev_total,
            f"不停讓行人({latest_day})": ped_latest,
            "不停讓行人(9/1起累計)": ped_total,
            f"三項合計({latest_day})": red_latest + rev_latest + ped_latest,
            "三項合計(9/1起累計)": red_total + rev_total + ped_total,
        })

    df_summary = pd.DataFrame(overview_rows)

    # 計算合計列
    sum_vals = {"單位": "合計"}
    for col in df_summary.columns:
        if col != "單位":
            sum_vals[col] = df_summary[col].sum()

    df_summary_final = pd.concat([pd.DataFrame([sum_vals]), df_summary], ignore_index=True)

    # 畫面展示
    st.subheader("🚦 取締三項重點違規（闖紅燈、逆向行駛、不停讓行人）統計表")
    st.caption(f"📅 統計起日：115 年 9 月 1 日 ｜ 涵蓋統計期間：{start_label} 至 {latest_day} ｜ 最後統計日：{latest_day}")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🎯 三項違規累計總數", f"{df_summary_final.iloc[0]['三項合計(9/1起累計)']} 件")
    c2.metric(f"📅 最後一日新增 ({latest_day})", f"{df_summary_final.iloc[0][f'三項合計({latest_day})']} 件")
    c3.metric("闖紅燈 (單日 / 累計)", f"{df_summary_final.iloc[0][f'闖紅燈({latest_day})']} / {df_summary_final.iloc[0]['闖紅燈(9/1起累計)']} 件")
    c4.metric("不停讓行人 (單日 / 累計)", f"{df_summary_final.iloc[0][f'不停讓行人({latest_day})']} / {df_summary_final.iloc[0]['不停讓行人(9/1起累計)']} 件")

    st.write("📊 **各單位三項重點違規最後一日新增與 115 年 9 月 1 日起累計：**")
    st.dataframe(df_summary_final, hide_index=True, use_container_width=True)

    with st.expander("🔍 檢視個別項目各日明細表（點擊展開）"):
        for cat_name, df_cat in cat_tables.items():
            st.write(f"**【{cat_name}】各單位每日件數表：**")
            st.dataframe(df_cat.reset_index().rename(columns={"index": "單位"}), hide_index=True, use_container_width=True)

    # 同步 Google Sheets
    if sh:
        try:
            ws_name = "三項重點違規-每日績效"
            ws = get_or_create_ws(sh, ws_name, rows=40, cols=20)
            ensure_ws_capacity(ws, len(df_summary_final) + 8, len(df_summary_final.columns) + 2)
            _ws_clear(ws)

            sheet_title = f"桃園市政府警察局龍潭分局 取締三項重點違規各項最後一日({latest_day})及累計(115年9月1日起)統計表"
            out_grid = (
                [[sheet_title] + [""] * (len(df_summary_final.columns) - 1)]
                + [df_summary_final.columns.tolist()]
                + df_summary_final.values.tolist()
            )

            _ws_update(ws, "A1", out_grid)

            reqs = [
                {
                    "mergeCells": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": len(df_summary_final.columns),
                        },
                        "mergeType": "MERGE_ALL",
                    }
                },
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "horizontalAlignment": "CENTER",
                                "verticalAlignment": "MIDDLE",
                                "textFormat": {
                                    "fontFamily": "DFKai-SB",
                                    "fontSize": 15,
                                    "bold": True,
                                    "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.8},
                                },
                            }
                        },
                        "fields": "userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment,userEnteredFormat.textFormat",
                    }
                },
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 1,
                            "endRowIndex": 2,
                            "startColumnIndex": 0,
                            "endColumnIndex": len(df_summary_final.columns),
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "horizontalAlignment": "CENTER",
                                "verticalAlignment": "MIDDLE",
                                "textFormat": {
                                    "fontFamily": "DFKai-SB",
                                    "fontSize": 11,
                                    "bold": True,
                                },
                            }
                        },
                        "fields": "userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment,userEnteredFormat.textFormat",
                    }
                },
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 2,
                            "endRowIndex": 3,
                            "startColumnIndex": 0,
                            "endColumnIndex": len(df_summary_final.columns),
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {
                                    "fontFamily": "DFKai-SB",
                                    "fontSize": 11,
                                    "bold": True,
                                    "foregroundColor": {"red": 0.8, "green": 0.0, "blue": 0.0},
                                }
                            }
                        },
                        "fields": "userEnteredFormat.textFormat",
                    }
                },
            ]
            _sh_batch_update(sh, {"requests": reqs})
            st.success("✅ 三項重點違規（單日新增與累計對照表）已成功同步至 Google 試算表！")
        except Exception as e:
            st.error(f"同步 Google Sheets 出錯：{e}")

# ==========================================
# 5. 首頁與雙軌輸入控制中心
# ==========================================
try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    pass

st.header("📈 交通執法數據全自動批次處理中心")

source_mode = st.radio(
    "請選擇報表資料來源：",
    options=["💻 本機檔案拖曳上傳", "☁️ 從 Google 雲端硬碟讀取"],
    horizontal=True,
)

uploads = []

if source_mode == "💻 本機檔案拖曳上傳":
    st.info("💡 請將所需報表全選後，直接拖曳至下方區域即可自動分流處理。")
    uploads = st.file_uploader(
        "📂 拖入所有報表檔案",
        type=["xlsx", "csv", "xls"],
        accept_multiple_files=True,
        key="local_batch_uploader",
    )

elif source_mode == "☁️ 從 Google 雲端硬碟讀取":
    folder_id = st.secrets.get("DRIVE_FOLDER_ID", "").strip()
    if not folder_id:
        st.warning("⚠️ 尚未在 secrets.toml 設定 `DRIVE_FOLDER_ID`。")
    else:
        col_a, col_b = st.columns([3, 1])
        with col_a:
            st.info(f"📂 目前連線之共用資料夾 ID：`{folder_id}`")
        with col_b:
            st.button("🔄 重新整理雲端檔案")

        with st.spinner("正在連線 Google 雲端硬碟讀取報表檔案清單..."):
            try:
                drive_files = fetch_files_from_drive(folder_id)
                if drive_files:
                    st.success(f"✅ 成功自共用資料夾讀取到 {len(drive_files)} 個試算表檔案！")
                    file_map = {f.name: f for f in drive_files}
                    selected_names = st.multiselect(
                        "請確認欲參與批次分析的雲端報表（預設已全選）：",
                        options=list(file_map.keys()),
                        default=list(file_map.keys()),
                        key="drive_batch_multiselect",
                    )
                    uploads = [file_map[name] for name in selected_names]
            except Exception as e:
                st.error(f"❌ 讀取雲端硬碟失敗：{e}")

st.divider()
st.subheader("🚀 啟動全自動批次作業")

# ==========================================
# 6. 自動分流、執行與引導更新簡報
# ==========================================
if uploads:
    file_hash = sum([f.size for f in uploads]) + len(uploads)
    force_rerun = False
    if st.session_state.get("last_processed_hash") == file_hash:
        st.success("✅ 目前載入的檔案皆已全自動處理完畢！")
        st.info("💡 若要重新執行，請切換資料來源、放入新檔案，或勾選下方選項強制重跑。")
        force_rerun = st.checkbox(
            "🔁 強制重新執行（沿用相同檔案重新統計）", key="force_rerun_checkbox"
        )

    if st.session_state.get("last_processed_hash") != file_hash or force_rerun:
        cat_files = {
            "科技執法": [],
            "重大違規": [],
            "超載統計": [],
            "強化專案": [],
            "交通事故": [],
            "靜桃計畫": [],
            "三項重點": [],
        }

        for f in uploads:
            name = f.name.lower()
            if any(k in name for k in ["list", "地點", "科技"]):
                cat_files["科技執法"].append(f)
            elif any(k in name for k in ["三項", "闖紅燈", "逆向", "行人", "禮讓", "停讓", "每日績效"]):
                cat_files["三項重點"].append(f)
            elif any(k in name for k in ["stone", "超載"]):
                cat_files["超載統計"].append(f)
            elif any(k in name for k in ["重大", "重點"]):
                cat_files["重大違規"].append(f)
                cat_files["三項重點"].append(f)
            elif any(k in name for k in ["強化", "專案", "砂石", "大貨", "r17", "法條", "自選匯出"]):
                cat_files["強化專案"].append(f)
            elif any(k in name for k in ["a1", "a2", "事故", "案件統計"]):
                cat_files["交通事故"].append(f)
            elif any(k in name for k in ["靜桃", "噪音", "改裝車", "總表", "詳細資料"]):
                cat_files["靜桃計畫"].append(f)

        try:
            sh = get_gsheet_connection()

            if cat_files["科技執法"]:
                with st.status("📸 處理【科技執法】...", expanded=True):
                    process_tech_enforcement(cat_files["科技執法"], sh)
                    time.sleep(1.0)

            if cat_files["超載統計"]:
                with st.status("🚛 處理【超載統計】...", expanded=True):
                    process_overload(cat_files["超載統計"], sh)
                    time.sleep(1.0)

            if cat_files["重大違規"]:
                with st.status("🚨 處理【重大交通違規】...", expanded=True):
                    process_major(cat_files["重大違規"], sh)
                    time.sleep(1.0)

            if cat_files["三項重點"]:
                with st.status("🚦 處理【三項重點違規（闖紅燈、逆向、不停讓行人）每日與累計統計】...", expanded=True):
                    process_three_major_daily(cat_files["三項重點"], sh)
                    time.sleep(1.0)

            if cat_files["強化專案"]:
                with st.status("🔥 處理【強化專案】...", expanded=True):
                    process_project(cat_files["強化專案"], sh)
                    time.sleep(1.0)

            if cat_files["交通事故"]:
                with st.status("🚑 處理【交通事故】...", expanded=True):
                    process_accident(cat_files["交通事故"], sh)
                    time.sleep(1.0)

            if cat_files["靜桃計畫"]:
                with st.status("🤫 處理【靜桃計畫】...", expanded=True):
                    process_jing_tao(cat_files["靜桃計畫"], sh)
                    time.sleep(1.0)

            st.session_state["last_processed_hash"] = file_hash
            st.balloons()
            st.success("🎉 全自動批次數據分析與 Google 試算表同步完成！")

            weekday = datetime.now().weekday()
            is_mon = weekday in [4, 5, 6, 0]
            rec_url = (
                "https://docs.google.com/presentation/d/1YPVp-PFiQhaJrkaMfBmLQ60ErqrQdOU4BLQMp_pDsXA/edit"
                if is_mon
                else "https://docs.google.com/presentation/d/1l3_HtTKHO5uHof1eBCsm_a_orjHIGrJrtY_E5hql5d4/edit"
            )
            rec_name = "週一主管會報簡報母本" if is_mon else "週四主管會報簡報母本"

            st.markdown(
                f"### 📑 接下來請執行以下步驟：\n\n👉 **[點此直接開啟 {rec_name}]({rec_url})**\n\n"
                "1. 點擊簡報畫面右上方的 **「全部更新」**（載入最新數據）。\n"
                "2. 點擊上方選單 **【📂 會議歸檔工具】>【🚀 建立當次會議副本並存檔】**，按一下 Enter 即可自動完成副本歸檔並寄發郵件通知！"
            )

        except Exception as e:
            st.error(f"⚠️ 批次處理發生錯誤：{e}")
            st.write(traceback.format_exc())
