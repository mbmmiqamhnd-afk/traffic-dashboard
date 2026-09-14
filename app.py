import io
import re
import time
import traceback
from datetime import datetime, timedelta

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import gspread
import numpy as np
import pandas as pd
import streamlit as st

# ==========================================
# 0. 系統初始化與格式套件
# ==========================================
st.set_page_config(
    page_title="交通執法自動化分析引擎", page_icon="🚓", layout="wide"
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

# ==========================================
# 2. Google Sheets & Drive 連線層
# ==========================================


def _gsheet_call_with_retry(fn, *args, max_retries=4, base_delay=5, **kwargs):
  for attempt in range(max_retries):
    try:
      return fn(*args, **kwargs)
    except gspread.exceptions.APIError as e:
      if "429" in str(e) and attempt < max_retries - 1:
        wait = base_delay * (2**attempt)
        st.warning(
            f"⏳ Google Sheets API 限速 (429)，等待 {wait} 秒後重試... (第"
            f" {attempt+1} 次)"
        )
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
  """限定資料夾範圍查詢並下載 Excel/CSV 檔案"""
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
      f
      for f in items
      if any(f["name"].lower().endswith(ext) for ext in [".xlsx", ".xls", ".csv"])
  ]

  if not valid_items:
    st.warning(
        f"⚠️ 資料夾內找到 {len(items)} 個檔案，但都不是 .xlsx / .xls / .csv 報表。"
    )
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


def _ws_batch_clear(ws, ranges):
  _gsheet_call_with_retry(ws.batch_clear, ranges)


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


def get_ws_by_index(sh, idx):
  cached = getattr(sh, "_cached_worksheets", [])
  if idx < len(cached):
    return cached[idx]
  return sh.get_worksheet(idx)


# --- [業務常數] ---
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

# 三項重點違規（115年9月1日起）
THREE_MAJOR_CATS = ["闖紅燈", "逆向行駛", "不停讓行人"]
THREE_MAJOR_START_ROC = 1150901  # 統計起日

# ==========================================
# 3. 輔助工具區
# ==========================================


def get_gsheet_rich_text_req(sheet_id, row_idx, col_idx, text):
  text = str(text)
  pattern = r"([0-9\(\)\/\-]+)"
  tokens = re.split(pattern, text)
  runs = []
  current_pos = 0
  for token in tokens:
    if not token:
      continue
    color = (
        {"red": 1.0, "green": 0.0, "blue": 0.0}
        if re.match(pattern, token)
        else {"red": 0.0, "green": 0.0, "blue": 0.0}
    )
    runs.append({
        "startIndex": current_pos,
        "format": {"foregroundColor": color, "bold": True},
    })
    current_pos += len(token)
  return {
      "updateCells": {
          "rows": [{
              "values": [{
                  "userEnteredValue": {"stringValue": text},
                  "textFormatRuns": runs,
              }]
          }],
          "fields": "userEnteredValue,textFormatRuns",
          "range": {
              "sheetId": sheet_id,
              "startRowIndex": row_idx,
              "endRowIndex": row_idx + 1,
              "startColumnIndex": col_idx,
              "endColumnIndex": col_idx + 1,
          },
      }
  }


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


# ==========================================
# 4. 業務邏輯處理區
# ==========================================

# 1. 科技執法
# (略，保持原有 process_tech_enforcement 邏輯不變)
# 2. 超載統計
# (略，保持原有 process_overload 邏輯不變)
# 3. 重大交通違規
# (略，保持原有 process_major 邏輯不變)
# 4. 強化專案
# (略，保持原有 process_project 邏輯不變)
# 5. 交通事故
# (略，保持原有 process_accident 邏輯不變)
# 6. 靜桃計畫
# (略，保持原有 process_jing_tao 邏輯不變)

# 7. 取締三項重點違規（闖紅燈、逆向行駛、不停讓行人）每日統計 (115年9月1日起)


def process_three_major_daily(files, sh):
  """統計各單位 115 年 9 月 1 日起，每日新增的件數及累計（闖紅燈、逆向行駛、不停讓行人）"""
  if not files:
    st.warning("⚠️ 未偵測到可供統計三項重點違規之報表檔案。")
    return

  daily_records = []  # 存放結構化記錄: dict(date_code, date_label, unit, cat, count)

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
      y, m, d = (
          int(m_ce.group(1)) - 1911,
          int(m_ce.group(2)),
          int(m_ce.group(3)),
      )
      return y * 10000 + m * 100 + d, f"{m:02d}/{d:02d}"
    return None, None

  for f in files:
    f.seek(0)
    is_csv = f.name.lower().endswith(".csv")
    try:
      if is_csv:
        try:
          df_raw = pd.read_csv(f, header=None)
        except Exception:
          f.seek(0)
          df_raw = pd.read_csv(f, encoding="cp950", header=None)
      else:
        df_raw = pd.read_excel(f, header=None)
    except Exception as e:
      st.warning(f"檔案 {f.name} 讀取失敗: {e}")
      continue

    # --- 判斷模式 A：是否為明細清冊型檔案（含單據/違規/入案明細） ---
    is_detail = False
    detail_header_idx = -1
    for i in range(min(20, len(df_raw))):
      row_strs = [str(x).strip() for x in df_raw.iloc[i].values if pd.notna(x)]
      if any("單位" in x or "所別" in x for x in row_strs) and any(
          "日" in x for x in row_strs
      ):
        if any(
            k in "".join(row_strs) for k in ["條", "法條", "違規", "單號", "事實"]
        ):
          detail_header_idx = i
          is_detail = True
          break

    if is_detail:
      f.seek(0)
      df_det = (
          pd.read_csv(
              f,
              skiprows=detail_header_idx,
              encoding="cp950" if is_csv else None,
          )
          if is_csv
          else pd.read_excel(f, skiprows=detail_header_idx)
      )
      df_det.columns = [str(c).strip() for c in df_det.columns]

      unit_col = next(
          (c for c in df_det.columns if any(k in c for k in ["單位", "所別"])),
          None,
      )
      date_col = next(
          (
              c
              for c in df_det.columns
              if any(k in c for k in ["入案", "違規日", "舉發日", "單據日", "日期"])
          ),
          None,
      )
      law_col = next(
          (
              c
              for c in df_det.columns
              if any(k in c for k in ["條款", "法條", "法規", "條"])
          ),
          None,
      )
      fact_col = next(
          (
              c
              for c in df_det.columns
              if any(k in c for k in ["違規事實", "事實", "違規項目", "取締項目"])
          ),
          None,
      )

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
          elif (
              any(k in str(r.get(law_col, "")) for k in ["45101", "45103", "45條"])
              or any(k in text_check for k in ["逆向", "來車道", "不按遵行"])
          ):
            target_cat = "逆向行駛"
          elif (
              any(
                  k in str(r.get(law_col, ""))
                  for k in ["442", "443", "444", "44條", "482", "48條"]
              )
              or any(k in text_check for k in ["停讓", "行人", "車不讓人"])
          ):
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

    # --- 判斷模式 B：彙總型報表（如重點違規統計表） ---
    header_idx = -1
    for i in range(min(15, len(df_raw))):
      row_strs = [str(x) for x in df_raw.iloc[i].values if pd.notna(x)]
      if any("闖紅燈" in v for v in row_strs) and any("逆向" in v for v in row_strs):
        header_idx = i
        break

    if header_idx != -1:
      # 從表頭前幾列或檔名解析日期
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
        # 單一日期匹配
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
        elif h in [
            "酒駕",
            "嚴重超速",
            "轉彎未依規定",
            "蛇行惡意逼車",
            "本年度",
            "去年度",
        ]:
          curr_c = None

        if curr_c:
          cat_cols[curr_c].append(c_idx)

      for r_idx in range(header_idx + 2, len(df_raw)):
        row = df_raw.iloc[r_idx]
        u = clean_unit_name(row.iloc[0])
        if u and "合計" not in str(row.iloc[0]):
          for cat, cols in cat_cols.items():
            cnt = sum([
                int(pd.to_numeric(row.iloc[c], errors="coerce") or 0)
                for c in cols
                if c < len(row)
            ])
            daily_records.append({
                "date_code": file_d_code,
                "date_label": file_d_label,
                "unit": u,
                "cat": cat,
                "count": cnt,
            })

  if not daily_records:
    st.error(
        "❌ 未能自所選檔案中解析出 115 年 9 月 1 日起之三項重點違規數據！"
    )
    return

  df_all = pd.DataFrame(daily_records)
  # 排序日期
  unique_dates = (
      df_all[["date_code", "date_label"]]
      .drop_duplicates()
      .sort_values("date_code")
  )
  date_cols = unique_dates["date_label"].tolist()
  start_label = date_cols[0]
  end_label = date_cols[-1]

  # 1. 產生三項重點違規總表 (每日各日新增 + 9/1起累計)
  p_tot = df_all.pivot_table(
      index="unit",
      columns="date_label",
      values="count",
      aggfunc="sum",
      fill_value=0,
  )
  p_tot = p_tot.reindex(columns=date_cols, fill_value=0)
  p_tot = p_tot.reindex(MAJOR_UNIT_ORDER, fill_value=0)
  p_tot["累計 (0901起)"] = p_tot.sum(axis=1)

  # 合計列
  sum_row = pd.DataFrame([p_tot.sum(axis=0)], index=["合計"])
  df_final_daily = pd.concat([sum_row, p_tot]).reset_index()
  df_final_daily.rename(columns={"index": "單位"}, inplace=True)

  # 2. 產生三大項細分統計表 (闖紅燈、逆向行駛、不停讓行人)
  cat_tables = {}
  for cat in THREE_MAJOR_CATS:
    df_c = df_all[df_all["cat"] == cat]
    if not df_c.empty:
      p_c = df_c.pivot_table(
          index="unit",
          columns="date_label",
          values="count",
          aggfunc="sum",
          fill_value=0,
      )
      p_c = p_c.reindex(columns=date_cols, fill_value=0)
      p_c = p_c.reindex(MAJOR_UNIT_ORDER, fill_value=0)
      p_c["累計"] = p_c.sum(axis=1)
      s_r = pd.DataFrame([p_c.sum(axis=0)], index=["合計"])
      p_final = pd.concat([s_r, p_c]).reset_index()
      p_final.rename(columns={"index": "單位"}, inplace=True)
      cat_tables[cat] = p_final

  # 3. 彙總一覽表（最新一日 vs 9/1起累計）
  latest_day = end_label
  summary_rows = []
  for u in ["合計"] + MAJOR_UNIT_ORDER:
    r_u = df_final_daily[df_final_daily["單位"] == u]
    r_red = cat_tables.get("闖紅燈", pd.DataFrame())
    r_rev = cat_tables.get("逆向行駛", pd.DataFrame())
    r_ped = cat_tables.get("不停讓行人", pd.DataFrame())

    def get_cnt(df_t, d_col):
      if df_t.empty or d_col not in df_t.columns:
        return 0
      sub = df_t[df_t["單位"] == u]
      return int(sub[d_col].values[0]) if not sub.empty else 0

    summary_rows.append({
        "單位": u,
        f"闖紅燈({latest_day})": get_cnt(r_red, latest_day),
        "闖紅燈(累計)": get_cnt(r_red, "累計"),
        f"逆向行駛({latest_day})": get_cnt(r_rev, latest_day),
        "逆向行駛(累計)": get_cnt(r_rev, "累計"),
        f"不停讓行人({latest_day})": get_cnt(r_ped, latest_day),
        "不停讓行人(累計)": get_cnt(r_ped, "累計"),
        f"三項合計({latest_day})": get_cnt(df_final_daily, latest_day),
        "三項合計(累計)": get_cnt(df_final_daily, "累計 (0901起)"),
    })
  df_summary_overview = pd.DataFrame(summary_rows)

  # --- 畫面呈現 ---
  st.subheader(
      "🚦 取締三項重點違規（闖紅燈、逆向行駛、不停讓行人）每日件數及累計統計表"
  )
  st.caption(
      f"📅 統計起日：115 年 9 月 1 日 ｜ 涵蓋統計期間：{start_label} 至"
      f" {end_label} ｜ 最新統計日：{latest_day}"
  )

  # 關鍵指標卡片
  c_m1, c_m2, c_m3, c_m4 = st.columns(4)
  c_m1.metric(
      "🎯 9/1起三項累計總數", f"{df_summary_overview.iloc[0]['三項合計(累計)']} 件"
  )
  c_m2.metric(
      f"📅 最新單日新增 ({latest_day})",
      f"{df_summary_overview.iloc[0][f'三項合計({latest_day})']} 件",
  )
  c_m3.metric(
      "闖紅燈 / 逆向累計",
      f"{df_summary_overview.iloc[0]['闖紅燈(累計)']} /"
      f" {df_summary_overview.iloc[0]['逆向行駛(累計)']} 件",
  )
  c_m4.metric(
      "不停讓行人累計", f"{df_summary_overview.iloc[0]['不停讓行人(累計)']} 件"
  )

  st.write("📊 **【總表】各單位 115 年 9 月 1 日起每日新增及累計：**")
  st.dataframe(df_final_daily, hide_index=True, use_container_width=True)

  with st.expander("🔍 檢視三項違規最新當日總覽與個別項目細表（點擊展開）"):
    st.write("📋 **三項違規最新當日與 9/1 起累計總覽表：**")
    st.dataframe(df_summary_overview, hide_index=True, use_container_width=True)
    for cat_name, df_cat in cat_tables.items():
      st.write(f"**【{cat_name}】各單位每日新增及累計：**")
      st.dataframe(df_cat, hide_index=True, use_container_width=True)

  # --- 同步 Google Sheets ---
  if sh:
    try:
      ws_name = "三項重點違規-每日績效"
      ws = get_or_create_ws(sh, ws_name, rows=50, cols=30)
      _ws_clear(ws)

      title_str = (
          "桃園市政府警察局龍潭分局 取締三項重點違規（闖紅燈、逆向行駛、不停讓行人）每日績效統計表"
          f" (115年9月1日至{latest_day})"
      )
      grid_data = (
          [[title_str] + [""] * (len(df_final_daily.columns) - 1)]
          + [df_final_daily.columns.tolist()]
          + df_final_daily.values.tolist()
      )

      _ws_update(ws, "A1", grid_data)

      reqs = [
          # 標題列合併與格式化
          {
              "mergeCells": {
                  "range": {
                      "sheetId": ws.id,
                      "startRowIndex": 0,
                      "endRowIndex": 1,
                      "startColumnIndex": 0,
                      "endColumnIndex": len(df_final_daily.columns),
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
                              "fontSize": 18,
                              "bold": True,
                              "foregroundColor": {
                                  "red": 0.0,
                                  "green": 0.0,
                                  "blue": 0.8,
                              },
                          },
                      }
                  },
                  "fields": (
                      "userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment,userEnteredFormat.textFormat"
                  ),
              }
          },
          # 表頭格式化 (A2)
          {
              "repeatCell": {
                  "range": {
                      "sheetId": ws.id,
                      "startRowIndex": 1,
                      "endRowIndex": 2,
                      "startColumnIndex": 0,
                      "endColumnIndex": len(df_final_daily.columns),
                  },
                  "cell": {
                      "userEnteredFormat": {
                          "horizontalAlignment": "CENTER",
                          "verticalAlignment": "MIDDLE",
                          "textFormat": {
                              "fontFamily": "DFKai-SB",
                              "fontSize": 12,
                              "bold": True,
                          },
                      }
                  },
                  "fields": (
                      "userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment,userEnteredFormat.textFormat"
                  ),
              }
          },
          # 合計列加粗 (A3)
          {
              "repeatCell": {
                  "range": {
                      "sheetId": ws.id,
                      "startRowIndex": 2,
                      "endRowIndex": 3,
                      "startColumnIndex": 0,
                      "endColumnIndex": len(df_final_daily.columns),
                  },
                  "cell": {
                      "userEnteredFormat": {
                          "textFormat": {
                              "fontFamily": "DFKai-SB",
                              "fontSize": 12,
                              "bold": True,
                              "foregroundColor": {
                                  "red": 0.8,
                                  "green": 0.0,
                                  "blue": 0.0,
                              },
                          }
                      }
                  },
                  "fields": "userEnteredFormat.textFormat",
              }
          },
      ]
      _sh_batch_update(sh, {"requests": reqs})
      st.write("✅ 三項重點違規每日績效已成功同步至 Google 試算表！")
    except Exception as e:
      st.error(f"雲端同步出錯：{e}")


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
          st.success(
              f"✅ 成功自共用資料夾讀取到 {len(drive_files)} 個試算表檔案！"
          )
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
      elif any(
          k in name
          for k in ["三項", "闖紅燈", "逆向", "行人", "禮讓", "停讓", "每日績效"]
      ):
        cat_files["三項重點"].append(f)
      elif any(k in name for k in ["stone", "超載"]):
        cat_files["超載統計"].append(f)
      elif any(k in name for k in ["重大", "重點"]):
        cat_files["重大違規"].append(f)
        cat_files["三項重點"].append(f)  # 重點違規統計報表亦自動支援三項重點每日萃取
      elif any(
          k in name
          for k in ["強化", "專案", "砂石", "大貨", "r17", "法條", "自選匯出"]
      ):
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
          time.sleep(1.5)

      if cat_files["超載統計"]:
        with st.status("🚛 處理【超載統計】...", expanded=True):
          process_overload(cat_files["超載統計"], sh)
          time.sleep(1.5)

      if cat_files["重大違規"]:
        with st.status("🚨 處理【重大交通違規】...", expanded=True):
          process_major(cat_files["重大違規"], sh)
          time.sleep(1.5)

      # 執行三項重點違規每日績效統計
      if cat_files["三項重點"]:
        with st.status(
            "🚦 處理【三項重點違規（闖紅燈、逆向、不停讓行人）每日統計】...",
            expanded=True,
        ):
          process_three_major_daily(cat_files["三項重點"], sh)
          time.sleep(1.5)

      if cat_files["強化專案"]:
        with st.status("🔥 處理【強化專案】...", expanded=True):
          process_project(cat_files["強化專案"], sh)
          time.sleep(1.5)

      if cat_files["交通事故"]:
        with st.status("🚑 處理【交通事故】...", expanded=True):
          process_accident(cat_files["交通事故"], sh)
          time.sleep(1.5)

      if cat_files["靜桃計畫"]:
        with st.status("🤫 處理【靜桃計畫】...", expanded=True):
          process_jing_tao(cat_files["靜桃計畫"], sh)
          time.sleep(1.5)

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
          f"### 📑 接下來請執行以下步驟：\n\n👉 **[點此直接開啟"
          f" {rec_name}]({rec_url})**\n\n1. 點擊簡報畫面右上方的"
          " **「全部更新」**（載入最新數據）。\n2. 點擊上方選單 **【📂 會議歸檔工具】>【🚀"
          " 建立當次會議副本並存檔】**，按一下 Enter"
          " 即可自動完成副本歸檔並寄發郵件通知！"
      )

    except Exception as e:
      st.error(f"⚠️ 批次處理發生錯誤：{e}")
      st.write(traceback.format_exc())
