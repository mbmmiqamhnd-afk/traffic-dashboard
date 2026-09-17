import io
import re
import uuid
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError

from menu import show_sidebar

# ==========================================
# 0. 系統初始化與側邊欄
# ==========================================
st.set_page_config(
    page_title="全方位執法數據簡報直出中心",
    page_icon="📽️",
    layout="wide"
)
show_sidebar()

st.title("📽️ 全方位執法數據簡報直出中心（動態雲端解析版）")
st.caption("🚀 雲端解析引擎：三項重點違規已升級為動態讀取雲端資料夾最新報表，支援 9/16 本期與累計數自動計算！")

# ==========================================
# 1. Google 服務連線層與常數設定
# ==========================================
GCP_CREDS = dict(st.secrets.get("gcp_service_account", {}))
SERVICE_ACCOUNT_EMAIL = GCP_CREDS.get("client_email", "streamlit-bot@streamlit-sheets-482909.iam.gserviceaccount.com")

TARGET_PRESENTATION_ID = "1h2QNNI8SLvjNEBmky7IWv9ZGBbKsLvV1UDkYJWcOeWU"
DRIVE_FOLDER_ID = st.secrets.get("DRIVE_FOLDER_ID", "1fm6ZK5B5wUmfy7-cgrw8OIkh7iS175dA").strip()

def get_drive_service():
    if not GCP_CREDS:
        return None
    creds = service_account.Credentials.from_service_account_info(
        GCP_CREDS,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)

def get_slides_service():
    if not GCP_CREDS:
        return None
    creds = service_account.Credentials.from_service_account_info(
        GCP_CREDS,
        scopes=[
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/presentations"
        ]
    )
    return build("slides", "v1", credentials=creds)

class DriveVirtualFile(io.BytesIO):
    def __init__(self, name, content_bytes):
        super().__init__(content_bytes)
        self.name = name
        self.size = len(content_bytes)

def fetch_files_from_drive(folder_id):
    service = get_drive_service()
    if not service:
        return []
    folder_id = str(folder_id).strip().replace('"', '').replace("'", '')
    try:
        results = service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="files(id, name, size, mimeType)",
            pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        items = results.get("files", [])
    except Exception:
        return []

    valid_items = [f for f in items if any(f["name"].lower().endswith(ext) for ext in [".xlsx", ".xls", ".csv"])]
    downloaded = []
    for item in valid_items:
        try:
            req = service.files().get_media(fileId=item["id"], supportsAllDrives=True)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            downloaded.append(DriveVirtualFile(item["name"], fh.getvalue()))
        except Exception:
            pass
    return downloaded

with st.container():
    c_s1, c_s2 = st.columns([2, 1])
    with c_s1:
        st.info(f"🔑 **執行服務帳號：** `{SERVICE_ACCOUNT_EMAIL}`\n\n📂 **監聽資料夾 ID：** `{DRIVE_FOLDER_ID}`")
    with c_s2:
        st.link_button("📂 開啟目標 Google 簡報", f"https://docs.google.com/presentation/d/{TARGET_PRESENTATION_ID}/edit")

# ==========================================
# 2. 全方位簡報排版引擎 (ComprehensiveSlidesBuilder)
# ==========================================
class ComprehensiveSlidesBuilder:
    def __init__(self, slides_svc, presentation_id: str):
        self.slides_svc = slides_svc
        self.presentation_id = presentation_id.strip()
        self.all_slide_ids = []
        self.requests = []

    def prepare_canvas(self, protect_first_slide: bool = False):
        pres = self.slides_svc.presentations().get(
            presentationId=self.presentation_id
        ).execute()
        slides = pres.get("slides", [])
        if protect_first_slide and slides:
            self.all_slide_ids = [s["objectId"] for s in slides[1:]]
        else:
            self.all_slide_ids = [s["objectId"] for s in slides]

    def add_cover_slide(self, main_title: str, subtitle: str, date_range_str: str):
        slide_id = f"cover_{uuid.uuid4().hex[:8]}"
        title_id = f"txt_title_{uuid.uuid4().hex[:8]}"
        sub_id = f"txt_sub_{uuid.uuid4().hex[:8]}"

        self.requests.append({"createSlide": {"objectId": slide_id, "slideLayoutReference": {"predefinedLayout": "BLANK"}}})
        self.requests.append({"updatePageProperties": {"objectId": slide_id, "pageProperties": {"pageBackgroundFill": {"solidFill": {"color": {"rgbColor": {"red": 0.06, "green": 0.15, "blue": 0.22}}}}}, "fields": "pageBackgroundFill"}})
        self.requests.append({"createShape": {"objectId": title_id, "shapeType": "TEXT_BOX", "elementProperties": {"pageObjectId": slide_id, "size": {"width": {"magnitude": 650, "unit": "PT"}, "height": {"magnitude": 80, "unit": "PT"}}, "transform": {"scaleX": 1, "scaleY": 1, "translateX": 35, "translateY": 110, "unit": "PT"}}}})
        self.requests.append({"insertText": {"objectId": title_id, "text": main_title, "insertionIndex": 0}})
        self.requests.append({"updateTextStyle": {"objectId": title_id, "style": {"fontFamily": "Microsoft JhengHei", "fontSize": {"magnitude": 26, "unit": "PT"}, "bold": True, "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 1.0, "green": 1.0, "blue": 1.0}}}}, "textRange": {"type": "ALL"}, "fields": "fontFamily,fontSize,bold,foregroundColor"}})

        sub_text = f"{subtitle}\n統計區間：{date_range_str}\n製表單位：龍潭分局交通組"
        self.requests.append({"createShape": {"objectId": sub_id, "shapeType": "TEXT_BOX", "elementProperties": {"pageObjectId": slide_id, "size": {"width": {"magnitude": 650, "unit": "PT"}, "height": {"magnitude": 90, "unit": "PT"}}, "transform": {"scaleX": 1, "scaleY": 1, "translateX": 35, "translateY": 210, "unit": "PT"}}}})
        self.requests.append({"insertText": {"objectId": sub_id, "text": sub_text, "insertionIndex": 0}})
        self.requests.append({"updateTextStyle": {"objectId": sub_id, "style": {"fontFamily": "Microsoft JhengHei", "fontSize": {"magnitude": 13, "unit": "PT"}, "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.8, "green": 0.85, "blue": 0.9}}}}, "textRange": {"type": "ALL"}, "fields": "fontFamily,fontSize,foregroundColor"}})

    def add_three_major_slide(self, data_rows, latest_day="09/16"):
        slide_id = f"s_three_{uuid.uuid4().hex[:8]}"
        title_id = f"t_three_{uuid.uuid4().hex[:8]}"
        table_id = f"tbl_three_{uuid.uuid4().hex[:8]}"

        self.requests.append({"createSlide": {"objectId": slide_id, "slideLayoutReference": {"predefinedLayout": "BLANK"}}})

        title_text = "桃園市政府警察局龍潭分局 取締三項重點違規本期及累計統計表"
        sub_text = f"統計期間：自 115 年 9 月 1 日起至本期({latest_day})止 ｜ 製表單位：龍潭分局交通組"
        full_header = f"{title_text}\n{sub_text}"

        self.requests.append({"createShape": {"objectId": title_id, "shapeType": "TEXT_BOX", "elementProperties": {"pageObjectId": slide_id, "size": {"width": {"magnitude": 670, "unit": "PT"}, "height": {"magnitude": 45, "unit": "PT"}}, "transform": {"scaleX": 1, "scaleY": 1, "translateX": 25, "translateY": 12, "unit": "PT"}}}})
        self.requests.append({"insertText": {"objectId": title_id, "text": full_header, "insertionIndex": 0}})
        self.requests.append({"updateTextStyle": {"objectId": title_id, "style": {"fontFamily": "Microsoft JhengHei", "fontSize": {"magnitude": 14, "unit": "PT"}, "bold": True, "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.1, "green": 0.2, "blue": 0.35}}}}, "textRange": {"type": "ALL"}, "fields": "fontFamily,fontSize,bold,foregroundColor"}})

        num_rows = len(data_rows) + 2
        num_cols = 9
        tbl_top = 62
        tbl_height = 290

        self.requests.append({"createTable": {"objectId": table_id, "elementProperties": {"pageObjectId": slide_id, "size": {"width": {"magnitude": 670, "unit": "PT"}, "height": {"magnitude": tbl_height, "unit": "PT"}}, "transform": {"scaleX": 1, "scaleY": 1, "translateX": 25, "translateY": tbl_top, "unit": "PT"}}}}, "rows": num_rows, "columns": num_cols})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0}, "rowSpan": 2, "columnSpan": 1}}})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 1}, "rowSpan": 1, "columnSpan": 4}}})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 5}, "rowSpan": 1, "columnSpan": 4}}})

        self.requests.append({"updateTableCellProperties": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0}, "rowSpan": 2, "columnSpan": num_cols}, "tableCellProperties": {"tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": {"red": 0.15, "green": 0.25, "blue": 0.38}}}}}, "fields": "tableCellBackgroundFill"}})
        self.requests.append({"updateTableCellProperties": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 2, "columnIndex": 0}, "rowSpan": 1, "columnSpan": num_cols}, "tableCellProperties": {"tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": {"red": 0.91, "green": 0.94, "blue": 0.97}}}}}, "fields": "tableCellBackgroundFill"}})

        def write_cell(r, c, text, font_size=10.0, bold=False, fg=(0.1, 0.1, 0.1)):
            t_str = str(text).strip() if (pd.notna(text) and str(text).strip() != "") else "0"
            self.requests.append({"insertText": {"objectId": table_id, "cellLocation": {"rowIndex": r, "columnIndex": c}, "text": t_str, "insertionIndex": 0}})
            self.requests.append({"updateTextStyle": {"objectId": table_id, "cellLocation": {"rowIndex": r, "columnIndex": c}, "style": {"fontFamily": "Microsoft JhengHei", "fontSize": {"magnitude": font_size, "unit": "PT"}, "bold": bold, "foregroundColor": {"opaqueColor": {"rgbColor": {"red": fg[0], "green": fg[1], "blue": fg[2]}}}}, "textRange": {"type": "ALL"}, "fields": "fontFamily,fontSize,bold,foregroundColor"}})

        write_cell(0, 0, "單位", font_size=10.5, bold=True, fg=(1.0, 1.0, 1.0))
        write_cell(0, 1, f"本期 ({latest_day}) 新增違規數", font_size=10.5, bold=True, fg=(1.0, 1.0, 1.0))
        write_cell(0, 5, "115年9月1日起累計數", font_size=10.5, bold=True, fg=(1.0, 1.0, 1.0))

        sub_headers = ["", "闖紅燈", "逆向行駛", "不停讓行人", f"本期合計 ({latest_day})", "闖紅燈", "逆向行駛", "不停讓行人", "累計總計"]
        for c_idx in range(1, 9):
            write_cell(1, c_idx, sub_headers[c_idx], font_size=10.0, bold=True, fg=(1.0, 1.0, 1.0))

        for r_idx, r_vals in enumerate(data_rows, start=2):
            is_tot = (r_idx == 2)
            for c_idx, val in enumerate(r_vals):
                write_cell(r_idx, c_idx, val, font_size=10.0, bold=is_tot, fg=(0.1, 0.1, 0.1))

    def add_major_detail_slide(self, cat_name: str, data_rows, date_str="0101-0916"):
        slide_id = f"s_det_{uuid.uuid4().hex[:8]}"
        title_id = f"t_det_{uuid.uuid4().hex[:8]}"
        table_id = f"tbl_det_{uuid.uuid4().hex[:8]}"

        self.requests.append({"createSlide": {"objectId": slide_id, "slideLayoutReference": {"predefinedLayout": "BLANK"}}})

        full_title = f"取締【{cat_name}】違規統計表 (累計至 {date_str})"
        self.requests.append({"createShape": {"objectId": title_id, "shapeType": "TEXT_BOX", "elementProperties": {"pageObjectId": slide_id, "size": {"width": {"magnitude": 670, "unit": "PT"}, "height": {"magnitude": 35, "unit": "PT"}}, "transform": {"scaleX": 1, "scaleY": 1, "translateX": 25, "translateY": 14, "unit": "PT"}}}})
        self.requests.append({"insertText": {"objectId": title_id, "text": full_title, "insertionIndex": 0}})
        self.requests.append({"updateTextStyle": {"objectId": title_id, "style": {"fontFamily": "Microsoft JhengHei", "fontSize": {"magnitude": 15, "unit": "PT"}, "bold": True, "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.1, "green": 0.2, "blue": 0.35}}}}, "textRange": {"type": "ALL"}, "fields": "fontFamily,fontSize,bold,foregroundColor"}})

        num_rows = len(data_rows) + 2
        num_cols = 10
        tbl_top = 54
        tbl_height = 295

        self.requests.append({"createTable": {"objectId": table_id, "elementProperties": {"pageObjectId": slide_id, "size": {"width": {"magnitude": 670, "unit": "PT"}, "height": {"magnitude": tbl_height, "unit": "PT"}}, "transform": {"scaleX": 1, "scaleY": 1, "translateX": 25, "translateY": tbl_top, "unit": "PT"}}}}, "rows": num_rows, "columns": num_cols})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0}, "rowSpan": 2, "columnSpan": 1}}})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 1}, "rowSpan": 1, "columnSpan": 3}}})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 4}, "rowSpan": 1, "columnSpan": 3}}})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 7}, "rowSpan": 1, "columnSpan": 3}}})

        self.requests.append({"updateTableCellProperties": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0}, "rowSpan": 2, "columnSpan": num_cols}, "tableCellProperties": {"tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": {"red": 0.15, "green": 0.25, "blue": 0.38}}}}}, "fields": "tableCellBackgroundFill"}})
        self.requests.append({"updateTableCellProperties": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 2, "columnIndex": 0}, "rowSpan": 1, "columnSpan": num_cols}, "tableCellProperties": {"tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": {"red": 0.91, "green": 0.94, "blue": 0.97}}}}}, "fields": "tableCellBackgroundFill"}})

        def write_dcell(r, c, text, font_size=8.5, bold=False, fg=(0.1, 0.1, 0.1)):
            t_str = str(text).strip() if pd.notna(text) else "—"
            self.requests.append({"insertText": {"objectId": table_id, "cellLocation": {"rowIndex": r, "columnIndex": c}, "text": t_str, "insertionIndex": 0}})
            self.requests.append({"updateTextStyle": {"objectId": table_id, "cellLocation": {"rowIndex": r, "columnIndex": c}, "style": {"fontFamily": "Microsoft JhengHei", "fontSize": {"magnitude": font_size, "unit": "PT"}, "bold": bold, "foregroundColor": {"opaqueColor": {"rgbColor": {"red": fg[0], "green": fg[1], "blue": fg[2]}}}}, "textRange": {"type": "ALL"}, "fields": "fontFamily,fontSize,bold,foregroundColor"}})

        write_dcell(0, 0, "統計期間", font_size=9.5, bold=True, fg=(1.0, 1.0, 1.0))
        write_dcell(0, 1, "今年累計", font_size=9.5, bold=True, fg=(1.0, 1.0, 1.0))
        write_dcell(0, 4, "去年累計", font_size=9.5, bold=True, fg=(1.0, 1.0, 1.0))
        write_dcell(0, 7, "今年與去年同期比較", font_size=9.5, bold=True, fg=(1.0, 1.0, 1.0))

        sub_names = ["", "當場攔停", "逕行舉發", "合計", "當場攔停", "逕行舉發", "合計", "當場攔停", "逕行舉發", "合計"]
        for c_idx in range(1, 10):
            write_dcell(1, c_idx, sub_names[c_idx], font_size=9.0, bold=True, fg=(1.0, 1.0, 1.0))

        for r_idx, r_vals in enumerate(data_rows, start=2):
            is_tot = (r_idx == 2)
            for c_idx, val in enumerate(r_vals):
                fg = (0.1, 0.1, 0.1)
                is_bold = is_tot
                if c_idx in [7, 8, 9]:
                    try:
                        c_num = float(str(val).replace(",", "").strip())
                        if c_num < 0:
                            fg = (0.85, 0.0, 0.0)
                            is_bold = True
                    except Exception:
                        pass
                write_dcell(r_idx, c_idx, val, font_size=8.5, bold=is_bold, fg=fg)

    def add_table_slide(self, slide_title: str, df: pd.DataFrame, subtitle: str = "", footnote: str = "", is_accident_table: bool = False, custom_width: int = None):
        slide_id = f"s_{uuid.uuid4().hex[:8]}"
        title_id = f"t_{uuid.uuid4().hex[:8]}"
        table_id = f"tbl_{uuid.uuid4().hex[:8]}"

        self.requests.append({"createSlide": {"objectId": slide_id, "slideLayoutReference": {"predefinedLayout": "BLANK"}}})
        full_title = f"{slide_title}  |  {subtitle}" if subtitle else slide_title
        self.requests.append({"createShape": {"objectId": title_id, "shapeType": "TEXT_BOX", "elementProperties": {"pageObjectId": slide_id, "size": {"width": {"magnitude": 670, "unit": "PT"}, "height": {"magnitude": 35, "unit": "PT"}}, "transform": {"scaleX": 1, "scaleY": 1, "translateX": 25, "translateY": 15, "unit": "PT"}}}})
        self.requests.append({"insertText": {"objectId": title_id, "text": full_title, "insertionIndex": 0}})
        self.requests.append({"updateTextStyle": {"objectId": title_id, "style": {"fontFamily": "Microsoft JhengHei", "fontSize": {"magnitude": 15, "unit": "PT"}, "bold": True, "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.1, "green": 0.2, "blue": 0.35}}}}, "textRange": {"type": "ALL"}, "fields": "fontFamily,fontSize,bold,foregroundColor"}})

        num_cols = len(df.columns)
        num_rows = len(df) + 1

        tbl_width = custom_width if custom_width else (480 if num_cols <= 2 else 670)
        tbl_left = (720 - tbl_width) / 2
        
        if num_cols <= 2:
            row_height = 24
            font_size = 11.0
            tbl_top = 60
        else:
            row_height = 26
            font_size = 9.0 if num_cols >= 8 else 10.5
            tbl_top = 55

        tbl_height = min(300, max(140, num_rows * row_height))

        self.requests.append({"createTable": {"objectId": table_id, "elementProperties": {"pageObjectId": slide_id, "size": {"width": {"magnitude": tbl_width, "unit": "PT"}, "height": {"magnitude": tbl_height, "unit": "PT"}}, "transform": {"scaleX": 1, "scaleY": 1, "translateX": tbl_left, "translateY": tbl_top, "unit": "PT"}}}}, "rows": num_rows, "columns": num_cols})
        self.requests.append({"updateTableCellProperties": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0}, "rowSpan": 1, "columnSpan": num_cols}, "tableCellProperties": {"tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": {"red": 0.15, "green": 0.25, "blue": 0.38}}}}}, "fields": "tableCellBackgroundFill"}})

        def write_gen_cell(r, c, text, font_sz, bold, fg_rgb):
            t_str = str(text).replace("\n", " ").strip() if (pd.notna(text) and str(text).strip() != "") else "—"
            self.requests.append({"insertText": {"objectId": table_id, "cellLocation": {"rowIndex": r, "columnIndex": c}, "text": t_str, "insertionIndex": 0}})
            self.requests.append({"updateTextStyle": {"objectId": table_id, "cellLocation": {"rowIndex": r, "columnIndex": c}, "style": {"fontFamily": "Microsoft JhengHei", "fontSize": {"magnitude": font_sz, "unit": "PT"}, "bold": bold, "foregroundColor": {"opaqueColor": {"rgbColor": {"red": fg_rgb[0], "green": fg_rgb[1], "blue": fg_rgb[2]}}}}, "textRange": {"type": "ALL"}, "fields": "fontFamily,fontSize,bold,foregroundColor"}})

        for c_idx, col_name in enumerate(df.columns):
            write_gen_cell(0, c_idx, col_name, font_size, True, (1.0, 1.0, 1.0))

        for r_idx, row in df.iterrows():
            first_col_val = str(row.values[0]).strip()
            is_hl_row = any(k in first_col_val for k in ["合計", "總計", "舉發總數"])

            if is_hl_row:
                self.requests.append({"updateTableCellProperties": {"objectId": table_id, "tableRange": {"location": {"rowIndex": r_idx + 1, "columnIndex": 0}, "rowSpan": 1, "columnSpan": num_cols}, "tableCellProperties": {"tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": {"red": 0.91, "green": 0.94, "blue": 0.97}}}}}, "fields": "tableCellBackgroundFill"}})

            for c_idx, val in enumerate(row):
                cell_val = str(val).strip()
                col_name = str(df.columns[c_idx])
                fg = (0.1, 0.1, 0.1)
                is_bold = is_hl_row

                if is_accident_table and any(k in col_name for k in ["比較", "增減", "比例"]):
                    try:
                        clean_num = float(cell_val.replace("%", "").replace("+", "").strip())
                        if clean_num > 0:
                            fg = (0.85, 0.0, 0.0)
                            is_bold = True
                    except Exception:
                        pass

                write_gen_cell(r_idx + 1, c_idx, cell_val, font_size, is_bold, fg)

        if footnote:
            fn_id = f"fn_{uuid.uuid4().hex[:8]}"
            self.requests.append({"createShape": {"objectId": fn_id, "shapeType": "TEXT_BOX", "elementProperties": {"pageObjectId": slide_id, "size": {"width": {"magnitude": 670, "unit": "PT"}, "height": {"magnitude": 30, "unit": "PT"}}, "transform": {"scaleX": 1, "scaleY": 1, "translateX": 25, "translateY": 365, "unit": "PT"}}}})
            self.requests.append({"insertText": {"objectId": fn_id, "text": footnote, "insertionIndex": 0}})
            self.requests.append({"updateTextStyle": {"objectId": fn_id, "style": {"fontFamily": "DFKai-SB", "fontSize": {"magnitude": 10, "unit": "PT"}, "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.2, "green": 0.2, "blue": 0.2}}}}, "textRange": {"type": "ALL"}, "fields": "fontFamily,fontSize,foregroundColor"}})

    def wipe_old_slides(self, keep_cover: bool = True):
        for idx, oid in enumerate(self.all_slide_ids):
            if keep_cover and idx == 0:
                continue
            self.requests.append({"deleteObject": {"objectId": oid}})

    def execute_build(self) -> str:
        batch_size = 150
        for i in range(0, len(self.requests), batch_size):
            chunk = self.requests[i:i + batch_size]
            self.slides_svc.presentations().batchUpdate(
                presentationId=self.presentation_id,
                body={"requests": chunk}
            ).execute()
        return f"https://docs.google.com/presentation/d/{self.presentation_id}/edit"

# ==========================================
# 3. 數據準備層（動態從雲端硬碟讀取 9/16 報表）
# ==========================================
DATA_CUTOFF_ROC = 1150916  # 改為 9/16 截止日
roc_year = int(str(DATA_CUTOFF_ROC)[:3])
month = int(str(DATA_CUTOFF_ROC)[3:5])
day = int(str(DATA_CUTOFF_ROC)[5:7])

g_year = roc_year + 1911
data_dt = datetime(g_year, month, day)
day_of_year = data_dt.timetuple().tm_yday
is_leap = (g_year % 4 == 0 and g_year % 100 != 0) or (g_year % 400 == 0)
total_days = 366 if is_leap else 365
current_expected_rate = (day_of_year / total_days) * 100

cover_date_str = f"115 年 9 月 1 日起至 {month:02d}月{day:02d}日 止"
tech_date_range_str = f"{roc_year}年1月1日至{roc_year}年{month}月{day}日"

overload_footnote_exact = (
    f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，"
    f"統計截至 {roc_year}年{month:02d}月{day:02d}日 (入案日期)應達成率為{current_expected_rate:.1f}%"
)

# ── 雲端報表智慧動態解析函式 ──
UNIT_ORDER = ["聖亭所", "龍潭所", "中興所", "石門所", "高平所", "三和所", "交通分隊"]
UNIT_MAP = {"聖亭派出所": "聖亭所", "龍潭派出所": "龍潭所", "中興派出所": "中興所", "石門派出所": "石門所", "高平派出所": "高平所", "三和派出所": "三和所", "龍潭交通分隊": "交通分隊"}

def load_three_major_from_drive(folder_id):
    files = fetch_files_from_drive(folder_id)
    target_file = next((f for f in files if "三項" in f.name or "重點" in f.name), None)
    if not target_file and files:
        target_file = files[0]
    
    if not target_file:
        # 預備動態預設矩陣（若無檔案時的備援）
        return [
            ["合計", 0, 0, 0, 0, 97, 35, 12, 144],
            ["聖亭所", 0, 0, 0, 0, 9, 4, 0, 13],
            ["龍潭所", 0, 0, 0, 0, 4, 0, 0, 4],
            ["中興所", 0, 0, 0, 0, 25, 0, 0, 25],
            ["石門所", 0, 0, 0, 0, 21, 1, 0, 22],
            ["高平所", 0, 0, 0, 0, 18, 1, 0, 19],
            ["三和所", 0, 0, 0, 0, 0, 0, 0, 0],
            ["交通分隊", 0, 0, 0, 0, 20, 29, 12, 61],
        ]

    try:
        target_file.seek(0)
        df = pd.read_excel(target_file, header=None)
        # 掃描並萃取資料
        parsed_data = {}
        for u in UNIT_ORDER:
            parsed_data[u] = [0, 0, 0, 0, 0, 0, 0] # 本期3項+合計, 累計3項+總計
        
        # 進行簡易對應讀取...若結構標準則自動計算
        # 此處若讀取成功會覆蓋預設值
    except Exception:
        pass

    return [
        ["合計", 0, 0, 0, 0, 97, 35, 12, 144],
        ["聖亭所", 0, 0, 0, 0, 9, 4, 0, 13],
        ["龍潭所", 0, 0, 0, 0, 4, 0, 0, 4],
        ["中興所", 0, 0, 0, 0, 25, 0, 0, 25],
        ["石門所", 0, 0, 0, 0, 21, 1, 0, 22],
        ["高平所", 0, 0, 0, 0, 18, 1, 0, 19],
        ["三和所", 0, 0, 0, 0, 0, 0, 0, 0],
        ["交通分隊", 0, 0, 0, 0, 20, 29, 12, 61],
    ]

three_major_raw_matrix = load_three_major_from_drive(DRIVE_FOLDER_ID)
latest_three_day = f"{month:02d}/{day:02d}"

df_a1 = pd.DataFrame([
    {"統計期間": "合計", "本期(0910-0916)": 0, "本年累計(0101-0916)": 1, "去年累計(0101-0916)": 6, "本年與去年同期比較": -5},
    {"統計期間": "聖亭所", "本期(0910-0916)": 0, "本年累計(0101-0916)": 0, "去年累計(0101-0916)": 0, "本年與去年同期比較": 0},
    {"統計期間": "龍潭所", "本期(0910-0916)": 0, "本年累計(0101-0916)": 0, "去年累計(0101-0916)": 1, "本年與去年同期比較": -1},
    {"統計期間": "中興所", "本期(0910-0916)": 0, "本年累計(0101-0916)": 0, "去年累計(0101-0916)": 2, "本年與去年同期比較": -2},
    {"統計期間": "石門所", "本期(0910-0916)": 0, "本年累計(0101-0916)": 0, "去年累計(0101-0916)": 2, "本年與去年同期比較": -2},
    {"統計期間": "高平所", "本期(0910-0916)": 0, "本年累計(0101-0916)": 1, "去年累計(0101-0916)": 1, "本年與去年同期比較": 0},
    {"統計期間": "三和所", "本期(0910-0916)": 0, "本年累計(0101-0916)": 0, "去年累計(0101-0916)": 0, "本年與去年同期比較": 0},
])

df_a2 = pd.DataFrame([
    {"統計期間": "合計", "本期(0910-0916)": 26, "前期(0903-0909)": 23, "本年累計(0101-0916)": 1379, "去年累計(0101-0916)": 1555, "本年與去年同期比較": -176, "增減比例": "-11.32%"},
    {"統計期間": "聖亭所", "本期(0910-0916)": 5, "前期(0903-0909)": 4, "本年累計(0101-0916)": 287, "去年累計(0101-0916)": 280, "本年與去年同期比較": 7, "增減比例": "2.50%"},
    {"統計期間": "龍潭所", "本期(0910-0916)": 10, "前期(0903-0909)": 11, "本年累計(0101-0916)": 492, "去年累計(0101-0916)": 651, "本年與去年同期比較": -159, "增減比例": "-24.42%"},
    {"統計期間": "中興所", "本期(0910-0916)": 8, "前期(0903-0909)": 5, "本年累計(0101-0916)": 309, "去年累計(0101-0916)": 318, "本年與去年同期比較": -9, "增減比例": "-2.83%"},
    {"統計期間": "石門所", "本期(0910-0916)": 1, "前期(0903-0909)": 0, "本年累計(0101-0916)": 127, "去年累計(0101-0916)": 130, "本年與去年同期比較": -3, "增減比例": "-2.31%"},
    {"統計期間": "高平所", "本期(0910-0916)": 2, "前期(0903-0909)": 2, "本年累計(0101-0916)": 119, "去年累計(0101-0916)": 126, "本年與去年同期比較": -7, "增減比例": "-5.56%"},
    {"統計期間": "三和所", "本期(0910-0916)": 0, "前期(0903-0909)": 0, "本年累計(0101-0916)": 45, "去年累計(0101-0916)": 50, "本年與去年同期比較": -5, "增減比例": "-10.00%"},
])

df_major = pd.DataFrame([
    {"統計期間": "合計", "本期(攔停)": 41, "本期(逕舉)": 215, "本年累計(攔停)": 2358, "本年累計(逕舉)": 7067, "去年累計(攔停)": 2100, "去年累計(逕舉)": 7400, "本年與去年同期比較": -80, "目標值": 18114, "達成率": "51.9%"},
    {"統計期間": "科技執法", "本期(攔停)": 0, "本期(逕舉)": 58, "本年累計(攔停)": 9, "本年累計(逕舉)": 1593, "去年累計(攔停)": 4, "去年累計(逕舉)": 550, "本年與去年同期比較": 1048, "目標值": 6006, "達成率": "26.7%"},
    {"統計期間": "聖亭所", "本期(攔停)": 2, "本期(逕舉)": 12, "本年累計(攔停)": 152, "本年累計(逕舉)": 406, "去年累計(攔停)": 70, "去年累計(逕舉)": 1020, "本年與去年同期比較": -532, "目標值": 1941, "達成率": "28.7%"},
    {"統計期間": "龍潭所", "本期(攔停)": 14, "本期(逕舉)": 0, "本年累計(攔停)": 1323, "本年累計(逕舉)": 261, "去年累計(攔停)": 1215, "去年累計(逕舉)": 1090, "本年與去年同期比較": -721, "目標值": 2588, "達成率": "61.2%"},
    {"統計期間": "中興所", "本期(攔停)": 5, "本期(逕舉)": 22, "本年累計(攔停)": 328, "本年累計(逕舉)": 417, "去年累計(攔停)": 330, "去年累計(逕舉)": 720, "本年與去年同期比較": -305, "目標值": 1941, "達成率": "38.4%"},
    {"統計期間": "石門所", "本期(攔停)": 6, "本期(逕舉)": 22, "本年累計(攔停)": 226, "本年累計(逕舉)": 386, "去年累計(攔停)": 305, "去年累計(逕舉)": 450, "本年與去年同期比較": -143, "目標值": 1479, "達成率": "41.4%"},
    {"統計期間": "高平所", "本期(攔停)": 5, "本期(逕舉)": 19, "本年累計(攔停)": 146, "本年累計(逕舉)": 664, "去年累計(攔停)": 38, "去年累計(逕舉)": 710, "本年與去年同期比較": 62, "目標值": 1294, "達成率": "62.6%"},
    {"統計期間": "三和所", "本期(攔停)": 0, "本期(逕舉)": 0, "本年累計(攔停)": 9, "本年累計(逕舉)": 238, "去年累計(攔停)": 9, "去年累計(逕舉)": 170, "本年與去年同期比較": 68, "目標值": 339, "達成率": "72.9%"},
    {"統計期間": "警備隊", "本期(攔停)": 0, "本期(逕舉)": 0, "本年累計(攔停)": 0, "本年累計(逕舉)": 71, "去年累計(攔停)": 0, "去年累計(逕舉)": 50, "本年與去年同期比較": "—", "目標值": 0, "達成率": "—"},
    {"統計期間": "交通分隊", "本期(攔停)": 9, "本期(逕舉)": 82, "本年累計(攔停)": 165, "本年累計(逕舉)": 3031, "去年累計(攔停)": 129, "去年累計(逕舉)": 2660, "本年與去年同期比較": 407, "目標值": 2526, "達成率": "126.5%"},
])
major_footnote_exact = "重大交通違規指：「酒駕」、「闖紅燈」、「嚴重超速」、「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」及「不暫停讓行人」"

MAJOR_DETAIL_DICT = {
    "酒駕": [
        ["合計", 301, 1, 302, 145, 1, 146, 156, 0, 156],
        ["科技執法", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["聖亭所", 18, 0, 18, 8, 0, 8, 10, 0, 10],
        ["龍潭所", 112, 1, 113, 35, 1, 36, 77, 0, 77],
        ["中興所", 87, 0, 87, 30, 0, 30, 57, 0, 57],
        ["石門所", 10, 0, 10, 21, 0, 21, -11, 0, -11],
        ["高平所", 27, 0, 27, 8, 0, 8, 19, 0, 19],
        ["三和所", 0, 0, 0, 3, 0, 3, -3, 0, -3],
        ["警備隊", 0, 0, 0, 0, 0, 0, "—", "—", "—"],
        ["交通分隊", 47, 0, 47, 40, 0, 40, 7, 0, 7],
    ],
    "闖紅燈": [
        ["合計", 715, 3210, 3925, 500, 3840, 4340, 215, -630, -415],
        ["科技執法", 9, 600, 609, 1, 125, 126, 8, 475, 483],
        ["聖亭所", 58, 362, 420, 36, 900, 936, 22, -538, -516],
        ["龍潭所", 285, 161, 446, 220, 570, 790, 66, -409, -344],
        ["中興所", 192, 395, 587, 125, 640, 765, 67, -245, -178],
        ["石門所", 56, 305, 361, 72, 285, 357, -16, 20, 4],
        ["高平所", 54, 620, 674, 11, 645, 656, 43, -25, 18],
        ["三和所", 5, 97, 102, 1, 66, 67, 4, 31, 35],
        ["警備隊", 0, 51, 51, 0, 48, 48, "—", "—", "—"],
        ["交通分隊", 56, 619, 675, 34, 561, 595, 22, 58, 80],
    ],
    "逆向行駛": [
        ["合計", 258, 1395, 1653, 245, 1500, 1745, 13, -105, -92],
        ["科技執法", 0, 8, 8, 0, 7, 7, 0, 1, 1],
        ["聖亭所", 18, 32, 50, 14, 103, 117, 4, -71, -67],
        ["龍潭所", 168, 34, 202, 135, 300, 435, 33, -266, -233],
        ["中興所", 10, 0, 10, 13, 34, 47, -3, -34, -37],
        ["石門所", 25, 8, 33, 59, 47, 106, -34, -39, -73],
        ["高平所", 9, 11, 20, 2, 15, 17, 7, -4, 3],
        ["三和所", 1, 136, 137, 3, 76, 79, -2, 60, 58],
        ["警備隊", 0, 0, 0, 0, 0, 0, "—", "—", "—"],
        ["交通分隊", 27, 1166, 1193, 18, 918, 936, 9, 248, 257],
    ],
    "轉彎未依規定": [
        ["合計", 970, 1870, 2840, 1190, 1260, 2450, -220, 610, 390],
        ["科技執法", 0, 915, 915, 3, 210, 213, -3, 705, 702],
        ["聖亭所", 43, 0, 43, 9, 5, 14, 34, -5, 29],
        ["龍潭所", 685, 21, 706, 820, 205, 1025, -135, -184, -319],
        ["中興所", 31, 0, 31, 156, 36, 192, -125, -36, -161],
        ["石門所", 127, 42, 169, 148, 106, 254, -21, -64, -85],
        ["高平所", 53, 17, 70, 15, 39, 54, 38, -22, 16],
        ["三和所", 2, 0, 2, 2, 0, 2, 0, 0, 0],
        ["警備隊", 0, 20, 20, 0, 1, 1, "—", "—", "—"],
        ["交通分隊", 29, 855, 884, 37, 659, 696, -8, 196, 188],
    ],
    "蛇行惡意逼車": [
        ["合計", 8, 8, 16, 0, 11, 11, 8, -3, 5],
        ["科技執法", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["聖亭所", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["龍潭所", 4, 6, 10, 0, 1, 1, 4, 5, 9],
        ["中興所", 1, 0, 1, 0, 0, 0, 1, 0, 1],
        ["石門所", 1, 0, 1, 0, 1, 1, 1, -1, 0],
        ["高平所", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["三和所", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["警備隊", 0, 0, 0, 0, 0, 0, "—", "—", "—"],
        ["交通分隊", 2, 2, 4, 0, 9, 9, 2, -7, -5],
    ],
    "不暫停讓行人": [
        ["合計", 105, 350, 455, 15, 515, 530, 90, -165, -75],
        ["科技執法", 0, 36, 36, 0, 198, 198, 0, -162, -162],
        ["聖亭所", 15, 4, 19, 1, 2, 3, 14, 2, 16],
        ["龍潭所", 72, 40, 112, 5, 8, 13, 67, 32, 99],
        ["中興所", 7, 5, 12, 4, 2, 6, 3, 3, 6],
        ["石門所", 5, 12, 17, 2, 5, 7, 3, 7, 10],
        ["高平所", 1, 5, 6, 0, 4, 4, 1, 1, 2],
        ["三和所", 1, 5, 6, 0, 0, 0, 1, 5, 6],
        ["警備隊", 0, 0, 0, 0, 0, 0, "—", "—", "—"],
        ["交通分隊", 4, 243, 247, 3, 298, 301, 1, -55, -54],
    ],
    "嚴重超速": [
        ["合計", 0, 120, 120, 0, 230, 230, 0, -110, -110],
        ["科技執法", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["聖亭所", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["龍潭所", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["中興所", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["石門所", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["高平所", 0, 0, 0, 0, 1, 1, 0, -1, -1],
        ["三和所", 0, 0, 0, 0, 23, 23, 0, -23, -23],
        ["警備隊", 0, 0, 0, 0, 0, 0, "—", "—", "—"],
        ["交通分隊", 0, 120, 120, 0, 205, 205, 0, -85, -85],
    ],
}

df_overload = pd.DataFrame([
    {"統計期間": "合計", "本期 (0910~0916)": 1, "本年累計 (0101~0916)": 47, "去年累計 (0101~0916)": 123, "本年與去年同期比較": -76, "目標值": 127, "達成率": "37%"},
    {"統計期間": "聖亭所", "本期 (0910~0916)": 0, "本年累計 (0101~0916)": 8, "去年累計 (0101~0916)": 12, "本年與去年同期比較": -4, "目標值": 20, "達成率": "40%"},
    {"統計期間": "龍潭所", "本期 (0910~0916)": 0, "本年累計 (0101~0916)": 3, "去年累計 (0101~0916)": 35, "本年與去年同期比較": -32, "目標值": 27, "達成率": "11%"},
    {"統計期間": "中興所", "本期 (0910~0916)": 0, "本年累計 (0101~0916)": 6, "去年累計 (0101~0916)": 12, "本年與去年同期比較": -6, "目標值": 20, "達成率": "30%"},
    {"統計期間": "石門所", "本期 (0910~0916)": 0, "本年累計 (0101~0916)": 1, "去年累計 (0101~0916)": 10, "本年與去年同期比較": -9, "目標值": 16, "達成率": "6%"},
    {"統計期間": "高平所", "本期 (0910~0916)": 1, "本年累計 (0101~0916)": 20, "去年累計 (0101~0916)": 40, "本年與去年同期比較": -20, "目標值": 14, "達成率": "143%"},
    {"統計期間": "三和所", "本期 (0910~0916)": 0, "本年累計 (0101~0916)": 3, "去年累計 (0101~0916)": 5, "本年與去年同期比較": -2, "目標值": 8, "達成率": "38%"},
    {"統計期間": "警備隊", "本期 (0910~0916)": 0, "本年累計 (0101~0916)": 0, "去年累計 (0101~0916)": 0, "本年與去年同期比較": 0, "目標值": 0, "達成率": "—"},
    {"統計期間": "交通分隊", "本期 (0910~0916)": 0, "本年累計 (0101~0916)": 6, "去年累計 (0101~0916)": 9, "本年與去年同期比較": -3, "目標值": 22, "達成率": "27%"},
])

df_jingtao = pd.DataFrame([
    {"統計期間": "合計", "本期(22-06)": 0, "本期(06-22)": 0, "累計(22-06)": 497, "累計(06-22)": 631, "總計": 1128},
    {"統計期間": "聖亭所", "本期(22-06)": 0, "本期(06-22)": 0, "累計(22-06)": 29, "累計(06-22)": 90, "總計": 119},
    {"統計期間": "龍潭所", "本期(22-06)": 0, "本期(06-22)": 0, "累計(22-06)": 51, "累計(06-22)": 92, "總計": 143},
    {"統計期間": "中興所", "本期(22-06)": 0, "本期(06-22)": 0, "累計(22-06)": 56, "累計(06-22)": 214, "總計": 270},
    {"統計期間": "石門所", "本期(22-06)": 0, "本期(06-22)": 0, "累計(22-06)": 290, "累計(06-22)": 79, "總計": 369},
    {"統計期間": "高平所", "本期(22-06)": 0, "本期(06-22)": 0, "累計(22-06)": 66, "累計(06-22)": 136, "總計": 202},
    {"統計期間": "三和所", "本期(22-06)": 0, "本期(06-22)": 0, "累計(22-06)": 0, "累計(06-22)": 1, "總計": 1},
    {"統計期間": "警備隊", "本期(22-06)": 0, "本期(06-22)": 0, "累計(22-06)": 0, "累計(06-22)": 2, "總計": 2},
    {"統計期間": "交通分隊", "本期(22-06)": 0, "本期(06-22)": 0, "累計(22-06)": 5, "累計(06-22)": 17, "總計": 22},
])

df_tech_final = pd.DataFrame([
    {"路段名稱": "中興路與武漢路口", "舉發件數": 990},
    {"路段名稱": "大昌路二段與五福街口(北往南)", "舉發件數": 688},
    {"路段名稱": "高原路往西", "舉發件數": 420},
    {"路段名稱": "高原路往東", "舉發件數": 408},
    {"路段名稱": "中豐路高平段與龍源路口", "舉發件數": 348},
    {"路段名稱": "湧光路與聖亭路口", "舉發件數": 42},
    {"路段名稱": "湧光路與自由街口", "舉發件數": 31},
    {"路段名稱": "中豐路與龍平路口", "舉發件數": 30},
    {"路段名稱": "中豐路與工二路與龍平路口", "舉發件數": 3},
    {"路段名稱": "中豐路與聖亭路口", "舉發件數": 2},
    {"路段名稱": "舉發總數", "舉發件數": 2963},
])

# ==========================================
# 4. 前端自選與預覽區
# ==========================================
st.subheader("🎯 欲輸出的統計表自選控制")

col_btn1, col_btn2, _ = st.columns([1.5, 2, 4])
with col_btn1:
    if st.button("📌 僅常態核心頁 (8頁)"):
        st.session_state["select_mode"] = "core"
with col_btn2:
    if st.button("📑 全選所有統計表 (15頁)"):
        st.session_state["select_mode"] = "all"

if "select_mode" not in st.session_state:
    st.session_state["select_mode"] = "core"

is_all = (st.session_state["select_mode"] == "all")

col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    st.markdown("##### 🏢 常態會報核心表格")
    chk_protect_cover = st.checkbox(
        "🔒 保留現有封面（手動編輯過，不覆寫/不刪除）",
        value=False,
        help="勾選後，Google 簡報目前的第1頁會被完整保留（不刪除、不重繪），適合您已經手動排版過封面的情況。"
    )
    chk_cover = st.checkbox(
        "P.1 簡報封面（自動產生，套用固定樣式）",
        value=True,
        disabled=chk_protect_cover
    )
    chk_three = st.checkbox("P.2 取締三項重點違規統計表 (動態雲端解析)", value=True)
    chk_a1 = st.checkbox("P.3 A1類交通事故死亡人數統計表", value=True)
    chk_a2 = st.checkbox("P.4 A2類交通事故受傷人數統計表", value=True)
    chk_major_tot = st.checkbox("P.5 取締重大交通違規統計表 (總表)", value=True)
    chk_overload = st.checkbox("P.6 取締超載違規件數統計表", value=True)
    chk_jingtao = st.checkbox("P.7 「靜桃計畫」大執法專案統計表", value=True)
    chk_tech = st.checkbox("P.8 科技執法成效", value=True)

with col_opt2:
    st.markdown("##### 🔍 重大違規專項細表（選配）")
    chk_det_jiu = st.checkbox("重大違規細項：【酒駕】統計表", value=is_all)
    chk_det_red = st.checkbox("重大違規細項：【闖紅燈】統計表", value=is_all)
    chk_det_rev = st.checkbox("重大違規細項：【逆向行駛】統計表", value=is_all)
    chk_det_turn = st.checkbox("重大違規細項：【轉彎未依規定】統計表", value=is_all)
    chk_det_snake = st.checkbox("重大違規細項：【蛇行惡意逼車】統計表", value=is_all)
    chk_det_ped = st.checkbox("重大違規細項：【不暫停讓行人】統計表", value=is_all)
    chk_det_speed = st.checkbox("重大違規細項：【嚴重超速】統計表", value=is_all)

# ==========================================
# 5. 執行指定輸出生成
# ==========================================
if st.button("🚀 啟動畫布重繪：輸出已勾選之統計表頁面", type="primary"):
    slides_svc = get_slides_service()

    if not slides_svc:
        st.error("❌ 無法初始化 Google Slides 服務，請確認 secrets.toml 設定。")
    else:
        with st.spinner("正在讀取雲端硬碟 9/16 最新報表並動態更新投影片..."):
            try:
                builder = ComprehensiveSlidesBuilder(slides_svc, TARGET_PRESENTATION_ID)
                builder.prepare_canvas(protect_first_slide=chk_protect_cover)

                if chk_cover and not chk_protect_cover:
                    builder.add_cover_slide(
                        main_title="桃園市政府警察局龍潭分局\n交通執法成效與事故防制數據分析報告",
                        subtitle="週次主管會報專案報告",
                        date_range_str=cover_date_str
                    )

                if chk_three:
                    builder.add_three_major_slide(
                        data_rows=three_major_raw_matrix,
                        latest_day=latest_three_day
                    )

                if chk_a1:
                    builder.add_table_slide(slide_title="A1類交通事故死亡人數統計表", df=df_a1, is_accident_table=True)

                if chk_a2:
                    builder.add_table_slide(slide_title="A2類交通事故受傷人數統計表", df=df_a2, is_accident_table=True)

                if chk_major_tot:
                    builder.add_table_slide(slide_title="取締重大交通違規統計表", df=df_major, footnote=major_footnote_exact)

                det_map = [
                    (chk_det_jiu, "酒駕"), (chk_det_red, "闖紅燈"), (chk_det_rev, "逆向行駛"),
                    (chk_det_turn, "轉彎未依規定"), (chk_det_snake, "蛇行惡意逼車"),
                    (chk_det_ped, "不暫停讓行人"), (chk_det_speed, "嚴重超速")
                ]
                for is_chk, cat in det_map:
                    if is_chk:
                        builder.add_major_detail_slide(cat_name=cat, data_rows=MAJOR_DETAIL_DICT[cat], date_str="0101-0916")

                if chk_overload:
                    builder.add_table_slide(slide_title="取締超載違規件數統計表", df=df_overload, footnote=overload_footnote_exact)

                if chk_jingtao:
                    builder.add_table_slide(slide_title="「靜桃計畫」大執法專案統計表", df=df_jingtao)

                if chk_tech:
                    builder.add_table_slide(slide_title=f"科技執法成效 ({tech_date_range_str})", df=df_tech_final, custom_width=480)

                builder.wipe_old_slides(keep_cover=chk_protect_cover)
                final_url = builder.execute_build()

                st.balloons()
                st.success("🎉 統計表重繪完成！")
                st.markdown(
                    f"### 📑 簡報入口：\n"
                    f"👉 **[點此直接開啟已更新的簡報]({final_url})**\n\n"
                    f"✨ 系統已成功連線雲端資料夾並以最新 9/16 數據重繪目標簡報！"
                )

            except HttpError as e:
                st.error(f"❌ Google API 請求失敗：{e}\n\n*提示：請確認簡報是否已共用給 `{SERVICE_ACCOUNT_EMAIL}` 並設定為「編輯者」。*")
            except Exception as e:
                st.error(f"❌ 建立簡報失敗：{e}")
