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

st.title("📽️ 全方位執法數據簡報直出中心（原版規格 15 頁完整版）")
st.caption("🚀 完整涵蓋：封面、三項重點、A1/A2事故、重大違規總表、7大項重大違規專項細表、超載、靜桃計畫及科技執法成效。")

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

with st.container():
    c_s1, c_s2 = st.columns([2, 1])
    with c_s1:
        st.info(f"🔑 **執行服務帳號：** `{SERVICE_ACCOUNT_EMAIL}`\n\n📂 **報表資料夾 ID：** `{DRIVE_FOLDER_ID}`")
    with c_s2:
        st.link_button("📂 開啟目標 Google 簡報", f"https://docs.google.com/presentation/d/{TARGET_PRESENTATION_ID}/edit")

# ==========================================
# 2. 全方位簡報排版引擎 (ComprehensiveSlidesBuilder)
# ==========================================
class ComprehensiveSlidesBuilder:
    def __init__(self, slides_svc, presentation_id: str):
        self.slides_svc = slides_svc
        self.presentation_id = presentation_id.strip()
        self.old_slide_ids = []
        self.requests = []

    def prepare_canvas(self):
        """記錄簡報所有既有頁面 ID（稍後全數安全抹除）"""
        pres = self.slides_svc.presentations().get(
            presentationId=self.presentation_id
        ).execute()
        slides = pres.get("slides", [])
        self.old_slide_ids = [s["objectId"] for s in slides]

    def add_cover_slide(self, main_title: str, subtitle: str, date_range_str: str):
        """【第 1 頁】警政深藍色專業封面"""
        slide_id = f"cover_{uuid.uuid4().hex[:8]}"
        title_id = f"txt_title_{uuid.uuid4().hex[:8]}"
        sub_id = f"txt_sub_{uuid.uuid4().hex[:8]}"

        self.requests.append({
            "createSlide": {
                "objectId": slide_id,
                "slideLayoutReference": {"predefinedLayout": "BLANK"}
            }
        })

        self.requests.append({
            "updatePageProperties": {
                "objectId": slide_id,
                "pageProperties": {
                    "pageBackgroundFill": {
                        "solidFill": {"color": {"rgbColor": {"red": 0.06, "green": 0.15, "blue": 0.22}}}
                    }
                },
                "fields": "pageBackgroundFill"
            }
        })

        self.requests.append({
            "createShape": {
                "objectId": title_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {"width": {"magnitude": 650, "unit": "PT"}, "height": {"magnitude": 80, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": 35, "translateY": 110, "unit": "PT"}
                }
            }
        })
        self.requests.append({"insertText": {"objectId": title_id, "text": main_title, "insertionIndex": 0}})
        self.requests.append({
            "updateTextStyle": {
                "objectId": title_id,
                "style": {
                    "fontFamily": "Microsoft JhengHei",
                    "fontSize": {"magnitude": 26, "unit": "PT"},
                    "bold": True,
                    "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 1.0, "green": 1.0, "blue": 1.0}}}
                },
                "textRange": {"type": "ALL"},
                "fields": "fontFamily,fontSize,bold,foregroundColor"
            }
        })

        sub_text = f"{subtitle}\n統計區間：{date_range_str}\n製表單位：龍潭分局交通組"
        self.requests.append({
            "createShape": {
                "objectId": sub_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {"width": {"magnitude": 650, "unit": "PT"}, "height": {"magnitude": 90, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": 35, "translateY": 210, "unit": "PT"}
                }
            }
        })
        self.requests.append({"insertText": {"objectId": sub_id, "text": sub_text, "insertionIndex": 0}})
        self.requests.append({
            "updateTextStyle": {
                "objectId": sub_id,
                "style": {
                    "fontFamily": "Microsoft JhengHei",
                    "fontSize": {"magnitude": 13, "unit": "PT"},
                    "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.8, "green": 0.85, "blue": 0.9}}}
                },
                "textRange": {"type": "ALL"},
                "fields": "fontFamily,fontSize,foregroundColor"
            }
        })

    def add_three_major_slide(self, data_rows, latest_day="09/15"):
        """【第 2 頁】三項重點違規（母本雙層合併表頭，原子化安全寫入）"""
        slide_id = f"s_three_{uuid.uuid4().hex[:8]}"
        title_id = f"t_three_{uuid.uuid4().hex[:8]}"
        table_id = f"tbl_three_{uuid.uuid4().hex[:8]}"

        self.requests.append({
            "createSlide": {
                "objectId": slide_id,
                "slideLayoutReference": {"predefinedLayout": "BLANK"}
            }
        })

        title_text = "桃園市政府警察局龍潭分局 取締三項重點違規本期及累計統計表"
        sub_text = f"統計期間：自 115 年 9 月 1 日起至本期({latest_day})止 ｜ 製表單位：龍潭分局交通組"
        full_header = f"{title_text}\n{sub_text}"

        self.requests.append({
            "createShape": {
                "objectId": title_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {"width": {"magnitude": 670, "unit": "PT"}, "height": {"magnitude": 45, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": 25, "translateY": 12, "unit": "PT"}
                }
            }
        })
        self.requests.append({"insertText": {"objectId": title_id, "text": full_header, "insertionIndex": 0}})
        self.requests.append({
            "updateTextStyle": {
                "objectId": title_id,
                "style": {
                    "fontFamily": "Microsoft JhengHei",
                    "fontSize": {"magnitude": 14, "unit": "PT"},
                    "bold": True,
                    "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.1, "green": 0.2, "blue": 0.35}}}
                },
                "textRange": {"type": "ALL"},
                "fields": "fontFamily,fontSize,bold,foregroundColor"
            }
        })

        num_rows = len(data_rows) + 2
        num_cols = 9
        tbl_top = 62
        tbl_height = 290

        self.requests.append({
            "createTable": {
                "objectId": table_id,
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {"width": {"magnitude": 670, "unit": "PT"}, "height": {"magnitude": tbl_height, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": 25, "translateY": tbl_top, "unit": "PT"}
                },
                "rows": num_rows,
                "columns": num_cols
            }
        })

        # 1. 表頭儲存格合併
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0}, "rowSpan": 2, "columnSpan": 1}}})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 1}, "rowSpan": 1, "columnSpan": 4}}})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 5}, "rowSpan": 1, "columnSpan": 4}}})

        # 2. 表頭底色（深藍）與合計列底色（淡藍）
        self.requests.append({"updateTableCellProperties": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0}, "rowSpan": 2, "columnSpan": num_cols}, "tableCellProperties": {"tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": {"red": 0.15, "green": 0.25, "blue": 0.38}}}}}, "fields": "tableCellBackgroundFill"}})
        self.requests.append({"updateTableCellProperties": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 2, "columnIndex": 0}, "rowSpan": 1, "columnSpan": num_cols}, "tableCellProperties": {"tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": {"red": 0.91, "green": 0.94, "blue": 0.97}}}}}, "fields": "tableCellBackgroundFill"}})

        # 3. 原子化寫入文字
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

    def add_major_detail_slide(self, cat_name: str, data_rows, date_str="0101-0915"):
        """【P.6~P.12 重大交通違規專項細表】雙層表頭原子化排版"""
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

        self.requests.append({"createTable": {"objectId": table_id, "elementProperties": {"pageObjectId": slide_id, "size": {"width": {"magnitude": 670, "unit": "PT"}, "height": {"magnitude": tbl_height, "unit": "PT"}}, "transform": {"scaleX": 1, "scaleY": 1, "translateX": 25, "translateY": tbl_top, "unit": "PT"}}, "rows": num_rows, "columns": num_cols}})

        # 1. 表頭合併
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0}, "rowSpan": 2, "columnSpan": 1}}})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 1}, "rowSpan": 1, "columnSpan": 3}}})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 4}, "rowSpan": 1, "columnSpan": 3}}})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 7}, "rowSpan": 1, "columnSpan": 3}}})

        # 2. 表頭底色與合計列底色
        self.requests.append({"updateTableCellProperties": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0}, "rowSpan": 2, "columnSpan": num_cols}, "tableCellProperties": {"tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": {"red": 0.15, "green": 0.25, "blue": 0.38}}}}}, "fields": "tableCellBackgroundFill"}})
        self.requests.append({"updateTableCellProperties": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 2, "columnIndex": 0}, "rowSpan": 1, "columnSpan": num_cols}, "tableCellProperties": {"tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": {"red": 0.91, "green": 0.94, "blue": 0.97}}}}}, "fields": "tableCellBackgroundFill"}})

        # 3. 寫入儲存格
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
                # 比較欄位負值標紅
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
        """【標準通用表格頁】"""
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

        self.requests.append({
            "createTable": {
                "objectId": table_id,
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {"width": {"magnitude": tbl_width, "unit": "PT"}, "height": {"magnitude": tbl_height, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": tbl_left, "translateY": tbl_top, "unit": "PT"}
                },
                "rows": num_rows,
                "columns": num_cols
            }
        })

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

    def wipe_old_slides(self):
        """抹除所有舊頁面，僅保留剛編譯完成的 15 頁"""
        for oid in self.old_slide_ids:
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
# 3. 數據準備層（✅ 全面鎖定資料截止日 115/09/15）
# ==========================================

DATA_CUTOFF_ROC = 1150915
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

# ── 1. 三項重點違規（母本規範矩陣，基準日 09/15，合計 144） ──
latest_three_day = f"{month:02d}/{day:02d}"

three_major_raw_matrix = [
    ["合計", 0, 0, 0, 0, 97, 35, 12, 144],
    ["聖亭所", 0, 0, 0, 0, 9, 4, 0, 13],
    ["龍潭所", 0, 0, 0, 0, 4, 0, 0, 4],
    ["中興所", 0, 0, 0, 0, 25, 0, 0, 25],
    ["石門所", 0, 0, 0, 0, 21, 1, 0, 22],
    ["高平所", 0, 0, 0, 0, 18, 1, 0, 19],
    ["三和所", 0, 0, 0, 0, 0, 0, 0, 0],
    ["交通分隊", 0, 0, 0, 0, 20, 29, 12, 61],
]

preview_cols = pd.MultiIndex.from_tuples([
    ("單位", ""),
    (f"本期 ({latest_three_day}) 新增違規數", "闖紅燈"),
    (f"本期 ({latest_three_day}) 新增違規數", "逆向行駛"),
    (f"本期 ({latest_three_day}) 新增違規數", "不停讓行人"),
    (f"本期 ({latest_three_day}) 新增違規數", f"本期合計 ({latest_three_day})"),
    ("115年9月1日起累計數", "闖紅燈"),
    ("115年9月1日起累計數", "逆向行駛"),
    ("115年9月1日起累計數", "不停讓行人"),
    ("115年9月1日起累計數", "累計總計")
])
df_three_preview = pd.DataFrame(three_major_raw_matrix, columns=preview_cols)

# ── 2. A1類交通事故死亡人數統計表 ──
df_a1 = pd.DataFrame([
    {"統計期間": "合計", "本期(0909-0915)": 0, "本年累計(0101-0915)": 1, "去年累計(0101-0915)": 6, "本年與去年同期比較": -5},
    {"統計期間": "聖亭所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 0, "本年與去年同期比較": 0},
    {"統計期間": "龍潭所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 1, "本年與去年同期比較": -1},
    {"統計期間": "中興所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 2, "本年與去年同期比較": -2},
    {"統計期間": "石門所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 2, "本年與去年同期比較": -2},
    {"統計期間": "高平所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 1, "去年累計(0101-0915)": 1, "本年與去年同期比較": 0},
    {"統計期間": "三和所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 0, "本年與去年同期比較": 0},
])

# ── 3. A2類交通事故受傷人數統計表 ──
df_a2 = pd.DataFrame([
    {"統計期間": "合計", "本期(0909-0915)": 25, "前期(0902-0908)": 22, "本年累計(0101-0915)": 1353, "去年累計(0101-0915)": 1532, "本年與去年同期比較": -179, "增減比例": "-11.68%"},
    {"統計期間": "聖亭所", "本期(0909-0915)": 5, "前期(0902-0908)": 4, "本年累計(0101-0915)": 282, "去年累計(0101-0915)": 276, "本年與去年同期比較": 6, "增減比例": "2.17%"},
    {"統計期間": "龍潭所", "本期(0909-0915)": 9, "前期(0902-0908)": 11, "本年累計(0101-0915)": 482, "去年累計(0101-0915)": 641, "本年與去年同期比較": -159, "增減比例": "-24.80%"},
    {"統計期間": "中興所", "本期(0909-0915)": 8, "前期(0902-0908)": 5, "本年累計(0101-0915)": 301, "去年累計(0101-0915)": 313, "本年與去年同期比較": -12, "增減比例": "-3.83%"},
    {"統計期間": "石門所", "本期(0909-0915)": 1, "前期(0902-0908)": 0, "本年累計(0101-0915)": 126, "去年累計(0101-0915)": 129, "本年與去年同期比較": -3, "增減比例": "-2.33%"},
    {"統計期間": "高平所", "本期(0909-0915)": 2, "前期(0902-0908)": 2, "本年累計(0101-0915)": 117, "去年累計(0101-0915)": 124, "本年與去年同期比較": -7, "增減比例": "-5.65%"},
    {"統計期間": "三和所", "本期(0909-0915)": 0, "前期(0902-0908)": 0, "本年累計(0101-0915)": 45, "去年累計(0101-0915)": 49, "本年與去年同期比較": -4, "增減比例": "-8.16%"},
])

# ── 4. 取締重大交通違規統計表 (總表) ──
df_major = pd.DataFrame([
    {"統計期間": "合計", "本期(攔停)": 39, "本期(逕舉)": 210, "本年累計(攔停)": 2317, "本年累計(逕舉)": 6852, "去年累計(攔停)": 2065, "去年累計(逕舉)": 7268, "本年與去年同期比較": -186, "目標值": 18114, "達成率": "50.6%"},
    {"統計期間": "科技執法", "本期(攔停)": 0, "本期(逕舉)": 56, "本年累計(攔停)": 9, "本年累計(逕舉)": 1535, "去年累計(攔停)": 4, "去年累計(逕舉)": 525, "本年與去年同期比較": 1015, "目標值": 6006, "達成率": "25.7%"},
    {"統計期間": "聖亭所", "本期(攔停)": 2, "本期(逕舉)": 11, "本年累計(攔停)": 150, "本年累計(逕舉)": 394, "去年累計(攔停)": 67, "去年累計(逕舉)": 1005, "本年與去年同期比較": -528, "目標值": 1941, "達成率": "28.0%"},
    {"統計期間": "龍潭所", "本期(攔停)": 13, "本期(逕舉)": 0, "本年累計(攔停)": 1309, "本年累計(逕舉)": 261, "去年累計(攔停)": 1201, "去年累計(逕舉)": 1075, "本年與去年同期比較": -706, "目標值": 2588, "達成率": "60.7%"},
    {"統計期間": "中興所", "本期(攔停)": 5, "本期(逕舉)": 21, "本年累計(攔停)": 323, "本年累計(逕舉)": 395, "去年累計(攔停)": 325, "去年累計(逕舉)": 705, "本年與去年同期比較": -312, "目標值": 1941, "達成率": "37.0%"},
    {"統計期間": "石門所", "本期(攔停)": 6, "本期(逕舉)": 21, "本年累計(攔停)": 220, "本年累計(逕舉)": 364, "去年累計(攔停)": 298, "去年累計(逕舉)": 439, "本年與去年同期比較": -153, "目標值": 1479, "達成率": "39.5%"},
    {"統計期間": "高平所", "本期(攔停)": 5, "本期(逕舉)": 19, "本年累計(攔停)": 141, "本年累計(逕舉)": 645, "去年累計(攔停)": 36, "去年累計(逕舉)": 697, "本年與去年同期比較": 53, "目標值": 1294, "達成率": "60.7%"},
    {"統計期間": "三和所", "本期(攔停)": 0, "本期(逕舉)": 0, "本年累計(攔停)": 9, "本年累計(逕舉)": 238, "去年累計(攔停)": 9, "去年累計(逕舉)": 165, "本年與去年同期比較": 73, "目標值": 339, "達成率": "72.9%"},
    {"統計期間": "警備隊", "本期(攔停)": 0, "本期(逕舉)": 0, "本年累計(攔停)": 0, "本年累計(逕舉)": 71, "去年累計(攔停)": 0, "去年累計(逕舉)": 49, "本年與去年同期比較": "—", "目標值": 0, "達成率": "—"},
    {"統計期間": "交通分隊", "本期(攔停)": 8, "本期(逕舉)": 82, "本年累計(攔停)": 156, "本年累計(逕舉)": 2949, "去年累計(攔停)": 125, "去年累計(逕舉)": 2608, "本年與去年同期比較": 372, "目標值": 2526, "達成率": "122.9%"},
])
major_footnote_exact = "重大交通違規指：「酒駕」、「闖紅燈」、「嚴重超速」、「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」及「不暫停讓行人」"

# ── 5. 重大交通違規 7 大專項細表矩陣（真實資料庫數據） ──
MAJOR_DETAIL_DICT = {
    "酒駕": [
        ["合計", 295, 1, 296, 141, 1, 142, 154, 0, 154],
        ["科技執法", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["聖亭所", 18, 0, 18, 8, 0, 8, 10, 0, 10],
        ["龍潭所", 109, 1, 110, 34, 1, 35, 75, 0, 75],
        ["中興所", 85, 0, 85, 29, 0, 29, 56, 0, 56],
        ["石門所", 10, 0, 10, 21, 0, 21, -11, 0, -11],
        ["高平所", 26, 0, 26, 8, 0, 8, 18, 0, 18],
        ["三和所", 0, 0, 0, 3, 0, 3, -3, 0, -3],
        ["警備隊", 0, 0, 0, 0, 0, 0, "—", "—", "—"],
        ["交通分隊", 47, 0, 47, 38, 0, 38, 9, 0, 9],
    ],
    "闖紅燈": [
        ["合計", 703, 3164, 3867, 490, 3790, 4280, 213, -629, -416],
        ["科技執法", 9, 589, 598, 1, 120, 121, 8, 469, 477],
        ["聖亭所", 57, 359, 416, 35, 896, 931, 22, -537, -515],
        ["龍潭所", 281, 160, 441, 214, 565, 779, 67, -405, -338],
        ["中興所", 189, 390, 579, 123, 633, 756, 66, -243, -177],
        ["石門所", 55, 303, 358, 71, 283, 354, -16, 20, 4],
        ["高平所", 53, 612, 665, 11, 638, 649, 42, -26, 16],
        ["三和所", 5, 97, 102, 1, 66, 67, 4, 31, 35],
        ["警備隊", 0, 51, 51, 0, 48, 48, "—", "—", "—"],
        ["交通分隊", 54, 603, 657, 34, 541, 575, 20, 62, 82],
    ],
    "逆向行駛": [
        ["合計", 252, 1375, 1627, 239, 1487, 1726, 13, -112, -99],
        ["科技執法", 0, 8, 8, 0, 7, 7, 0, 1, 1],
        ["聖亭所", 18, 31, 49, 14, 102, 116, 4, -71, -67],
        ["龍潭所", 165, 34, 199, 133, 298, 431, 32, -264, -232],
        ["中興所", 10, 0, 10, 13, 34, 47, -3, -34, -37],
        ["石門所", 24, 8, 32, 58, 46, 104, -34, -38, -72],
        ["高平所", 9, 11, 20, 2, 15, 17, 7, -4, 3],
        ["三和所", 1, 136, 137, 3, 76, 79, -2, 60, 58],
        ["警備隊", 0, 0, 0, 0, 0, 0, "—", "—", "—"],
        ["交通分隊", 25, 1147, 1172, 16, 909, 925, 9, 238, 247],
    ],
    "轉彎未依規定": [
        ["合計", 957, 1845, 2802, 1181, 1244, 2425, -224, 582, 358],
        ["科技執法", 0, 903, 903, 3, 203, 206, -3, 700, 697],
        ["聖亭所", 42, 0, 42, 9, 5, 14, 33, -5, 28],
        ["龍潭所", 680, 21, 701, 815, 202, 1017, -135, -181, -316],
        ["中興所", 31, 0, 31, 156, 36, 192, -125, -36, -161],
        ["石門所", 125, 41, 166, 146, 104, 250, -21, -63, -84],
        ["高平所", 52, 17, 69, 15, 39, 54, 37, -22, 15],
        ["三和所", 2, 0, 2, 2, 0, 2, 0, 0, 0],
        ["警備隊", 0, 20, 20, 0, 1, 1, "—", "—", "—"],
        ["交通分隊", 25, 843, 868, 35, 654, 689, -10, 189, 179],
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
        ["合計", 102, 343, 445, 14, 508, 522, 88, -165, -77],
        ["科技執法", 0, 35, 35, 0, 195, 195, 0, -160, -160],
        ["聖亭所", 15, 4, 19, 1, 2, 3, 14, 2, 16],
        ["龍潭所", 70, 39, 109, 5, 8, 13, 65, 31, 96],
        ["中興所", 7, 5, 12, 4, 2, 6, 3, 3, 6],
        ["石門所", 5, 12, 17, 2, 5, 7, 3, 7, 10],
        ["高平所", 1, 5, 6, 0, 4, 4, 1, 1, 2],
        ["三和所", 1, 5, 6, 0, 0, 0, 1, 5, 6],
        ["警備隊", 0, 0, 0, 0, 0, 0, "—", "—", "—"],
        ["交通分隊", 3, 238, 241, 2, 292, 294, 1, -54, -53],
    ],
    "嚴重超速": [
        ["合計", 0, 116, 116, 0, 227, 227, 0, -111, -111],
        ["科技執法", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["聖亭所", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["龍潭所", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["中興所", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["石門所", 0, 0, 0, 0, 0, 0, 0, 0, 0],
        ["高平所", 0, 0, 0, 0, 1, 1, 0, -1, -1],
        ["三和所", 0, 0, 0, 0, 23, 23, 0, -23, -23],
        ["警備隊", 0, 0, 0, 0, 0, 0, "—", "—", "—"],
        ["交通分隊", 0, 116, 116, 0, 203, 203, 0, -87, -87],
    ],
}

# ── 6. 取締超載違規件數統計表 ──
df_overload = pd.DataFrame([
    {"統計期間": "合計", "本期 (0909~0915)": 1, "本年累計 (0101~0915)": 46, "去年累計 (0101~0915)": 121, "本年與去年同期比較": -75, "目標值": 127, "達成率": "36%"},
    {"統計期間": "聖亭所", "本期 (0909~0915)": 0, "本年累計 (0101~0915)": 8, "去年累計 (0101~0915)": 12, "本年與去年同期比較": -4, "目標值": 20, "達成率": "40%"},
    {"統計期間": "龍潭所", "本期 (0909~0915)": 0, "本年累計 (0101~0915)": 3, "去年累計 (0101~0915)": 35, "本年與去年同期比較": -32, "目標值": 27, "達成率": "11%"},
    {"統計期間": "中興所", "本期 (0909~0915)": 0, "本年累計 (0101~0915)": 6, "去年累計 (0101~0915)": 12, "本年與去年同期比較": -6, "目標值": 20, "達成率": "30%"},
    {"統計期間": "石門所", "本期 (0909~0915)": 0, "本年累計 (0101~0915)": 1, "去年累計 (0101~0915)": 10, "本年與去年同期比較": -9, "目標值": 16, "達成率": "6%"},
    {"統計期間": "高平所", "本期 (0909~0915)": 1, "本年累計 (0101~0915)": 19, "去年累計 (0101~0915)": 39, "本年與去年同期比較": -20, "目標值": 14, "達成率": "136%"},
    {"統計期間": "三和所", "本期 (0909~0915)": 0, "本年累計 (0101~0915)": 3, "去年累計 (0101~0915)": 5, "本年與去年同期比較": -2, "目標值": 8, "達成率": "38%"},
    {"統計期間": "警備隊", "本期 (0909~0915)": 0, "本年累計 (0101~0915)": 0, "去年累計 (0101~0915)": 0, "本年與去年同期比較": 0, "目標值": 0, "達成率": "—"},
    {"統計期間": "交通分隊", "本期 (0909~0915)": 0, "本年累計 (0101~0915)": 6, "去年累計 (0101~0915)": 8, "本年與去年同期比較": -2, "目標值": 22, "達成率": "27%"},
])

# ── 7. 「靜桃計畫」大執法專案統計表 ──
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

# ── 8. 科技執法成效 ──
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
# 4. 前端預覽區
# ==========================================
with st.expander("👀 點擊展開預覽完整 15 頁業務表格內容"):
    t1, t2, t3, t4, t5, t6, t7, t8 = st.tabs(["三項重點", "A1事故死亡", "A2事故受傷", "重大違規(總表)", "重大違規(7大細項)", "超載取締", "靜桃計畫", "科技執法成效"])
    with t1: st.dataframe(df_three_preview, hide_index=True)
    with t2: st.dataframe(df_a1, hide_index=True)
    with t3: st.dataframe(df_a2, hide_index=True)
    with t4: st.dataframe(df_major, hide_index=True)
    with t5:
        sub_t = st.selectbox("選擇欲檢視的專項細表：", list(MAJOR_DETAIL_DICT.keys()))
        det_cols = pd.MultiIndex.from_tuples([
            ("統計期間", ""),
            ("今年累計", "當場攔停"), ("今年累計", "逕行舉發"), ("今年累計", "合計"),
            ("去年累計", "當場攔停"), ("去年累計", "逕行舉發"), ("去年累計", "合計"),
            ("今年與去年同期比較", "當場攔停"), ("今年與去年同期比較", "逕行舉發"), ("今年與去年同期比較", "合計")
        ])
        st.dataframe(pd.DataFrame(MAJOR_DETAIL_DICT[sub_t], columns=det_cols), hide_index=True)
    with t6: 
        st.dataframe(df_overload, hide_index=True)
        st.caption(f"📝 {overload_footnote_exact}")
    with t7: st.dataframe(df_jingtao, hide_index=True)
    with t8: st.dataframe(df_tech_final, hide_index=True)

st.write("")

# ==========================================
# 5. 執行生成
# ==========================================
if st.button("🚀 啟動畫布清空重繪：全新產出【原版規格 15 頁完整會報簡報】", type="primary"):
    slides_svc = get_slides_service()

    if not slides_svc:
        st.error("❌ 無法初始化 Google Slides 服務，請確認 secrets.toml 設定。")
    else:
        with st.spinner("正在清空母本畫布、動態編譯 15 頁原版規格投影片並整批覆蓋..."):
            try:
                builder = ComprehensiveSlidesBuilder(slides_svc, TARGET_PRESENTATION_ID)

                # 1. 記錄舊頁面 ID
                builder.prepare_canvas()

                # 2. P.1 封面頁
                builder.add_cover_slide(
                    main_title="桃園市政府警察局龍潭分局\n交通執法成效與事故防制數據分析報告",
                    subtitle="週次主管會報專案報告",
                    date_range_str=cover_date_str
                )

                # 3. P.2 取締三項重點違規統計表 (09/15)
                builder.add_three_major_slide(
                    data_rows=three_major_raw_matrix,
                    latest_day=latest_three_day
                )

                # 4. P.3 A1類交通事故死亡人數統計表
                builder.add_table_slide(
                    slide_title="A1類交通事故死亡人數統計表",
                    df=df_a1,
                    is_accident_table=True
                )

                # 5. P.4 A2類交通事故受傷人數統計表
                builder.add_table_slide(
                    slide_title="A2類交通事故受傷人數統計表",
                    df=df_a2,
                    is_accident_table=True
                )

                # 6. P.5 取締重大交通違規統計表 (總表)
                builder.add_table_slide(
                    slide_title="取締重大交通違規統計表",
                    df=df_major,
                    footnote=major_footnote_exact
                )

                # 7. P.6~P.12 重大交通違規 7 大專項細表
                for cat in ["酒駕", "闖紅燈", "逆向行駛", "轉彎未依規定", "蛇行惡意逼車", "不暫停讓行人", "嚴重超速"]:
                    builder.add_major_detail_slide(
                        cat_name=cat,
                        data_rows=MAJOR_DETAIL_DICT[cat],
                        date_str="0101-0915"
                    )

                # 8. P.13 取締超載違規件數統計表
                builder.add_table_slide(
                    slide_title="取締超載違規件數統計表",
                    df=df_overload,
                    footnote=overload_footnote_exact
                )

                # 9. P.14 「靜桃計畫」大執法專案統計表
                builder.add_table_slide(
                    slide_title="「靜桃計畫」大執法專案統計表",
                    df=df_jingtao
                )

                # 10. P.15 科技執法成效
                tech_slide_title = f"科技執法成效 ({tech_date_range_str})"
                builder.add_table_slide(
                    slide_title=tech_slide_title,
                    df=df_tech_final,
                    custom_width=480
                )

                # 11. 徹底清除原有舊頁面
                builder.wipe_old_slides()

                # 12. 整批傳送執行
                final_url = builder.execute_build()

                st.balloons()
                st.success("🎉 全套 15 頁 Google Slides 簡報已依原版完整規格全自動重繪完成！")
                st.markdown(
                    f"### 📑 簡報入口：\n"
                    f"👉 **[點此直接開啟全新會報簡報]({final_url})**\n\n"
                    f"✅ **15 頁原版規格全數歸位**：\n"
                    f"1. **P.6~P.12 重大違規 7 大細表完整收錄**：酒駕、闖紅燈、逆向、轉彎、逼車、禮讓行人、嚴重超速，均採原版雙層表頭排版。\n"
                    f"2. **全套數據與日期嚴格錨定 9/15**，超載法定達成率精準為 **70.7%**。\n"
                    f"3. 徹底排除空儲存格 400 報錯，一鍵產出完整、標準、無格式殘留的正式會報簡報！"
                )

            except HttpError as e:
                st.error(f"❌ Google API 請求失敗：{e}\n\n*提示：請確認簡報是否已共用給 `{SERVICE_ACCOUNT_EMAIL}` 並設定為「編輯者」。*")
            except Exception as e:
                st.error(f"❌ 建立簡報失敗：{e}")
