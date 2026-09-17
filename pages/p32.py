import io
import re
import uuid
import smtplib
from datetime import datetime, timedelta
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

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

st.title("📽️ 全方位執法數據簡報直出中心（自選頁面版）")
st.caption("🚀 自由勾選機制：可任意指定欲輸出的統計表，系統動態按需編譯並覆蓋目標簡報，並可自訂副本名稱與自動寄發通知郵件。")

# ==========================================
# 1. Google 服務連線層與常數設定
# ==========================================
GCP_CREDS = dict(st.secrets.get("gcp_service_account", {}))
SERVICE_ACCOUNT_EMAIL = GCP_CREDS.get("client_email", "streamlit-bot@streamlit-sheets-482909.iam.gserviceaccount.com")

TARGET_PRESENTATION_ID = "1h2QNNI8SLvjNEBmky7IWv9ZGBbKsLvV1UDkYJWcOeWU"
DRIVE_FOLDER_ID = st.secrets.get("DRIVE_FOLDER_ID", "1fm6ZK5B5wUmfy7-cgrw8OIkh7iS175dA").strip()

# 郵件 SMTP 設定（由 secrets.toml 提供）
SMTP_USER = st.secrets.get("SMTP_USER", "")
SMTP_PASSWORD = st.secrets.get("SMTP_PASSWORD", "")
DEFAULT_RECIPIENT = "mbmmiqamhnd@gmail.com"

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

def create_presentation_copy(drive_svc, source_file_id: str, title: str, folder_id: str = None) -> dict:
    """利用 Google Drive API 複製簡報並設定指定檔名與存放資料夾"""
    body = {"name": title}
    if folder_id:
        body["parents"] = [folder_id]

    copied_file = drive_svc.files().copy(
        fileId=source_file_id,
        body=body,
        fields="id, name, webViewLink"
    ).execute()
    return copied_file

def send_report_email(to_email: str, subject: str, body_html: str, from_email: str = None) -> tuple:
    """使用 Gmail SMTP 寄送通報郵件"""
    sender = from_email or SMTP_USER
    if not sender or not SMTP_PASSWORD:
        return False, "未於 secrets.toml 設定 SMTP_USER 或 SMTP_PASSWORD"

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = Header(subject, "utf-8")
        msg["From"] = f"交通戰情室自動通報 <{sender}>"
        msg["To"] = to_email

        msg.attach(MIMEText(body_html, "html", "utf-8"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, SMTP_PASSWORD)
            server.sendmail(sender, [to_email], msg.as_string())
        return True, "寄送成功"
    except Exception as e:
        return False, str(e)

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

    def prepare_canvas(self, protect_first_slide: bool = False):
        pres = self.slides_svc.presentations().get(
            presentationId=self.presentation_id
        ).execute()
        slides = pres.get("slides", [])
        if protect_first_slide and slides:
            self.old_slide_ids = [s["objectId"] for s in slides[1:]]
        else:
            self.old_slide_ids = [s["objectId"] for s in slides]

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

    def add_three_major_slide(self, data_rows, latest_day="09/15"):
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

        self.requests.append({"createTable": {"objectId": table_id, "elementProperties": {"pageObjectId": slide_id, "size": {"width": {"magnitude": 670, "unit": "PT"}, "height": {"magnitude": tbl_height, "unit": "PT"}}, "transform": {"scaleX": 1, "scaleY": 1, "translateX": 25, "translateY": tbl_top, "unit": "PT"}}, "rows": num_rows, "columns": num_cols}})
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

    def add_major_detail_slide(self, cat_name: str, data_rows, date_str="0101-0915"):
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

        self.requests.append({"createTable": {"objectId": table_id, "elementProperties": {"pageObjectId": slide_id, "size": {"width": {"magnitude": tbl_width, "unit": "PT"}, "height": {"magnitude": tbl_height, "unit": "PT"}}, "transform": {"scaleX": 1, "scaleY": 1, "translateX": tbl_left, "translateY": tbl_top, "unit": "PT"}}, "rows": num_rows, "columns": num_cols}})
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
# 3. 數據準備層（鎖定資料截止日 115/09/15）
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

# 1. 三項重點違規
latest_three_day = "09/16"
three_major_raw_matrix = [
    ["合計", 5, 0, 1, 6, 213, 111, 25, 349],
    ["聖亭所", 2, 0, 0, 2, 15, 5, 0, 20],
    ["龍潭所", 1, 0, 1, 2, 11, 0, 1, 12],
    ["中興所", 0, 0, 0, 0, 38, 0, 0, 38],
    ["石門所", 2, 0, 0, 2, 40, 1, 0, 41],
    ["高平所", 0, 0, 0, 0, 28, 1, 0, 29],
    ["三和所", 0, 0, 0, 0, 32, 34, 0, 66],
    ["交通分隊", 0, 0, 0, 0, 49, 70, 24, 143],
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

# 2. A1 死亡
df_a1 = pd.DataFrame([
    {"統計期間": "合計", "本期(0909-0915)": 0, "本年累計(0101-0915)": 1, "去年累計(0101-0915)": 6, "本年與去年同期比較": -5},
    {"統計期間": "聖亭所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 0, "本年與去年同期比較": 0},
    {"統計期間": "龍潭所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 1, "本年與去年同期比較": -1},
    {"統計期間": "中興所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 2, "本年與去年同期比較": -2},
    {"統計期間": "石門所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 2, "本年與去年同期比較": -2},
    {"統計期間": "高平所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 1, "去年累計(0101-0915)": 1, "本年與去年同期比較": 0},
    {"統計期間": "三和所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 0, "本年與去年同期比較": 0},
])

# 3. A2 受傷
df_a2 = pd.DataFrame([
    {"統計期間": "合計", "本期(0909-0915)": 25, "前期(0902-0908)": 22, "本年累計(0101-0915)": 1353, "去年累計(0101-0915)": 1532, "本年與去年同期比較": -179, "增減比例": "-11.68%"},
    {"統計期間": "聖亭所", "本期(0909-0915)": 5, "前期(0902-0908)": 4, "本年累計(0101-0915)": 282, "去年累計(0101-0915)": 276, "本年與去年同期比較": 6, "增減比例": "2.17%"},
    {"統計期間": "龍潭所", "本期(0909-0915)": 9, "前期(0902-0908)": 11, "本年累計(0101-0915)": 482, "去年累計(0101-0915)": 641, "本年與去年同期比較": -159, "增減比例": "-24.80%"},
    {"統計期間": "中興所", "本期(0909-0915)": 8, "前期(0902-0908)": 5, "本年累計(0101-0915)": 301, "去年累計(0101-0915)": 313, "本年與去年同期比較": -12, "增減比例": "-3.83%"},
    {"統計期間": "石門所", "本期(0909-0915)": 1, "前期(0902-0908)": 0, "本年累計(0101-0915)": 126, "去年累計(0101-0915)": 129, "本年與去年同期比較": -3, "增減比例": "-2.33%"},
    {"統計期間": "高平所", "本期(0909-0915)": 2, "前期(0902-0908)": 2, "本年累計(0101-0915)": 117, "去年累計(0101-0915)": 124, "本年與去年同期比較": -7, "增減比例": "-5.65%"},
    {"統計期間": "三和所", "本期(0909-0915)": 0, "前期(0902-0908)": 0, "本年累計(0101-0915)": 45, "去年累計(0101-0915)": 49, "本年與去年同期比較": -4, "增減比例": "-8.16%"},
])

# 4. 重大違規總表
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

# 5. 重大違規 7 大專項細表
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

# 6. 超載
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

# 7. 靜桃
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

# 8. 科技執法
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

# 快捷選擇按鈕
col_btn1, col_btn2, _ = st.columns([1.5, 2, 4])
if "select_mode" not in st.session_state:
    st.session_state["select_mode"] = "core"

with col_btn1:
    if st.button("📌 僅常態核心頁 (8頁)"):
        st.session_state["select_mode"] = "core"
with col_btn2:
    if st.button("📑 全選所有統計表 (15頁)"):
        st.session_state["select_mode"] = "all"

is_all = (st.session_state["select_mode"] == "all")

# 區塊勾選選項
col_opt1, col_opt2 = st.columns(2)

with col_opt1:
    st.markdown("##### 🏢 常態會報核心表格")

    chk_protect_cover = st.checkbox(
        "🔒 保留現有封面（手動編輯過，不覆寫/不刪除）",
        value=False,
        help="勾選後，Google 簡報目前的第1頁會被完整保留（不刪除、不重繪），適合您已手動排版過封面的情況。"
    )
    chk_cover = st.checkbox(
        "P.1 簡報封面（自動產生，套用固定樣式）",
        value=True,
        disabled=chk_protect_cover
    )

    chk_three = st.checkbox("P.2 取締三項重點違規統計表 (母本雙層)", value=True)
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

# 計算總勾選頁數
selected_pages = []
if chk_protect_cover:
    selected_pages.append("封面(保留手動版本)")
elif chk_cover:
    selected_pages.append("封面")
if chk_three: selected_pages.append("三項重點")
if chk_a1: selected_pages.append("A1事故死亡")
if chk_a2: selected_pages.append("A2事故受傷")
if chk_major_tot: selected_pages.append("重大違規總表")
if chk_det_jiu: selected_pages.append("細表-酒駕")
if chk_det_red: selected_pages.append("細表-闖紅燈")
if chk_det_rev: selected_pages.append("細表-逆向")
if chk_det_turn: selected_pages.append("細表-轉彎")
if chk_det_snake: selected_pages.append("細表-逼車")
if chk_det_ped: selected_pages.append("細表-讓行人")
if chk_det_speed: selected_pages.append("細表-嚴重超速")
if chk_overload: selected_pages.append("超載統計")
if chk_jingtao: selected_pages.append("靜桃計畫")
if chk_tech: selected_pages.append("科技執法")

if chk_protect_cover:
    st.caption(f"📊 目前共勾選 **{len(selected_pages)}** 個頁面待編譯輸出（封面將維持現況，不重新產生）。")
else:
    st.caption(f"📊 目前共勾選 **{len(selected_pages)}** 個頁面待編譯輸出。")

# ==========================================
# 4.1 副本命名與郵件寄送自訂配置
# ==========================================
st.markdown("---")
st.markdown("#### 📁 副本命名與郵件寄送設定")
col_cfg1, col_cfg2 = st.columns([3, 2])

default_copy_title = f"龍潭分局執法數據簡報_{datetime.now().strftime('%Y%m%d_%H%M')}"

with col_cfg1:
    custom_copy_title = st.text_input(
        "✏️ 自訂副本簡報名稱（可隨意修改）：",
        value=default_copy_title,
        help="產出後系統將以此名稱複製一份獨立的 Google 簡報存檔於資料夾中。"
    )

with col_cfg2:
    recipient_email = st.text_input(
        "✉️ 接收副本連結信箱：",
        value=DEFAULT_RECIPIENT,
        help="產出完成後將自動寄出包含此副本超連結的通報信件。"
    )

with st.expander("👀 點擊展開預覽待輸出業務數據"):
    t1, t2, t3, t4, t5, t6, t7 = st.tabs(["三項重點", "A1事故死亡", "A2事故受傷", "重大違規", "超載取締", "靜桃計畫", "科技執法成效"])
    with t1: st.dataframe(df_three_preview, hide_index=True)
    with t2: st.dataframe(df_a1, hide_index=True)
    with t3: st.dataframe(df_a2, hide_index=True)
    with t4: st.dataframe(df_major, hide_index=True)
    with t5:
        st.dataframe(df_overload, hide_index=True)
        st.caption(f"📝 {overload_footnote_exact}")
    with t6: st.dataframe(df_jingtao, hide_index=True)
    with t7: st.dataframe(df_tech_final, hide_index=True)

st.write("")

# ==========================================
# 5. 執行指定輸出生成
# ==========================================
btn_label = f"🚀 立即編譯產出【已選定的 {len(selected_pages)} 個統計表頁面】"

if st.button(btn_label, type="primary"):
    if not selected_pages:
        st.warning("⚠️ 請至少勾選一個統計表頁面！")
    elif not custom_copy_title.strip():
        st.warning("⚠️ 副本簡報名稱不可為空白！")
    else:
        slides_svc = get_slides_service()
        drive_svc = get_drive_service()

        if not slides_svc or not drive_svc:
            st.error("❌ 無法初始化 Google 服務，請確認 secrets.toml 設定。")
        else:
            spinner_msg = f"正在{'（保留現有封面）' if chk_protect_cover else '清空母本畫布、'}動態編譯已勾選的 {len(selected_pages)} 頁投影片..."
            with st.spinner(spinner_msg):
                try:
                    builder = ComprehensiveSlidesBuilder(slides_svc, TARGET_PRESENTATION_ID)

                    # 1. 記錄舊頁面 ID（稍後整批抹除；若保留封面，第1頁不在此清單）
                    builder.prepare_canvas(protect_first_slide=chk_protect_cover)

                    # 2. 依勾選順序動態注入
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
                        builder.add_table_slide(
                            slide_title="A1類交通事故死亡人數統計表",
                            df=df_a1,
                            is_accident_table=True
                        )

                    if chk_a2:
                        builder.add_table_slide(
                            slide_title="A2類交通事故受傷人數統計表",
                            df=df_a2,
                            is_accident_table=True
                        )

                    if chk_major_tot:
                        builder.add_table_slide(
                            slide_title="取締重大交通違規統計表",
                            df=df_major,
                            footnote=major_footnote_exact
                        )

                    # 專項細表
                    det_map = [
                        (chk_det_jiu, "酒駕"), (chk_det_red, "闖紅燈"), (chk_det_rev, "逆向行駛"),
                        (chk_det_turn, "轉彎未依規定"), (chk_det_snake, "蛇行惡意逼車"),
                        (chk_det_ped, "不暫停讓行人"), (chk_det_speed, "嚴重超速")
                    ]
                    for is_chk, cat in det_map:
                        if is_chk:
                            builder.add_major_detail_slide(
                                cat_name=cat,
                                data_rows=MAJOR_DETAIL_DICT[cat],
                                date_str="0101-0915"
                            )

                    if chk_overload:
                        builder.add_table_slide(
                            slide_title="取締超載違規件數統計表",
                            df=df_overload,
                            footnote=overload_footnote_exact
                        )

                    if chk_jingtao:
                        builder.add_table_slide(
                            slide_title="「靜桃計畫」大執法專案統計表",
                            df=df_jingtao
                        )

                    if chk_tech:
                        tech_slide_title = f"科技執法成效 ({tech_date_range_str})"
                        builder.add_table_slide(
                            slide_title=tech_slide_title,
                            df=df_tech_final,
                            custom_width=480
                        )

                    # 3. 抹除舊頁面
                    builder.wipe_old_slides()

                    # 4. 整批送出更新目標母本
                    final_url = builder.execute_build()

                    # 5. 建立自訂檔名之獨立副本並移入資料夾
                    final_copy_title = custom_copy_title.strip()
                    copied_info = create_presentation_copy(
                        drive_svc=drive_svc,
                        source_file_id=TARGET_PRESENTATION_ID,
                        title=final_copy_title,
                        folder_id=DRIVE_FOLDER_ID
                    )
                    copy_id = copied_info.get("id")
                    copy_url = copied_info.get("webViewLink", f"https://docs.google.com/presentation/d/{copy_id}/edit")

                    # 6. 自動寄送通知信件
                    email_html = f"""
                    <div style="font-family: Arial, 'Microsoft JhengHei', sans-serif; line-height: 1.6; color: #333; max-width: 600px; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;">
                        <h2 style="color: #1e3a5f; margin-top: 0;">📊 龍潭分局執法數據簡報產出通報</h2>
                        <p>您好，全方位執法數據簡報已自動編譯完成，並已產生指定檔名之獨立副本存檔：</p>
                        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
                            <tr><td style="padding: 6px 0; color: #666; width: 100px;"><b>產出時間：</b></td><td>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</td></tr>
                            <tr><td style="padding: 6px 0; color: #666;"><b>副本檔名：</b></td><td><b>{final_copy_title}</b></td></tr>
                            <tr><td style="padding: 6px 0; color: #666;"><b>包含頁數：</b></td><td>共 {len(selected_pages)} 個指定頁面</td></tr>
                        </table>
                        <div style="margin: 25px 0;">
                            <a href="{copy_url}" style="background-color: #0b57d0; color: #ffffff; padding: 12px 22px; text-decoration: none; border-radius: 6px; font-weight: bold; display: inline-block;">
                                📑 開啟當期簡報副本
                            </a>
                            &nbsp;&nbsp;
                            <a href="{final_url}" style="background-color: #f1f3f4; color: #1f1f1f; padding: 12px 18px; text-decoration: none; border-radius: 6px; display: inline-block;">
                                📂 檢視母本
                            </a>
                        </div>
                        <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
                        <small style="color: #888;">此郵件由 Streamlit 龍潭分局執法數據簡報直出系統自動發送。</small>
                    </div>
                    """

                    mail_success, mail_msg = False, ""
                    target_email = recipient_email.strip()
                    if target_email:
                        mail_success, mail_msg = send_report_email(
                            to_email=target_email,
                            subject=f"【簡報副本通報】{final_copy_title}",
                            body_html=email_html
                        )

                    st.balloons()
                    st.success(f"🎉 指定的 {len(selected_pages)} 個統計表頁面已重繪完成，並已成功建立自訂副本！")

                    if mail_success:
                        st.info(f"📧 副本簡報連結已同步發送至信箱：`{target_email}`")
                    elif target_email and (not SMTP_USER or not SMTP_PASSWORD):
                        st.warning("⚠️ 副本已建立，但尚未設定 secrets 中的 SMTP 帳密，暫未寄發通知信。若需寄信請在 `.streamlit/secrets.toml` 配置 SMTP_USER 與 SMTP_PASSWORD。")
                    elif target_email and not mail_success:
                        st.warning(f"⚠️ 信件寄送未成功：{mail_msg}")

                    protect_note = "（第1頁封面已依設定保留，未受影響）\n\n" if chk_protect_cover else ""
                    st.markdown(
                        f"### 📑 簡報入口連結：\n"
                        f"- 📎 **[點此直接開啟全新獨立副本：{final_copy_title}]({copy_url})**\n"
                        f"- 👉 **[開啟目標母本簡報]({final_url})**\n\n"
                        f"{protect_note}"
                    )

                except HttpError as e:
                    st.error(f"❌ Google API 請求失敗：{e}\n\n*提示：請確認簡報與雲端硬碟資料夾是否已共用給 `{SERVICE_ACCOUNT_EMAIL}` 並設定為「編輯者」。*")
                except Exception as e:
                    st.error(f"❌ 建立簡報失敗：{e}")
