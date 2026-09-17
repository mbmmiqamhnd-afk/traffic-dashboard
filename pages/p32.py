import io
import re
import uuid
from datetime import datetime
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

st.title("📽️ 全方位執法數據簡報直出中心")
st.caption("支援三項重點違規動態上傳核算、手動封面保留與指定頁面匯出。")

# ==========================================
# 1. Google 服務連線層與常數設定
# ==========================================
GCP_CREDS = dict(st.secrets.get("gcp_service_account", {}))
SERVICE_ACCOUNT_EMAIL = GCP_CREDS.get("client_email", "streamlit-bot@streamlit-sheets-482909.iam.gserviceaccount.com")

TARGET_PRESENTATION_ID = "1h2QNNI8SLvjNEBmky7IWv9ZGBbKsLvV1UDkYJWcOeWU"
DRIVE_FOLDER_ID = str(st.secrets.get("DRIVE_FOLDER_ID", "1fm6ZK5B5wUmfy7-cgrw8OIkh7iS175dA")).strip()

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
    f_id = str(folder_id).strip().replace('"', '').replace("'", '')
    try:
        results = service.files().list(
            q=f"'{f_id}' in parents and trashed = false",
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

    def add_three_major_slide(self, data_rows, latest_day="09/16"):
        slide_id = f"s_three_{uuid.uuid4().hex[:8]}"
        title_id = f"t_three_{uuid.uuid4().hex[:8]}"
        table_id = f"tbl_three_{uuid.uuid4().hex[:8]}"

        self.requests.append({"createSlide": {"objectId": slide_id, "slideLayoutReference": {"predefinedLayout": "BLANK"}}})

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
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0}, "rowSpan": 2, "columnSpan": 1}}})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 1}, "rowSpan": 1, "columnSpan": 4}}})
        self.requests.append({"mergeTableCells": {"objectId": table_id, "tableRange": {"location": {"rowIndex": 0, "columnIndex": 5}, "rowSpan": 1, "columnSpan": 4}}})

        self.requests.append({
            "updateTableCellProperties": {
                "objectId": table_id,
                "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0}, "rowSpan": 2, "columnSpan": num_cols},
                "tableCellProperties": {
                    "tableCellBackgroundFill": {
                        "solidFill": {"color": {"rgbColor": {"red": 0.15, "green": 0.25, "blue": 0.38}}}
                    }
                },
                "fields": "tableCellBackgroundFill"
            }
        })
        self.requests.append({
            "updateTableCellProperties": {
                "objectId": table_id,
                "tableRange": {"location": {"rowIndex": 2, "columnIndex": 0}, "rowSpan": 1, "columnSpan": num_cols},
                "tableCellProperties": {
                    "tableCellBackgroundFill": {
                        "solidFill": {"color": {"rgbColor": {"red": 0.91, "green": 0.94, "blue": 0.97}}}
                    }
                },
                "fields": "tableCellBackgroundFill"
            }
        })

        def write_cell(r, c, text, font_size=10.0, bold=False, fg=(0.1, 0.1, 0.1)):
            t_str = str(text).strip() if (pd.notna(text) and str(text).strip() != "") else "0"
            self.requests.append({
                "insertText": {
                    "objectId": table_id,
                    "cellLocation": {"rowIndex": r, "columnIndex": c},
                    "text": t_str,
                    "insertionIndex": 0
                }
            })
            self.requests.append({
                "updateTextStyle": {
                    "objectId": table_id,
                    "cellLocation": {"rowIndex": r, "columnIndex": c},
                    "style": {
                        "fontFamily": "Microsoft JhengHei",
                        "fontSize": {"magnitude": font_size, "unit": "PT"},
                        "bold": bold,
                        "foregroundColor": {"opaqueColor": {"rgbColor": {"red": fg[0], "green": fg[1], "blue": fg[2]}}}
                    },
                    "textRange": {"type": "ALL"},
                    "fields": "fontFamily,fontSize,bold,foregroundColor"
                }
            })

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

    def add_table_slide(self, slide_title: str, df: pd.DataFrame, subtitle: str = "", footnote: str = "", is_accident_table: bool = False, custom_width: int = None):
        slide_id = f"s_{uuid.uuid4().hex[:8]}"
        title_id = f"t_{uuid.uuid4().hex[:8]}"
        table_id = f"tbl_{uuid.uuid4().hex[:8]}"

        self.requests.append({"createSlide": {"objectId": slide_id, "slideLayoutReference": {"predefinedLayout": "BLANK"}}})
        full_title = f"{slide_title} | {subtitle}" if subtitle else slide_title
        self.requests.append({
            "createShape": {
                "objectId": title_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {"width": {"magnitude": 670, "unit": "PT"}, "height": {"magnitude": 35, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": 25, "translateY": 15, "unit": "PT"}
                }
            }
        })
        self.requests.append({"insertText": {"objectId": title_id, "text": full_title, "insertionIndex": 0}})
        self.requests.append({
            "updateTextStyle": {
                "objectId": title_id,
                "style": {
                    "fontFamily": "Microsoft JhengHei",
                    "fontSize": {"magnitude": 15, "unit": "PT"},
                    "bold": True,
                    "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.1, "green": 0.2, "blue": 0.35}}}
                },
                "textRange": {"type": "ALL"},
                "fields": "fontFamily,fontSize,bold,foregroundColor"
            }
        })

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
        self.requests.append({
            "updateTableCellProperties": {
                "objectId": table_id,
                "tableRange": {"location": {"rowIndex": 0, "columnIndex": 0}, "rowSpan": 1, "columnSpan": num_cols},
                "tableCellProperties": {
                    "tableCellBackgroundFill": {
                        "solidFill": {"color": {"rgbColor": {"red": 0.15, "green": 0.25, "blue": 0.38}}}
                    }
                },
                "fields": "tableCellBackgroundFill"
            }
        })

        def write_gen_cell(r, c, text, font_sz, bold, fg_rgb):
            t_str = str(text).replace("\n", " ").strip() if (pd.notna(text) and str(text).strip() != "") else "—"
            self.requests.append({
                "insertText": {
                    "objectId": table_id,
                    "cellLocation": {"rowIndex": r, "columnIndex": c},
                    "text": t_str,
                    "insertionIndex": 0
                }
            })
            self.requests.append({
                "updateTextStyle": {
                    "objectId": table_id,
                    "cellLocation": {"rowIndex": r, "columnIndex": c},
                    "style": {
                        "fontFamily": "Microsoft JhengHei",
                        "fontSize": {"magnitude": font_sz, "unit": "PT"},
                        "bold": bold,
                        "foregroundColor": {"opaqueColor": {"rgbColor": {"red": fg_rgb[0], "green": fg_rgb[1], "blue": fg_rgb[2]}}}
                    },
                    "textRange": {"type": "ALL"},
                    "fields": "fontFamily,fontSize,bold,foregroundColor"
                }
            })

        for c_idx, col_name in enumerate(df.columns):
            write_gen_cell(0, c_idx, col_name, font_size, True, (1.0, 1.0, 1.0))

        for r_idx, row in df.iterrows():
            first_col_val = str(row.values[0]).strip()
            is_hl_row = any(k in first_col_val for k in ["合計", "總計", "舉發總數"])

            if is_hl_row:
                self.requests.append({
                    "updateTableCellProperties": {
                        "objectId": table_id,
                        "tableRange": {"location": {"rowIndex": r_idx + 1, "columnIndex": 0}, "rowSpan": 1, "columnSpan": num_cols},
                        "tableCellProperties": {
                            "tableCellBackgroundFill": {
                                "solidFill": {"color": {"rgbColor": {"red": 0.91, "green": 0.94, "blue": 0.97}}}
                            }
                        },
                        "fields": "tableCellBackgroundFill"
                    }
                })

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
            self.requests.append({
                "createShape": {
                    "objectId": fn_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {"width": {"magnitude": 670, "unit": "PT"}, "height": {"magnitude": 30, "unit": "PT"}},
                        "transform": {"scaleX": 1, "scaleY": 1, "translateX": 25, "translateY": 365, "unit": "PT"}
                    }
                }
            })
            self.requests.append({"insertText": {"objectId": fn_id, "text": footnote, "insertionIndex": 0}})
            self.requests.append({
                "updateTextStyle": {
                    "objectId": fn_id,
                    "style": {
                        "fontFamily": "DFKai-SB",
                        "fontSize": {"magnitude": 10, "unit": "PT"},
                        "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.2, "green": 0.2, "blue": 0.2}}}
                    },
                    "textRange": {"type": "ALL"},
                    "fields": "fontFamily,fontSize,foregroundColor"
                }
            })

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
# 3. 三項重點違規：解析邏輯
# ==========================================
UNIT_ORDER = ["聖亭所", "龍潭所", "中興所", "石門所", "高平所", "三和所", "交通分隊"]
UNIT_MAP = {
    "聖亭": "聖亭所", "龍潭": "龍潭所", "中興": "中興所",
    "石門": "石門所", "高平": "高平所", "三和": "三和所",
    "分隊": "交通分隊"
}

def parse_three_major_file(f_obj):
    counts = {u: {"闖紅燈": 0, "逆向": 0, "行人": 0} for u in UNIT_ORDER}
    if not f_obj:
        return counts
    try:
        f_obj.seek(0)
        df = pd.read_excel(f_obj) if f_obj.name.endswith(('.xlsx', '.xls')) else pd.read_csv(f_obj, encoding="cp950")
        df.columns = [str(c).strip() for c in df.columns]

        has_summary = any("闖紅燈" in c for c in df.columns) and any("逆向" in c for c in df.columns)
        if has_summary:
            u_col = next((c for c in df.columns if any(k in c for k in ["單位", "所別", "隊別"])), df.columns[0])
            red_c = next((c for c in df.columns if "闖紅燈" in c), None)
            rev_c = next((c for c in df.columns if "逆向" in c), None)
            ped_c = next((c for c in df.columns if any(k in c for k in ["行人", "車不讓"])), None)

            for _, r in df.iterrows():
                u_str = str(r[u_col]).strip()
                target_u = next((v for k, v in UNIT_MAP.items() if k in u_str), None)
                if target_u and target_u in UNIT_ORDER and "科技" not in u_str and "警備" not in u_str:
                    counts[target_u]["闖紅燈"] += int(pd.to_numeric(r[red_c], errors="coerce") or 0)
                    counts[target_u]["逆向"] += int(pd.to_numeric(r[rev_c], errors="coerce") or 0)
                    counts[target_u]["行人"] += int(pd.to_numeric(r[ped_c], errors="coerce") or 0)
            return counts

        u_col = next((c for c in df.columns if any(k in c for k in ["單位", "所別", "隊別", "局署"])), None)
        f_col = next((c for c in df.columns if any(k in c for k in ["違規事實", "法條", "項目", "條款", "案由"])), None)

        if u_col and f_col:
            for _, r in df.iterrows():
                u_str = str(r[u_col]).strip()
                if "科技" in u_str or "警備" in u_str:
                    continue
                target_u = next((v for k, v in UNIT_MAP.items() if k in u_str), None)
                if target_u and target_u in UNIT_ORDER:
                    fact = str(r[f_col])
                    if any(k in fact for k in ["53條1項", "5310001", "5310002"]) or ("闖紅燈" in fact and "右轉" not in fact):
                        counts[target_u]["闖紅燈"] += 1
                    elif any(k in fact for k in ["45條1項1款", "45條1項3款", "4510101", "4510301", "逆向"]):
                        counts[target_u]["逆向"] += 1
                    elif any(k in fact for k in ["44條2項", "44條4項", "4420002", "4420003", "4420004", "不停讓行人", "車不讓"]):
                        counts[target_u]["行人"] += 1
    except Exception as e:
        st.warning(f"檔案解析提醒：{e}")
    return counts

# ==========================================
# 4. 前端資料輸入與即時運算
# ==========================================
st.markdown("### 📥 三項重點違規：來源表動態核算")

col_up1, col_up2 = st.columns(2)
with col_up1:
    up_wk = st.file_uploader("📂 1. 上傳【本期來源表】", type=["xlsx", "xls", "csv"], key="up_wk")
with col_up2:
    up_cumu = st.file_uploader("📂 2. 上傳【累計來源表】", type=["xlsx", "xls", "csv"], key="up_cumu")

drive_files = []
if not up_wk or not up_cumu:
    drive_files = fetch_files_from_drive(DRIVE_FOLDER_ID)

file_wk_obj = up_wk or next((f for f in drive_files if any(k in f.name for k in ["本期", "0916", "日報"])), None)
file_cm_obj = up_cumu or next((f for f in drive_files if any(k in f.name for k in ["累計", "0901", "9月"])), None)

wk_counts = parse_three_major_file(file_wk_obj)
cm_counts = parse_three_major_file(file_cm_obj)

has_dynamic_data = any(sum(d.values()) > 0 for d in cm_counts.values())

if has_dynamic_data:
    st.success(f"✅ 成功動態解析最新數據！本期檔案：`{file_wk_obj.name if file_wk_obj else '無'}` ｜ 累計檔案：`{file_cm_obj.name if file_cm_obj else '無'}`")
else:
    st.info("ℹ️ 尚未偵測到 9/16 最新來源表，目前顯示基準數據（上傳後將即時重算）。")

three_matrix = []
tot_wk_r = sum(wk_counts[u]["闖紅燈"] for u in UNIT_ORDER)
tot_wk_v = sum(wk_counts[u]["逆向"] for u in UNIT_ORDER)
tot_wk_p = sum(wk_counts[u]["行人"] for u in UNIT_ORDER)
tot_wk_s = tot_wk_r + tot_wk_v + tot_wk_p

tot_cm_r = sum(cm_counts[u]["闖紅燈"] for u in UNIT_ORDER)
tot_cm_v = sum(cm_counts[u]["逆向"] for u in UNIT_ORDER)
tot_cm_p = sum(cm_counts[u]["行人"] for u in UNIT_ORDER)
tot_cm_s = tot_cm_r + tot_cm_v + tot_cm_p

if has_dynamic_data:
    three_matrix.append(["合計", tot_wk_r, tot_wk_v, tot_wk_p, tot_wk_s, tot_cm_r, tot_cm_v, tot_cm_p, tot_cm_s])
    for u in UNIT_ORDER:
        w_r, w_v, w_p = wk_counts[u]["闖紅燈"], wk_counts[u]["逆向"], wk_counts[u]["行人"]
        c_r, c_v, c_p = cm_counts[u]["闖紅燈"], cm_counts[u]["逆向"], cm_counts[u]["行人"]
        three_matrix.append([u, w_r, w_v, w_p, w_r + w_v + w_p, c_r, c_v, c_p, c_r + c_v + c_p])
else:
    three_matrix = [
        ["合計", 0, 0, 0, 0, 97, 35, 12, 144],
        ["聖亭所", 0, 0, 0, 0, 9, 4, 0, 13],
        ["龍潭所", 0, 0, 0, 0, 4, 0, 0, 4],
        ["中興所", 0, 0, 0, 0, 25, 0, 0, 25],
        ["石門所", 0, 0, 0, 0, 21, 1, 0, 22],
        ["高平所", 0, 0, 0, 0, 18, 1, 0, 19],
        ["三和所", 0, 0, 0, 0, 0, 0, 0, 0],
        ["交通分隊", 0, 0, 0, 0, 20, 29, 12, 61],
    ]

cover_date_str = "115 年 9 月 1 日起至 09月16日 止"
tech_date_range_str = "115年1月1日至115年9月16日"

# 115/09/16 第 259 天 -> 70.8%
current_expected_rate = (259 / 366) * 100
overload_footnote_exact = f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，統計截至 115年09月16日 (入案日期)應達成率為{current_expected_rate:.1f}%"

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

preview_cols = pd.MultiIndex.from_tuples([
    ("單位", ""),
    ("本期 (09/16) 新增違規數", "闖紅燈"),
    ("本期 (09/16) 新增違規數", "逆向行駛"),
    ("本期 (09/16) 新增違規數", "不停讓行人"),
    ("本期 (09/16) 新增違規數", "本期合計 (09/16)"),
    ("115年9月1日起累計數", "闖紅燈"),
    ("115年9月1日起累計數", "逆向行駛"),
    ("115年9月1日起累計數", "不停讓行人"),
    ("115年9月1日起累計數", "累計總計")
])
df_three_preview = pd.DataFrame(three_matrix, columns=preview_cols)

# ==========================================
# 5. 前端自選與即時預覽區
# ==========================================
st.subheader("🎯 欲輸出的統計表自選控制")

col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    st.markdown("##### 🏢 常態會報核心表格")
    chk_protect_cover = st.checkbox("🔒 保留現有封面（手動編輯過，不覆寫/不刪除）", value=False)
    chk_cover = st.checkbox("P.1 簡報封面（自動產生）", value=True, disabled=chk_protect_cover)
    chk_three = st.checkbox("P.2 取締三項重點違規統計表 (即時動態核算)", value=True)
    chk_a1 = st.checkbox("P.3 A1類交通事故死亡人數統計表", value=True)
    chk_a2 = st.checkbox("P.4 A2類交通事故受傷人數統計表", value=True)
    chk_major_tot = st.checkbox("P.5 取締重大交通違規統計表 (總表)", value=True)
    chk_overload = st.checkbox("P.6 取締超載違規件數統計表", value=True)
    chk_jingtao = st.checkbox("P.7 「靜桃計畫」大執法專案統計表", value=True)
    chk_tech = st.checkbox("P.8 科技執法成效", value=True)

with col_opt2:
    st.markdown("##### 🔍 狀態檢視")
    st.info(f"📊 三項重點【本期 (09/16)】合計：**{three_matrix[0][4]}** 件\n\n📈 三項重點【9/1起累計】總計：**{three_matrix[0][8]}** 件")

with st.expander("👀 點擊展開預覽待輸出業務數據（確認為最新數值）", expanded=True):
    t1, t2, t3, t4, t5, t6, t7 = st.tabs(["三項重點 (最新核算)", "A1事故死亡", "A2事故受傷", "重大違規", "超載取締", "靜桃計畫", "科技執法成效"])
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
# 6. 執行指定輸出生成
# ==========================================
if st.button("🚀 立即直出簡報：將上方核算數值同步至 Google 簡報", type="primary"):
    slides_svc = get_slides_service()

    if not slides_svc:
        st.error("❌ 無法初始化 Google Slides 服務，請確認 secrets.toml 設定。")
    else:
        with st.spinner("正在動態編譯並更新目標 Google 簡報..."):
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
                        data_rows=three_matrix,
                        latest_day="09/16"
                    )

                if chk_a1:
                    builder.add_table_slide(slide_title="A1類交通事故死亡人數統計表", df=df_a1, is_accident_table=True)

                if chk_a2:
                    builder.add_table_slide(slide_title="A2類交通事故受傷人數統計表", df=df_a2, is_accident_table=True)

                if chk_major_tot:
                    builder.add_table_slide(slide_title="取締重大交通違規統計表", df=df_major, footnote=major_footnote_exact)

                if chk_overload:
                    builder.add_table_slide(slide_title="取締超載違規件數統計表", df=df_overload, footnote=overload_footnote_exact)

                if chk_jingtao:
                    builder.add_table_slide(slide_title="「靜桃計畫」大執法專案統計表", df=df_jingtao)

                if chk_tech:
                    builder.add_table_slide(slide_title=f"科技執法成效 ({tech_date_range_str})", df=df_tech_final, custom_width=480)

                builder.wipe_old_slides()
                final_url = builder.execute_build()

                st.balloons()
                st.success("🎉 Google 簡報直出重繪完成！")
                st.markdown(
                    f"### 📑 簡報入口：\n"
                    f"👉 **[點此直接開啟已更新的簡報]({final_url})**\n\n"
                    f"三項重點違規數值已與來源表完全連動更新！"
                )

            except HttpError as e:
                st.error(f"❌ Google API 請求失敗：{e}\n\n*提示：請確認簡報是否已共用給 `{SERVICE_ACCOUNT_EMAIL}` 並設定為「編輯者」。*")
            except Exception as e:
                st.error(f"❌ 建立簡報失敗：{e}")
