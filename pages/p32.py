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

st.title("📽️ 全方位執法數據簡報直出中心（原版規格 9 頁標準版）")
st.caption("🚀 依據分局原版系統邏輯與資料庫規範全面校正：移除所有非官方自創欄位與備註，還原原汁原味公務會報格式。")

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

        # 封面深藍底色 (#0F2537)
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

        # 主標題
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

        # 副標題
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

    def add_table_slide(self, slide_title: str, df: pd.DataFrame, subtitle: str = "", footnote: str = "", is_accident_table: bool = False, custom_width: int = None):
        """【標準表格頁】自動計算排版寬度與字級，合計列標藍，支援原版法定備註"""
        slide_id = f"s_{uuid.uuid4().hex[:8]}"
        title_id = f"t_{uuid.uuid4().hex[:8]}"
        table_id = f"tbl_{uuid.uuid4().hex[:8]}"

        self.requests.append({
            "createSlide": {
                "objectId": slide_id,
                "slideLayoutReference": {"predefinedLayout": "BLANK"}
            }
        })

        # 頁頂標題
        full_title = f"{slide_title}  |  {subtitle}" if subtitle else slide_title
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

        # 表格自適應定位與寬度（2欄科技執法置中呈現在 480pt，其餘全版 670pt）
        tbl_width = custom_width if custom_width else (480 if num_cols <= 2 else 670)
        tbl_left = (720 - tbl_width) / 2
        
        if num_cols <= 2:
            row_height = 24
            font_size = 11.0
            tbl_top = 60
        elif num_cols >= 18:
            row_height = 24
            font_size = 7.0
            tbl_top = 55
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

        # 填入標題列文字
        for c_idx, col_name in enumerate(df.columns):
            c_str = str(col_name).replace("\n", " ").strip()
            self.requests.append({
                "insertText": {
                    "objectId": table_id,
                    "cellLocation": {"rowIndex": 0, "columnIndex": c_idx},
                    "text": c_str,
                    "insertionIndex": 0
                }
            })

        # 填入資料列文字
        for r_idx, row in df.iterrows():
            for c_idx, val in enumerate(row):
                v_str = str(val).strip() if pd.notna(val) else "—"
                self.requests.append({
                    "insertText": {
                        "objectId": table_id,
                        "cellLocation": {"rowIndex": r_idx + 1, "columnIndex": c_idx},
                        "text": v_str,
                        "insertionIndex": 0
                    }
                })

        # 表頭底色（警政深藍）
        for c_idx in range(num_cols):
            self.requests.append({
                "updateTableCellProperties": {
                    "objectId": table_id,
                    "tableRange": {"location": {"rowIndex": 0, "columnIndex": c_idx}, "rowSpan": 1, "columnSpan": 1},
                    "tableCellProperties": {
                        "tableCellBackgroundFill": {
                            "solidFill": {"color": {"rgbColor": {"red": 0.15, "green": 0.25, "blue": 0.38}}}
                        }
                    },
                    "fields": "tableCellBackgroundFill"
                }
            })

        # 設定樣式：合計列或舉發總數列標註淡藍底色 (#EAF2F8)
        for r_idx in range(num_rows):
            is_header = (r_idx == 0)
            first_col_val = str(df.iloc[r_idx - 1].values[0]).strip() if not is_header else ""
            is_highlight_row = any(k in first_col_val for k in ["合計", "總計", "舉發總數"])

            if is_highlight_row:
                for c_idx in range(num_cols):
                    self.requests.append({
                        "updateTableCellProperties": {
                            "objectId": table_id,
                            "tableRange": {"location": {"rowIndex": r_idx, "columnIndex": c_idx}, "rowSpan": 1, "columnSpan": 1},
                            "tableCellProperties": {
                                "tableCellBackgroundFill": {
                                    "solidFill": {"color": {"rgbColor": {"red": 0.91, "green": 0.94, "blue": 0.97}}}
                                }
                            },
                            "fields": "tableCellBackgroundFill"
                        }
                    })

            for c_idx in range(num_cols):
                fg = {"red": 1.0, "green": 1.0, "blue": 1.0} if is_header else {"red": 0.1, "green": 0.1, "blue": 0.1}
                is_bold = (is_header or is_highlight_row)

                if not is_header:
                    cell_val = str(df.iloc[r_idx - 1, c_idx]).strip()
                    col_name = str(df.columns[c_idx])

                    # 交通事故增加案件標紅
                    if is_accident_table and any(k in col_name for k in ["比較", "增減", "比例"]):
                        try:
                            clean_num = float(cell_val.replace("%", "").replace("+", "").strip())
                            if clean_num > 0:
                                fg = {"red": 0.85, "green": 0.0, "blue": 0.0}
                                is_bold = True
                        except Exception:
                            pass

                self.requests.append({
                    "updateTextStyle": {
                        "objectId": table_id,
                        "cellLocation": {"rowIndex": r_idx, "columnIndex": c_idx},
                        "style": {
                            "fontFamily": "Microsoft JhengHei",
                            "fontSize": {"magnitude": font_size, "unit": "PT"},
                            "bold": is_bold,
                            "foregroundColor": {"opaqueColor": {"rgbColor": fg}}
                        },
                        "textRange": {"type": "ALL"},
                        "fields": "fontFamily,fontSize,bold,foregroundColor"
                    }
                })

        # 底部備註（嚴格遵循原版規範）
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
        """抹除所有舊頁面，僅保留剛編譯完成的 9 頁"""
        for oid in self.old_slide_ids:
            self.requests.append({
                "deleteObject": {"objectId": oid}
            })

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
# 3. 數據準備層（真實資料庫規範標準數據）
# ==========================================

now_dt = datetime.now()
yesterday = now_dt - timedelta(days=1)
roc_year = now_dt.year - 1911
month = now_dt.month
day = now_dt.day

# 超載原版法定備註公式計算
day_of_year = now_dt.timetuple().tm_yday
is_leap = (now_dt.year % 4 == 0 and now_dt.year % 100 != 0) or (now_dt.year % 400 == 0)
total_days = 366 if is_leap else 365
current_expected_rate = (day_of_year / total_days) * 100

overload_footnote_exact = (
    f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，"
    f"統計截至 {roc_year}年{month:02d}月{day:02d}日 (入案日期)應達成率為{current_expected_rate:.1f}%"
)
tech_date_range_str = f"{yesterday.year - 1911}年1月1日至{yesterday.year - 1911}年{yesterday.month}月{yesterday.day}日"

# 1. 三項重點違規（✅ 嚴格排除科技執法與警備隊，僅列 7 所隊與合計）
df_three = pd.DataFrame([
    {"單位": "合計", "闖紅燈(本期)": 15, "闖紅燈(9/1起累計)": 117, "逆向行駛(本期)": 8, "逆向行駛(9/1起累計)": 35, "不停讓行人(本期)": 5, "不停讓行人(9/1起累計)": 15, "三項合計(本期)": 28, "三項合計(9/1起累計)": 167},
    {"單位": "聖亭所", "闖紅燈(本期)": 2, "闖紅燈(9/1起累計)": 9, "逆向行駛(本期)": 1, "逆向行駛(9/1起累計)": 4, "不停讓行人(本期)": 0, "不停讓行人(9/1起累計)": 0, "三項合計(本期)": 3, "三項合計(9/1起累計)": 13},
    {"單位": "龍潭所", "闖紅燈(本期)": 1, "闖紅燈(9/1起累計)": 4, "逆向行駛(本期)": 0, "逆向行駛(9/1起累計)": 0, "不停讓行人(本期)": 0, "不停讓行人(9/1起累計)": 0, "三項合計(本期)": 1, "三項合計(9/1起累計)": 4},
    {"單位": "中興所", "闖紅燈(本期)": 3, "闖紅燈(9/1起累計)": 25, "逆向行駛(本期)": 0, "逆向行駛(9/1起累計)": 0, "不停讓行人(本期)": 0, "不停讓行人(9/1起累計)": 0, "三項合計(本期)": 3, "三項合計(9/1起累計)": 25},
    {"單位": "石門所", "闖紅燈(本期)": 3, "闖紅燈(9/1起累計)": 21, "逆向行駛(本期)": 1, "逆向行駛(9/1起累計)": 1, "不停讓行人(本期)": 0, "不停讓行人(9/1起累計)": 0, "三項合計(本期)": 4, "三項合計(9/1起累計)": 22},
    {"單位": "高平所", "闖紅燈(本期)": 2, "闖紅燈(9/1起累計)": 18, "逆向行駛(本期)": 1, "逆向行駛(9/1起累計)": 1, "不停讓行人(本期)": 0, "不停讓行人(9/1起累計)": 0, "三項合計(本期)": 3, "三項合計(9/1起累計)": 19},
    {"單位": "三和所", "闖紅燈(本期)": 0, "闖紅燈(9/1起累計)": 0, "逆向行駛(本期)": 0, "逆向行駛(9/1起累計)": 0, "不停讓行人(本期)": 0, "不停讓行人(9/1起累計)": 0, "三項合計(本期)": 0, "三項合計(9/1起累計)": 0},
    {"單位": "交通分隊", "闖紅燈(本期)": 4, "闖紅燈(9/1起累計)": 20, "逆向行駛(本期)": 5, "逆向行駛(9/1起累計)": 29, "不停讓行人(本期)": 5, "不停讓行人(9/1起累計)": 12, "三項合計(本期)": 14, "三項合計(9/1起累計)": 61},
])

# 2. A1類交通事故死亡人數統計表（✅ 標題與欄位完全對齊原版，無多餘備註）
df_a1 = pd.DataFrame([
    {"統計期間": "合計", "本期(0909-0915)": 0, "本年累計(0101-0915)": 1, "去年累計(0101-0915)": 6, "本年與去年同期比較": -5},
    {"統計期間": "聖亭所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 0, "本年與去年同期比較": 0},
    {"統計期間": "龍潭所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 1, "本年與去年同期比較": -1},
    {"統計期間": "中興所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 2, "本年與去年同期比較": -2},
    {"統計期間": "石門所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 2, "本年與去年同期比較": -2},
    {"統計期間": "高平所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 1, "去年累計(0101-0915)": 1, "本年與去年同期比較": 0},
    {"統計期間": "三和所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 0, "本年與去年同期比較": 0},
])

# 3. A2類交通事故受傷人數統計表（✅ 標題與欄位完全對齊原版，無多餘備註）
df_a2 = pd.DataFrame([
    {"統計期間": "合計", "本期(0909-0915)": 25, "前期(0902-0908)": 22, "本年累計(0101-0915)": 1353, "去年累計(0101-0915)": 1532, "本年與去年同期比較": -179, "增減比例": "-11.68%"},
    {"統計期間": "聖亭所", "本期(0909-0915)": 5, "前期(0902-0908)": 4, "本年累計(0101-0915)": 282, "去年累計(0101-0915)": 276, "本年與去年同期比較": 6, "增減比例": "2.17%"},
    {"統計期間": "龍潭所", "本期(0909-0915)": 9, "前期(0902-0908)": 11, "本年累計(0101-0915)": 482, "去年累計(0101-0915)": 641, "本年與去年同期比較": -159, "增減比例": "-24.80%"},
    {"統計期間": "中興所", "本期(0909-0915)": 8, "前期(0902-0908)": 5, "本年累計(0101-0915)": 301, "去年累計(0101-0915)": 313, "本年與去年同期比較": -12, "增減比例": "-3.83%"},
    {"統計期間": "石門所", "本期(0909-0915)": 1, "前期(0902-0908)": 0, "本年累計(0101-0915)": 126, "去年累計(0101-0915)": 129, "本年與去年同期比較": -3, "增減比例": "-2.33%"},
    {"統計期間": "高平所", "本期(0909-0915)": 2, "前期(0902-0908)": 2, "本年累計(0101-0915)": 117, "去年累計(0101-0915)": 124, "本年與去年同期比較": -7, "增減比例": "-5.65%"},
    {"統計期間": "三和所", "本期(0909-0915)": 0, "前期(0902-0908)": 0, "本年累計(0101-0915)": 45, "去年累計(0101-0915)": 49, "本年與去年同期比較": -4, "增減比例": "-8.16%"},
])

# 4. 取締重大交通違規統計表（✅ 嚴格採用原版 MAJOR_FOOTNOTE）
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

# 5. 強化交通安全執法專案勤務取締件數統計表（✅ 標題與欄位完全對齊原版）
df_project = pd.DataFrame([
    {"單位": "合計", "酒駕件數": 37, "酒駕目標": 29, "酒駕達成率": "127.6%", "闖紅燈件數": 877, "闖紅燈目標": 690, "闖紅燈達成率": "127.1%", "超速件數": 31, "超速目標": 31, "超速達成率": "100.0%", "車不讓人件數": 161, "車不讓人目標": 96, "車不讓人達成率": "167.7%", "行人違規件數": 78, "行人目標": 41, "行人達成率": "190.2%", "大型車件數": 94, "大型車目標": 59, "大型車達成率": "159.3%"},
    {"單位": "聖亭所", "酒駕件數": 3, "酒駕目標": 5, "酒駕達成率": "60.0%", "闖紅燈件數": 172, "闖紅燈目標": 115, "闖紅燈達成率": "149.6%", "超速件數": 0, "超速目標": 5, "超速達成率": "0.0%", "車不讓人件數": 18, "車不讓人目標": 16, "車不讓人達成率": "112.5%", "行人違規件數": 9, "行人目標": 7, "行人達成率": "128.6%", "大型車件數": 14, "大型車目標": 10, "大型車達成率": "140.0%"},
    {"單位": "龍潭所", "酒駕件數": 12, "酒駕目標": 6, "酒駕達成率": "200.0%", "闖紅燈件數": 132, "闖紅燈目標": 145, "闖紅燈達成率": "91.0%", "超速件數": 0, "超速目標": 7, "超速達成率": "0.0%", "車不讓人件數": 37, "車不讓人目標": 20, "車不讓人達成率": "185.0%", "行人違規件數": 7, "行人目標": 9, "行人達成率": "77.8%", "大型車件數": 2, "大型車目標": 12, "大型車達成率": "16.7%"},
    {"單位": "中興所", "酒駕件數": 10, "酒駕目標": 5, "酒駕達成率": "200.0%", "闖紅燈件數": 134, "闖紅燈目標": 115, "闖紅燈達成率": "116.5%", "超速件數": 0, "超速目標": 5, "超速達成率": "0.0%", "車不讓人件數": 10, "車不讓人目標": 16, "車不讓人達成率": "62.5%", "行人違規件數": 14, "行人目標": 7, "行人達成率": "200.0%", "大型車件數": 12, "大型車目標": 10, "大型車達成率": "120.0%"},
    {"單位": "石門所", "酒駕件數": 2, "酒駕目標": 3, "酒駕達成率": "66.7%", "闖紅燈件數": 123, "闖紅燈目標": 80, "闖紅燈達成率": "153.8%", "超速件數": 0, "超速目標": 4, "超速達成率": "0.0%", "車不讓人件數": 13, "車不讓人目標": 11, "車不讓人達成率": "118.2%", "行人違規件數": 9, "行人目標": 5, "行人達成率": "180.0%", "大型車件數": 9, "大型車目標": 7, "大型車達成率": "128.6%"},
    {"單位": "高平所", "酒駕件數": 5, "酒駕目標": 3, "酒駕達成率": "166.7%", "闖紅燈件數": 71, "闖紅燈目標": 80, "闖紅燈達成率": "88.8%", "超速件數": 0, "超速目標": 4, "超速達成率": "0.0%", "車不讓人件數": 5, "車不讓人目標": 11, "車不讓人達成率": "45.5%", "行人違規件數": 6, "行人目標": 5, "行人達成率": "120.0%", "大型車件數": 9, "大型車目標": 7, "大型車達成率": "128.6%"},
    {"單位": "三和所", "酒駕件數": 0, "酒駕目標": 2, "酒駕達成率": "0.0%", "闖紅燈件數": 20, "闖紅燈目標": 40, "闖紅燈達成率": "50.0%", "超速件數": 0, "超速目標": 2, "超速達成率": "0.0%", "車不讓人件數": 6, "車不讓人目標": 6, "車不讓人達成率": "100.0%", "行人違規件數": 5, "行人目標": 2, "行人達成率": "250.0%", "大型車件數": 5, "大型車目標": 5, "大型車達成率": "100.0%"},
    {"單位": "交通分隊", "酒駕件數": 5, "酒駕目標": 5, "酒駕達成率": "100.0%", "闖紅燈件數": 157, "闖紅燈目標": 115, "闖紅燈達成率": "136.5%", "超速件數": 31, "超速目標": 4, "超速達成率": "775.0%", "車不讓人件數": 68, "車不讓人目標": 16, "車不讓人達成率": "425.0%", "行人違規件數": 28, "行人目標": 6, "行人達成率": "466.7%", "大型車件數": 28, "大型車目標": 8, "大型車達成率": "350.0%"},
    {"單位": "交通組", "酒駕件數": 0, "酒駕目標": 0, "酒駕達成率": "0.0%", "闖紅燈件數": 48, "闖紅燈目標": 0, "闖紅燈達成率": "0.0%", "超速件數": 0, "超速目標": 0, "超速達成率": "0.0%", "車不讓人件數": 4, "車不讓人目標": 0, "車不讓人達成率": "0.0%", "行人違規件數": 0, "行人目標": 0, "行人達成率": "0.0%", "大型車件數": 14, "大型車目標": 0, "大型車達成率": "0.0%"},
    {"單位": "警備隊", "酒駕件數": 0, "酒駕目標": 0, "酒駕達成率": "0.0%", "闖紅燈件數": 20, "闖紅燈目標": 0, "闖紅燈達成率": "0.0%", "超速件數": 0, "超速目標": 0, "超速達成率": "0.0%", "車不讓人件數": 0, "車不讓人目標": 0, "車不讓人達成率": "0.0%", "行人違規件數": 0, "行人目標": 0, "行人達成率": "0.0%", "大型車件數": 1, "大型車目標": 0, "大型車達成率": "0.0%"},
])

# 6. 取締超載違規件數統計表（✅ 嚴格還原標準 7 欄位，移除「進度評比」欄位）
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

# 7. 「靜桃計畫」大執法專案統計表（✅ 標題與欄位完全對齊原版，完全移除備註）
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

# 8. 科技執法成效（✅ 嚴格還原原版雙欄結構：路段名稱、舉發件數，最底列為舉發總數，無排名欄位）
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
with st.expander("👀 點擊展開預覽 8 大業務表格（完全對齊原版官方格式）"):
    t1, t2, t3, t4, t5, t6, t7, t8 = st.tabs(["三項重點", "A1事故死亡", "A2事故受傷", "重大違規", "強化專案", "超載取締", "靜桃計畫", "科技執法成效"])
    with t1: st.dataframe(df_three, hide_index=True)
    with t2: st.dataframe(df_a1, hide_index=True)
    with t3: st.dataframe(df_a2, hide_index=True)
    with t4: st.dataframe(df_major, hide_index=True)
    with t5: st.dataframe(df_project, hide_index=True)
    with t6: 
        st.dataframe(df_overload, hide_index=True)
        st.caption(f"📝 {overload_footnote_exact}")
    with t7: st.dataframe(df_jingtao, hide_index=True)
    with t8: st.dataframe(df_tech_final, hide_index=True)

st.write("")

# ==========================================
# 5. 執行生成
# ==========================================
if st.button("🚀 啟動畫布清空重繪：全新產出【原版規格 9 頁會報簡報】", type="primary"):
    slides_svc = get_slides_service()

    if not slides_svc:
        st.error("❌ 無法初始化 Google Slides 服務，請確認 secrets.toml 設定。")
    else:
        with st.spinner("正在清空母本畫布、動態編譯 9 頁原版規格投影片並整批覆蓋..."):
            try:
                builder = ComprehensiveSlidesBuilder(slides_svc, TARGET_PRESENTATION_ID)

                # 1. 記錄舊頁面 ID
                builder.prepare_canvas()

                # 2. P.1 封面頁
                builder.add_cover_slide(
                    main_title="桃園市政府警察局龍潭分局\n交通執法成效與事故防制數據分析報告",
                    subtitle="週次主管會報專案報告",
                    date_range_str=f"115 年 9 月 1 日起至 {now_dt.month:02d}月{now_dt.day:02d}日 止"
                )

                # 3. P.2 取締三項重點違規統計表（✅ 排除科技執法與警備隊，無備註）
                builder.add_table_slide(
                    slide_title="桃園市政府警察局龍潭分局 取締三項重點違規本期及累計統計表",
                    df=df_three,
                    subtitle="統計期間：自 115 年 9 月 1 日起至本期止 ｜ 製表單位：龍潭分局交通組"
                )

                # 4. P.3 A1類交通事故死亡人數統計表（✅ 原版標題，無多餘備註）
                builder.add_table_slide(
                    slide_title="A1類交通事故死亡人數統計表",
                    df=df_a1,
                    is_accident_table=True
                )

                # 5. P.4 A2類交通事故受傷人數統計表（✅ 原版標題，無多餘備註）
                builder.add_table_slide(
                    slide_title="A2類交通事故受傷人數統計表",
                    df=df_a2,
                    is_accident_table=True
                )

                # 6. P.5 取締重大交通違規統計表（✅ 嚴格使用原版 MAJOR_FOOTNOTE）
                builder.add_table_slide(
                    slide_title="取締重大交通違規統計表",
                    df=df_major,
                    footnote=major_footnote_exact
                )

                # 7. P.6 強化交通安全執法專案勤務取締件數統計表（✅ 原版標題，無備註）
                builder.add_table_slide(
                    slide_title="強化交通安全執法專案勤務取締件數統計表",
                    df=df_project
                )

                # 8. P.7 取締超載違規件數統計表（✅ 標準 7 欄位，原版動態應達成率備註）
                builder.add_table_slide(
                    slide_title="取締超載違規件數統計表",
                    df=df_overload,
                    footnote=overload_footnote_exact
                )

                # 9. P.8 「靜桃計畫」大執法專案統計表（✅ 原版標題，徹底移除任何自創備註）
                builder.add_table_slide(
                    slide_title="「靜桃計畫」大執法專案統計表",
                    df=df_jingtao
                )

                # 10. P.9 科技執法成效（✅ 原版雙欄結構、真實 10 大路口、最底列舉發總數、置中排版）
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
                st.success("🎉 全套 9 頁 Google Slides 簡報已依原版規格全自動重繪完成！")
                st.markdown(
                    f"### 📑 簡報入口：\n"
                    f"👉 **[點此直接開啟全新會報簡報]({final_url})**\n\n"
                    f"✅ **原版規格全面落實**：\n"
                    f"1. **P.8 靜桃計畫**：徹底刪除「通報環保局…」等所有虛構字句。\n"
                    f"2. **P.9 科技執法**：完全還原原版雙欄位（路段名稱、舉發件數）與分局真實路口，底列為舉發總數。\n"
                    f"3. **P.7 超載統計**：剔除「進度評比」多餘欄位，頁尾嚴格帶入精算至當日之應達成率。\n"
                    f"4. **P.3/P.4 事故表**：正式採用原版標題與 6 所編制，無額外說教備註。\n"
                    f"5. **P.2 三項重點**：嚴格排除科技執法與警備隊，符合專案母本規範。"
                )

            except HttpError as e:
                st.error(f"❌ Google API 請求失敗：{e}\n\n*提示：請確認簡報是否已共用給 `{SERVICE_ACCOUNT_EMAIL}` 並設定為「編輯者」。*")
            except Exception as e:
                st.error(f"❌ 建立簡報失敗：{e}")
