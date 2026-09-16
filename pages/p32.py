import uuid
from datetime import datetime
import pandas as pd
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build

from menu import show_sidebar

# ==========================================
# 0. 頁面初始化
# ==========================================
st.set_page_config(page_title="全方位執法數據簡報直出中心", page_icon="📽️", layout="wide")
show_sidebar()

st.title("📽️ 全方位執法數據簡報直出中心（免試算表、涵蓋全統計）")
st.info("💡 本引擎整合科技執法、超載、重大違規、三項重點、強化專案、交通事故與靜桃計畫共 7 大模組，免經 Google Sheets，直接產出高規格 Google Slides 會報簡報。")

# ==========================================
# 1. Google 服務連線層
# ==========================================
GCP_CREDS = dict(st.secrets.get("gcp_service_account", {}))
DRIVE_FOLDER_ID = st.secrets.get("DRIVE_FOLDER_ID", "").strip()

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

def get_drive_service():
    if not GCP_CREDS:
        return None
    creds = service_account.Credentials.from_service_account_info(
        GCP_CREDS,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)

# ==========================================
# 2. 全方位簡報編譯器 (ComprehensiveSlidesBuilder)
# ==========================================
class ComprehensiveSlidesBuilder:
    def __init__(self, slides_svc, drive_svc=None):
        self.slides_svc = slides_svc
        self.drive_svc = drive_svc
        self.presentation_id = None
        self.requests = []

    def create_presentation(self, title: str, parent_folder_id: str = None) -> str:
        """建立全新空白簡報並移動至指定資料夾"""
        body = {"title": title}
        pres = self.slides_svc.presentations().create(body=body).execute()
        self.presentation_id = pres.get("presentationId")

        if parent_folder_id and self.drive_svc:
            try:
                fid = parent_folder_id.replace('"', '').replace("'", '').strip()
                file = self.drive_svc.files().get(
                    fileId=self.presentation_id, fields="parents", supportsAllDrives=True
                ).execute()
                previous_parents = ",".join(file.get("parents", []))
                
                self.drive_svc.files().update(
                    fileId=self.presentation_id,
                    addParents=fid,
                    removeParents=previous_parents,
                    fields="id, parents",
                    supportsAllDrives=True
                ).execute()
            except Exception as e:
                st.warning(f"⚠️ 簡報建立成功，但移至資料夾失敗：{e}")

        # 移除預設的空白第一頁
        initial_slides = pres.get("slides", [])
        if initial_slides:
            self.requests.append({
                "deleteObject": {"objectId": initial_slides[0]["objectId"]}
            })

        return self.presentation_id

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
        # 背景色：警政深藍 (#0F2537)
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
                "fields": "fontFamily,fontSize,bold,foregroundColor"
            }
        })
        # 副標題與備註
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
                "fields": "fontFamily,fontSize,foregroundColor"
            }
        })

    def add_table_slide(self, slide_title: str, df: pd.DataFrame, subtitle: str = "", footnote: str = ""):
        """【標準表格頁】自動計算欄寬、字級與表頭底色，支援註腳"""
        slide_id = f"s_{uuid.uuid4().hex[:8]}"
        title_id = f"t_{uuid.uuid4().hex[:8]}"
        table_id = f"tbl_{uuid.uuid4().hex[:8]}"

        self.requests.append({
            "createSlide": {
                "objectId": slide_id,
                "slideLayoutReference": {"predefinedLayout": "BLANK"}
            }
        })

        # 頂部標題列
        full_title = f"{slide_title}  |  {subtitle}" if subtitle else slide_title
        self.requests.append({
            "createShape": {
                "objectId": title_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {"width": {"magnitude": 660, "unit": "PT"}, "height": {"magnitude": 35, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": 30, "translateY": 15, "unit": "PT"}
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
                "fields": "fontFamily,fontSize,bold,foregroundColor"
            }
        })

        # 動態字型微調（欄位 > 10 欄時自動微縮至 8-9pt，避免折行）
        num_cols = len(df.columns)
        num_rows = len(df) + 1
        font_size = 7 if num_cols >= 18 else (9 if num_cols >= 10 else 11)

        tbl_top = 55
        tbl_height = min(300, max(140, num_rows * 24))

        self.requests.append({
            "createTable": {
                "objectId": table_id,
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {"width": {"magnitude": 660, "unit": "PT"}, "height": {"magnitude": tbl_height, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": 30, "translateY": tbl_top, "unit": "PT"}
                },
                "rows": num_rows,
                "columns": num_cols
            }
        })

        # 填寫標題列文字
        for c_idx, col_name in enumerate(df.columns):
            c_str = str(col_name).replace("\n", " ")
            self.requests.append({
                "insertText": {
                    "objectId": table_id,
                    "cellLocation": {"rowIndex": 0, "columnIndex": c_idx},
                    "text": c_str
                }
            })

        # 填寫資料列
        for r_idx, row in df.iterrows():
            for c_idx, val in enumerate(row):
                self.requests.append({
                    "insertText": {
                        "objectId": table_id,
                        "cellLocation": {"rowIndex": r_idx + 1, "columnIndex": c_idx},
                        "text": str(val) if pd.notna(val) else "—"
                    }
                })

        # 表頭美化（深藍底色、白字）
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

        # 全表字級與字型統一
        for r_idx in range(num_rows):
            for c_idx in range(num_cols):
                fg = {"red": 1.0, "green": 1.0, "blue": 1.0} if r_idx == 0 else {"red": 0.1, "green": 0.1, "blue": 0.1}
                self.requests.append({
                    "updateTextStyle": {
                        "objectId": table_id,
                        "cellLocation": {"rowIndex": r_idx, "columnIndex": c_idx},
                        "style": {
                            "fontFamily": "Microsoft JhengHei",
                            "fontSize": {"magnitude": font_size, "unit": "PT"},
                            "bold": (r_idx == 0 or r_idx == 1),
                            "foregroundColor": {"opaqueColor": {"rgbColor": fg}}
                        },
                        "fields": "fontFamily,fontSize,bold,foregroundColor"
                    }
                })

        # 底部註腳說明
        if footnote:
            fn_id = f"fn_{uuid.uuid4().hex[:8]}"
            self.requests.append({
                "createShape": {
                    "objectId": fn_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {"width": {"magnitude": 660, "unit": "PT"}, "height": {"magnitude": 30, "unit": "PT"}},
                        "transform": {"scaleX": 1, "scaleY": 1, "translateX": 30, "translateY": 365, "unit": "PT"}
                    }
                }
            })
            self.requests.append({"insertText": {"objectId": fn_id, "text": f"📝 {footnote}", "insertionIndex": 0}})
            self.requests.append({
                "updateTextStyle": {
                    "objectId": fn_id,
                    "style": {
                        "fontFamily": "Microsoft JhengHei",
                        "fontSize": {"magnitude": 9.5, "unit": "PT"},
                        "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.4, "green": 0.4, "blue": 0.4}}}
                    },
                    "fields": "fontFamily,fontSize,foregroundColor"
                }
            })

    def add_side_by_side_tables(self, slide_title: str, df_left: pd.DataFrame, title_left: str, df_right: pd.DataFrame, title_right: str, subtitle: str = ""):
        """【雙表並排頁】專供 A1 死亡 與 A2 受傷 在同一頁直接對比"""
        slide_id = f"s_dual_{uuid.uuid4().hex[:8]}"
        title_id = f"t_dual_{uuid.uuid4().hex[:8]}"

        self.requests.append({
            "createSlide": {
                "objectId": slide_id,
                "slideLayoutReference": {"predefinedLayout": "BLANK"}
            }
        })

        # 主標題
        full_title = f"{slide_title}  |  {subtitle}" if subtitle else slide_title
        self.requests.append({
            "createShape": {
                "objectId": title_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {"width": {"magnitude": 660, "unit": "PT"}, "height": {"magnitude": 35, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": 30, "translateY": 15, "unit": "PT"}
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
                "fields": "fontFamily,fontSize,bold,foregroundColor"
            }
        })

        # 繪製左右兩表格
        def build_one_tbl(df, t_label, x_offset, w):
            t_box_id = f"subt_{uuid.uuid4().hex[:8]}"
            tbl_id = f"tbl_{uuid.uuid4().hex[:8]}"

            # 子標題
            self.requests.append({
                "createShape": {
                    "objectId": t_box_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {"width": {"magnitude": w, "unit": "PT"}, "height": {"magnitude": 25, "unit": "PT"}},
                        "transform": {"scaleX": 1, "scaleY": 1, "translateX": x_offset, "translateY": 50, "unit": "PT"}
                    }
                }
            })
            self.requests.append({"insertText": {"objectId": t_box_id, "text": t_label, "insertionIndex": 0}})
            self.requests.append({
                "updateTextStyle": {
                    "objectId": t_box_id,
                    "style": {"fontFamily": "Microsoft JhengHei", "fontSize": {"magnitude": 12, "unit": "PT"}, "bold": True, "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.2, "green": 0.3, "blue": 0.45}}}},
                    "fields": "fontFamily,fontSize,bold,foregroundColor"
                }
            })

            # 表格主體
            n_rows, n_cols = len(df) + 1, len(df.columns)
            h = min(280, n_rows * 28)
            self.requests.append({
                "createTable": {
                    "objectId": tbl_id,
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {"width": {"magnitude": w, "unit": "PT"}, "height": {"magnitude": h, "unit": "PT"}},
                        "transform": {"scaleX": 1, "scaleY": 1, "translateX": x_offset, "translateY": 78, "unit": "PT"}
                    },
                    "rows": n_rows,
                    "columns": n_cols
                }
            })
            for c, name in enumerate(df.columns):
                self.requests.append({"insertText": {"objectId": tbl_id, "cellLocation": {"rowIndex": 0, "columnIndex": c}, "text": str(name)}})
            for r, row in df.iterrows():
                for c, val in enumerate(row):
                    self.requests.append({"insertText": {"objectId": tbl_id, "cellLocation": {"rowIndex": r + 1, "columnIndex": c}, "text": str(val) if pd.notna(val) else "—"}})
            for c in range(n_cols):
                self.requests.append({
                    "updateTableCellProperties": {
                        "objectId": tbl_id,
                        "tableRange": {"location": {"rowIndex": 0, "columnIndex": c}, "rowSpan": 1, "columnSpan": 1},
                        "tableCellProperties": {"tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": {"red": 0.18, "green": 0.28, "blue": 0.42}}}}},
                        "fields": "tableCellBackgroundFill"
                    }
                })
            for r in range(n_rows):
                for c in range(n_cols):
                    self.requests.append({
                        "updateTextStyle": {
                            "objectId": tbl_id,
                            "cellLocation": {"rowIndex": r, "columnIndex": c},
                            "style": {
                                "fontFamily": "Microsoft JhengHei",
                                "fontSize": {"magnitude": 8.5 if n_cols > 5 else 9.5, "unit": "PT"},
                                "bold": (r == 0 or r == 1),
                                "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 1.0, "green": 1.0, "blue": 1.0} if r == 0 else {"red": 0.1, "green": 0.1, "blue": 0.1}}}
                            },
                            "fields": "fontFamily,fontSize,bold,foregroundColor"
                        }
                    })

        # 左表寬 315 pt，右表寬 335 pt，間距 10 pt
        build_one_tbl(df_left, title_left, x_offset=30, w=315)
        build_one_tbl(df_right, title_right, x_offset=355, w=335)

    def execute_build(self) -> str:
        """整批執行請求並回傳檢視網址"""
        batch_size = 200
        for i in range(0, len(self.requests), batch_size):
            chunk = self.requests[i:i + batch_size]
            self.slides_svc.presentations().batchUpdate(
                presentationId=self.presentation_id,
                body={"requests": chunk}
            ).execute()
        return f"https://docs.google.com/presentation/d/{self.presentation_id}/edit"

# ==========================================
# 3. 測試操作介面（內建 7 大業務真實結構數據）
# ==========================================
st.subheader("⚙️ 簡報匯出參數設定")
c_p1, c_p2 = st.columns([2, 1])
with c_p1:
    target_folder = st.text_input("雲端硬碟共用資料夾 ID (DRIVE_FOLDER_ID)：", value=DRIVE_FOLDER_ID)
with c_p2:
    report_title = st.text_input("輸出簡報標題：", value=f"龍潭分局主管會報交通執法與事故防制專案報告_{datetime.now().strftime('%m%d')}")

include_details = st.checkbox("📑 一併產出附錄：重大違規 7 大項細表（酒駕、超速、轉彎等，簡報總頁數約 15 頁）", value=False)

st.divider()

# --- 準備 7 大業務完整結構數據 ---
# 1. 三項重點違規
df_three = pd.DataFrame([
    {"單位": "合計", "闖紅燈(本期)": 15, "闖紅燈(累計)": 128, "逆向(本期)": 8, "逆向(累計)": 72, "不停讓行人(本期)": 5, "不停讓行人(累計)": 43, "三項合計(本期)": 28, "三項合計(累計)": 243},
    {"單位": "聖亭所", "闖紅燈(本期)": 3, "闖紅燈(累計)": 24, "逆向(本期)": 1, "逆向(累計)": 15, "不停讓行人(本期)": 1, "不停讓行人(累計)": 8, "三項合計(本期)": 5, "三項合計(累計)": 47},
    {"單位": "龍潭所", "闖紅燈(本期)": 4, "闖紅燈(累計)": 38, "逆向(本期)": 2, "逆向(累計)": 20, "不停讓行人(本期)": 2, "不停讓行人(累計)": 14, "三項合計(本期)": 8, "三項合計(累計)": 72},
    {"單位": "中興所", "闖紅燈(本期)": 2, "闖紅燈(累計)": 21, "逆向(本期)": 1, "逆向(累計)": 12, "不停讓行人(本期)": 1, "不停讓行人(累計)": 7, "三項合計(本期)": 4, "三項合計(累計)": 40},
    {"單位": "石門所", "闖紅燈(本期)": 2, "闖紅燈(累計)": 16, "逆向(本期)": 1, "逆向(累計)": 9, "不停讓行人(本期)": 0, "不停讓行人(累計)": 5, "三項合計(本期)": 3, "三項合計(累計)": 30},
    {"單位": "高平所", "闖紅燈(本期)": 1, "闖紅燈(累計)": 12, "逆向(本期)": 1, "逆向(累計)": 7, "不停讓行人(本期)": 0, "不停讓行人(累計)": 3, "三項合計(本期)": 2, "三項合計(累計)": 22},
    {"單位": "三和所", "闖紅燈(本期)": 0, "闖紅燈(累計)": 4, "逆向(本期)": 0, "逆向(累計)": 3, "不停讓行人(本期)": 0, "不停讓行人(累計)": 1, "三項合計(本期)": 0, "三項合計(累計)": 8},
    {"單位": "交通分隊", "闖紅燈(本期)": 3, "闖紅燈(累計)": 13, "逆向(本期)": 2, "逆向(累計)": 6, "不停讓行人(本期)": 1, "不停讓行人(累計)": 5, "三項合計(本期)": 6, "三項合計(累計)": 24},
])

# 2. 交通事故 (A1 / A2)
df_a1 = pd.DataFrame([
    {"統計期間": "合計", "本期(0901-0907)": 0, "本年累計(0101-0907)": 3, "去年累計(0101-0907)": 4, "本年與去年同期比較": -1},
    {"統計期間": "聖亭所", "本期(0901-0907)": 0, "本年累計(0101-0907)": 1, "去年累計(0101-0907)": 1, "本年與去年同期比較": 0},
    {"統計期間": "龍潭所", "本期(0901-0907)": 0, "本年累計(0101-0907)": 1, "去年累計(0101-0907)": 2, "本年與去年同期比較": -1},
    {"統計期間": "中興所", "本期(0901-0907)": 0, "本年累計(0101-0907)": 0, "去年累計(0101-0907)": 0, "本年與去年同期比較": 0},
    {"統計期間": "石門所", "本期(0901-0907)": 0, "本年累計(0101-0907)": 1, "去年累計(0101-0907)": 1, "本年與去年同期比較": 0},
    {"統計期間": "高平所", "本期(0901-0907)": 0, "本年累計(0101-0907)": 0, "去年累計(0101-0907)": 0, "本年與去年同期比較": 0},
    {"統計期間": "三和所", "本期(0901-0907)": 0, "本年累計(0101-0907)": 0, "去年累計(0101-0907)": 0, "本年與去年同期比較": 0},
])

df_a2 = pd.DataFrame([
    {"統計期間": "合計", "本期": 32, "前期": 35, "本年累計": 1284, "去年累計": 1390, "比較": -106, "增減率": "-7.63%"},
    {"統計期間": "聖亭所", "本期": 7, "前期": 8, "本年累計": 312, "去年累計": 330, "比較": -18, "增減率": "-5.45%"},
    {"統計期間": "龍潭所", "本期": 11, "前期": 12, "本年累計": 445, "去年累計": 472, "比較": -27, "增減率": "-5.72%"},
    {"統計期間": "中興所", "本期": 6, "前期": 7, "本年累計": 268, "去年累計": 290, "比較": -22, "增減率": "-7.59%"},
    {"統計期間": "石門所", "本期": 4, "前期": 5, "本年累計": 142, "去年累計": 160, "比較": -18, "增減率": "-11.25%"},
    {"統計期間": "高平所", "本期": 3, "前期": 2, "本年累計": 92, "去年累計": 105, "比較": -13, "增減率": "-12.38%"},
    {"統計期間": "三和所", "本期": 1, "前期": 1, "本年累計": 25, "去年累計": 33, "比較": -8, "增減率": "-24.24%"},
])

# 3. 重大交通違規 (總表)
df_major = pd.DataFrame([
    {"單位": "合計", "本期(攔停)": 48, "本期(逕舉)": 152, "本年(攔停)": 1840, "本年(逕舉)": 6420, "去年同期": 7950, "增減比較": 310, "目標值": 18115, "達成率": "45.6%"},
    {"單位": "科技執法", "本期(攔停)": 0, "本期(逕舉)": 88, "本年(攔停)": 0, "本年(逕舉)": 3120, "去年同期": 2900, "增減比較": 220, "目標值": 6006, "達成率": "51.9%"},
    {"單位": "聖亭所", "本期(攔停)": 8, "本期(逕舉)": 12, "本年(攔停)": 340, "本年(逕舉)": 590, "去年同期": 910, "增減比較": 20, "目標值": 1941, "達成率": "47.9%"},
    {"單位": "龍潭所", "本期(攔停)": 12, "本期(逕舉)": 16, "本年(攔停)": 480, "本年(逕舉)": 780, "去年同期": 1210, "增減比較": 50, "目標值": 2588, "達成率": "48.7%"},
    {"單位": "中興所", "本期(攔停)": 7, "本期(逕舉)": 10, "本年(攔停)": 310, "本年(逕舉)": 540, "去年同期": 820, "增減比較": 30, "目標值": 1941, "達成率": "43.8%"},
    {"單位": "石門所", "本期(攔停)": 5, "本期(逕舉)": 8, "本年(攔停)": 210, "本年(逕舉)": 430, "去年同期": 610, "增減比較": 30, "目標值": 1479, "達成率": "43.3%"},
    {"單位": "高平所", "本期(攔停)": 4, "本期(逕舉)": 6, "本年(攔停)": 180, "本年(逕舉)": 380, "去年同期": 540, "增減比較": 20, "目標值": 1294, "達成率": "43.3%"},
    {"單位": "三和所", "本期(攔停)": 2, "本期(逕舉)": 2, "本年(攔停)": 60, "本年(逕舉)": 90, "去年同期": 140, "增減比較": 10, "目標值": 339, "達成率": "44.2%"},
    {"單位": "交通分隊", "本期(攔停)": 10, "本期(逕舉)": 10, "本年(攔停)": 260, "本年(逕舉)": 490, "去年同期": 820, "增減比較": -70, "目標值": 2526, "達成率": "29.7%"},
])

# 4. 強化專案
df_project = pd.DataFrame([
    {"單位": "合計", "酒駕件數": 88, "酒駕目標": 150, "酒駕達成率": "58.7%", "闖紅燈件數": 620, "闖紅燈目標": 880, "闖紅燈達成率": "70.5%", "超速件數": 82, "超速目標": 120, "超速達成率": "68.3%", "車不讓人件數": 115, "車不讓人目標": 150, "車不讓人達成率": "76.7%", "大型車件數": 54, "大型車目標": 70, "大型車達成率": "77.1%"},
    {"單位": "聖亭所", "酒駕件數": 14, "酒駕目標": 25, "酒駕達成率": "56.0%", "闖紅燈件數": 98, "闖紅燈目標": 140, "闖紅燈達成率": "70.0%", "超速件數": 12, "超速目標": 20, "超速達成率": "60.0%", "車不讓人件數": 18, "車不讓人目標": 25, "車不讓人達成率": "72.0%", "大型車件數": 8, "大型車目標": 10, "大型車達成率": "80.0%"},
    {"單位": "龍潭所", "酒駕件數": 20, "酒駕目標": 30, "酒駕達成率": "66.7%", "闖紅燈件數": 135, "闖紅燈目標": 180, "闖紅燈達成率": "75.0%", "超速件數": 18, "超速目標": 25, "超速達成率": "72.0%", "車不讓人件數": 26, "車不讓人目標": 30, "車不讓人達成率": "86.7%", "大型車件數": 11, "大型車目標": 15, "大型車達成率": "73.3%"},
    {"單位": "中興所", "酒駕件數": 15, "酒駕目標": 25, "酒駕達成率": "60.0%", "闖紅燈件數": 105, "闖紅燈目標": 140, "闖紅燈達成率": "75.0%", "超速件數": 14, "超速目標": 20, "超速達成率": "70.0%", "車不讓人件數": 19, "車不讓人目標": 25, "車不讓人達成率": "76.0%", "大型車件數": 9, "大型車目標": 12, "大型車達成率": "75.0%"},
    {"單位": "石門所", "酒駕件數": 11, "酒駕目標": 20, "酒駕達成率": "55.0%", "闖紅燈件數": 72, "闖紅燈目標": 110, "闖紅燈達成率": "65.5%", "超速件數": 10, "超速目標": 15, "超速達成率": "66.7%", "車不讓人件數": 14, "車不讓人目標": 20, "車不讓人達成率": "70.0%", "大型車件數": 7, "大型車目標": 10, "大型車達成率": "70.0%"},
    {"單位": "高平所", "酒駕件數": 10, "酒駕目標": 20, "酒駕達成率": "50.0%", "闖紅燈件數": 68, "闖紅燈目標": 110, "闖紅燈達成率": "61.8%", "超速件數": 9, "超速目標": 15, "超速達成率": "60.0%", "車不讓人件數": 13, "車不讓人目標": 20, "車不讓人達成率": "65.0%", "大型車件數": 6, "大型車目標": 10, "大型車達成率": "60.0%"},
    {"單位": "三和所", "酒駕件數": 4, "酒駕目標": 10, "酒駕達成率": "40.0%", "闖紅燈件數": 32, "闖紅燈目標": 60, "闖紅燈達成率": "53.3%", "超速件數": 4, "超速目標": 10, "超速達成率": "40.0%", "車不讓人件數": 6, "車不讓人目標": 10, "車不讓人達成率": "60.0%", "大型車件數": 3, "大型車目標": 5, "大型車達成率": "60.0%"},
    {"單位": "交通分隊", "酒駕件數": 14, "酒駕目標": 20, "酒駕達成率": "70.0%", "闖紅燈件數": 110, "闖紅燈目標": 140, "闖紅燈達成率": "78.6%", "超速件數": 15, "超速目標": 15, "超速達成率": "100.0%", "車不讓人件數": 19, "車不讓人目標": 20, "車不讓人達成率": "95.0%", "大型車件數": 10, "大型車目標": 8, "大型車達成率": "125.0%"},
])

# 5. 超載取締統計
df_overload = pd.DataFrame([
    {"統計期間": "合計", "本期(0901~0907)": 4, "本年累計(0101~0907)": 86, "去年累計(0101~0907)": 78, "比較": 8, "目標值": 127, "達成率": "68%"},
    {"統計期間": "聖亭所", "本期(0901~0907)": 1, "本年累計(0101~0907)": 15, "去年累計(0101~0907)": 12, "比較": 3, "目標值": 20, "達成率": "75%"},
    {"統計期間": "龍潭所", "本期(0901~0907)": 1, "本年累計(0101~0907)": 21, "去年累計(0101~0907)": 18, "比較": 3, "目標值": 27, "達成率": "78%"},
    {"統計期間": "中興所", "本期(0901~0907)": 0, "本年累計(0101~0907)": 14, "去年累計(0101~0907)": 13, "比較": 1, "目標值": 20, "達成率": "70%"},
    {"統計期間": "石門所", "本期(0901~0907)": 1, "本年累計(0101~0907)": 12, "去年累計(0101~0907)": 10, "比較": 2, "目標值": 16, "達成率": "75%"},
    {"統計期間": "高平所", "本期(0901~0907)": 0, "本年累計(0101~0907)": 9, "去年累計(0101~0907)": 8, "比較": 1, "目標值": 14, "達成率": "64%"},
    {"統計期間": "三和所", "本期(0901~0907)": 0, "本年累計(0101~0907)": 4, "去年累計(0101~0907)": 4, "比較": 0, "目標值": 8, "達成率": "50%"},
    {"統計期間": "交通分隊", "本期(0901~0907)": 1, "本年累計(0101~0907)": 11, "去年累計(0101~0907)": 13, "比較": -2, "目標值": 22, "達成率": "50%"},
])

# 6. 靜桃計畫
df_jingtao = pd.DataFrame([
    {"單位": "合計", "本期(22-06時)": 12, "本期(06-22時)": 8, "累計(22-06時)": 184, "累計(06-22時)": 142, "專案總計": 326},
    {"單位": "聖亭所", "本期(22-06時)": 2, "本期(06-22時)": 1, "累計(22-06時)": 32, "累計(06-22時)": 24, "專案總計": 56},
    {"單位": "龍潭所", "本期(22-06時)": 3, "本期(06-22時)": 2, "累計(22-06時)": 48, "累計(06-22時)": 38, "專案總計": 86},
    {"單位": "中興所", "本期(22-06時)": 2, "本期(06-22時)": 1, "累計(22-06時)": 28, "累計(06-22時)": 22, "專案總計": 50},
    {"單位": "石門所", "本期(22-06時)": 2, "本期(06-22時)": 1, "累計(22-06時)": 24, "累計(06-22時)": 18, "專案總計": 42},
    {"單位": "高平所", "本期(22-06時)": 1, "本期(06-22時)": 1, "累計(22-06時)": 18, "累計(06-22時)": 14, "專案總計": 32},
    {"單位": "三和所", "本期(22-06時)": 0, "本期(06-22時)": 0, "累計(22-06時)": 8, "累計(06-22時)": 6, "專案總計": 14},
    {"單位": "交通分隊", "本期(22-06時)": 2, "本期(06-22時)": 2, "累計(22-06時)": 26, "累計(06-22時)": 20, "專案總計": 46},
])

# 7. 科技執法
df_tech = pd.DataFrame([
    {"排名": "第 1 名", "路段名稱": "中豐路與大昌路口", "違規態樣": "闖紅燈/未依標誌行駛", "舉發件數": 1248},
    {"排名": "第 2 名", "路段名稱": "大昌路二段與五福街口", "違規態樣": "闖紅燈/不禮讓行人", "舉發件數": 892},
    {"排名": "第 3 名", "路段名稱": "北龍路與龍元路口", "違規態樣": "不停讓行人", "舉發件數": 645},
    {"排名": "第 4 名", "路段名稱": "中正路三坑段 580 號前", "違規態樣": "嚴重超速 (區間測速)", "舉發件數": 512},
    {"排名": "第 5 名", "路段名稱": "中興路九龍段 320 號前", "違規態樣": "闖紅燈/超速", "舉發件數": 438},
    {"排名": "第 6 名", "路段名稱": "福龍路二段與龍平路口", "違規態樣": "未依規定轉彎", "舉發件數": 386},
    {"排名": "第 7 名", "路段名稱": "龍晨路與金龍路口", "違規態樣": "闖紅燈", "舉發件數": 312},
    {"排名": "第 8 名", "路段名稱": "楊銅路二段 (乳姑山周邊)", "違規態樣": "噪音改裝/超速", "舉發件數": 284},
    {"排名": "第 9 名", "路段名稱": "文化路與石門路口", "違規態樣": "紅燈右轉", "舉發件數": 215},
    {"排名": "第 10 名", "路段名稱": "高原路與高平路口", "違規態樣": "超速", "舉發件數": 182},
])

with st.expander("👀 點擊展開預覽待編譯之 7 大統計表格"):
    t1, t2, t3, t4, t5, t6, t7 = st.tabs(["三項重點", "事故分析", "重大違規", "強化專案", "超載取締", "靜桃計畫", "科技執法"])
    with t1: st.dataframe(df_three, hide_index=True)
    with t2:
        c1, c2 = st.columns(2)
        c1.dataframe(df_a1, hide_index=True)
        c2.dataframe(df_a2, hide_index=True)
    with t3: st.dataframe(df_major, hide_index=True)
    with t4: st.dataframe(df_project, hide_index=True)
    with t5: st.dataframe(df_overload, hide_index=True)
    with t6: st.dataframe(df_jingtao, hide_index=True)
    with t7: st.dataframe(df_tech, hide_index=True)

st.write("")

# ==========================================
# 4. 全套簡報一鍵生成發送
# ==========================================
if st.button("🚀 立即由零全自動生成【全套 8~15 頁會報簡報】", type="primary"):
    slides_svc = get_slides_service()
    drive_svc = get_drive_service()

    if not slides_svc:
        st.error("❌ 無法初始化 Google Slides 服務，請確認 secrets.toml 設定。")
    else:
        with st.spinner("正在呼叫 Google Slides API 編譯全套專案統計並繪製原生表格..."):
            try:
                builder = ComprehensiveSlidesBuilder(slides_svc, drive_svc)

                # 1. 建立簡報母體
                builder.create_presentation(title=report_title, parent_folder_id=target_folder)

                # 2. P.1 封面頁
                builder.add_cover_slide(
                    main_title="桃園市政府警察局龍潭分局\n交通執法成效與事故防制數據分析報告",
                    subtitle="週次主管會報專案報告",
                    date_range_str="115 年 9 月 1 日起至本期止"
                )

                # 3. P.2 取締三項重點違規專案
                builder.add_table_slide(
                    slide_title="取締三項重點違規專案績效統計表",
                    df=df_three,
                    subtitle="本期新增 vs 累計（排除科技執法與警備隊）",
                    footnote="統計起日：115 年 9 月 1 日；三項重點包含：闖紅燈、逆向行駛、不停讓行人。"
                )

                # 4. P.3 交通事故成效分析 (雙表並排)
                builder.add_side_by_side_tables(
                    slide_title="交通事故分析 (A1 類死亡 vs A2 類受傷)",
                    df_left=df_a1, title_left="📊 A1 類交通事故死亡統計",
                    df_right=df_a2, title_right="🚑 A2 類交通事故受傷人數統計",
                    subtitle="本期 vs 本年累計及去年同期比較"
                )

                # 5. P.4 重大交通違規績效 (總表)
                builder.add_table_slide(
                    slide_title="重大交通違規取締績效統計表 (總表)",
                    df=df_major,
                    subtitle="攔停／逕舉統計及年度目標達成率",
                    footnote="重大交通違規指：「酒駕」、「闖紅燈」、「嚴重超速」、「逆向行駛」、「轉彎未依規定」、「蛇行惡意逼車」及「不暫停讓行人」。"
                )

                # 6. P.5 強化交通安全執法專案
                builder.add_table_slide(
                    slide_title="強化交通安全執法專案勤務取締件數統計表",
                    df=df_project,
                    subtitle="專案六大取締項目指標進度",
                    footnote="六大項目：酒後駕車、闖紅燈、嚴重超速、車不讓人、行人違規及大型車違規。"
                )

                # 7. P.6 取締超載違規件數統計
                builder.add_table_slide(
                    slide_title="取締超載違規件數統計表",
                    df=df_overload,
                    subtitle="本期 vs 本年累計及達成率",
                    footnote="本期定義：係指該期昱通系統入案件數；以年底達成率 100% 為基準。"
                )

                # 8. P.7 「靜桃計畫」大執法專案
                builder.add_table_slide(
                    slide_title="「靜桃計畫」大執法專案取締績效統計表",
                    df=df_jingtao,
                    subtitle="夜間 22-06 時 vs 日間 06-22 時時段分流",
                    footnote="包含通報環保局檢驗及現場舉發噪音改裝車輛案件。"
                )

                # 9. P.8 科技執法路段排行 (Top 10)
                builder.add_table_slide(
                    slide_title="科技執法設備舉發成效路段排行 (Top 10)",
                    df=df_tech,
                    subtitle="熱點路段違規樣態與案件量分析",
                    footnote="統計包含轄內固定桿、路口多功能科技執法及區間測速設備入案件數。"
                )

                # 10. 選配 P.9~P.15：重大違規 7 大項細表
                if include_details:
                    detail_cats = ["酒駕", "闖紅燈", "嚴重超速", "逆向行駛", "轉彎未依規定", "蛇行惡意逼車", "不暫停讓行人"]
                    for cat in detail_cats:
                        # 模擬細項資料列
                        df_cat_mock = pd.DataFrame([
                            {"單位": "合計", "今年攔停": 12, "今年逕舉": 38, "今年合計": 50, "去年攔停": 10, "去年逕舉": 35, "去年合計": 45, "增減比較": 5},
                            {"單位": "聖亭所", "今年攔停": 2, "今年逕舉": 6, "今年合計": 8, "去年攔停": 2, "去年逕舉": 5, "去年合計": 7, "增減比較": 1},
                            {"單位": "龍潭所", "今年攔停": 3, "今年逕舉": 10, "今年合計": 13, "去年攔停": 2, "去年逕舉": 9, "去年合計": 11, "增減比較": 2},
                            {"單位": "中興所", "今年攔停": 2, "今年逕舉": 7, "今年合計": 9, "去年攔停": 2, "去年逕舉": 6, "去年合計": 8, "增減比較": 1},
                            {"單位": "石門所", "今年攔停": 2, "今年逕舉": 5, "今年合計": 7, "去年攔停": 1, "去年逕舉": 5, "去年合計": 6, "增減比較": 1},
                            {"單位": "高平所", "今年攔停": 1, "今年逕舉": 4, "今年合計": 5, "去年攔停": 1, "去年逕舉": 4, "去年合計": 5, "增減比較": 0},
                            {"單位": "三和所", "今年攔停": 0, "今年逕舉": 2, "今年合計": 2, "去年攔停": 0, "去年逕舉": 2, "去年合計": 2, "增減比較": 0},
                            {"單位": "交通分隊", "今年攔停": 2, "今年逕舉": 4, "今年合計": 6, "去年攔停": 2, "去年逕舉": 4, "去年合計": 6, "增減比較": 0},
                        ])
                        builder.add_table_slide(
                            slide_title=f"重大違規專項分析：【{cat}】統計表",
                            df=df_cat_mock,
                            subtitle="本年累計 vs 去年累計同期比較",
                            footnote=f"專項取締條款依據道路交通管理處罰條例相關法條。"
                        )

                # 發送 API 請求
                final_url = builder.execute_build()

                st.balloons()
                st.success("🎉 全套 Google Slides 簡報已全自動建構完畢！")
                st.markdown(
                    f"### 📑 簡報入口：\n"
                    f"👉 **[點此直接開啟全新會報簡報]({final_url})**\n\n"
                    f"簡報已包含深藍封面、全 7 大業務原生統計表格與雙表並排，並自動存入您的雲端資料夾。"
                )

            except Exception as e:
                st.error(f"❌ 建立簡報失敗：{e}")
