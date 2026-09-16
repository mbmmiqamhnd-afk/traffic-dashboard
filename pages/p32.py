import uuid
from datetime import datetime
import pandas as pd
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
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

st.title("📽️ 全方位執法數據簡報直出中心（免試算表、完整 7 大統計）")
st.caption("🚀 畫布清空重繪機制：由 Python 從零動態繪製全新 8 頁投影片並整批覆蓋，完全不使用舊物件搜尋替換，徹底解決 403 空間不足問題。")

# ==========================================
# 1. Google 服務連線層與常數設定
# ==========================================
GCP_CREDS = dict(st.secrets.get("gcp_service_account", {}))
SERVICE_ACCOUNT_EMAIL = GCP_CREDS.get("client_email", "streamlit-bot@streamlit-sheets-482909.iam.gserviceaccount.com")

# 已鎖定您的專屬簡報容器 ID
TARGET_PRESENTATION_ID = "1h2QNNI8SLvjNEBmky7IWv9ZGBbKsLvV1UDkYJWcOeWU"

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
        st.info(f"🔑 **執行服務帳號：** `{SERVICE_ACCOUNT_EMAIL}`")
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

        # 主標題方塊
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

        # 副標題方塊
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

    def add_table_slide(self, slide_title: str, df: pd.DataFrame, subtitle: str = "", footnote: str = "", highlight_below_target: float = None):
        """【標準表格頁】自動計算排版、支援合計列淡藍底色高亮、備註及落後指標紅字預警"""
        slide_id = f"s_{uuid.uuid4().hex[:8]}"
        title_id = f"t_{uuid.uuid4().hex[:8]}"
        table_id = f"tbl_{uuid.uuid4().hex[:8]}"

        self.requests.append({
            "createSlide": {
                "objectId": slide_id,
                "slideLayoutReference": {"predefinedLayout": "BLANK"}
            }
        })

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
        font_size = 7.5 if num_cols >= 12 else (8.5 if num_cols >= 8 else 10)

        tbl_top = 55
        tbl_height = min(295, max(140, num_rows * 24))

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

        # 標題列文字
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

        # 資料列文字
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

        # 設定樣式：合計列淡藍底色 (#EAF2F8)、落後項目自動標紅
        for r_idx in range(num_rows):
            is_header = (r_idx == 0)
            is_total_row = (r_idx == 1 and str(df.iloc[0].values[0]).strip() in ["合計", "總計"])

            if is_total_row:
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
                is_bold = (is_header or is_total_row)

                # 智慧色彩預警：達成率低於目前標準或進度負值時標為紅色
                if not is_header:
                    cell_val = str(df.iloc[r_idx - 1, c_idx]).strip()
                    col_name = str(df.columns[c_idx])
                    
                    if "落後" in cell_val or cell_val.startswith("-"):
                        fg = {"red": 0.85, "green": 0.0, "blue": 0.0}
                        is_bold = True
                    elif highlight_below_target is not None and "達成率" in col_name:
                        try:
                            rate_num = float(cell_val.replace("%", "").strip())
                            if rate_num < highlight_below_target:
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
            self.requests.append({"insertText": {"objectId": fn_id, "text": f"📝 {footnote}", "insertionIndex": 0}})
            self.requests.append({
                "updateTextStyle": {
                    "objectId": fn_id,
                    "style": {
                        "fontFamily": "Microsoft JhengHei",
                        "fontSize": {"magnitude": 9.5, "unit": "PT"},
                        "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.4, "green": 0.4, "blue": 0.4}}}
                    },
                    "textRange": {"type": "ALL"},
                    "fields": "fontFamily,fontSize,foregroundColor"
                }
            })

    def add_side_by_side_tables(self, slide_title: str, df_left: pd.DataFrame, title_left: str, df_right: pd.DataFrame, title_right: str, subtitle: str = ""):
        """【雙表並排頁】A1 與 A2 左右對稱對照"""
        slide_id = f"s_dual_{uuid.uuid4().hex[:8]}"
        title_id = f"t_dual_{uuid.uuid4().hex[:8]}"

        self.requests.append({
            "createSlide": {
                "objectId": slide_id,
                "slideLayoutReference": {"predefinedLayout": "BLANK"}
            }
        })

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

        def build_one_tbl(df, t_label, x_offset, w):
            t_box_id = f"subt_{uuid.uuid4().hex[:8]}"
            tbl_id = f"tbl_{uuid.uuid4().hex[:8]}"

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
                    "style": {"fontFamily": "Microsoft JhengHei", "fontSize": {"magnitude": 11, "unit": "PT"}, "bold": True, "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 0.2, "green": 0.3, "blue": 0.45}}}},
                    "textRange": {"type": "ALL"},
                    "fields": "fontFamily,fontSize,bold,foregroundColor"
                }
            })

            n_rows, n_cols = len(df) + 1, len(df.columns)
            h = min(280, n_rows * 27)
            self.requests.append({
                "createTable": {
                    "objectId": tbl_id,
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {"width": {"magnitude": w, "unit": "PT"}, "height": {"magnitude": h, "unit": "PT"}},
                        "transform": {"scaleX": 1, "scaleY": 1, "translateX": x_offset, "translateY": 75, "unit": "PT"}
                    },
                    "rows": n_rows,
                    "columns": n_cols
                }
            })
            for c, name in enumerate(df.columns):
                self.requests.append({"insertText": {"objectId": tbl_id, "cellLocation": {"rowIndex": 0, "columnIndex": c}, "text": str(name), "insertionIndex": 0}})
            for r, row in df.iterrows():
                for c, val in enumerate(row):
                    self.requests.append({"insertText": {"objectId": tbl_id, "cellLocation": {"rowIndex": r + 1, "columnIndex": c}, "text": str(val) if pd.notna(val) else "—", "insertionIndex": 0}})
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
                            "textRange": {"type": "ALL"},
                            "fields": "fontFamily,fontSize,bold,foregroundColor"
                        }
                    })

        build_one_tbl(df_left, title_left, x_offset=25, w=325)
        build_one_tbl(df_right, title_right, x_offset=365, w=330)

    def wipe_old_slides(self):
        """抹除所有舊頁面，僅保留剛編譯完成的 8 頁"""
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
# 3. 數據準備層（含動態「目前應達成率」計算）
# ==========================================

# ── 動態計算「目前應達成率 (以年底 100% 為基準)」 ──
now_dt = datetime.now()
roc_year = now_dt.year - 1911
month = now_dt.month
day = now_dt.day
day_of_year = now_dt.timetuple().tm_yday
is_leap = (now_dt.year % 4 == 0 and now_dt.year % 100 != 0) or (now_dt.year % 400 == 0)
total_days = 366 if is_leap else 365
current_expected_rate = round((day_of_year / total_days) * 100, 1)

overload_footnote_dynamic = (
    f"本期定義：係指該期昱通系統入案件數；以年底達成率100%為基準，"
    f"統計截至 {roc_year}年{month:02d}月{day:02d}日(入案日期)目前應達成率為 {current_expected_rate:.1f}%"
)
overload_subtitle_dynamic = f"本期 vs 本年累計 ｜ 目前應達成率：{current_expected_rate:.1f}%"

# 1. 三項重點違規
df_three = st.session_state.get("df_three", pd.DataFrame([
    {"單位": "合計", "闖紅燈(本期)": 15, "闖紅燈(累計)": 128, "逆向(本期)": 8, "逆向(累計)": 72, "不停讓行人(本期)": 5, "不停讓行人(累計)": 43, "三項合計(本期)": 28, "三項合計(累計)": 243},
    {"單位": "聖亭所", "闖紅燈(本期)": 3, "闖紅燈(累計)": 24, "逆向(本期)": 1, "逆向(累計)": 15, "不停讓行人(本期)": 1, "不停讓行人(累計)": 8, "三項合計(本期)": 5, "三項合計(累計)": 47},
    {"單位": "龍潭所", "闖紅燈(本期)": 4, "闖紅燈(累計)": 38, "逆向(本期)": 2, "逆向(累計)": 20, "不停讓行人(本期)": 2, "不停讓行人(累計)": 14, "三項合計(本期)": 8, "三項合計(累計)": 72},
    {"單位": "中興所", "闖紅燈(本期)": 2, "闖紅燈(累計)": 21, "逆向(本期)": 1, "逆向(累計)": 12, "不停讓行人(本期)": 1, "不停讓行人(累計)": 7, "三項合計(本期)": 4, "三項合計(累計)": 40},
    {"單位": "石門所", "闖紅燈(本期)": 2, "闖紅燈(累計)": 16, "逆向(本期)": 1, "逆向(累計)": 9, "不停讓行人(本期)": 0, "不停讓行人(累計)": 5, "三項合計(本期)": 3, "三項合計(累計)": 30},
    {"單位": "高平所", "闖紅燈(本期)": 1, "闖紅燈(累計)": 12, "逆向(本期)": 1, "逆向(累計)": 7, "不停讓行人(本期)": 0, "不停讓行人(累計)": 3, "三項合計(本期)": 2, "三項合計(累計)": 22},
    {"單位": "三和所", "闖紅燈(本期)": 0, "闖紅燈(累計)": 4, "逆向(本期)": 0, "逆向(累計)": 3, "不停讓行人(本期)": 0, "不停讓行人(累計)": 1, "三項合計(本期)": 0, "三項合計(累計)": 8},
    {"單位": "交通分隊", "闖紅燈(本期)": 3, "闖紅燈(累計)": 13, "逆向(本期)": 2, "逆向(累計)": 6, "不停讓行人(本期)": 1, "不停讓行人(累計)": 5, "三項合計(本期)": 6, "三項合計(累計)": 24},
]))

# 2. 交通事故 (A1 / A2)
df_a1 = st.session_state.get("df_a1", pd.DataFrame([
    {"統計期間": "合計", "本期": 0, "本年累計": 3, "去年同期": 4, "增減比較": -1},
    {"統計期間": "聖亭所", "本期": 0, "本年累計": 1, "去年同期": 1, "增減比較": 0},
    {"統計期間": "龍潭所", "本期": 0, "本年累計": 1, "去年同期": 2, "增減比較": -1},
    {"統計期間": "中興所", "本期": 0, "本年累計": 0, "去年同期": 0, "增減比較": 0},
    {"統計期間": "石門所", "本期": 0, "本年累計": 1, "去年同期": 1, "增減比較": 0},
    {"統計期間": "高平所", "本期": 0, "本年累計": 0, "去年同期": 0, "增減比較": 0},
    {"統計期間": "三和所", "本期": 0, "本年累計": 0, "去年同期": 0, "增減比較": 0},
]))

df_a2 = st.session_state.get("df_a2", pd.DataFrame([
    {"統計期間": "合計", "本期": 32, "前期": 35, "本年累計": 1284, "去年累計": 1390, "比較": -106, "增減比例": "-7.63%"},
    {"統計期間": "聖亭所", "本期": 7, "前期": 8, "本年累計": 312, "去年累計": 330, "比較": -18, "增減比例": "-5.45%"},
    {"統計期間": "龍潭所", "本期": 11, "前期": 12, "本年累計": 445, "去年累計": 472, "比較": -27, "增減比例": "-5.72%"},
    {"統計期間": "中興所", "本期": 6, "前期": 7, "本年累計": 268, "抽取": 290, "比較": -22, "增減比例": "-7.59%"},
    {"統計期間": "石門所", "本期": 4, "前期": 5, "本年累計": 142, "去年累計": 160, "比較": -18, "增減比例": "-11.25%"},
    {"統計期間": "高平所", "本期": 3, "前期": 2, "本年累計": 92, "去年累計": 105, "比較": -13, "增減比例": "-12.38%"},
    {"統計期間": "三和所", "本期": 1, "前期": 1, "本年累計": 25, "去年累計": 33, "比較": -8, "增減比例": "-24.24%"},
]))

# 3. 重大交通違規 (總表)
df_major = st.session_state.get("df_major", pd.DataFrame([
    {"單位": "合計", "本期(攔停)": 48, "本期(逕舉)": 152, "本年(攔停)": 1840, "本年(逕舉)": 6420, "去年同期": 7950, "增減比較": 310, "目標值": 18115, "達成率": "45.6%"},
    {"單位": "科技執法", "本期(攔停)": 0, "本期(逕舉)": 88, "本年(攔停)": 0, "本年(逕舉)": 3120, "去年同期": 2900, "增減比較": 220, "目標值": 6006, "達成率": "51.9%"},
    {"單位": "聖亭所", "本期(攔停)": 8, "本期(逕舉)": 12, "本年(攔停)": 340, "本年(逕舉)": 590, "去年同期": 910, "增減比較": 20, "目標值": 1941, "達成率": "47.9%"},
    {"單位": "龍潭所", "本期(攔停)": 12, "本期(逕舉)": 16, "本年(攔停)": 480, "本年(逕舉)": 780, "去年同期": 1210, "增減比較": 50, "目標值": 2588, "達成率": "48.7%"},
    {"單位": "中興所", "本期(攔停)": 7, "本期(逕舉)": 10, "本年(攔停)": 310, "本年(逕舉)": 540, "去年同期": 820, "增減比較": 30, "目標值": 1941, "達成率": "43.8%"},
    {"單位": "石門所", "本期(攔停)": 5, "本期(逕舉)": 8, "本年(攔停)": 210, "本年(逕舉)": 430, "去年同期": 610, "增減比較": 30, "目標值": 1479, "達成率": "43.3%"},
    {"單位": "高平所", "本期(攔停)": 4, "本期(逕舉)": 6, "本年(攔停)": 180, "本年(逕舉)": 380, "去年同期": 540, "增減比較": 20, "目標值": 1294, "達成率": "43.3%"},
    {"單位": "三和所", "本期(攔停)": 2, "本期(逕舉)": 2, "本年(攔停)": 60, "本年(逕舉)": 90, "去年同期": 140, "增減比較": 10, "目標值": 339, "達成率": "44.2%"},
    {"單位": "交通分隊", "本期(攔停)": 10, "本期(逕舉)": 10, "本年(攔停)": 260, "本年(逕舉)": 490, "去年同期": 820, "增減比較": -70, "目標值": 2526, "達成率": "29.7%"},
]))

# 4. 強化專案
df_project = st.session_state.get("df_project", pd.DataFrame([
    {"單位": "合計", "酒駕件數": 88, "酒駕目標": 150, "酒駕達成率": "58.7%", "闖紅燈件數": 620, "闖紅燈目標": 880, "闖紅燈達成率": "70.5%", "超速件數": 82, "超速目標": 120, "超速達成率": "68.3%", "車不讓人件數": 115, "車不讓人目標": 150, "車不讓人達成率": "76.7%", "大型車件數": 54, "大型車目標": 70, "大型車達成率": "77.1%"},
    {"單位": "聖亭所", "酒駕件數": 14, "酒駕目標": 25, "酒駕達成率": "56.0%", "闖紅燈件數": 98, "闖紅燈目標": 140, "闖紅燈達成率": "70.0%", "超速件數": 12, "超速目標": 20, "超速達成率": "60.0%", "車不讓人件數": 18, "車不讓人目標": 25, "車不讓人達成率": "72.0%", "大型車件數": 8, "大型車目標": 10, "大型車達成率": "80.0%"},
    {"單位": "龍潭所", "酒駕件數": 20, "酒駕目標": 30, "酒駕達成率": "66.7%", "闖紅燈件數": 135, "闖紅燈目標": 180, "闖紅燈達成率": "75.0%", "超速件數": 18, "超速目標": 25, "超速達成率": "72.0%", "車不讓人件數": 26, "車不讓人目標": 30, "車不讓人達成率": "86.7%", "大型車件數": 11, "大型車目標": 15, "大型車達成率": "73.3%"},
    {"單位": "中興所", "酒駕件數": 15, "酒駕目標": 25, "酒駕達成率": "60.0%", "闖紅燈件數": 105, "闖紅燈目標": 140, "闖紅燈達成率": "75.0%", "超速件數": 14, "超速目標": 20, "超速達成率": "70.0%", "車不讓人件數": 19, "車不讓人目標": 25, "車不讓人達成率": "76.0%", "大型車件數": 9, "大型車目標": 12, "大型車達成率": "75.0%"},
    {"單位": "石門所", "酒駕件數": 11, "酒駕目標": 20, "酒駕達成率": "55.0%", "闖紅燈件數": 72, "闖紅燈目標": 110, "闖紅燈達成率": "65.5%", "超速件數": 10, "超速目標": 15, "超速達成率": "66.7%", "車不讓人件數": 14, "車不讓人目標": 20, "車不讓人達成率": "70.0%", "大型車件數": 7, "大型車目標": 10, "大型車達成率": "70.0%"},
    {"單位": "高平所", "酒駕件數": 10, "酒駕目標": 20, "酒駕達成率": "50.0%", "闖紅燈件數": 68, "闖紅燈目標": 110, "闖紅燈達成率": "61.8%", "超速件數": 9, "超速目標": 15, "超速達成率": "60.0%", "車不讓人件數": 13, "車不讓人目標": 20, "車不讓人達成率": "65.0%", "大型車件數": 6, "大型車目標": 10, "大型車達成率": "60.0%"},
    {"單位": "三和所", "酒駕件數": 4, "酒駕目標": 10, "酒駕達成率": "40.0%", "闖紅燈件數": 32, "闖紅燈目標": 60, "闖紅燈達成率": "53.3%", "超速件數": 4, "超速目標": 10, "超速達成率": "40.0%", "車不讓人件數": 6, "車不讓人目標": 10, "車不讓人達成率": "60.0%", "大型車件數": 3, "大型車目標": 5, "大型車達成率": "60.0%"},
    {"單位": "交通分隊", "酒駕件數": 14, "酒駕目標": 20, "酒駕達成率": "70.0%", "闖紅燈件數": 110, "闖紅燈目標": 140, "闖紅燈達成率": "78.6%", "超速件數": 15, "超速目標": 15, "超速達成率": "100.0%", "車不讓人件數": 19, "車不讓人目標": 20, "車不讓人達成率": "95.0%", "大型車件數": 10, "大型車目標": 8, "大型車達成率": "125.0%"},
]))

# 5. 超載取締統計（動態精算「達成率」與「進度差距」）
raw_overload = [
    {"統計期間": "合計", "本期": 4, "本年累計": 86, "去年同期": 78, "比較": 8, "目標值": 127},
    {"統計期間": "聖亭所", "本期": 1, "本年累計": 15, "去年同期": 12, "比較": 3, "目標值": 20},
    {"統計期間": "龍潭所", "本期": 1, "本年累計": 21, "去年同期": 18, "比較": 3, "目標值": 27},
    {"統計期間": "中興所", "本期": 0, "本年累計": 14, "去年同期": 13, "比較": 1, "目標值": 20},
    {"統計期間": "石門所", "本期": 1, "本年累計": 12, "去年同期": 10, "比較": 2, "目標值": 16},
    {"統計期間": "高平所", "本期": 0, "本年累計": 9, "去年同期": 8, "比較": 1, "目標值": 14},
    {"統計期間": "三和所", "本期": 0, "本年累計": 4, "去年同期": 4, "比較": 0, "目標值": 8},
    {"統計期間": "交通分隊", "本期": 1, "本年累計": 11, "去年同期": 13, "比較": -2, "目標值": 22},
]

# 若 session_state 內已有真實數據則取用，否則以 raw_overload 進行動態結算
if "df_overload" in st.session_state and isinstance(st.session_state["df_overload"], pd.DataFrame):
    df_overload = st.session_state["df_overload"].copy()
else:
    df_overload = pd.DataFrame(raw_overload)
    # 動態計算「達成率」與「進度差距」
    rates = []
    diffs = []
    for _, r in df_overload.iterrows():
        tgt = float(r["目標值"])
        cumu = float(r["本年累計"])
        if tgt > 0:
            calc_rate = round((cumu / tgt) * 100, 1)
            diff_from_target = round(calc_rate - current_expected_rate, 1)
            rates.append(f"{calc_rate:.0f}%")
            if diff_from_target >= 0:
                diffs.append(f"🟢 達標 (+{diff_from_target:.1f}%)")
            else:
                diffs.append(f"🔴 落後 ({diff_from_target:.1f}%)")
        else:
            rates.append("—")
            diffs.append("—")
            
    df_overload["達成率"] = rates
    df_overload["進度評比"] = diffs

# 6. 靜桃計畫
df_jingtao = st.session_state.get("df_jingtao", pd.DataFrame([
    {"單位": "合計", "本期(22-06時)": 12, "本期(06-22時)": 8, "累計(22-06時)": 184, "累計(06-22時)": 142, "專案總計": 326},
    {"單位": "聖亭所", "本期(22-06時)": 2, "本期(06-22時)": 1, "累計(22-06時)": 32, "累計(06-22時)": 24, "專案總計": 56},
    {"單位": "龍潭所", "本期(22-06時)": 3, "本期(06-22時)": 2, "累計(22-06時)": 48, "累計(06-22時)": 38, "專案總計": 86},
    {"單位": "中興所", "本期(22-06時)": 2, "本期(06-22時)": 1, "累計(22-06時)": 28, "累計(06-22時)": 22, "專案總計": 50},
    {"單位": "石門所", "本期(22-06時)": 2, "本期(06-22時)": 1, "累計(22-06時)": 24, "累計(06-22時)": 18, "專案總計": 42},
    {"單位": "高平所", "本期(22-06時)": 1, "本期(06-22時)": 1, "累計(22-06時)": 18, "累計(06-22時)": 14, "專案總計": 32},
    {"單位": "三和所", "本期(22-06時)": 0, "本期(06-22時)": 0, "累計(22-06時)": 8, "累計(06-22時)": 6, "專案總計": 14},
    {"單位": "交通分隊", "本期(22-06時)": 2, "本期(06-22時)": 2, "累計(22-06時)": 26, "累計(06-22時)": 20, "專案總計": 46},
]))

# 7. 科技執法
df_tech = st.session_state.get("df_tech", pd.DataFrame([
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
]))

# ==========================================
# 4. 前端檢視區
# ==========================================
with st.expander("👀 點擊展開預覽 7 大統計業務表格內容（含超載動態達成率）"):
    t1, t2, t3, t4, t5, t6, t7 = st.tabs(["三項重點", "事故分析", "重大違規", "強化專案", "超載取締 (新)", "靜桃計畫", "科技執法"])
    with t1: st.dataframe(df_three, hide_index=True)
    with t2:
        c1, c2 = st.columns(2)
        c1.dataframe(df_a1, hide_index=True)
        c2.dataframe(df_a2, hide_index=True)
    with t3: st.dataframe(df_major, hide_index=True)
    with t4: st.dataframe(df_project, hide_index=True)
    with t5: 
        st.caption(f"🎯 **{overload_subtitle_dynamic}**")
        st.dataframe(df_overload, hide_index=True)
        st.caption(f"📝 {overload_footnote_dynamic}")
    with t6: st.dataframe(df_jingtao, hide_index=True)
    with t7: st.dataframe(df_tech, hide_index=True)

st.write("")

# ==========================================
# 5. 啟動重繪
# ==========================================
if st.button("🚀 啟動畫布清空重繪：全新產出【完整 8 頁會報簡報】", type="primary"):
    slides_svc = get_slides_service()

    if not slides_svc:
        st.error("❌ 無法初始化 Google Slides 服務，請確認 secrets.toml 設定。")
    else:
        with st.spinner("正在讀取簡報畫布、動態編譯 8 頁投影片並整批覆蓋..."):
            try:
                builder = ComprehensiveSlidesBuilder(slides_svc, TARGET_PRESENTATION_ID)

                # 1. 記錄舊頁面 ID
                builder.prepare_canvas()

                # 2. P.1 封面頁
                builder.add_cover_slide(
                    main_title="桃園市政府警察局龍潭分局\n交通執法成效與事故防制數據分析報告",
                    subtitle="週次主管會報專案報告",
                    date_range_str=f"115 年 9 月 1 日起至 {month:02d}月{day:02d}日 止"
                )

                # 3. P.2 三項重點違規專案
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

                # 7. P.6 取締超載違規件數統計（✅ 動態帶入「目前應達成率」與落後紅字標註）
                builder.add_table_slide(
                    slide_title="取締超載違規件數統計表",
                    df=df_overload,
                    subtitle=overload_subtitle_dynamic,
                    footnote=overload_footnote_dynamic,
                    highlight_below_target=current_expected_rate
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

                # 10. 徹底清除原有舊頁面
                builder.wipe_old_slides()

                # 11. 整批傳送執行
                final_url = builder.execute_build()

                st.balloons()
                st.success("🎉 全套 Google Slides 簡報已全自動重繪完成！")
                st.markdown(
                    f"### 📑 簡報入口：\n"
                    f"👉 **[點此直接開啟全新會報簡報]({final_url})**\n\n"
                    f"✅ **更新亮點**：P.6 超載統計表已精確計算「目前應達成率（{current_expected_rate:.1f}%）」與進度差距，落後單位將在投影片中自動以紅字突顯，頁尾亦動態生成完整法定備註說明。"
                )

            except HttpError as e:
                st.error(f"❌ Google API 請求失敗：{e}\n\n*提示：請確認簡報是否已共用給 `{SERVICE_ACCOUNT_EMAIL}` 並設定為「編輯者」。*")
            except Exception as e:
                st.error(f"❌ 建立簡報失敗：{e}")
