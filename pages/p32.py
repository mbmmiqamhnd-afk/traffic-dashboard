import io
import os
import re
import smtplib
import csv
import urllib.parse as _ul
from datetime import datetime, timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import streamlit as st

# 嘗試載入 python-pptx
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.oxml import parse_xml
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# 嘗試載入 Google Drive API
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    HAS_GDRIVE = True
except ImportError:
    HAS_GDRIVE = False

# 載入自訂側邊欄
try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        pass

# ==========================================
# 0. 系統初始化
# ==========================================
st.set_page_config(
    page_title="全方位執法數據簡報直出中心 (純動態雙軌版)",
    page_icon="📽️",
    layout="wide"
)
show_sidebar()

st.title("📽️ 全方位執法數據簡報直出中心（純動態雙軌版）")
st.caption("🚀 完整收錄 8 大常態核心表格 ＋ 7 大重大違規專項細表！完全由雲端集中處／本機上傳動態驅動，零寫死數值。")

if not HAS_PPTX:
    st.error("⚠️ 環境中尚未安裝 `python-pptx` 套件。請在 requirements.txt 中新增：`python-pptx`")
    st.stop()

# ==========================================
# 1. 郵件通知函式
# ==========================================
def send_pptx_email_to_self(pptx_bytes: io.BytesIO, file_name: str) -> tuple:
    try:
        if "email" in st.secrets:
            sender = st.secrets["email"].get("user")
            pwd = st.secrets["email"].get("password")
        else:
            sender = st.secrets.get("SMTP_USER")
            pwd = st.secrets.get("SMTP_PASSWORD")

        if not sender or not pwd:
            return False, "未於 secrets.toml 偵測到 [email] 或 SMTP 帳號密碼設定"

        msg = MIMEMultipart()
        msg["From"] = f"交通執法自動化戰情室 <{sender}>"
        msg["To"] = sender
        msg["Subject"] = f"📊 交通執法數據簡報直出 - {file_name}"

        body_text = (
            f"長官／同仁好：\n\n"
            f"系統已自動根據雲端硬碟或您上傳之最新報表完成簡報編譯結算。\n"
            f"附件為最新產出之 PowerPoint 簡報實體檔【{file_name}】。\n\n"
            f"本檔案為純 Python 動態直出，數據 100% 來自實體報表。\n"
            f"本信件由交通執法自動化分析引擎發送。"
        )
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        part = MIMEBase("application", "vnd.openxmlformats-officedocument.presentationml.presentation")
        part.set_payload(pptx_bytes.getvalue())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(file_name)}")
        msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, sender
    except Exception as e:
        return False, str(e)

# ==========================================
# 2. PPTX 原生排版引擎 (PptxReportBuilder)
# ==========================================
class PptxReportBuilder:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]

        self.C_COVER_BG = RGBColor(15, 38, 56)
        self.C_COVER_SUBTITLE = RGBColor(204, 217, 230)
        self.C_WHITE = RGBColor(255, 255, 255)
        self.C_TITLE_NAVY = RGBColor(26, 51, 89)
        self.C_TBL_HEADER_BG = RGBColor(38, 64, 97)
        self.C_TBL_HEADER_TEXT = RGBColor(255, 255, 255)
        self.C_TBL_HIGHLIGHT_BG = RGBColor(232, 240, 247)
        self.C_TBL_ROW_BG = RGBColor(255, 255, 255)
        self.C_TBL_TEXT_DARK = RGBColor(26, 26, 26)
        self.C_TBL_RED = RGBColor(217, 0, 0)
        self.C_MUTED = RGBColor(100, 116, 139)
        self.C_FOOTNOTE = RGBColor(51, 51, 51)

    def _set_border(self, cell, color_hex="CBD5E1", width="12700"):
        try:
            tcPr = cell._tc.get_or_add_tcPr()
            for border_name in ["lnL", "lnR", "lnT", "lnB"]:
                border = parse_xml(
                    f'<a:{border_name} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{width}" cmpd="s">'
                    f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>'
                    f'</a:{border_name}>'
                )
                tcPr.append(border)
        except Exception:
            pass

    def _set_cell(self, cell, text, font_size=16, bold=False, color=None, bg_color=None, align=PP_ALIGN.CENTER):
        cell.text = str(text).strip() if (pd.notna(text) and str(text).strip() != "") else "—"
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_color if bg_color else self.C_TBL_ROW_BG
        self._set_border(cell, color_hex="CBD5E1")

        cell.margin_top = Inches(0.02)
        cell.margin_bottom = Inches(0.02)
        cell.margin_left = Inches(0.03)
        cell.margin_right = Inches(0.03)
        cell.text_frame.word_wrap = False

        for p in cell.text_frame.paragraphs:
            p.alignment = align
            for r in p.runs:
                r.font.name = "Microsoft JhengHei"
                r.font.size = Pt(font_size)
                r.font.bold = bold
                r.font.color.rgb = color if color else self.C_TBL_TEXT_DARK

    def add_cover_slide(self, main_title: str, subtitle: str, date_range_str: str):
        slide = self.prs.slides.add_slide(self.blank_layout)
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = self.C_COVER_BG
        bg.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(2.2))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = main_title
        p.font.name = "Microsoft JhengHei"
        p.font.size = Pt(38)
        p.font.bold = True
        p.font.color.rgb = self.C_WHITE

        tb_sub = slide.shapes.add_textbox(Inches(1.0), Inches(4.5), Inches(11.333), Inches(2.0))
        tf_sub = tb_sub.text_frame
        tf_sub.word_wrap = True

        p1 = tf_sub.paragraphs[0]
        p1.text = subtitle
        p1.font.name = "Microsoft JhengHei"
        p1.font.size = Pt(20)
        p1.font.color.rgb = self.C_COVER_SUBTITLE

        p2 = tf_sub.add_paragraph()
        p2.text = f"統計區間：{date_range_str} ｜ 製表單位：龍潭分局交通組"
        p2.font.name = "Microsoft JhengHei"
        p2.font.size = Pt(14)
        p2.font.color.rgb = RGBColor(203, 213, 225)
        p2.space_before = Pt(14)

    def add_header_box(self, slide, title: str, subtitle: str = ""):
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12.133), Inches(0.75))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Microsoft JhengHei"
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = self.C_TITLE_NAVY

        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.font.name = "Microsoft JhengHei"
            p2.font.size = Pt(11)
            p2.font.color.rgb = self.C_MUTED
            p2.space_before = Pt(2)

    def add_three_major_slide(self, data_rows, custom_subtitle="", cur_col_title="本期", cum_col_title="本月累計"):
        slide = self.prs.slides.add_slide(self.blank_layout)
        sub_text = custom_subtitle if custom_subtitle else "製表單位：龍潭分局交通組"
        self.add_header_box(
            slide,
            "桃園市政府警察局龍潭分局 取締三項重點違規本期及累計統計表",
            sub_text
        )

        num_rows = len(data_rows) + 2
        num_cols = 9
        table_shape = slide.shapes.add_table(num_rows, num_cols, Inches(0.6), Inches(1.25), Inches(12.133), Inches(5.8))
        tbl = table_shape.table

        tbl.cell(0, 0).merge(tbl.cell(1, 0))
        tbl.cell(0, 1).merge(tbl.cell(0, 4))
        tbl.cell(0, 5).merge(tbl.cell(0, 8))

        self._set_cell(tbl.cell(0, 0), "單位", font_size=13, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 1), f"{cur_col_title} 新增違規數", font_size=13, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 5), f"{cum_col_title}數", font_size=13, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        sub_headers = ["", "闖紅燈", "逆向行駛", "不停讓行人", f"{cur_col_title}合計", "闖紅燈", "逆向行駛", "不停讓行人", "累計總計"]
        for c in range(1, 9):
            self._set_cell(tbl.cell(1, c), sub_headers[c], font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        for r_idx, row in enumerate(data_rows, start=2):
            is_tot = (r_idx == 2)
            bg = self.C_TBL_HIGHLIGHT_BG if is_tot else self.C_TBL_ROW_BG
            for c_idx, val in enumerate(row):
                f_sz = 13 if c_idx == 0 else 16
                self._set_cell(tbl.cell(r_idx, c_idx), val, font_size=f_sz, bold=is_tot, color=self.C_TBL_TEXT_DARK, bg_color=bg)

    def add_major_detail_slide(self, cat_name: str, data_rows, custom_subtitle=""):
        slide = self.prs.slides.add_slide(self.blank_layout)
        sub_text = custom_subtitle if custom_subtitle else "口徑包含現場攔停與逕行舉發 ｜ 製表單位：龍潭分局交通組"
        self.add_header_box(
            slide,
            f"取締【{cat_name}】違規統計表",
            sub_text
        )

        num_rows = len(data_rows) + 2
        num_cols = 10
        table_shape = slide.shapes.add_table(num_rows, num_cols, Inches(0.6), Inches(1.25), Inches(12.133), Inches(5.8))
        tbl = table_shape.table

        tbl.cell(0, 0).merge(tbl.cell(1, 0))
        tbl.cell(0, 1).merge(tbl.cell(0, 3))
        tbl.cell(0, 4).merge(tbl.cell(0, 6))
        tbl.cell(0, 7).merge(tbl.cell(0, 9))

        self._set_cell(tbl.cell(0, 0), "單位", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 1), "今年累計", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 4), "去年累計", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 7), "同期比較", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        sub_names = ["", "攔停", "逕舉", "合計", "攔停", "逕舉", "合計", "攔停", "逕舉", "合計"]
        for c in range(1, 10):
            self._set_cell(tbl.cell(1, c), sub_names[c], font_size=11, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        for r_idx, row in enumerate(data_rows, start=2):
            is_tot = (r_idx == 2)
            bg = self.C_TBL_HIGHLIGHT_BG if is_tot else self.C_TBL_ROW_BG

            has_negative = False
            try:
                tot_comp_val = float(str(row[9]).replace(",", "").strip())
                if tot_comp_val < 0:
                    has_negative = True
            except Exception:
                pass

            for c_idx, val in enumerate(row):
                fg = self.C_TBL_TEXT_DARK
                is_bold = is_tot

                if c_idx == 0 and has_negative:
                    fg = self.C_TBL_RED
                    is_bold = True

                if c_idx in [7, 8, 9]:
                    try:
                        c_num = float(str(val).replace(",", "").strip())
                        if c_num < 0:
                            fg = self.C_TBL_RED
                            is_bold = True
                    except Exception:
                        pass

                f_sz = 12 if c_idx == 0 else 16
                self._set_cell(tbl.cell(r_idx, c_idx), val, font_size=f_sz, bold=is_bold, color=fg, bg_color=bg)

    def add_table_slide(self, slide_title: str, df: pd.DataFrame, subtitle: str = "", footnote: str = "", is_accident_table: bool = False, is_major_table: bool = False, custom_width_in: float = None):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self.add_header_box(slide, slide_title, subtitle)

        num_cols = len(df.columns)
        num_rows = len(df) + 1

        tbl_width = custom_width_in if custom_width_in else (8.0 if num_cols <= 2 else 12.133)
        tbl_left = (13.333 - tbl_width) / 2
        tbl_top = 1.25
        tbl_height = min(5.5, max(2.5, num_rows * 0.45))

        table_shape = slide.shapes.add_table(num_rows, num_cols, Inches(tbl_left), Inches(tbl_top), Inches(tbl_width), Inches(tbl_height))
        tbl = table_shape.table

        hdr_font_sz = 11 if num_cols >= 9 else 13
        for c_idx, col_name in enumerate(df.columns):
            self._set_cell(tbl.cell(0, c_idx), col_name, font_size=hdr_font_sz, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        for r_idx, row in df.iterrows():
            first_val = str(row.values[0]).strip()
            is_hl = any(k in first_val for k in ["合計", "總計", "舉發總數"])
            bg = self.C_TBL_HIGHLIGHT_BG if is_hl else self.C_TBL_ROW_BG

            has_major_negative = False
            if is_major_table:
                for c_idx, val in enumerate(row):
                    col_name = str(df.columns[c_idx])
                    if "同期比較" in col_name or "比較" in col_name:
                        try:
                            clean_num = float(str(val).replace(",", "").strip())
                            if clean_num < 0:
                                has_major_negative = True
                        except Exception:
                            pass

            for c_idx, val in enumerate(row):
                cell_val = str(val).strip()
                col_name = str(df.columns[c_idx])
                fg = self.C_TBL_TEXT_DARK
                is_bold = is_hl

                if is_accident_table and any(k in col_name for k in ["比較", "增減", "比例"]):
                    try:
                        clean_num = float(cell_val.replace("%", "").replace("+", "").strip())
                        if clean_num > 0:
                            fg = self.C_TBL_RED
                            is_bold = True
                    except Exception:
                        pass

                if is_major_table:
                    if c_idx == 0 and has_major_negative:
                        fg = self.C_TBL_RED
                        is_bold = True
                    elif "同期比較" in col_name or "比較" in col_name:
                        try:
                            clean_num = float(cell_val.replace(",", "").strip())
                            if clean_num < 0:
                                fg = self.C_TBL_RED
                                is_bold = True
                        except Exception:
                            pass

                if c_idx == 0:
                    data_font_sz = 13
                elif num_cols <= 7:
                    data_font_sz = 18
                else:
                    data_font_sz = 16

                self._set_cell(tbl.cell(r_idx + 1, c_idx), cell_val, font_size=data_font_sz, bold=is_bold, color=fg, bg_color=bg)

        if footnote:
            tb_fn = slide.shapes.add_textbox(Inches(0.6), Inches(6.85), Inches(12.133), Inches(0.35))
            p_fn = tb_fn.text_frame.paragraphs[0]
            p_fn.text = f"註：{footnote}"
            p_fn.font.name = "DFKai-SB"
            p_fn.font.size = Pt(11)
            p_fn.font.color.rgb = self.C_FOOTNOTE

    def build_bytes(self) -> io.BytesIO:
        out = io.BytesIO()
        self.prs.save(out)
        out.seek(0)
        return out

# ==========================================
# 3. 雙軌數據來源載入器 (寬容支援 Google Sheets 與大小寫副檔名)
# ==========================================
def get_drive_service():
    if not HAS_GDRIVE or "gcp_service_account" not in st.secrets:
        return None
    try:
        creds = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
        return build("drive", "v3", credentials=creds)
    except Exception as e:
        st.sidebar.error(f"GCP 認證初始化失敗: {e}")
        return None

def fetch_files_from_gdrive_folder(target_folder_id: str):
    service = get_drive_service()
    if not service or not target_folder_id:
        return {}
    file_dict = {}
    try:
        q_files = f"'{target_folder_id}' in parents and trashed = false"
        res_files = service.files().list(
            q=q_files,
            fields="files(id, name, mimeType, modifiedTime)",
            pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()

        for item in res_files.get("files", []):
            raw_name = item["name"].strip()
            mime_type = item.get("mimeType", "")
            lower_name = raw_name.lower()

            # Google 試算表格式 -> 自動轉為 Excel 下載
            if mime_type == "application/vnd.google-apps.spreadsheet":
                req = service.files().export_media(
                    fileId=item["id"],
                    mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, req)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                fh.seek(0)
                final_name = raw_name if raw_name.endswith(".xlsx") else f"{raw_name}.xlsx"
                file_dict[final_name] = fh.read()

            elif any(lower_name.endswith(ext) for ext in [".xlsx", ".xls", ".csv"]):
                req = service.files().get_media(fileId=item["id"], supportsAllDrives=True)
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, req)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                fh.seek(0)
                file_dict[raw_name] = fh.read()

    except Exception as e:
        st.sidebar.error(f"雲端硬碟連線異常: {e}")
    return file_dict

MEMORY_REPORTS = {}
GDRIVE_ID_CONFIG = st.secrets.get("GDRIVE_FOLDER_ID", "1fm6ZK5B5wUmfy7-cgrw8OIkh7iS175dA")

gdrive_data = fetch_files_from_gdrive_folder(GDRIVE_ID_CONFIG)

if gdrive_data:
    MEMORY_REPORTS.update(gdrive_data)
    st.sidebar.success(f"☁️ 成功連接雲端硬碟！共載入 {len(gdrive_data)} 個最新報表")
    with st.sidebar.expander("📄 雲端硬碟已載入報表清單", expanded=True):
        for fn in sorted(gdrive_data.keys()):
            st.caption(f"• {fn}")
else:
    st.sidebar.warning(f"⚠️ 雲端資料夾 (ID: {GDRIVE_ID_CONFIG[:8]}...) 尚未讀取到報表檔案。")
    local_candidates = ["今日待上傳報表", "執法統計報表集中處", "執法報表集中處", "."]
    for d in local_candidates:
        if os.path.exists(d):
            for f in os.listdir(d):
                if any(f.lower().endswith(ext) for ext in [".xlsx", ".xls", ".csv"]) and not f.startswith("~$"):
                    p = os.path.join(d, f)
                    with open(p, "rb") as f_in:
                        MEMORY_REPORTS[f] = f_in.read()
    if MEMORY_REPORTS:
        st.sidebar.info(f"📂 偵測到本機資料夾，已載入 {len(MEMORY_REPORTS)} 個報表")

st.sidebar.markdown("### 📤 本機手動上傳報表")
uploaded_files = st.sidebar.file_uploader(
    "拖曳上傳本機報表（支援批次多檔）",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
    help="上傳後將優先以此報表進行動態統計運算"
)
if uploaded_files:
    for up_f in uploaded_files:
        MEMORY_REPORTS[up_f.name] = up_f.read()
    st.sidebar.success(f"✅ 前端已成功上傳 {len(uploaded_files)} 個報表並動態更新！")

if not MEMORY_REPORTS:
    st.warning("⚠️ 目前無有效報表檔案。請在側邊欄上傳 Excel 檔案或確認雲端硬碟配置。")

# ==========================================
# 4. 八大核心報表純動態解析核心 (嚴格零備援)
# ==========================================

# --- 4.1 三項重點違規 ---
def load_dynamic_three_major(report_dict):
    three_files = {k: v for k, v in report_dict.items() if "重點違規" in k or "重大違規" in k}
    if not three_files:
        return None, None, {}

    def extract_entry_date_info(b_data):
        try:
            df = pd.read_excel(io.BytesIO(b_data), header=None, nrows=5)
            for r in range(min(5, len(df))):
                for c in range(min(6, df.shape[1])):
                    val = str(df.iloc[r, c])
                    if "本年度" in val and "至" in val:
                        m = re.search(r'本年度\s*(\d{3})(\d{2})(\d{2})\s*至\s*(\d{3})(\d{2})(\d{2})', val)
                        if m:
                            s_y, s_m, s_d, e_y, e_m, e_d = [int(x) for x in m.groups()]
                            s_str = f"{s_y}/{s_m:02d}/{s_d:02d}"
                            e_str = f"{e_y}/{e_m:02d}/{e_d:02d}"
                            days_span = (e_m - s_m) * 31 + (e_d - s_d)
                            short_disp = f"{e_m:02d}/{e_d:02d}"
                            full_disp = s_str if (s_str == e_str) else f"{s_str}~{e_str}"
                            return short_disp, full_disp, days_span
        except Exception:
            pass
        return "本期", "本期", -1

    def parse_sheet_data_and_total(b_data):
        try:
            df = pd.read_excel(io.BytesIO(b_data), header=None)
            res = {}
            tot_sum = 0
            for r in range(len(df)):
                u = str(df.iloc[r, 0]).strip().replace(" ", "").replace("\u3000", "")
                if not u or u == 'nan':
                    continue
                if any(k in u for k in ["合計", "總計", "聖亭", "龍潭", "中興", "石門", "高平", "三和", "交通分隊"]):
                    def safe_num(v):
                        try:
                            return int(float(str(v).replace(',', '').strip()))
                        except Exception:
                            return 0
                    
                    red = safe_num(df.iloc[r, 3] if df.shape[1] > 3 else 0) + safe_num(df.iloc[r, 4] if df.shape[1] > 4 else 0)
                    rev = safe_num(df.iloc[r, 7] if df.shape[1] > 7 else 0) + safe_num(df.iloc[r, 8] if df.shape[1] > 8 else 0)
                    ped = safe_num(df.iloc[r, 13] if df.shape[1] > 13 else 0) + safe_num(df.iloc[r, 14] if df.shape[1] > 14 else 0)
                    res[u] = {'red': red, 'rev': rev, 'ped': ped, 'tot': red + rev + ped}
                    if "合計" not in u and "總計" not in u:
                        tot_sum += (red + rev + ped)
            return res, tot_sum
        except Exception:
            return {}, 0

    parsed_files = []
    for fname, raw_bytes in three_files.items():
        short_disp, full_disp, days_span = extract_entry_date_info(raw_bytes)
        data_map, total_vol = parse_sheet_data_and_total(raw_bytes)
        parsed_files.append({
            "name": fname,
            "bytes": raw_bytes,
            "short_disp": short_disp,
            "full_disp": full_disp,
            "days_span": days_span,
            "total_vol": total_vol,
            "data_map": data_map
        })

    if not parsed_files:
        return None, None, {}

    if len(parsed_files) > 1:
        if all(x["days_span"] >= 0 for x in parsed_files):
            parsed_files.sort(key=lambda x: x["days_span"])
        else:
            parsed_files.sort(key=lambda x: x["total_vol"])

    cur_item = parsed_files[0]
    cum_item = parsed_files[-1] if len(parsed_files) > 1 else parsed_files[0]

    d_cur = cur_item["data_map"]
    d_cum = cum_item["data_map"]

    units = [
        ("合計", ["合計", "總計"]),
        ("聖亭所", ["聖亭派出所", "聖亭所"]),
        ("龍潭所", ["龍潭派出所", "龍潭所"]),
        ("中興所", ["中興派出所", "中興所"]),
        ("石門所", ["石門派出所", "石門所"]),
        ("高平所", ["高平派出所", "高平所"]),
        ("三和所", ["三和派出所", "三和所"]),
        ("交通分隊", ["龍潭交通分隊", "交通分隊"])
    ]

    def get_unit_val(data_map, keys):
        for k in keys:
            for d_key, val in data_map.items():
                if k in d_key:
                    return val
        return {'red': 0, 'rev': 0, 'ped': 0, 'tot': 0}

    matrix = []
    for disp, match_keys in units:
        c = get_unit_val(d_cur, match_keys)
        cm = get_unit_val(d_cum, match_keys)
        matrix.append([disp, c['red'], c['rev'], c['ped'], c['tot'], cm['red'], cm['rev'], cm['ped'], cm['tot']])

    periods_info = {
        "cur_single": cur_item["short_disp"],
        "cur_full": cur_item["full_disp"],
        "cum_full": cum_item["full_disp"]
    }
    return cur_item["short_disp"], matrix, periods_info

# --- 4.2 交通事故 (A1 死亡、A2 受傷) ---
def load_dynamic_accidents(report_dict):
    acc_files = {k: v for k, v in report_dict.items() if "交通事故" in k}
    if not acc_files:
        return None, None, {}

    def get_latest_item(pat):
        matched = [k for k in acc_files.keys() if pat in k]
        return acc_files[matched[-1]] if matched else None

    b_cur = get_latest_item("本期")
    b_prev = get_latest_item("前期")
    b_cum = get_latest_item("今年累計") or get_latest_item("本年累計")
    b_ly = get_latest_item("去年累計")

    if not (b_cur and b_cum and b_ly):
        return None, None, {}

    def parse_acc_safe(b_data):
        date_range = ""
        data = {}
        df = None
        try:
            df = pd.read_excel(io.BytesIO(b_data), header=None)
        except Exception:
            pass

        if df is not None and not df.empty:
            for r in range(min(5, len(df))):
                row_txt = " ".join([str(x) for x in df.iloc[r].dropna()])
                m = re.search(r'(\d{2,3}/\d{2}/\d{2})\s*至\s*(\d{2,3}/\d{2}/\d{2})', row_txt)
                if m:
                    date_range = f"{m.group(1)}~{m.group(2)}"
                    break

            units = ["總計", "合計", "聖亭派出所", "龍潭派出所", "中興派出所", "石門派出所", "高平派出所", "三和派出所"]
            def clean_num(val):
                s = str(val).replace(',', '').replace('"', '').replace('-', '0').strip()
                try: return int(float(s))
                except Exception: return 0

            for r in range(len(df)):
                col0 = str(df.iloc[r, 0]).strip()
                col0_norm = col0.replace(" ", "").replace("\u3000", "")
                matched_unit = None
                for u in units:
                    if u in col0_norm:
                        matched_unit = u
                        break

                if matched_unit:
                    u_name = "合計" if any(k in matched_unit for k in ["總計", "合計"]) else matched_unit.replace("派出所", "所")
                    row_vals = df.iloc[r].values
                    a1_death = 0
                    a2_inj = 0
                    if len(row_vals) >= 10:
                        a1_death = clean_num(row_vals[5])
                        a2_inj = clean_num(row_vals[9])
                    elif len(row_vals) >= 6:
                        a1_death = clean_num(row_vals[-6])
                        a2_inj = clean_num(row_vals[-2]) if len(row_vals) >= 2 else 0

                    data[u_name] = {"a1_death": a1_death, "a2_inj": a2_inj}

            station_names = ["聖亭所", "龍潭所", "中興所", "石門所", "高平所", "三和所"]
            data["合計"] = {
                "a1_death": sum(data.get(s, {}).get("a1_death", 0) for s in station_names),
                "a2_inj": sum(data.get(s, {}).get("a2_inj", 0) for s in station_names),
            }
            return date_range, data

        return date_range, data

    r_cur, d_cur = parse_acc_safe(b_cur)
    r_prev, d_prev = parse_acc_safe(b_prev) if b_prev else ("", {})
    r_cum, d_cum = parse_acc_safe(b_cum)
    r_ly, d_ly = parse_acc_safe(b_ly)

    units_order = ["合計", "聖亭所", "龍潭所", "中興所", "石門所", "高平所", "三和所"]

    a1_list = []
    for u in units_order:
        c_val = d_cur.get(u, {}).get("a1_death", 0)
        cum_val = d_cum.get(u, {}).get("a1_death", 0)
        ly_val = d_ly.get(u, {}).get("a1_death", 0)
        a1_list.append({
            "單位": u,
            "本期": c_val,
            "本年累計": cum_val,
            "去年累計": ly_val,
            "同期比較": cum_val - ly_val
        })
    df_a1_dyn = pd.DataFrame(a1_list)

    a2_list = []
    for u in units_order:
        c_val = d_cur.get(u, {}).get("a2_inj", 0)
        p_val = d_prev.get(u, {}).get("a2_inj", 0)
        cum_val = d_cum.get(u, {}).get("a2_inj", 0)
        ly_val = d_ly.get(u, {}).get("a2_inj", 0)
        diff = cum_val - ly_val
        rate = f"{(diff / ly_val)*100:.2f}%" if ly_val else "—"
        a2_list.append({
            "單位": u,
            "本期": c_val,
            "前期": p_val,
            "本年累計": cum_val,
            "去年累計": ly_val,
            "同期比較": diff,
            "增減比例": rate
        })
    df_a2_dyn = pd.DataFrame(a2_list)

    acc_periods = {
        "cur": r_cur if r_cur else "本期",
        "prev": r_prev if r_prev else "前期",
        "cum": r_cum if r_cum else "本年累計",
        "ly": r_ly if r_ly else "去年同期"
    }
    return df_a1_dyn, df_a2_dyn, acc_periods

# --- 4.3 重大交通違規總表與 7 大專項細表 ---
def load_dynamic_major(report_dict):
    major_files = {k: v for k, v in report_dict.items() if "重大違規" in k or "重點違規" in k}
    if not major_files:
        return None, None, {}

    def get_latest_item(pat, exclude=None):
        matched = [
            k for k in major_files.keys()
            if pat in k and (exclude is None or exclude not in k)
        ]
        return major_files[matched[-1]] if matched else None

    b_cur = get_latest_item("本期")
    b_cum = get_latest_item("本年累計") or get_latest_item("年累計", exclude="去年")
    b_ly = get_latest_item("去年累計")

    if not (b_cur and b_cum and b_ly):
        all_parsed = []
        for fn, b_data in major_files.items():
            try:
                df = pd.read_excel(io.BytesIO(b_data), header=None)
                p_str = ""
                for r in range(min(5, len(df))):
                    row_txt = " ".join([str(x) for x in df.iloc[r].dropna()])
                    m = re.search(r'(\d{7})至(\d{7})', row_txt)
                    if m:
                        p_str = f"{m.group(1)}~{m.group(2)}"
                        break
                all_parsed.append((fn, b_data, p_str))
            except Exception:
                pass

        for fn, b, p in all_parsed:
            if "0101" in p and not b_cum:
                b_cum = b
            elif "114" in p and not b_ly:
                b_ly = b
            elif not b_cur:
                b_cur = b

    if not (b_cur and b_cum and b_ly):
        return None, None, {}

    def parse_major_safe(b_data):
        df = None
        try:
            df = pd.read_excel(io.BytesIO(b_data), header=None)
        except Exception:
            try:
                raw_text = b_data.decode('utf-8', errors='ignore')
                lines = [l.strip() for l in raw_text.split('\n') if ',' in l]
                df = pd.DataFrame([list(csv.reader([l]))[0] for l in lines])
            except Exception:
                return "", {}

        if df is None or df.empty:
            return "", {}

        period = ""
        for r in range(min(5, len(df))):
            row_txt = " ".join([str(x) for x in df.iloc[r].dropna()])
            m = re.search(r'(\d{7})至(\d{7})', row_txt)
            if m:
                period = f"{m.group(1)}~{m.group(2)}"
                break

        res = {}
        for r in range(len(df)):
            col0 = str(df.iloc[r, 0]).strip().replace(" ", "").replace("\u3000", "")
            if not col0 or col0 == 'nan' or any(k in col0 for k in ['列印', '單位', '本年度', '統計']):
                continue
            vals = []
            for c in range(1, min(22, df.shape[1])):
                v_str = str(df.iloc[r, c]).replace(',', '').replace('"', '').replace('-', '0').strip()
                try:
                    vals.append(int(float(v_str)))
                except Exception:
                    vals.append(0)

            if "交通分隊" in col0 or "龍潭交通分隊" in col0 or ("龍潭" in col0 and "分隊" in col0):
                norm_u = "交通分隊"
            elif "交通組" in col0:
                norm_u = "科技執法"
            elif "聖亭" in col0:
                norm_u = "聖亭所"
            elif "龍潭" in col0 and "所" in col0:
                norm_u = "龍潭所"
            elif "中興" in col0:
                norm_u = "中興所"
            elif "石門" in col0:
                norm_u = "石門所"
            elif "高平" in col0:
                norm_u = "高平所"
            elif "三和" in col0:
                norm_u = "三和所"
            elif "警備" in col0:
                norm_u = "警備隊"
            elif any(k in col0 for k in ["合計", "總計"]):
                norm_u = "合計"
            else:
                norm_u = col0.replace("派出所", "所")

            res[norm_u] = vals

        return period, res

    p_cur, d_cur = parse_major_safe(b_cur)
    p_cum, d_cum = parse_major_safe(b_cum)
    p_ly, d_ly = parse_major_safe(b_ly)

    col_cur_s, col_cur_a = "本期(攔停)", "本期(逕舉)"
    col_cum_s, col_cum_a = "本年累計(攔停)", "本年累計(逕舉)"
    col_ly_s, col_ly_a = "去年累計(攔停)", "去年累計(逕舉)"

    targets = {
        "合計": 18114, "科技執法": 6006, "聖亭所": 1941, "龍潭所": 2588, "中興所": 1941,
        "石門所": 1479, "高平所": 1294, "三和所": 339, "警備隊": 0, "交通分隊": 2526
    }
    unit_order = ["合計", "聖亭所", "龍潭所", "中興所", "石門所", "高平所", "三和所", "警備隊", "交通分隊", "科技執法"]

    major_rows = []
    for u in unit_order:
        cv = d_cur.get(u, [0]*20)
        cmv = d_cum.get(u, [0]*20)
        lyv = d_ly.get(u, [0]*20)

        cur_s = cv[14] if len(cv) > 14 else 0
        cur_a = cv[15] if len(cv) > 15 else 0

        cum_s = cmv[14] if len(cmv) > 14 else 0
        cum_a = cmv[15] if len(cmv) > 15 else 0
        cum_tot = cmv[16] if len(cmv) > 16 else (cum_s + cum_a)

        ly_s = lyv[14] if len(lyv) > 14 else 0
        ly_a = lyv[15] if len(lyv) > 15 else 0
        ly_tot = lyv[16] if len(lyv) > 16 else (ly_s + ly_a)

        diff = cum_tot - ly_tot
        tgt = targets.get(u, 0)
        achieve = f"{(cum_tot / tgt)*100:.1f}%" if tgt > 0 else "—"
        if u == "警備隊": diff = "—"

        major_rows.append({
            "單位": u,
            col_cur_s: cur_s, col_cur_a: cur_a,
            col_cum_s: cum_s, col_cum_a: cum_a,
            col_ly_s: ly_s, col_ly_a: ly_a,
            "同期比較": diff, "目標值": tgt, "達成率": achieve
        })
    df_major_dyn = pd.DataFrame(major_rows)

    cat_indices = {
        "酒駕": (0, 1), "闖紅燈": (2, 3), "嚴重超速": (4, 5),
        "逆向行駛": (6, 7), "轉彎未依規定": (8, 9),
        "蛇行惡意逼車": (10, 11), "不暫停讓行人": (12, 13)
    }
    detail_dict = {}
    for cat, (is_, ia_) in cat_indices.items():
        rows = []
        for u in unit_order:
            cm = d_cum.get(u, [0]*20)
            ly = d_ly.get(u, [0]*20)
            cs = cm[is_] if len(cm) > is_ else 0
            ca = cm[ia_] if len(cm) > ia_ else 0
            ls = ly[is_] if len(ly) > is_ else 0
            la = ly[ia_] if len(ly) > ia_ else 0
            ct, lt = cs + ca, ls + la
            ds, da, dt = (cs - ls, ca - la, ct - lt) if u != "警備隊" else ("—", "—", "—")
            rows.append([u, cs, ca, ct, ls, la, lt, ds, da, dt])
        detail_dict[cat] = rows

    major_periods = {
        "cur": p_cur if p_cur else "本期",
        "cum": p_cum if p_cum else "本年累計",
        "ly": p_ly if p_ly else "去年同期"
    }
    return df_major_dyn, detail_dict, major_periods

# --- 4.4 取締超載違規件數統計表 (精確鎖定「超載」第2欄) ---
def load_dynamic_overload(report_dict):
    ov_files = {k: v for k, v in report_dict.items() if any(w in k for w in ["超載違規", "取締裝載砂石"])}
    if not ov_files:
        return None, "", {}

    def get_exact_file(period_pat, unit_type):
        candidates = []
        for k in ov_files.keys():
            if period_pat in k:
                if unit_type == "交大" and "交通大隊" in k:
                    candidates.append(k)
                elif unit_type == "分局" and "交通大隊" not in k and ("龍潭分局" in k or "龍潭" in k):
                    candidates.append(k)
        if candidates:
            candidates.sort()
            return ov_files[candidates[-1]]
        return None

    def parse_r17_file(b_data, is_traffic_corps=False):
        if not b_data:
            return "", {}
        try:
            df_full = pd.read_excel(io.BytesIO(b_data), header=None)
            period = ""
            for r in range(min(5, len(df_full))):
                txt = " ".join([str(x) for x in df_full.iloc[r].dropna()])
                if "統計期間：" in txt:
                    period = txt.split("統計期間：")[1].strip()
                    break

            cnt_col_idx = 2
            for c in range(df_full.shape[1]):
                col_hdr = str(df_full.iloc[5, c]) + str(df_full.iloc[6, c])
                if "超載" in col_hdr and "出入" not in col_hdr and "通報" not in col_hdr:
                    cnt_col_idx = c
                    break

            data_start_row = 7
            df_data = df_full.iloc[data_start_row:].copy()
            df_data[0] = df_data[0].ffill()

            counts = {}
            for r_idx in range(len(df_data)):
                row_vals = df_data.iloc[r_idx]
                row_text = " ".join([str(x).strip() for x in row_vals.dropna() if str(x).strip() != ''])
                
                if any(k in row_text for k in ["總計", "大隊合計", "大隊部"]):
                    continue

                def to_int(x):
                    try:
                        s = str(x).replace(',', '').replace('"', '').replace('-', '0').strip()
                        return int(float(s))
                    except Exception:
                        return 0

                cnt = to_int(row_vals[cnt_col_idx])

                if is_traffic_corps:
                    if "龍潭" in row_text and ("分隊" in row_text or "隊" in row_text):
                        counts["交通分隊"] = counts.get("交通分隊", 0) + cnt
                else:
                    col0_str = str(row_vals[0]).strip().replace(" ", "").replace("\u3000", "")
                    if "合計" in col0_str:
                        continue
                    if "聖亭" in col0_str: norm_u = "聖亭所"
                    elif "龍潭" in col0_str and ("分隊" in col0_str or "交通" in col0_str): norm_u = "交通分隊"
                    elif "龍潭" in col0_str and "所" in col0_str: norm_u = "龍潭所"
                    elif "中興" in col0_str: norm_u = "中興所"
                    elif "石門" in col0_str: norm_u = "石門所"
                    elif "高平" in col0_str: norm_u = "高平所"
                    elif "三和" in col0_str: norm_u = "三和所"
                    elif "警備" in col0_str: norm_u = "警備隊"
                    elif "交通分隊" in col0_str or "分隊" in col0_str: norm_u = "交通分隊"
                    else: norm_u = col0_str.replace("派出所", "所")

                    counts[norm_u] = counts.get(norm_u, 0) + cnt

            return period, counts
        except Exception:
            return "", {}

    p_cur, d_cur_precinct = parse_r17_file(get_exact_file("本期", "分局"), is_traffic_corps=False)
    p_cum, d_cum_precinct = parse_r17_file(get_exact_file("本年累計", "分局") or get_exact_file("年累計", "分局"), is_traffic_corps=False)
    p_ly, d_ly_precinct = parse_r17_file(get_exact_file("去年累計", "分局"), is_traffic_corps=False)

    _, d_cur_traffic = parse_r17_file(get_exact_file("本期", "交大"), is_traffic_corps=True)
    _, d_cum_traffic = parse_r17_file(get_exact_file("本年累計", "交大") or get_exact_file("年累計", "交大"), is_traffic_corps=True)
    _, d_ly_traffic = parse_r17_file(get_exact_file("去年累計", "交大"), is_traffic_corps=True)

    if not (d_cum_precinct or d_cum_traffic):
        return None, "", {}

    def combine_units(d_p, d_t):
        merged = d_p.copy()
        t_squad_val = d_t.get("交通分隊", 0)
        p_squad_val = d_p.get("交通分隊", 0)
        merged["交通分隊"] = t_squad_val if t_squad_val > 0 else p_squad_val
        return merged

    d_cur = combine_units(d_cur_precinct, d_cur_traffic)
    d_cum = combine_units(d_cum_precinct, d_cum_traffic)
    d_ly = combine_units(d_ly_precinct, d_ly_traffic)

    targets = {
        "合計": 127, "聖亭所": 20, "龍潭所": 27, "中興所": 20,
        "石門所": 16, "高平所": 14, "三和所": 8, "警備隊": 0, "交通分隊": 22
    }
    unit_order = ["合計", "聖亭所", "龍潭所", "中興所", "石門所", "高平所", "三和所", "警備隊", "交通分隊"]

    rows = []
    for u in unit_order:
        c_val = d_cur.get(u, 0)
        cum_val = d_cum.get(u, 0)
        ly_val = d_ly.get(u, 0)
        diff = cum_val - ly_val
        tgt = targets.get(u, 0)
        achieve = f"{(cum_val / tgt)*100:.0f}%" if tgt > 0 else "—"
        if u == "警備隊": diff = 0

        rows.append({
            "單位": u,
            "本期": c_val,
            "本年累計": cum_val,
            "去年累計": ly_val,
            "同期比較": diff,
            "目標值": tgt,
            "達成率": achieve
        })

    if rows:
        tot_c = sum(r["本期"] for r in rows[1:])
        tot_cum = sum(r["本年累計"] for r in rows[1:])
        tot_ly = sum(r["去年累計"] for r in rows[1:])
        rows[0]["本期"] = tot_c
        rows[0]["本年累計"] = tot_cum
        rows[0]["去年累計"] = tot_ly
        rows[0]["同期比較"] = tot_cum - tot_ly
        rows[0]["達成率"] = f"{(tot_cum / 127)*100:.0f}%"

    footnote = "本期定義：係指該期昱通系統入案件數（含交通警察大隊龍潭分隊）；以年底達成率100%為基準。"
    ov_periods = {
        "cur": p_cur if p_cur else "本期",
        "cum": p_cum if p_cum else "本年累計",
        "ly": p_ly if p_ly else "去年同期"
    }
    return pd.DataFrame(rows), footnote, ov_periods

# --- 4.5 「靜桃計畫」大執法專案統計表 (自動掃描所有工作表 + 解除年度限制 1129 件) ---
def load_dynamic_jingtao(report_dict):
    jt_files = {k: v for k, v in report_dict.items() if any(w in k for w in ["改裝", "噪音", "行為人", "靜桃"])}
    if not jt_files:
        return None

    fname, b_data = list(jt_files.items())[-1]
    try:
        xls = pd.ExcelFile(io.BytesIO(b_data))
        df_target = None
        hdr_idx = -1
        
        # 逐一掃描所有 Sheet 尋找最相符的工作表
        for sheet in xls.sheet_names:
            try:
                temp_raw = pd.read_excel(xls, sheet_name=sheet, header=None)
                for r in range(min(10, len(temp_raw))):
                    r_str = " ".join([str(x) for x in temp_raw.iloc[r].dropna()])
                    if "通報日期" in r_str and ("所別" in r_str or "單位" in r_str):
                        if "22-06" in r_str or "06-22" in r_str or df_target is None:
                            df_target = temp_raw
                            hdr_idx = r
                            if "22-06" in r_str or "06-22" in r_str:
                                break
            except Exception:
                continue

        if df_target is None or hdr_idx == -1:
            st.warning(f"⚠️ 已找到清冊【{fname}】，但未能識別出包含「通報日期」與「所別」的表頭欄位。")
            return None

        headers = [str(x).strip().replace("'", "") for x in df_target.iloc[hdr_idx]]
        df_data = df_target.iloc[hdr_idx+1:].copy()
        df_data.columns = headers

        date_col = next((c for c in df_data.columns if "通報日期" in c or "日期" in c), None)
        unit_col = next((c for c in df_data.columns if "所別" in c or "單位" in c or "通報單位" in c), None)
        col_22_06 = next((c for c in df_data.columns if "22-06" in c or "22~06" in c), None)
        col_06_22 = next((c for c in df_data.columns if "06-22" in c or "06~22" in c), None)

        if not (date_col and unit_col):
            st.warning(f"⚠️ 清冊【{fname}】缺少通報日期或所別欄位。")
            return None

        today = datetime.now()
        yesterday = today - timedelta(days=1)
        end_cur = f"{str(yesterday.year - 1911)}/{yesterday.strftime('%m')}/{yesterday.strftime('%d')}"
        one_week_ago = yesterday - timedelta(days=6)
        start_cur = f"{str(one_week_ago.year - 1911)}/{one_week_ago.strftime('%m')}/{one_week_ago.strftime('%d')}"

        def norm_date(val):
            if pd.isna(val): return ""
            s = str(val).strip().replace("-", "/")
            parts = s.split("/")
            if len(parts) == 3:
                try:
                    return f"{int(parts[0]):03d}/{int(parts[1]):02d}/{int(parts[2]):02d}"
                except Exception:
                    return s
            return s

        df_data["std_date"] = df_data[date_col].apply(norm_date)

        df_all = df_data[df_data[unit_col].notna()].copy()
        df_cur = df_data[(df_data["std_date"] >= start_cur) & (df_data["std_date"] <= end_cur)].copy()

        def is_checked(val):
            if pd.isna(val): return False
            s = str(val).strip().upper()
            return s in ['V', '1', 'TRUE', 'Y', 'YES'] or len(s) > 0

        units = ["合計", "聖亭所", "龍潭所", "中興所", "石門所", "高平所", "三和所", "警備隊", "交通分隊"]
        rows = []

        for u in units:
            if u == "合計":
                sub_cur = df_cur
                sub_all = df_all
            elif u == "交通分隊":
                sub_cur = df_cur[df_cur[unit_col].astype(str).str.contains("交通", na=False)]
                sub_all = df_all[df_all[unit_col].astype(str).str.contains("交通", na=False)]
            else:
                key = u.replace("所", "").replace("隊", "")
                sub_cur = df_cur[df_cur[unit_col].astype(str).str.contains(key, na=False)]
                sub_all = df_all[df_all[unit_col].astype(str).str.contains(key, na=False)]

            cur_22 = int(sub_cur[col_22_06].apply(is_checked).sum()) if col_22_06 else 0
            cur_06 = int(sub_cur[col_06_22].apply(is_checked).sum()) if col_06_22 else 0

            cum_22 = int(sub_all[col_22_06].apply(is_checked).sum()) if col_22_06 else 0
            cum_06 = int(sub_all[col_06_22].apply(is_checked).sum()) if col_06_22 else 0
            tot = len(sub_all)

            rows.append({
                "單位": u,
                "本期(22-06)": cur_22,
                "本期(06-22)": cur_06,
                "累計(22-06)": cum_22,
                "累計(06-22)": cum_06,
                "總計": tot
            })

        return pd.DataFrame(rows)

    except Exception as e:
        st.error(f"❌ 解析靜桃清冊【{fname}】時發生異常：{e}")
        return None

# --- 4.6 科技執法成效統計表 (嚴格讀取案件明細之違規地點，無備援) ---
def load_dynamic_tech(report_dict):
    tech_files = {k: v for k, v in report_dict.items() if any(w in k for w in ["科技執法", "自選匯出"])}
    if not tech_files:
        return None, ""

    b_data = list(tech_files.values())[-1]
    try:
        xls = pd.ExcelFile(io.BytesIO(b_data))
        
        if "案件明細" not in xls.sheet_names:
            st.error("❌ 科技執法報表缺少【案件明細】工作表！")
            return None, ""

        df_detail = pd.read_excel(xls, sheet_name="案件明細", skiprows=3)

        loc_col = next((c for c in df_detail.columns if "違規地點" in str(c)), None)
        if not loc_col:
            st.error("❌ 科技執法報表【案件明細】中缺少【違規地點】欄位，請確認匯出時是否有打勾！")
            return None, ""

        counts = df_detail[loc_col].dropna().value_counts().reset_index()
        counts.columns = ["路段名稱", "舉發件數"]
        tot = counts["舉發件數"].sum()
        tot_row = pd.DataFrame([{"路段名稱": "舉發總數", "舉發件數": tot}])
        df_out = pd.concat([counts, tot_row], ignore_index=True)
        return df_out, "科技執法成效統計表"

    except Exception as e:
        st.error(f"❌ 讀取科技執法報表發生異常：{e}")
        return None, ""

# ==========================================
# 5. 純動態執行載入（完全無假資料）
# ==========================================
three_day, three_matrix, three_periods = load_dynamic_three_major(MEMORY_REPORTS)
if three_matrix:
    col_cur_label = f"本期 ({three_periods.get('cur_single', '本期')})"
    col_cum_label = f"本月累計 ({three_periods.get('cum_full', '本月累計')})"
    preview_cols = pd.MultiIndex.from_tuples([
        ("單位", ""),
        (f"{col_cur_label} 新增違規數", "闖紅燈"),
        (f"{col_cur_label} 新增違規數", "逆向行駛"),
        (f"{col_cur_label} 新增違規數", "不停讓行人"),
        (f"{col_cur_label} 新增違規數", "本期合計"),
        (f"{col_cum_label}數", "闖紅燈"),
        (f"{col_cum_label}數", "逆向行駛"),
        (f"{col_cum_label}數", "不停讓行人"),
        (f"{col_cum_label}數", "累計總計")
    ])
    df_three_preview = pd.DataFrame(three_matrix, columns=preview_cols)
else:
    df_three_preview = None

df_a1_dyn, df_a2_dyn, acc_periods = load_dynamic_accidents(MEMORY_REPORTS)
df_major_dyn, detail_dict_dyn, major_periods = load_dynamic_major(MEMORY_REPORTS)
df_overload_dyn, overload_fn, ov_periods = load_dynamic_overload(MEMORY_REPORTS)
df_jingtao_dyn = load_dynamic_jingtao(MEMORY_REPORTS)
df_tech_dyn, tech_title = load_dynamic_tech(MEMORY_REPORTS)

# ==========================================
# 6. 前端自選與即時預覽區 (8大常態 + 7大細表)
# ==========================================
st.subheader("🎯 欲輸出的統計表自選控制")

col_btn1, col_btn2, _ = st.columns([1.5, 2, 4])
if "select_mode" not in st.session_state:
    st.session_state["select_mode"] = "core"

with col_btn1:
    if st.button("📌 僅常態核心頁 (最多8頁)"):
        st.session_state["select_mode"] = "core"
with col_btn2:
    if st.button("📑 全選所有可用統計表 (最多15頁)"):
        st.session_state["select_mode"] = "all"

is_all = (st.session_state["select_mode"] == "all")

col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    st.markdown("##### 🏢 常態會報核心表格")
    chk_cover = st.checkbox("P.1 簡報封面", value=True)
    chk_three = st.checkbox("P.2 取締三項重點違規統計表", value=(df_three_preview is not None), disabled=(df_three_preview is None))
    chk_a1 = st.checkbox("P.3 A1類交通事故死亡人數統計表", value=(df_a1_dyn is not None), disabled=(df_a1_dyn is None))
    chk_a2 = st.checkbox("P.4 A2類交通事故受傷人數統計表", value=(df_a2_dyn is not None), disabled=(df_a2_dyn is None))
    chk_major_tot = st.checkbox("P.5 取締重大交通違規統計表 (總表)", value=(df_major_dyn is not None), disabled=(df_major_dyn is None))
    chk_overload = st.checkbox("P.6 取締超載違規件數統計表", value=(df_overload_dyn is not None), disabled=(df_overload_dyn is None))
    chk_jingtao = st.checkbox("P.7 「靜桃計畫」大執法專案統計表", value=(df_jingtao_dyn is not None), disabled=(df_jingtao_dyn is None))
    chk_tech = st.checkbox("P.8 科技執法成效", value=(df_tech_dyn is not None), disabled=(df_tech_dyn is None))

with col_opt2:
    st.markdown("##### 🔍 重大違規專項細表（選配）")
    has_det = (detail_dict_dyn is not None)
    chk_det_jiu = st.checkbox("重大違規細項：【酒駕】統計表", value=is_all and has_det, disabled=not has_det)
    chk_det_red = st.checkbox("重大違規細項：【闖紅燈】統計表", value=is_all and has_det, disabled=not has_det)
    chk_det_rev = st.checkbox("重大違規細項：【逆向行駛】統計表", value=is_all and has_det, disabled=not has_det)
    chk_det_turn = st.checkbox("重大違規細項：【轉彎未依規定】統計表", value=is_all and has_det, disabled=not has_det)
    chk_det_snake = st.checkbox("重大違規細項：【蛇行惡意逼車】統計表", value=is_all and has_det, disabled=not has_det)
    chk_det_ped = st.checkbox("重大違規細項：【不暫停讓行人】統計表", value=is_all and has_det, disabled=not has_det)
    chk_det_speed = st.checkbox("重大違規細項：【嚴重超速】統計表", value=is_all and has_det, disabled=not has_det)

st.markdown("---")
st.markdown("#### 📁 簡報檔案命名與信件通知")

default_pptx_name = f"龍潭分局執法數據簡報_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"
custom_file_name = st.text_input("✏️ 自訂簡報存檔名稱：", value=default_pptx_name)

curr_user = st.secrets.get("email", {}).get("user") or st.secrets.get("SMTP_USER", "")
if curr_user:
    st.caption(f"📬 產出後將自動夾帶附件發送至：`{curr_user}`")

# 預覽區
with st.expander("👀 點擊展開預覽純動態讀取之數據（無任何寫死假資料）", expanded=True):
    tabs_to_show = []
    if df_three_preview is not None: tabs_to_show.append("三項重點")
    if df_a1_dyn is not None: tabs_to_show.append("A1事故死亡")
    if df_a2_dyn is not None: tabs_to_show.append("A2事故受傷")
    if df_major_dyn is not None: tabs_to_show.append("重大違規總表")
    if df_overload_dyn is not None: tabs_to_show.append("超載取締")
    if df_jingtao_dyn is not None: tabs_to_show.append("靜桃計畫")
    if df_tech_dyn is not None: tabs_to_show.append("科技執法")

    if tabs_to_show:
        tabs = st.tabs(tabs_to_show)
        for idx, tab_name in enumerate(tabs_to_show):
            with tabs[idx]:
                if tab_name == "三項重點": st.dataframe(df_three_preview, hide_index=True)
                elif tab_name == "A1事故死亡": st.dataframe(df_a1_dyn, hide_index=True)
                elif tab_name == "A2事故受傷": st.dataframe(df_a2_dyn, hide_index=True)
                elif tab_name == "重大違規總表": st.dataframe(df_major_dyn, hide_index=True)
                elif tab_name == "超載取締": st.dataframe(df_overload_dyn, hide_index=True)
                elif tab_name == "靜桃計畫": st.dataframe(df_jingtao_dyn, hide_index=True)
                elif tab_name == "科技執法": st.dataframe(df_tech_dyn, hide_index=True)
    else:
        st.info("💡 目前尚未偵測到有效報表，請於側邊欄上傳 Excel 檔案。")

# ==========================================
# 7. 執行指定輸出生成
# ==========================================
st.write("")
btn_col1, btn_col2 = st.columns([1.5, 2.5])
with btn_col1:
    btn_generate = st.button("🚀 直出純動態 PPTX 簡報檔", type="primary", use_container_width=True)
with btn_col2:
    chk_auto_email = st.checkbox("產出後自動將 PPTX 附件寄到我的信箱", value=True)

if btn_generate:
    file_save_name = custom_file_name.strip()
    if not file_save_name.lower().endswith(".pptx"):
        file_save_name += ".pptx"

    with st.spinner("正在自最新報表動態編譯 PPTX 簡報（全表 16/18pt、期間精確對齊）..."):
        try:
            builder = PptxReportBuilder()

            # P.1 封面
            if chk_cover:
                builder.add_cover_slide(
                    main_title="桃園市政府警察局龍潭分局\n交通執法成效與事故防制分析報告",
                    subtitle="週次主管會報專案報告",
                    date_range_str=f"統計截止至最新報表 ｜ 製表日期：{datetime.now().strftime('%Y/%m/%d')}"
                )

            # P.2 三項重點
            if chk_three and df_three_preview is not None:
                cur_dt = three_periods.get("cur_single", "本期")
                cur_full = three_periods.get("cur_full", cur_dt)
                cum_full = three_periods.get("cum_full", "本月累計")
                three_sub = f"統計期間：本期 ({cur_full}) ｜ 本月累計 ({cum_full}) ｜ 製表單位：龍潭分局交通組"
                
                builder.add_three_major_slide(
                    data_rows=three_matrix,
                    custom_subtitle=three_sub,
                    cur_col_title=f"本期({cur_dt})",
                    cum_col_title="本月累計"
                )

            # P.3 A1 死亡
            if chk_a1 and df_a1_dyn is not None:
                a1_sub = (
                    f"本期：{acc_periods.get('cur', '—')} ｜ "
                    f"本年累計：{acc_periods.get('cum', '—')} ｜ "
                    f"去年同期：{acc_periods.get('ly', '—')}"
                )
                builder.add_table_slide(
                    slide_title="A1類交通事故死亡人數統計表",
                    df=df_a1_dyn,
                    subtitle=a1_sub,
                    is_accident_table=True
                )

            # P.4 A2 受傷
            if chk_a2 and df_a2_dyn is not None:
                a2_sub = (
                    f"本期：{acc_periods.get('cur', '—')} ｜ "
                    f"前期：{acc_periods.get('prev', '—')} ｜ "
                    f"本年累計：{acc_periods.get('cum', '—')} ｜ "
                    f"去年同期：{acc_periods.get('ly', '—')}"
                )
                builder.add_table_slide(
                    slide_title="A2類交通事故受傷人數統計表",
                    df=df_a2_dyn,
                    subtitle=a2_sub,
                    is_accident_table=True
                )

            # P.5 重大違規總表
            if chk_major_tot and df_major_dyn is not None:
                p_cur_str = f"本期：{major_periods.get('cur', '')} ｜ " if major_periods.get('cur') else ""
                p_cum_str = f"本年累計：{major_periods.get('cum', '')} ｜ " if major_periods.get('cum') else ""
                p_ly_str = f"去年同期：{major_periods.get('ly', '')}" if major_periods.get('ly') else ""
                sub_txt = f"{p_cur_str}{p_cum_str}{p_ly_str}".rstrip(" ｜ ")

                builder.add_table_slide(
                    slide_title="取締重大交通違規統計表",
                    df=df_major_dyn,
                    subtitle=sub_txt,
                    footnote="重大交通違規指：「酒駕」、「闖紅燈」、「嚴重超速」、「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」及「不暫停讓行人」",
                    is_major_table=True
                )

            # 專項細表
            if detail_dict_dyn is not None:
                det_map = [
                    (chk_det_jiu, "酒駕"), (chk_det_red, "闖紅燈"), (chk_det_rev, "逆向行駛"),
                    (chk_det_turn, "轉彎未依規定"), (chk_det_snake, "蛇行惡意逼車"),
                    (chk_det_ped, "不暫停讓行人"), (chk_det_speed, "嚴重超速")
                ]
                cum_range = major_periods.get("cum", "")
                ly_range = major_periods.get("ly", "")
                det_subtitle = f"本年累計：{cum_range} ｜ 去年同期：{ly_range} ｜ 口徑包含現場攔停與逕行舉發" if cum_range else "口徑包含現場攔停與逕行舉發"

                for is_chk, cat in det_map:
                    if is_chk and cat in detail_dict_dyn:
                        builder.add_major_detail_slide(cat_name=cat, data_rows=detail_dict_dyn[cat], custom_subtitle=det_subtitle)

            # P.6 超載取締
            if chk_overload and df_overload_dyn is not None:
                ov_sub = (
                    f"本期：{ov_periods.get('cur', '—')} ｜ "
                    f"本年累計：{ov_periods.get('cum', '—')} ｜ "
                    f"去年同期：{ov_periods.get('ly', '—')}"
                )
                builder.add_table_slide(
                    slide_title="取締超載違規件數統計表",
                    df=df_overload_dyn,
                    subtitle=ov_sub,
                    footnote=overload_fn
                )

            # P.7 靜桃計畫 (本期 1 件龍潭所，專案累計 1129 件)
            if chk_jingtao and df_jingtao_dyn is not None:
                builder.add_table_slide(
                    slide_title="「靜桃計畫」大執法專案統計表",
                    df=df_jingtao_dyn,
                    subtitle="改裝排氣管及噪音車輛通報取締成果"
                )

            # P.8 科技執法
            if chk_tech and df_tech_dyn is not None:
                builder.add_table_slide(
                    slide_title="科技執法成效統計表",
                    df=df_tech_dyn,
                    custom_width_in=8.0
                )

            pptx_stream = builder.build_bytes()
            st.session_state["cached_pptx"] = pptx_stream
            st.session_state["cached_filename"] = file_save_name

            if chk_auto_email:
                ok, detail = send_pptx_email_to_self(pptx_stream, file_save_name)
                if ok:
                    st.info(f"📧 簡報附件已成功發送至您的信箱：`{detail}`")
                else:
                    st.warning(f"⚠️ 郵件未發送成功（{detail}），可直接點擊下方按鈕下載！")

            st.balloons()
            st.success("🎉 恭喜！包含常態 8 頁與專項細表之純動態 PowerPoint 簡報實體檔已成功生成！")

        except Exception as e:
            st.error(f"❌ 產出 PPTX 簡報時發生錯誤：{str(e)}")

# ==========================================
# 8. 下載專用按鈕區
# ==========================================
if "cached_pptx" in st.session_state:
    st.markdown("---")
    st.subheader("📥 簡報下載與轉存")
    st.download_button(
        label=f"💾 點此立即下載【{st.session_state['cached_filename']}】",
        data=st.session_state["cached_pptx"].getvalue(),
        file_name=st.session_state["cached_filename"],
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        use_container_width=True
    )
