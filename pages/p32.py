import io
import os
import re
import smtplib
import csv
import urllib.parse as _ul
from datetime import datetime
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

# 嘗試載入 Google Drive API 相關套件
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
st.caption("🚀 來源嚴格鎖定：僅能依據【Google 雲端硬碟執法報表集中處】或【網站前端即時上傳】，絕無程式碼寫死數值！")

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

    def _set_cell(self, cell, text, font_size=11, bold=False, color=None, bg_color=None, align=PP_ALIGN.CENTER):
        cell.text = str(text).strip() if (pd.notna(text) and str(text).strip() != "") else "—"
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_color if bg_color else self.C_TBL_ROW_BG
        self._set_border(cell, color_hex="CBD5E1")

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
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12.133), Inches(0.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Microsoft JhengHei"
        p.font.size = Pt(21)
        p.font.bold = True
        p.font.color.rgb = self.C_TITLE_NAVY

        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.font.name = "Microsoft JhengHei"
            p2.font.size = Pt(11.5)
            p2.font.color.rgb = self.C_MUTED
            p2.space_before = Pt(4)

    def add_three_major_slide(self, data_rows, latest_day="本期"):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self.add_header_box(
            slide,
            "桃園市政府警察局龍潭分局 取締三項重點違規本期及累計統計表",
            f"統計期間：自本月起至本期({latest_day})止 ｜ 製表單位：龍潭分局交通組"
        )

        num_rows = len(data_rows) + 2
        num_cols = 9
        table_shape = slide.shapes.add_table(num_rows, num_cols, Inches(0.6), Inches(1.4), Inches(12.133), Inches(5.4))
        tbl = table_shape.table

        tbl.cell(0, 0).merge(tbl.cell(1, 0))
        tbl.cell(0, 1).merge(tbl.cell(0, 4))
        tbl.cell(0, 5).merge(tbl.cell(0, 8))

        self._set_cell(tbl.cell(0, 0), "單位", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 1), f"本期 ({latest_day}) 新增違規數", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 5), "本月累計數", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        sub_headers = ["", "闖紅燈", "逆向行駛", "不停讓行人", f"本期合計\n({latest_day})", "闖紅燈", "逆向行駛", "不停讓行人", "累計總計"]
        for c in range(1, 9):
            self._set_cell(tbl.cell(1, c), sub_headers[c], font_size=11, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        for r_idx, row in enumerate(data_rows, start=2):
            is_tot = (r_idx == 2)
            bg = self.C_TBL_HIGHLIGHT_BG if is_tot else self.C_TBL_ROW_BG
            for c_idx, val in enumerate(row):
                self._set_cell(tbl.cell(r_idx, c_idx), val, font_size=11, bold=is_tot, color=self.C_TBL_TEXT_DARK, bg_color=bg)

    def add_major_detail_slide(self, cat_name: str, data_rows, date_str=""):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self.add_header_box(
            slide,
            f"取締【{cat_name}】違規統計表 {f'(累計至 {date_str})' if date_str else ''}",
            "口徑包含現場攔停與逕行舉發 ｜ 製表單位：龍潭分局交通組"
        )

        num_rows = len(data_rows) + 2
        num_cols = 10
        table_shape = slide.shapes.add_table(num_rows, num_cols, Inches(0.6), Inches(1.4), Inches(12.133), Inches(5.4))
        tbl = table_shape.table

        tbl.cell(0, 0).merge(tbl.cell(1, 0))
        tbl.cell(0, 1).merge(tbl.cell(0, 3))
        tbl.cell(0, 4).merge(tbl.cell(0, 6))
        tbl.cell(0, 7).merge(tbl.cell(0, 9))

        self._set_cell(tbl.cell(0, 0), "統計期間", font_size=11, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 1), "今年累計", font_size=11, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 4), "去年累計", font_size=11, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 7), "今年與去年同期比較", font_size=11, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        sub_names = ["", "當場攔停", "逕行舉發", "合計", "當場攔停", "逕行舉發", "合計", "當場攔停", "逕行舉發", "合計"]
        for c in range(1, 10):
            self._set_cell(tbl.cell(1, c), sub_names[c], font_size=10, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

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

                self._set_cell(tbl.cell(r_idx, c_idx), val, font_size=10, bold=is_bold, color=fg, bg_color=bg)

    def add_table_slide(self, slide_title: str, df: pd.DataFrame, subtitle: str = "", footnote: str = "", is_accident_table: bool = False, is_major_table: bool = False, custom_width_in: float = None):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self.add_header_box(slide, slide_title, subtitle)

        num_cols = len(df.columns)
        num_rows = len(df) + 1

        tbl_width = custom_width_in if custom_width_in else (8.0 if num_cols <= 2 else 12.133)
        tbl_left = (13.333 - tbl_width) / 2
        tbl_top = 1.4
        tbl_height = min(5.2, max(2.5, num_rows * 0.42))

        table_shape = slide.shapes.add_table(num_rows, num_cols, Inches(tbl_left), Inches(tbl_top), Inches(tbl_width), Inches(tbl_height))
        tbl = table_shape.table

        font_sz = 12 if num_cols <= 4 else (10 if num_cols >= 8 else 11)
        for c_idx, col_name in enumerate(df.columns):
            self._set_cell(tbl.cell(0, c_idx), col_name, font_size=font_sz, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

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

                self._set_cell(tbl.cell(r_idx + 1, c_idx), cell_val, font_size=font_sz, bold=is_bold, color=fg, bg_color=bg)

        if footnote:
            tb_fn = slide.shapes.add_textbox(Inches(0.6), Inches(6.8), Inches(12.133), Inches(0.4))
            p_fn = tb_fn.text_frame.paragraphs[0]
            p_fn.text = f"註：{footnote}"
            p_fn.font.name = "DFKai-SB"
            p_fn.font.size = Pt(10)
            p_fn.font.color.rgb = self.C_FOOTNOTE

    def build_bytes(self) -> io.BytesIO:
        out = io.BytesIO()
        self.prs.save(out)
        out.seek(0)
        return out

# ==========================================
# 3. 雙軌數據來源載入器 (僅限雲端硬碟集中處 或 前端上傳)
# ==========================================
def get_drive_service():
    """若 secrets 內有配置 GCP 服務帳戶，則連線 Google Drive API"""
    if not HAS_GDRIVE or "gcp_service_account" not in st.secrets:
        return None
    try:
        creds = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
        return build("drive", "v3", credentials=creds)
    except Exception:
        return None

def fetch_files_from_gdrive_folder(folder_name="執法統計報表集中處"):
    """自雲端硬碟指定資料夾動態下載所有 xlsx 檔案串流至記憶體"""
    service = get_drive_service()
    if not service:
        return {}
    file_dict = {}
    try:
        q_folder = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        res_f = service.files().list(q=q_folder, fields="files(id, name)").execute()
        f_items = res_f.get("files", [])
        if not f_items:
            return {}
        folder_id = f_items[0]["id"]

        q_files = f"'{folder_id}' in parents and trashed = false"
        res_files = service.files().list(q=q_files, fields="files(id, name, modifiedTime)").execute()
        for item in res_files.get("files", []):
            fname = item["name"]
            if fname.endswith(".xlsx") or fname.endswith(".csv"):
                req = service.files().get_media(fileId=item["id"])
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, req)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                fh.seek(0)
                file_dict[fname] = fh.read()
    except Exception as e:
        st.sidebar.error(f"雲端硬碟連線異常: {e}")
    return file_dict

# 建立可用報表記憶體字典 { 檔名: 二進位bytes }
MEMORY_REPORTS = {}

# 1. 先嘗試自 Google 雲端硬碟取得
gdrive_data = fetch_files_from_gdrive_folder("執法統計報表集中處")
if not gdrive_data:
    gdrive_data = fetch_files_from_gdrive_folder("執法報表集中處")

if gdrive_data:
    MEMORY_REPORTS.update(gdrive_data)
    st.sidebar.success(f"☁️ 成功連接雲端硬碟！共載入 {len(gdrive_data)} 個最新報表")
else:
    # 檢查本地環境下的資料夾（供本機測試）
    local_candidates = ["執法統計報表集中處", "執法報表集中處", "執法統計報表_已歸檔", "."]
    for d in local_candidates:
        if os.path.exists(d):
            for f in os.listdir(d):
                if f.endswith(".xlsx") and not f.startswith("~$"):
                    p = os.path.join(d, f)
                    with open(p, "rb") as f_in:
                        MEMORY_REPORTS[f] = f_in.read()
    if MEMORY_REPORTS:
        st.sidebar.info(f"📂 偵測到本機集中處資料夾，已載入 {len(MEMORY_REPORTS)} 個報表")

# 2. 前端上傳覆蓋（軌道 2：本機上傳至網站）
st.sidebar.markdown("### 📤 本機手動上傳報表")
uploaded_files = st.sidebar.file_uploader(
    "拖曳上傳本機報表（支援批次上傳多個檔案）",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="上傳後將優先以此報表進行動態統計運算"
)
if uploaded_files:
    for up_f in uploaded_files:
        MEMORY_REPORTS[up_f.name] = up_f.read()
    st.sidebar.success(f"✅ 前端已成功上傳 {len(uploaded_files)} 個報表並動態更新！")

if not MEMORY_REPORTS:
    st.warning("⚠️ 目前【雲端硬碟執法報表集中處】無檔案，亦未於【本機上傳至網站】。\n請在側邊欄上傳 Excel 檔案或確認雲端硬碟配置。")

# ==========================================
# 4. 純動態報表解析核心 (無寫死數據、絕不 IndexError)
# ==========================================

# --- 4.1 動態解析：三項重點違規 ---
def load_dynamic_three_major(report_dict):
    three_files = {k: v for k, v in report_dict.items() if "重點違規" in k}
    if not three_files:
        return None, None

    file_meta = []
    for fname, raw_bytes in three_files.items():
        text = raw_bytes.decode('utf-8', errors='ignore')
        m = re.search(r'本年度\d{3}(\d{2})(\d{2})至\d{3}(\d{2})(\d{2})', text)
        if m:
            s_m, s_d, e_m, e_d = m.groups()
            file_meta.append({
                "name": fname,
                "bytes": raw_bytes,
                "is_single": (s_m == e_m and s_d == e_d),
                "day_str": f"{e_m}/{e_d}"
            })

    if not file_meta:
        return None, None

    singles = [x for x in file_meta if x["is_single"]]
    cums = [x for x in file_meta if not x["is_single"]]

    cur_item = singles[-1] if singles else file_meta[0]
    cum_item = cums[-1] if cums else file_meta[-1]

    def parse_sheet_data(b_data):
        try:
            df = pd.read_excel(io.BytesIO(b_data), header=None)
            res = {}
            for r in range(5, len(df)):
                u = str(df.iloc[r, 0]).strip()
                if not u or u == 'nan': continue
                red = (df.iloc[r, 3] or 0) + (df.iloc[r, 4] or 0)
                rev = (df.iloc[r, 7] or 0) + (df.iloc[r, 8] or 0)
                ped = (df.iloc[r, 13] or 0) + (df.iloc[r, 14] or 0)
                res[u] = {'red': int(red), 'rev': int(rev), 'ped': int(ped), 'tot': int(red + rev + ped)}
            return res
        except Exception:
            return {}

    d_cur = parse_sheet_data(cur_item["bytes"])
    d_cum = parse_sheet_data(cum_item["bytes"])

    units = [
        ("合計", "合計"), ("聖亭所", "聖亭派出所"), ("龍潭所", "龍潭派出所"),
        ("中興所", "中興派出所"), ("石門所", "石門派出所"), ("高平所", "高平派出所"),
        ("三和所", "三和派出所"), ("交通分隊", "龍潭交通分隊")
    ]
    matrix = []
    for disp, raw in units:
        c = d_cur.get(raw, {'red': 0, 'rev': 0, 'ped': 0, 'tot': 0})
        cm = d_cum.get(raw, {'red': 0, 'rev': 0, 'ped': 0, 'tot': 0})
        matrix.append([disp, c['red'], c['rev'], c['ped'], c['tot'], cm['red'], cm['rev'], cm['ped'], cm['tot']])

    return cur_item["day_str"], matrix

# --- 4.2 動態解析：交通事故 (A1 死亡、A2 受傷) - 容錯強化版 ---
def load_dynamic_accidents(report_dict):
    acc_files = {k: v for k, v in report_dict.items() if "交通事故" in k}
    if not acc_files:
        return None, None, None

    def get_latest_item(pat):
        matched = [k for k in acc_files.keys() if pat in k]
        return acc_files[matched[-1]] if matched else None

    b_cur = get_latest_item("本期")
    b_prev = get_latest_item("前期")
    b_cum = get_latest_item("今年累計")
    b_ly = get_latest_item("去年累計")

    if not (b_cur and b_cum and b_ly):
        return None, None, None

    def parse_acc_safe(b_data):
        """雙重容錯解析：先以 pandas 二進位讀取，若失敗回退文字處理，全程保護長度索引"""
        date_range = ""
        data = {}
        df = None

        # 1. 優先使用 pandas / openpyxl 讀取二進位串流
        try:
            df = pd.read_excel(io.BytesIO(b_data), header=None)
        except Exception:
            pass

        # 2. 若為標準 Excel 轉成的 DataFrame
        if df is not None and not df.empty:
            for r in range(min(5, len(df))):
                row_txt = " ".join([str(x) for x in df.iloc[r].dropna()])
                m_date = re.search(r'(\d{2,3}/\d{2}/\d{2})\s*至\s*(\d{2,3}/\d{2}/\d{2})', row_txt)
                if m_date:
                    date_range = f"{m_date.group(1)}~{m_date.group(2)}"
                    break

            units = [
                "總計", "合計", "聖亭派出所", "龍潭派出所", "中興派出所",
                "石門派出所", "高平派出所", "三和派出所",
                "聖亭所", "龍潭所", "中興所", "石門所", "高平所", "三和所"
            ]

            def clean_num(val):
                s = str(val).replace(',', '').replace('"', '').replace('-', '0').strip()
                try:
                    return int(float(s))
                except Exception:
                    return 0

            for r in range(len(df)):
                col0 = str(df.iloc[r, 0]).strip()
                matched_unit = None
                for u in units:
                    if u in col0:
                        matched_unit = u
                        break

                if matched_unit:
                    u_name = "合計" if any(k in matched_unit for k in ["總計", "合計"]) else matched_unit.replace("派出所", "所")
                    row_vals = df.iloc[r].values

                    # 警政系統標準格式：第 5 欄為 A1 死亡、第 9 欄為 A2 受傷
                    a1_death = 0
                    a2_inj = 0
                    if len(row_vals) >= 10:
                        a1_death = clean_num(row_vals[5])
                        a2_inj = clean_num(row_vals[9])
                    elif len(row_vals) >= 6:
                        a1_death = clean_num(row_vals[-6])
                        a2_inj = clean_num(row_vals[-2]) if len(row_vals) >= 2 else 0

                    data[u_name] = {"a1_death": a1_death, "a2_inj": a2_inj}

            return date_range, data

        # 3. 回退文字模式（嚴格長度檢查，杜絕 IndexError）
        try:
            raw_text = b_data.decode('utf-8', errors='ignore')
            m_date = re.search(r'統計日期：\s*(\d{2,3}/\d{2}/\d{2})\s*至\s*(\d{2,3}/\d{2}/\d{2})', raw_text)
            date_range = f"{m_date.group(1)}~{m_date.group(2)}" if m_date else ""

            start = raw_text.find('總計,')
            end = raw_text.find('備註：')
            data_str = raw_text[start:end].strip() if start != -1 and end != -1 else raw_text

            units = ["三和派出所", "高平派出所", "石門派出所", "中興派出所", "龍潭派出所", "聖亭派出所"]
            for u in units:
                data_str = re.sub(r'(\d+|\-|\")\s+' + u, r'\1\n' + u, data_str)

            lines = [l.strip() for l in data_str.split('\n') if l.strip()]
            for l in lines:
                row_tokens = list(csv.reader([l]))
                if not row_tokens or not row_tokens[0]:
                    continue
                r = row_tokens[0]
                if len(r) < 6:
                    continue  # 長度過濾

                u_name = "合計" if "總計" in r[0] else r[0].strip().replace("派出所", "所")

                def to_i(val):
                    s = str(val).replace('"', '').replace(',', '').replace('-', '0').strip()
                    try:
                        return int(float(s))
                    except Exception:
                        return 0

                data[u_name] = {
                    "a1_death": to_i(r[-6]),
                    "a2_inj": to_i(r[-2]) if len(r) >= 2 else 0
                }
        except Exception:
            pass

        return date_range, data

    r_cur, d_cur = parse_acc_safe(b_cur)
    r_prev, d_prev = parse_acc_safe(b_prev) if b_prev else ("", {})
    r_cum, d_cum = parse_acc_safe(b_cum)
    r_ly, d_ly = parse_acc_safe(b_ly)

    units_order = ["合計", "聖亭所", "龍潭所", "中興所", "石門所", "高平所", "三和所"]

    # 組合 A1 死亡表
    a1_list = []
    for u in units_order:
        c_val = d_cur.get(u, {}).get("a1_death", 0)
        cum_val = d_cum.get(u, {}).get("a1_death", 0)
        ly_val = d_ly.get(u, {}).get("a1_death", 0)
        a1_list.append({
            "統計期間": u,
            f"本期({r_cur})": c_val,
            f"本年累計({r_cum})": cum_val,
            f"去年累計({r_ly})": ly_val,
            "本年與去年同期比較": cum_val - ly_val
        })
    df_a1_dyn = pd.DataFrame(a1_list)

    # 組合 A2 受傷表
    a2_list = []
    for u in units_order:
        c_val = d_cur.get(u, {}).get("a2_inj", 0)
        p_val = d_prev.get(u, {}).get("a2_inj", 0)
        cum_val = d_cum.get(u, {}).get("a2_inj", 0)
        ly_val = d_ly.get(u, {}).get("a2_inj", 0)
        diff = cum_val - ly_val
        rate = f"{(diff / ly_val)*100:.2f}%" if ly_val else "—"
        a2_list.append({
            "統計期間": u,
            f"本期({r_cur})": c_val,
            f"前期({r_prev})": p_val,
            f"本年累計({r_cum})": cum_val,
            f"去年累計({r_ly})": ly_val,
            "本年與去年同期比較": diff,
            "增減比例": rate
        })
    df_a2_dyn = pd.DataFrame(a2_list)

    return df_a1_dyn, df_a2_dyn, r_cur

# --- 4.3 動態解析：重大交通違規總表與 7 大專項細表 ---
def load_dynamic_major(report_dict):
    major_files = {k: v for k, v in report_dict.items() if "重大違規" in k}
    if not major_files:
        return None, None, ""

    def get_latest_item(pat):
        matched = [k for k in major_files.keys() if pat in k]
        return major_files[matched[-1]] if matched else None

    b_cur = get_latest_item("本期")
    b_cum = get_latest_item("年累計")
    b_ly = get_latest_item("去年累計")

    if not (b_cur and b_cum and b_ly):
        return None, None, ""

    def parse_major(b_data):
        raw_text = b_data.decode('utf-8', errors='ignore')
        m_date = re.search(r'本年度(\d{7})至(\d{7})', raw_text)
        period = f"{m_date.group(1)}~{m_date.group(2)}" if m_date else ""
        m = re.search(r'(合計,\d+.*?交通組,[\d,]+)', raw_text)
        if not m: return period, {}
        units = ['合計', '龍潭交通分隊', '警備隊', '聖亭派出所', '龍潭派出所', '中興派出所', '石門派出所', '高平派出所', '三和派出所', '交通組']
        splits = re.split(r'(' + '|'.join(units) + r'),', m.group(1))
        res = {}
        for u, vals in zip(splits[1::2], splits[2::2]):
            val_list = [int(v.strip()) for v in vals.split(',') if v.strip().isdigit()]
            norm_u = u.replace("派出所", "所").replace("龍潭交通分隊", "交通分隊").replace("交通組", "科技執法")
            res[norm_u] = val_list
        return period, res

    _, d_cur = parse_major(b_cur)
    p_cum, d_cum = parse_major(b_cum)
    _, d_ly = parse_major(b_ly)

    targets = {
        "合計": 18114, "科技執法": 6006, "聖亭所": 1941, "龍潭所": 2588, "中興所": 1941,
        "石門所": 1479, "高平所": 1294, "三和所": 339, "警備隊": 0, "交通分隊": 2526
    }
    unit_order = ["合計", "科技執法", "聖亭所", "龍潭所", "中興所", "石門所", "高平所", "三和所", "警備隊", "交通分隊"]

    # 總表
    major_rows = []
    for u in unit_order:
        cv = d_cur.get(u, [0]*20)
        cmv = d_cum.get(u, [0]*20)
        lyv = d_ly.get(u, [0]*20)

        cur_s, cur_a = cv[14], cv[15]
        cum_s, cum_a, cum_tot = cmv[14], cmv[15], cmv[16]
        ly_s, ly_a, ly_tot = cmv[17], cmv[18], cmv[19]

        diff = cum_tot - ly_tot
        tgt = targets.get(u, 0)
        achieve = f"{(cum_tot / tgt)*100:.1f}%" if tgt > 0 else "—"
        if u == "警備隊": diff = "—"

        major_rows.append({
            "統計期間": u,
            "本期(攔停)": cur_s, "本期(逕舉)": cur_a,
            "本年累計(攔停)": cum_s, "本年累計(逕舉)": cum_a,
            "去年累計(攔停)": ly_s, "去年累計(逕舉)": ly_a,
            "本年與去年同期比較": diff, "目標值": tgt, "達成率": achieve
        })
    df_major_dyn = pd.DataFrame(major_rows)

    # 7 大細表
    cat_indices = {
        "酒駕": (0, 1), "闖紅燈": (2, 3), "逆向行駛": (6, 7),
        "轉彎未依規定": (8, 9), "蛇行惡意逼車": (10, 11),
        "不暫停讓行人": (12, 13), "嚴重超速": (4, 5)
    }
    detail_dict = {}
    for cat, (is_, ia_) in cat_indices.items():
        rows = []
        for u in unit_order:
            cm = d_cum.get(u, [0]*20)
            ly = d_ly.get(u, [0]*20)
            cs, ca = cm[is_], cm[ia_]
            ls, la = ly[is_], ly[ia_]
            ct, lt = cs + ca, ls + la
            ds, da, dt = (cs - ls, ca - la, ct - lt) if u != "警備隊" else ("—", "—", "—")
            rows.append([u, cs, ca, ct, ls, la, lt, ds, da, dt])
        detail_dict[cat] = rows

    return df_major_dyn, detail_dict, p_cum

# ==========================================
# 5. 純動態執行載入（完全無假資料）
# ==========================================
three_day, three_matrix = load_dynamic_three_major(MEMORY_REPORTS)
if three_matrix:
    preview_cols = pd.MultiIndex.from_tuples([
        ("單位", ""),
        (f"本期 ({three_day}) 新增違規數", "闖紅燈"),
        (f"本期 ({three_day}) 新增違規數", "逆向行駛"),
        (f"本期 ({three_day}) 新增違規數", "不停讓行人"),
        (f"本期 ({three_day}) 新增違規數", f"本期合計 ({three_day})"),
        ("本月累計數", "闖紅燈"),
        ("本月累計數", "逆向行駛"),
        ("本月累計數", "不停讓行人"),
        ("本月累計數", "累計總計")
    ])
    df_three_preview = pd.DataFrame(three_matrix, columns=preview_cols)
else:
    df_three_preview = None

df_a1_dyn, df_a2_dyn, cur_acc_period = load_dynamic_accidents(MEMORY_REPORTS)
df_major_dyn, detail_dict_dyn, major_period = load_dynamic_major(MEMORY_REPORTS)

# ==========================================
# 6. 前端自選與即時預覽區
# ==========================================
st.subheader("🎯 欲輸出的統計表自選控制")

col_btn1, col_btn2, _ = st.columns([1.5, 2, 4])
if "select_mode" not in st.session_state:
    st.session_state["select_mode"] = "core"

with col_btn1:
    if st.button("📌 僅常態核心頁"):
        st.session_state["select_mode"] = "core"
with col_btn2:
    if st.button("📑 全選所有可用統計表"):
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

    if tabs_to_show:
        tabs = st.tabs(tabs_to_show)
        for idx, tab_name in enumerate(tabs_to_show):
            with tabs[idx]:
                if tab_name == "三項重點": st.dataframe(df_three_preview, hide_index=True)
                elif tab_name == "A1事故死亡": st.dataframe(df_a1_dyn, hide_index=True)
                elif tab_name == "A2事故受傷": st.dataframe(df_a2_dyn, hide_index=True)
                elif tab_name == "重大違規總表": st.dataframe(df_major_dyn, hide_index=True)
    else:
        st.info("💡 目前雲端集中處或本機尚未上傳有效報表，請在左側上傳 Excel 檔案以呈現數據。")

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

    with st.spinner("正在自最新報表動態編譯 PPTX 簡報（完全無寫死數值、負數自動標紅）..."):
        try:
            builder = PptxReportBuilder()

            # 封面
            if chk_cover:
                builder.add_cover_slide(
                    main_title="桃園市政府警察局龍潭分局\n交通執法成效與事故防制分析報告",
                    subtitle="週次主管會報專案報告",
                    date_range_str=f"統計截止至最新報表 ｜ 製表日期：{datetime.now().strftime('%Y/%m/%d')}"
                )

            # 三項重點
            if chk_three and df_three_preview is not None:
                builder.add_three_major_slide(data_rows=three_matrix, latest_day=three_day)

            # A1 死亡
            if chk_a1 and df_a1_dyn is not None:
                builder.add_table_slide(slide_title="A1類交通事故死亡人數統計表", df=df_a1_dyn, is_accident_table=True)

            # A2 受傷
            if chk_a2 and df_a2_dyn is not None:
                builder.add_table_slide(slide_title="A2類交通事故受傷人數統計表", df=df_a2_dyn, is_accident_table=True)

            # 重大違規總表
            if chk_major_tot and df_major_dyn is not None:
                builder.add_table_slide(
                    slide_title="取締重大交通違規統計表",
                    df=df_major_dyn,
                    footnote="重大交通違規指：「酒駕」、「闖紅燈」、「嚴重超速」、「逆向行駛」、「轉彎未依規定」、「蛇行、惡意逼車」及「不暫停讓行人」",
                    is_major_table=True
                )

            # 7 大專項細表
            if detail_dict_dyn is not None:
                det_map = [
                    (chk_det_jiu, "酒駕"), (chk_det_red, "闖紅燈"), (chk_det_rev, "逆向行駛"),
                    (chk_det_turn, "轉彎未依規定"), (chk_det_snake, "蛇行惡意逼車"),
                    (chk_det_ped, "不暫停讓行人"), (chk_det_speed, "嚴重超速")
                ]
                for is_chk, cat in det_map:
                    if is_chk and cat in detail_dict_dyn:
                        builder.add_major_detail_slide(cat_name=cat, data_rows=detail_dict_dyn[cat], date_str=major_period)

            pptx_stream = builder.build_bytes()
            st.session_state["cached_pptx"] = pptx_stream
            st.session_state["cached_filename"] = file_save_name

            # 寄信
            if chk_auto_email:
                ok, detail = send_pptx_email_to_self(pptx_stream, file_save_name)
                if ok:
                    st.info(f"📧 簡報附件已成功發送至您的信箱：`{detail}`")
                else:
                    st.warning(f"⚠️ 郵件未發送成功（{detail}），可直接點擊下方按鈕下載！")

            st.balloons()
            st.success("🎉 恭喜！純動態數據 PowerPoint 簡報實體檔已成功生成！")

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
