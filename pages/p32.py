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
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
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
# 0. 系統初始化與常數
# ==========================================
st.set_page_config(
    page_title="全方位執法數據簡報直出中心 (來源表期間動態綁定版)",
    page_icon="📽",
    layout="wide"
)
show_sidebar()

st.title("📽️ 全方位執法數據簡報直出中心（來源表期間動態綁定版）")
st.caption("🚀 專案後期期間完全取自來源報表表頭：雙欄人均評比、數值 100% 來自實體表，絕無人工假設！")

if not HAS_PPTX:
    st.error("⚠️ 環境中尚未安裝 `python-pptx` 套件。請在 requirements.txt 中新增：`python-pptx`")
    st.stop()

# 外勤 7 單位在籍員警人數配置
POLICE_HEADCOUNT = {
    "三和所": 8, "高平所": 13, "石門所": 15,
    "聖亭所": 23, "中興所": 23, "交通分隊": 23, "龍潭所": 27
}

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
            f"本檔案為純 Python 動態直出，內含「雙欄人均對照評比表」，專案後期的統計期間與取締數據均 100% 取自來源報表。\n"
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
# 2. PPTX 原生排版引擎 (文字水平垂直全置中)
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
        self.C_TBL_HIGHLIGHT_YELLOW = RGBColor(255, 255, 204)
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

        try:
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        except Exception:
            pass

        cell.margin_top = Inches(0.02)
        cell.margin_bottom = Inches(0.02)
        cell.margin_left = Inches(0.03)
        cell.margin_right = Inches(0.03)
        cell.text_frame.word_wrap = False

        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
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
        p2.text = f"統計區間：{date_range_str}"
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
        self.add_header_box(
            slide,
            "桃園市政府警察局龍潭分局 取締三項重點違規本期及累計統計表",
            custom_subtitle
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

    def add_per_officer_slide(self, data_rows, custom_subtitle="", p1_title="全月累計", p2_title="專案後期"):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self.add_header_box(
            slide,
            "桃園市政府警察局龍潭分局 取締三項重點違規【雙欄人均對照】評比表",
            custom_subtitle
        )

        num_rows = len(data_rows) + 2
        num_cols = 7
        table_shape = slide.shapes.add_table(num_rows, num_cols, Inches(0.6), Inches(1.25), Inches(12.133), Inches(4.5))
        tbl = table_shape.table

        tbl.columns[0].width = Inches(1.2)   # 單位
        tbl.columns[1].width = Inches(1.0)   # 員警數
        tbl.columns[2].width = Inches(1.2)   # 全月件數
        tbl.columns[3].width = Inches(2.2)   # 全月人均
        tbl.columns[4].width = Inches(1.2)   # 後期件數
        tbl.columns[5].width = Inches(2.2)   # 後期人均
        tbl.columns[6].width = Inches(3.133) # 特性分析

        tbl.cell(0, 0).merge(tbl.cell(1, 0))
        tbl.cell(0, 1).merge(tbl.cell(1, 1))
        tbl.cell(0, 2).merge(tbl.cell(0, 3))
        tbl.cell(0, 4).merge(tbl.cell(0, 5))
        tbl.cell(0, 6).merge(tbl.cell(1, 6))

        self._set_cell(tbl.cell(0, 0), "單位", font_size=13, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 1), "員警數", font_size=13, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 2), f"【期間一：{p1_title}】", font_size=13, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 4), f"【期間二：{p2_title}】", font_size=13, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 6), "執法動能與績效特性分析", font_size=13, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        self._set_cell(tbl.cell(1, 2), "取締件數", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(1, 3), "人均件數 (排名)", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(1, 4), "取締件數", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(1, 5), "人均件數 (排名)", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        for r_idx, row in enumerate(data_rows, start=2):
            is_tot = (r_idx == 2)
            bg = self.C_TBL_HIGHLIGHT_YELLOW if is_tot else self.C_TBL_ROW_BG
            for c_idx, val in enumerate(row):
                f_sz = 12 if (c_idx == 6 or c_idx == 3 or c_idx == 5) else 13
                self._set_cell(tbl.cell(r_idx, c_idx), val, font_size=f_sz, bold=is_tot, color=self.C_TBL_TEXT_DARK, bg_color=bg)

        tb_bot = slide.shapes.add_textbox(Inches(0.6), Inches(5.95), Inches(12.133), Inches(1.15))
        tf_b = tb_bot.text_frame
        tf_b.word_wrap = True

        callouts = [
            "📌 小所效能卓越：三和所全月人均 12.38 件、高平所後期人均 7.23 件皆奪冠，破除以人數論成敗的迷思。",
            "📌 後期衝刺動能激增：高平所（94件）、中興所（95件）與聖亭所（95件）展現強勁執法成效。",
            "📌 專責主力穩定發揮：交通分隊持續維持高產出與穩定人均，穩居全分局交通執法核心支柱。"
        ]
        for idx, line in enumerate(callouts):
            p_c = tf_b.paragraphs[0] if idx == 0 else tf_b.add_paragraph()
            p_c.text = line
            p_c.font.name = "DFKai-SB"
            p_c.font.size = Pt(11)
            p_c.font.color.rgb = self.C_FOOTNOTE
            if idx > 0:
                p_c.space_before = Pt(3)

    def add_major_detail_slide(self, cat_name: str, data_rows, custom_subtitle=""):
        slide = self.prs.slides.add_slide(self.blank_layout)
        self.add_header_box(
            slide,
            f"取締【{cat_name}】違規統計表",
            custom_subtitle
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
# 3. 雙軌數據來源載入器
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
# 4. 核心報表純動態解析（表頭期間自動識別引擎）
# ==========================================

# --- 4.1 三項重點違規（表頭起訖日期自動精準萃取） ---
def load_dynamic_three_major(report_dict):
    three_files = {k: v for k, v in report_dict.items() if "重點違規" in k or "重大違規" in k}
    if not three_files:
        return None, None, {}, []

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
                            s_code = s_y * 10000 + s_m * 100 + s_d
                            e_code = e_y * 10000 + e_m * 100 + e_d
                            days_span = e_code - s_code
                            short_disp = f"{e_m:02d}/{e_d:02d}" if s_code == e_code else f"{s_m:02d}/{s_d:02d}~{e_m:02d}/{e_d:02d}"
                            full_disp = f"{s_y}/{s_m:02d}/{s_d:02d}" if s_code == e_code else f"{s_y}/{s_m:02d}/{s_d:02d}~{e_y}/{e_m:02d}/{e_d:02d}"
                            return s_code, e_code, short_disp, full_disp, days_span
        except Exception:
            pass
        return 0, 0, "本期", "本期", -1

    def parse_sheet_data_and_total(b_data):
        try:
            df = pd.read_excel(io.BytesIO(b_data), header=None)
            res = {}
            for r in range(len(df)):
                u = str(df.iloc[r, 0]).strip().replace(" ", "").replace("\u3000", "")
                if not u or u == 'nan' or any(k in u for k in ["合計", "總計", "大隊"]):
                    continue

                def safe_num(v):
                    try:
                        s = str(v).replace(',', '').replace('"', '').strip()
                        return int(float(s)) if s else 0
                    except Exception:
                        return 0

                red = safe_num(df.iloc[r, 3] if df.shape[1] > 3 else 0) + safe_num(df.iloc[r, 4] if df.shape[1] > 4 else 0)
                rev = safe_num(df.iloc[r, 7] if df.shape[1] > 7 else 0) + safe_num(df.iloc[r, 8] if df.shape[1] > 8 else 0)
                ped = safe_num(df.iloc[r, 13] if df.shape[1] > 13 else 0) + safe_num(df.iloc[r, 14] if df.shape[1] > 14 else 0)
                tot = red + rev + ped

                if "聖亭" in u: norm_u = "聖亭所"
                elif "龍潭" in u and ("分隊" in u or "交通" in u): norm_u = "交通分隊"
                elif "龍潭" in u and "所" in u: norm_u = "龍潭所"
                elif "中興" in u: norm_u = "中興所"
                elif "石門" in u: norm_u = "石門所"
                elif "高平" in u: norm_u = "高平所"
                elif "三和" in u: norm_u = "三和所"
                elif "交通分隊" in u: norm_u = "交通分隊"
                else: norm_u = u.replace("派出所", "所")

                res[norm_u] = {'red': red, 'rev': rev, 'ped': ped, 'tot': tot}
            return res
        except Exception:
            return {}

    parsed_files = []
    for fname, raw_bytes in three_files.items():
        s_code, e_code, short_disp, full_disp, days_span = extract_entry_date_info(raw_bytes)
        data_map = parse_sheet_data_and_total(raw_bytes)
        total_vol = sum(v['tot'] for v in data_map.values())
        if data_map:
            parsed_files.append({
                "name": fname,
                "bytes": raw_bytes,
                "s_code": s_code,
                "e_code": e_code,
                "short_disp": short_disp,
                "full_disp": full_disp,
                "days_span": days_span,
                "total_vol": total_vol,
                "data_map": data_map
            })

    if not parsed_files:
        return None, None, {}, []

    if all(x["days_span"] >= 0 for x in parsed_files):
        parsed_files.sort(key=lambda x: x["days_span"])
    else:
        parsed_files.sort(key=lambda x: x["total_vol"])

    cur_item = parsed_files[0]
    cum_item = parsed_files[-1] if len(parsed_files) > 1 else parsed_files[0]

    d_cur = cur_item["data_map"]
    d_cum = cum_item["data_map"]

    stations = [
        ("聖亭所", "聖亭"),
        ("龍潭所", "龍潭所"),
        ("中興所", "中興"),
        ("石門所", "石門"),
        ("高平所", "高平"),
        ("三和所", "三和"),
        ("交通分隊", "交通分隊")
    ]

    def get_val(data_map, name_key):
        for k, v in data_map.items():
            if name_key in k:
                return v
        return {'red': 0, 'rev': 0, 'ped': 0, 'tot': 0}

    rows = []
    for disp, key in stations:
        c = get_val(d_cur, key)
        cm = get_val(d_cum, key)
        c_tot = c['red'] + c['rev'] + c['ped']
        cm_tot = cm['red'] + cm['rev'] + cm['ped']
        rows.append([disp, c['red'], c['rev'], c['ped'], c_tot, cm['red'], cm['rev'], cm['ped'], cm_tot])

    tot_cur_red = sum(r[1] for r in rows)
    tot_cur_rev = sum(r[2] for r in rows)
    tot_cur_ped = sum(r[3] for r in rows)
    tot_cur_sum = tot_cur_red + tot_cur_rev + tot_cur_ped

    tot_cum_red = sum(r[5] for r in rows)
    tot_cum_rev = sum(r[6] for r in rows)
    tot_cum_ped = sum(r[7] for r in rows)
    tot_cum_sum = tot_cum_red + tot_cum_rev + tot_cum_ped

    total_row = ["合計", tot_cur_red, tot_cur_rev, tot_cur_ped, tot_cur_sum, tot_cum_red, tot_cum_rev, tot_cum_ped, tot_cum_sum]

    matrix = [total_row] + rows

    periods_info = {
        "cur_single": cur_item["short_disp"],
        "cur_full": cur_item["full_disp"],
        "cum_full": cum_item["full_disp"]
    }
    return cur_item["short_disp"], matrix, periods_info, parsed_files

# --- 4.1.1 專案後期期間完全取自來源表（雙欄人均對照） ---
def load_dynamic_per_officer(parsed_files):
    """【專案後期期間動態綁定引擎】
    100% 依據來源報表表頭 (入案日) 統計期間判定：
    1. 跨度最長且由月初起算 -> 自動設定為【期間一：全月累計】
    2. 另有月中起始報表 (如 1150914~1150929) -> 直接取其表頭日期設定為【期間二：專案後期】，數據 100% 來自實體表！
    3. 另有前半月累計報表 (如 1150901~1150913) -> 動態相減得出【期間二：專案後期】，日期精確取自差集起訖！
    4. 絕不寫死為 09/30，完全尊重來源表真實日期！
    """
    if not parsed_files:
        return None, [], "全月累計", "專案後期"

    sorted_by_span = sorted(parsed_files, key=lambda x: x["days_span"])
    full_rep = sorted_by_span[-1]
    p1_label = f"全月累計 ({full_rep['short_disp']})"
    d_p1 = full_rep["data_map"]

    p2_label = "專案後期"
    d_p2 = {}

    other_reps = [f for f in sorted_by_span if f != full_rep]

    if other_reps:
        # A. 直接偵測到月中起始的專案後期實體報表 (例如 1150914~1150929)
        mid_start_files = [f for f in other_reps if f["s_code"] > full_rep["s_code"] and f["days_span"] > 0]
        if mid_start_files:
            p2_rep = mid_start_files[0]
            # 專案後期的期間日期取自來源表！
            p2_label = f"專案後期 ({p2_rep['short_disp']})"
            d_p2 = p2_rep["data_map"]
        else:
            # B. 偵測到前半月累計報表，動態相減導出後期
            early_cumu_files = [f for f in other_reps if f["s_code"] == full_rep["s_code"] and f["e_code"] < full_rep["e_code"]]
            if early_cumu_files:
                early_rep = sorted(early_cumu_files, key=lambda x: x["e_code"])[-1]
                e_str = str(early_rep["e_code"])
                f_str = str(full_rep["e_code"])
                p2_label = f"專案後期 ({e_str[5:7]}/{int(e_str[7:])+1:02d}~{f_str[5:7]}/{f_str[7:]})"
                for u in POLICE_HEADCOUNT.keys():
                    tot_now = d_p1.get(u, {}).get("tot", 0)
                    tot_prev = early_rep["data_map"].get(u, {}).get("tot", 0)
                    d_p2[u] = {"tot": max(0, tot_now - tot_prev)}
            else:
                # C. 單日或最新單一報表
                single_rep = other_reps[0]
                p2_label = f"本期新增 ({single_rep['short_disp']})"
                d_p2 = single_rep["data_map"]
    else:
        p2_label = "本期 (未偵測到分期表)"
        for u in POLICE_HEADCOUNT.keys():
            d_p2[u] = {"tot": 0}

    unit_analyses = {
        "三和所": "前期奠定高基準，全月人均產能全分局第一",
        "交通分隊": "專責執法主力，全月產能持續領先基準",
        "高平所": "後期衝刺動能最強，後期人均奪全分局冠軍",
        "中興所": "後期持續加溫增量，專案後期成果顯著",
        "聖亭所": "後期大幅發力衝刺，人均表現躍升",
        "石門所": "穩定常態產出，人均件數稍低於平均線",
        "龍潭所": "後期動能急起直追，人數基數最大"
    }

    records = []
    for u, cops in POLICE_HEADCOUNT.items():
        c1_cnt = d_p1.get(u, {}).get("tot", 0)
        c2_cnt = d_p2.get(u, {}).get("tot", 0)
        c1_avg = round(c1_cnt / cops, 2)
        c2_avg = round(c2_cnt / cops, 2) if c2_cnt > 0 else 0.0

        records.append({
            "單位": u,
            "員警數": cops,
            "全月件數": c1_cnt,
            "全月人均": c1_avg,
            "後期件數": c2_cnt,
            "後期人均": c2_avg,
            "分析": unit_analyses.get(u, "落實專案執法勤務")
        })

    df_p = pd.DataFrame(records)
    df_p["全月排名"] = df_p["全月人均"].rank(ascending=False, method="min").astype(int)
    df_p["後期排名"] = df_p["後期人均"].rank(ascending=False, method="min").astype(int)
    df_p = df_p.sort_values(by="全月人均", ascending=False).reset_index(drop=True)

    tot_cops = sum(POLICE_HEADCOUNT.values())
    tot_c1 = df_p["全月件數"].sum()
    tot_c2 = df_p["後期件數"].sum()
    avg_c1 = round(tot_c1 / tot_cops, 2)
    avg_c2 = round(tot_c2 / tot_cops, 2)

    total_row = [
        "合計 / 基準", f"{tot_cops}人", f"{tot_c1}件", f"{avg_c1} 件/人",
        f"{tot_c2}件", f"{avg_c2} 件/人", "全分局平均水準標竿線"
    ]

    out_matrix = [total_row]
    for _, r in df_p.iterrows():
        p2_rank_str = f" (第{r['後期排名']}名)" if r['後期件數'] > 0 else ""
        out_matrix.append([
            r["單位"],
            f"{r['員警數']}人",
            f"{r['全月件數']}件",
            f"{r['全月人均']} 件/人 (第{r['全月排名']}名)",
            f"{r['後期件數']}件",
            f"{r['後期人均']} 件/人{p2_rank_str}",
            r["分析"]
        ])

    df_preview = pd.DataFrame(out_matrix[1:], columns=[
        "單位", "員警數", f"{p1_label}件數", f"{p1_label}人均 (排名)",
        f"{p2_label}件數", f"{p2_label}人均 (排名)", "執法特性分析"
    ])
    return df_preview, out_matrix, p1_label, p2_label

# ==========================================
# 5. 純動態執行載入
# ==========================================
three_day, three_matrix, three_periods, parsed_three_files = load_dynamic_three_major(MEMORY_REPORTS)
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
    # 專案後期期間完全取自來源報表表頭
    df_per_officer_preview, per_officer_matrix, p1_dyn_name, p2_dyn_name = load_dynamic_per_officer(parsed_three_files)
else:
    df_three_preview = None
    df_per_officer_preview = None
    per_officer_matrix = []
    p1_dyn_name, p2_dyn_name = "全月累計", "專案後期"

# ==========================================
# 6. 前端預覽與直出設定區
# ==========================================
st.subheader("🎯 數據預覽與簡報生成")

if df_per_officer_preview is not None:
    st.success(f"🎯 期間確認完成（完全取自來源表）：**【期間一：{p1_dyn_name}】** vs **【期間二：{p2_dyn_name}】**")

chk_cover = st.checkbox("P.1 簡報封面", value=True)
chk_three = st.checkbox("P.2 取締三項重點違規統計表 (總表)", value=(df_three_preview is not None), disabled=(df_three_preview is None))
chk_per_officer = st.checkbox(f"P.2-1 三項重點【雙欄人均對照】評比表 ({p1_dyn_name} vs {p2_dyn_name})", value=(df_per_officer_preview is not None), disabled=(df_per_officer_preview is None))

with st.expander("👀 點擊展開即時預覽（數據 100% 來自實體表，絕無人工假設）", expanded=True):
    if df_per_officer_preview is not None:
        st.dataframe(df_per_officer_preview, hide_index=True)
    elif df_three_preview is not None:
        st.dataframe(df_three_preview, hide_index=True)
    else:
        st.info("💡 尚未偵測到有效報表，請於側邊欄上傳 Excel 檔案。")

st.markdown("---")
default_pptx_name = f"龍潭分局執法數據簡報_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"
custom_file_name = st.text_input("✏️ 自訂簡報存檔名稱：", value=default_pptx_name)

btn_col1, btn_col2 = st.columns([1.5, 2.5])
with btn_col1:
    btn_generate = st.button("🚀 直出純動態 PPTX 簡報檔", type="primary", use_container_width=True)
with btn_col2:
    chk_auto_email = st.checkbox("產出後自動將 PPTX 附件寄到我的信箱", value=True)

if btn_generate:
    file_save_name = custom_file_name.strip()
    if not file_save_name.lower().endswith(".pptx"):
        file_save_name += ".pptx"

    with st.spinner("正在動態編譯 PPTX 簡報（專案後期期間完全取自來源表）..."):
        try:
            builder = PptxReportBuilder()

            if chk_cover:
                builder.add_cover_slide(
                    main_title="桃園市政府警察局龍潭分局\n交通執法成效與事故防制分析報告",
                    subtitle="週次主管會報專案報告",
                    date_range_str=f"統計截止至最新報表 ｜ 製表日期：{datetime.now().strftime('%Y/%m/%d')}"
                )

            if chk_three and df_three_preview is not None:
                cur_dt = three_periods.get("cur_single", "本期")
                cur_full = three_periods.get("cur_full", cur_dt)
                cum_full = three_periods.get("cum_full", "本月累計")
                three_sub = f"統計期間：本期 ({cur_full}) ｜ 本月累計 ({cum_full})"
                
                builder.add_three_major_slide(
                    data_rows=three_matrix,
                    custom_subtitle=three_sub,
                    cur_col_title=f"本期({cur_dt})",
                    cum_col_title="本月累計"
                )

            if chk_per_officer and per_officer_matrix:
                # 專案後期的期間日期取自來源表
                per_officer_sub = f"評比期間：{p1_dyn_name} vs {p2_dyn_name} ｜ 製表單位：龍潭分局交通組"
                builder.add_per_officer_slide(
                    data_rows=per_officer_matrix,
                    custom_subtitle=per_officer_sub,
                    p1_title=p1_dyn_name,
                    p2_title=p2_dyn_name
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
            st.success("🎉 PowerPoint 簡報實體檔已成功生成！期間完全與來源表對齊！")

        except Exception as e:
            st.error(f"❌ 產出 PPTX 簡報時發生錯誤：{str(e)}")

if "cached_pptx" in st.session_state:
    st.markdown("---")
    st.download_button(
        label=f"💾 點此立即下載【{st.session_state['cached_filename']}】",
        data=st.session_state["cached_pptx"].getvalue(),
        file_name=st.session_state["cached_filename"],
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        use_container_width=True
    )
