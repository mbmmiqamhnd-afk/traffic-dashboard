import io
import os
import re
import smtplib
import urllib.parse as _ul
from datetime import datetime, timedelta
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import streamlit as st

# 嘗試載入 python-pptx，若環境尚未安裝則提供提示
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

# 載入自訂側邊欄
try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        pass

# ==========================================
# 0. 系統初始化與側邊欄
# ==========================================
st.set_page_config(
    page_title="全方位執法數據簡報直出中心 (PPTX 直出版)",
    page_icon="📽️",
    layout="wide"
)
show_sidebar()

st.title("📽️ 全方位執法數據簡報直出中心（PPTX 實體檔直出版）")
st.caption("🚀 做法 A 架構：純 Python 本地直出實體 PowerPoint (.pptx) 簡報，徹底擺脫 Google 母本與雲端空間配額限制，色彩 100% 精準還原！")

if not HAS_PPTX:
    st.error("⚠️ 偵測到環境中尚未安裝 `python-pptx` 套件。請在終端機或 requirements.txt 中執行：`pip install python-pptx`")
    st.stop()

# ==========================================
# 1. 郵件通知函式（直接夾帶附件寄給自己）
# ==========================================
def send_pptx_email_to_self(pptx_bytes: io.BytesIO, file_name: str) -> tuple:
    """
    讀取 st.secrets["email"] 或頂層 SMTP 設定，
    將產出之 PPTX 簡報檔夾帶為附件寄送至寄件者本人信箱。
    """
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
            f"系統已自動完成交通執法數據簡報之編譯結算。\n"
            f"附件為最新產出之 PowerPoint 簡報實體檔【{file_name}】。\n\n"
            f"本檔案可直接在電腦以 Microsoft PowerPoint、WPS 編輯，或直接拖拉上傳至 Google 雲端硬碟使用。\n\n"
            f"本信件由交通執法自動化分析引擎發送。"
        )
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        # 夾帶 PPTX 實體檔案
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
# 2. PPTX 原生簡報排版引擎 (PptxReportBuilder)
# ==========================================
class PptxReportBuilder:
    """專門負責將執法數據以 16:9 比例直出高品質 PPTX 簡報，色彩嚴格還原原版 Google 簡報設定"""
    def __init__(self):
        self.prs = Presentation()
        # 設定為標準 16:9 寬螢幕尺寸 (13.333 x 7.5 英吋)
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]  # 全空白版型

        # ====================================================
        # 🎨 原版精準色票定義
        # ====================================================
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
        """為儲存格設定細微精準邊框線，維持表格工整俐落"""
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
        """統整設定儲存格內容與樣式，100% 還原原版色彩與邊框"""
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
        """封面頁：原版深海軍藍大器全幅底色"""
        slide = self.prs.slides.add_slide(self.blank_layout)

        bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height
        )
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
        """投影片頂端標題區塊"""
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

    def add_three_major_slide(self, data_rows, latest_day="09/17"):
        """P.2 三項重點違規統計表 (雙層表頭 + 原版表頭藍與合計淺藍)"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self.add_header_box(
            slide,
            "桃園市政府警察局龍潭分局 取締三項重點違規本期及累計統計表",
            f"統計期間：自 115 年 9 月 1 日起至本期({latest_day})止 ｜ 製表單位：龍潭分局交通組"
        )

        num_rows = len(data_rows) + 2
        num_cols = 9
        table_shape = slide.shapes.add_table(
            num_rows, num_cols, Inches(0.6), Inches(1.4), Inches(12.133), Inches(5.4)
        )
        tbl = table_shape.table

        tbl.cell(0, 0).merge(tbl.cell(1, 0))
        tbl.cell(0, 1).merge(tbl.cell(0, 4))
        tbl.cell(0, 5).merge(tbl.cell(0, 8))

        self._set_cell(tbl.cell(0, 0), "單位", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 1), f"本期 ({latest_day}) 新增違規數", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)
        self._set_cell(tbl.cell(0, 5), "115年9月1日起累計數", font_size=12, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        sub_headers = ["", "闖紅燈", "逆向行駛", "不停讓行人", f"本期合計\n({latest_day})", "闖紅燈", "逆向行駛", "不停讓行人", "累計總計"]
        for c in range(1, 9):
            self._set_cell(tbl.cell(1, c), sub_headers[c], font_size=11, bold=True, color=self.C_TBL_HEADER_TEXT, bg_color=self.C_TBL_HEADER_BG)

        for r_idx, row in enumerate(data_rows, start=2):
            is_tot = (r_idx == 2)
            bg = self.C_TBL_HIGHLIGHT_BG if is_tot else self.C_TBL_ROW_BG
            for c_idx, val in enumerate(row):
                self._set_cell(tbl.cell(r_idx, c_idx), val, font_size=11, bold=is_tot, color=self.C_TBL_TEXT_DARK, bg_color=bg)

    def add_major_detail_slide(self, cat_name: str, data_rows, date_str="0101-0915"):
        """重大違規 7 大專項細表 (負數及該列單位名稱標紅)"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self.add_header_box(
            slide,
            f"取締【{cat_name}】違規統計表 (累計至 {date_str})",
            "口徑包含現場攔停與逕行舉發 ｜ 製表單位：龍潭分局交通組"
        )

        num_rows = len(data_rows) + 2
        num_cols = 10
        table_shape = slide.shapes.add_table(
            num_rows, num_cols, Inches(0.6), Inches(1.4), Inches(12.133), Inches(5.4)
        )
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
        """標準表格投影片 (A1, A2, 重大違規總表, 超載, 靜桃, 科技執法)"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self.add_header_box(slide, slide_title, subtitle)

        num_cols = len(df.columns)
        num_rows = len(df) + 1

        tbl_width = custom_width_in if custom_width_in else (8.0 if num_cols <= 2 else 12.133)
        tbl_left = (13.333 - tbl_width) / 2
        tbl_top = 1.4
        tbl_height = min(5.2, max(2.5, num_rows * 0.42))

        table_shape = slide.shapes.add_table(
            num_rows, num_cols, Inches(tbl_left), Inches(tbl_top), Inches(tbl_width), Inches(tbl_height)
        )
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
        """編譯並輸出記憶體 BytesIO 串流"""
        out = io.BytesIO()
        self.prs.save(out)
        out.seek(0)
        return out

# ==========================================
# 3. 做法 B：自動掃描雲端硬碟/本機資料夾報表引擎
# ==========================================
def scan_and_load_three_major():
    """
    自動在當前目錄及『執法報表集中處』資料夾掃描《重點違規統計表》
    自動辨識【月累計表】與【單日本期表】，並解析出最新統計數據
    """
    search_dirs = [
        ".",
        "執法報表集中處",
        os.path.expanduser("~/Google 雲端硬碟/執法報表集中處"),
        os.path.expanduser("~/Google Drive/執法報表集中處"),
        os.path.expanduser("~/我的雲端硬碟/執法報表集中處"),
    ]

    found_files = []
    for d in search_dirs:
        if os.path.exists(d):
            for f in os.listdir(d):
                if "重點違規" in f and f.endswith(".xlsx") and not f.startswith("~$"):
                    full_p = os.path.join(d, f)
                    if full_p not in found_files:
                        found_files.append(full_p)

    if not found_files:
        return None, None, "未在『執法報表集中處』或目錄下找到《重點違規統計表》Excel 檔案"

    # 解析每個檔案的統計期間
    file_info = []
    for fp in found_files:
        try:
            df_hdr = pd.read_excel(fp, header=None, nrows=4)
            period_str = str(df_hdr.iloc[2, 1]) if df_hdr.shape[0] > 2 and df_hdr.shape[1] > 1 else ""
            m = re.search(r'本年度\d{3}(\d{2})(\d{2})至\d{3}(\d{2})(\d{2})', period_str)
            if m:
                s_m, s_d, e_m, e_d = m.groups()
                is_single_day = (s_m == e_m and s_d == e_d)
                file_info.append({
                    "path": fp,
                    "period_str": period_str,
                    "is_single_day": is_single_day,
                    "end_day": f"{e_m}/{e_d}",
                    "start_day": f"{s_m}/{s_d}",
                    "mtime": os.path.getmtime(fp)
                })
        except Exception:
            pass

    if not file_info:
        return None, None, "無法解析檔案內的統計期間文字"

    # 分辨累計表與單日本期表
    single_files = [x for x in file_info if x["is_single_day"]]
    cum_files = [x for x in file_info if not x["is_single_day"]]

    if not single_files and len(file_info) >= 2:
        # 依照修改時間或檔名備援判定
        file_info.sort(key=lambda x: x["mtime"], reverse=True)
        cur_info = file_info[0]
        cum_info = file_info[1]
    else:
        cur_info = max(single_files, key=lambda x: x["mtime"]) if single_files else file_info[0]
        cum_info = max(cum_files, key=lambda x: x["mtime"]) if cum_files else file_info[-1]

    # 解析表格內容
    def parse_sheet(file_path):
        df = pd.read_excel(file_path, header=None)
        data = {}
        for r in range(5, len(df)):
            unit = str(df.iloc[r, 0]).strip()
            if not unit or unit == 'nan':
                continue
            # 闖紅燈(攔3,逕4), 逆向(攔7,逕8), 不讓行人(攔13,逕14)
            red = (df.iloc[r, 3] or 0) + (df.iloc[r, 4] or 0)
            rev = (df.iloc[r, 7] or 0) + (df.iloc[r, 8] or 0)
            ped = (df.iloc[r, 13] or 0) + (df.iloc[r, 14] or 0)
            data[unit] = {'red': int(red), 'rev': int(rev), 'ped': int(ped), 'tot': int(red + rev + ped)}
        return data

    d_cum = parse_sheet(cum_info["path"])
    d_cur = parse_sheet(cur_info["path"])

    unit_mapping = [
        ("合計", "合計"),
        ("聖亭所", "聖亭派出所"),
        ("龍潭所", "龍潭派出所"),
        ("中興所", "中興派出所"),
        ("石門所", "石門派出所"),
        ("高平所", "高平派出所"),
        ("三和所", "三和派出所"),
        ("交通分隊", "龍潭交通分隊")
    ]

    matrix = []
    for display_name, raw_name in unit_mapping:
        cur = d_cur.get(raw_name, {'red': 0, 'rev': 0, 'ped': 0, 'tot': 0})
        cum = d_cum.get(raw_name, {'red': 0, 'rev': 0, 'ped': 0, 'tot': 0})
        matrix.append([
            display_name,
            cur['red'], cur['rev'], cur['ped'], cur['tot'],
            cum['red'], cum['rev'], cum['ped'], cum['tot']
        ])

    info_msg = (
        f"🎯 自動載入成功！\n"
        f"・本期來源：{os.path.basename(cur_info['path'])} ({cur_info['end_day']})\n"
        f"・累計來源：{os.path.basename(cum_info['path'])} (9/1~{cum_info['end_day']})"
    )
    return cur_info["end_day"], matrix, info_msg

# ==========================================
# 4. 數據準備層（自動載入或安全備援）
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

# 執行方案 B：自動讀取
auto_day, auto_matrix, auto_msg = scan_and_load_three_major()

if auto_matrix:
    latest_three_day = auto_day
    three_major_raw_matrix = auto_matrix
    st.sidebar.success(auto_msg)
else:
    # 安全備用靜態數據
    latest_three_day = "09/17"
    three_major_raw_matrix = [
        ["合計", 30, 10, 5, 45, 291, 122, 34, 447],
        ["聖亭所", 0, 1, 1, 2, 15, 6, 1, 22],
        ["龍潭所", 3, 2, 1, 6, 17, 2, 3, 22],
        ["中興所", 0, 0, 0, 0, 38, 0, 0, 38],
        ["石門所", 10, 0, 3, 13, 50, 1, 3, 54],
        ["高平所", 14, 6, 0, 20, 45, 7, 0, 52],
        ["三和所", 0, 0, 0, 0, 32, 34, 0, 66],
        ["交通分隊", 0, 1, 0, 1, 49, 71, 24, 144],
    ]
    st.sidebar.info(f"💡 目前使用系統最新內建數據 ({latest_three_day})。")

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

df_a1 = pd.DataFrame([
    {"統計期間": "合計", "本期(0909-0915)": 0, "本年累計(0101-0915)": 1, "去年累計(0101-0915)": 6, "本年與去年同期比較": -5},
    {"統計期間": "聖亭所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 0, "本年與去年同期比較": 0},
    {"統計期間": "龍潭所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 1, "本年與去年同期比較": -1},
    {"統計期間": "中興所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 2, "本年與去年同期比較": -2},
    {"統計期間": "石門所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 2, "本年與去年同期比較": -2},
    {"統計期間": "高平所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 1, "去年累計(0101-0915)": 1, "本年與去年同期比較": 0},
    {"統計期間": "三和所", "本期(0909-0915)": 0, "本年累計(0101-0915)": 0, "去年累計(0101-0915)": 0, "本年與去年同期比較": 0},
])

df_a2 = pd.DataFrame([
    {"統計期間": "合計", "本期(0909-0915)": 25, "前期(0902-0908)": 22, "本年累計(0101-0915)": 1353, "去年累計(0101-0915)": 1532, "本年與去年同期比較": -179, "增減比例": "-11.68%"},
    {"統計期間": "聖亭所", "本期(0909-0915)": 5, "前期(0902-0908)": 4, "本年累計(0101-0915)": 282, "去年累計(0101-0915)": 276, "本年與去年同期比較": 6, "增減比例": "2.17%"},
    {"統計期間": "龍潭所", "本期(0909-0915)": 9, "前期(0902-0908)": 11, "本年累計(0101-0915)": 482, "去年累計(0101-0915)": 641, "本年與去年同期比較": -159, "增減比例": "-24.80%"},
    {"統計期間": "中興所", "本期(0909-0915)": 8, "前期(0902-0908)": 5, "本年累計(0101-0915)": 301, "去年累計(0101-0915)": 313, "本年與去年同期比較": -12, "增減比例": "-3.83%"},
    {"統計期間": "石門所", "本期(0909-0915)": 1, "前期(0902-0908)": 0, "本年累計(0101-0915)": 126, "去年累計(0101-0915)": 129, "本年與去年同期比較": -3, "增減比例": "-2.33%"},
    {"統計期間": "高平所", "本期(0909-0915)": 2, "前期(0902-0908)": 2, "本年累計(0101-0915)": 117, "去年累計(0101-0915)": 124, "本年與去年同期比較": -7, "增減比例": "-5.65%"},
    {"統計期間": "三和所", "本期(0909-0915)": 0, "前期(0902-0908)": 0, "本年累計(0101-0915)": 45, "去年累計(0101-0915)": 49, "本年與去年同期比較": -4, "增減比例": "-8.16%"},
])

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
# 5. 前端自選與預覽區
# ==========================================
st.subheader("🎯 欲輸出的統計表自選控制")

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

col_opt1, col_opt2 = st.columns(2)

with col_opt1:
    st.markdown("##### 🏢 常態會報核心表格")
    chk_cover = st.checkbox("P.1 簡報封面 (高對比海軍藍大器版型)", value=True)
    chk_three = st.checkbox("P.2 取締三項重點違規統計表 (雙層表頭)", value=True)
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

selected_pages = []
if chk_cover: selected_pages.append("封面")
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

st.caption(f"📊 目前共勾選 **{len(selected_pages)}** 個頁面待編譯輸出。")

# ==========================================
# 5.1 檔案命名設定
# ==========================================
st.markdown("---")
st.markdown("#### 📁 簡報檔案命名與信件通知")

default_pptx_name = f"龍潭分局執法數據簡報_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"
custom_file_name = st.text_input(
    "✏️ 自訂簡報存檔名稱（副檔名請保留 .pptx）：",
    value=default_pptx_name,
    help="產出後下載之檔案及郵件附件均會以此命名。"
)

curr_user = st.secrets.get("email", {}).get("user") or st.secrets.get("SMTP_USER", "")
if curr_user:
    st.caption(f"📬 點擊「直出簡報並寄給我」後，系統將自動夾帶 PPTX 附件發送至：`{curr_user}`")
else:
    st.caption("💡 提示：若需自動寄信，請確認 secrets.toml 是否已配置 `[email] user = ...` 與 `password = ...`。")

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
# 6. 執行指定輸出生成（做法 A：純 Python 記憶體編譯 PPTX）
# ==========================================
btn_col1, btn_col2 = st.columns([1.5, 2.5])

with btn_col1:
    btn_generate = st.button(f"🚀 直出 PPTX 簡報檔 ({len(selected_pages)} 頁)", type="primary", use_container_width=True)

with btn_col2:
    chk_auto_email = st.checkbox("產出後自動將 PPTX 附件寄到我的信箱", value=True)

if btn_generate:
    if not selected_pages:
        st.warning("⚠️ 請至少勾選一個統計表頁面！")
    elif not custom_file_name.strip():
        st.warning("⚠️ 簡報檔案名稱不可為空白！")
    else:
        file_save_name = custom_file_name.strip()
        if not file_save_name.lower().endswith(".pptx"):
            file_save_name += ".pptx"

        with st.spinner(f"正在純本地動態編譯已勾選的 {len(selected_pages)} 頁 PPTX 簡報（免雲端等待、零配額衝突、色彩 100% 還原）..."):
            try:
                builder = PptxReportBuilder()

                # 1. 依勾選動態加入頁面
                if chk_cover:
                    builder.add_cover_slide(
                        main_title="桃園市政府警察局龍潭分局\n交通執法成效與事故防制分析報告",
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

                # 重大違規總表：負數與單位紅字
                if chk_major_tot:
                    builder.add_table_slide(
                        slide_title="取締重大交通違規統計表",
                        df=df_major,
                        footnote=major_footnote_exact,
                        is_major_table=True
                    )

                # 專項細表：負數與單位紅字
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
                    builder.add_table_slide(
                        slide_title=f"科技執法成效 ({tech_date_range_str})",
                        df=df_tech_final,
                        custom_width_in=8.0
                    )

                # 2. 產出 BytesIO 串流
                pptx_stream = builder.build_bytes()

                st.session_state["cached_pptx"] = pptx_stream
                st.session_state["cached_filename"] = file_save_name

                # 3. 處理自動寄件
                email_sent_msg = None
                if chk_auto_email:
                    ok, detail = send_pptx_email_to_self(pptx_stream, file_save_name)
                    if ok:
                        email_sent_msg = f"📧 簡報附件已成功發送至您的信箱：`{detail}`"
                    else:
                        email_sent_msg = f"⚠️ 郵件未發送成功（{detail}），您依然可以點擊下方按鈕直接下載簡報！"

                st.balloons()
                st.success(f"🎉 恭喜！共 {len(selected_pages)} 頁的 PowerPoint 簡報實體檔案已成功生成！")

                if email_sent_msg:
                    if "📧" in email_sent_msg:
                        st.info(email_sent_msg)
                    else:
                        st.warning(email_sent_msg)

            except Exception as e:
                st.error(f"❌ 產出 PPTX 簡報時發生錯誤：{str(e)}")

# ==========================================
# 7. 下載專用按鈕區
# ==========================================
if "cached_pptx" in st.session_state:
    st.markdown("---")
    st.subheader("📥 簡報下載與轉存")
    c_dl, c_mail = st.columns([2, 2])

    with c_dl:
        st.download_button(
            label=f"💾 點此立即下載【{st.session_state['cached_filename']}】",
            data=st.session_state["cached_pptx"].getvalue(),
            file_name=st.session_state["cached_filename"],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
            use_container_width=True
        )

    with c_mail:
        if st.button("📧 再次補寄這份 PPTX 到我的信箱", use_container_width=True):
            with st.spinner("重新寄送中..."):
                ok, detail = send_pptx_email_to_self(
                    st.session_state["cached_pptx"],
                    st.session_state["cached_filename"]
                )
                if ok:
                    st.success(f"✅ 已成功再次補寄至：`{detail}`")
                else:
                    st.error(f"❌ 補寄失敗：{detail}")
