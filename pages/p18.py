import streamlit as st
import pandas as pd
import io
import sys
import os
import re
import smtplib
import json
import numpy as np
import urllib.parse as _ul
from collections import Counter
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import pdfplumber
from pdf2image import convert_from_bytes

# --- 導入 Gemini API（保留，未來額度恢復可用）---
try:
    import google.generativeai as genai
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
try:
    from app import show_sidebar
except ImportError:
    def show_sidebar():
        pass

# ==========================================
# 1. 自動寄信功能
# ==========================================
def send_report_email_auto(files, year, month):
    try:
        if "email" not in st.secrets:
            return False, "找不到 st.secrets 中的 email 設定"
        sender = st.secrets["email"]["user"]
        pwd    = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg['From']    = sender
        msg['To']      = sender
        msg['Subject'] = f"【系統備份】龍潭分局 {year}年{month}月 獎勵金點數統計表暨印領清冊"
        body = (
            f"郭同仁您好：\n\n"
            f"系統已自動完成 {year}年{month}月份的獎勵金點數彙整與印領清冊產出。\n"
            f"本次附件包含「點數統計表」與「印領清冊」共兩份 Excel 檔案，請查收。"
        )
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        for file_data, filename in files:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file_data)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f"attachment; filename*=UTF-8''{_ul.quote(filename)}")
            msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)


# ==========================================
# 2. PDF 分配表解析（多策略）
# ==========================================
def _init_vision_client():
    """初始化 Google Cloud Vision 客戶端，失敗回傳 (None, 錯誤訊息)。"""
    try:
        from google.cloud import vision as gvision
        from google.oauth2 import service_account
        if "gcp_service_account" not in st.secrets:
            return None, "找不到 gcp_service_account 設定，請檢查 Streamlit secrets"
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        credentials = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/cloud-platform"]
        )
        client = gvision.ImageAnnotatorClient(credentials=credentials)
        return client, None
    except ImportError:
        return None, "尚未安裝 google-cloud-vision，請在 requirements.txt 加入"
    except Exception as e:
        return None, str(e)


def _extract_longtan_from_text(ocr_text):
    """
    從文字中找出龍潭分局那行的 (point_val, direct, coworker)。
    找不到回傳 (None, None, None)。
    """
    lines = ocr_text.split('\n')
    for line_idx, line in enumerate(lines):
        if '龍潭' not in line:
            continue
        combined = line
        if line_idx + 1 < len(lines):
            combined += " " + lines[line_idx + 1]
        clean = combined.replace(',', '').replace('"', '')
        nums  = re.findall(r'\d+\.\d+|\d+', clean)
        for i, num in enumerate(nums):
            if '.' not in num:
                continue
            pv = float(num)
            if not (0.5 < pv < 10):
                continue
            after = [n for n in nums[i+1:] if '.' not in n]
            if len(after) >= 2:
                return pv, int(after[0]), int(after[1])
    return None, None, None


def parse_alloc_pdf(pdf_bytes):
    """
    三層策略解析分配表 PDF。
    回傳 (point_val, direct, coworker, method_msg)，
    任一數值無法取得時對應欄位為 None。
    """
    point_val  = None
    direct     = None
    coworker   = None
    method_msg = None

    # ── 策略一：pdfplumber 文字萃取 ──────────────────────────────
    try:
        pdf_text = ""
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text()
                if extracted:
                    pdf_text += extracted + "\n"

        if pdf_text.strip():
            pv, d, cw = _extract_longtan_from_text(pdf_text)
            if pv is not None:
                point_val, direct, coworker = pv, d, cw
                method_msg = "pdfplumber 文字萃取"
    except Exception:
        pass

    if point_val is not None:
        return point_val, direct, coworker, method_msg

    # ── 策略二：Google Cloud Vision OCR ──────────────────────────
    vision_client, vision_err = _init_vision_client()
    if vision_client is None:
        st.warning(f"⚠️ Google Cloud Vision 無法使用：{vision_err}，請手動輸入。")
        return None, None, None, None

    try:
        images = convert_from_bytes(pdf_bytes, dpi=200)
    except Exception as e:
        st.warning(f"⚠️ PDF 轉圖片失敗：{e}")
        return None, None, None, None

    with st.spinner("🔍 Google Cloud Vision OCR 辨識中..."):
        for page_num, page_img in enumerate(images, start=1):
            try:
                from google.cloud import vision as gvision
                img_buf = io.BytesIO()
                page_img.save(img_buf, format="PNG")
                vision_image = gvision.Image(content=img_buf.getvalue())
                response = vision_client.text_detection(image=vision_image)

                if response.error.message:
                    st.warning(f"⚠️ Vision API 第 {page_num} 頁錯誤：{response.error.message}")
                    continue

                ocr_text = response.full_text_annotation.text
                if not ocr_text or '龍潭' not in ocr_text:
                    continue

                pv, d, cw = _extract_longtan_from_text(ocr_text)
                if pv is not None:
                    point_val, direct, coworker = pv, d, cw
                    method_msg = f"Google Cloud Vision OCR（第 {page_num} 頁）"
                    break

            except Exception as page_e:
                st.warning(f"⚠️ 第 {page_num} 頁 OCR 失敗：{page_e}")
                continue

    return point_val, direct, coworker, method_msg


# ==========================================
# 3. 主頁面邏輯
# ==========================================
def p18_page():
    show_sidebar()

    st.title("💰 龍潭分局 - 獎勵金點數統計表暨印領清冊產生器")
    st.info("權重已固定 (A2:10, A3:5, 交整:5)。系統支援智慧解析分配表、自動裁剪、對帳防呆，並發送雙報表郵件。")

    P_A2, P_A3, P_TRAF = 10.0, 5.0, 5.0

    # ── 檔案上傳區 ────────────────────────────────────────────────
    st.subheader("📂 1. 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template  = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】", type=['xlsx'])
    file_acc       = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    file_traf_list = st.file_uploader(
        "3. 上傳當月【各單位_交通疏導統計】(可多選)",
        type=['xlsx'], accept_multiple_files=True
    )

    # ── 分配表上傳與解析 ──────────────────────────────────────────
    st.subheader("📝 2. 印領清冊設定與官方分配額度")
    file_alloc = st.file_uploader(
        "📥 (選用) 上傳【獎勵金分配表】(PDF) 自動抓取點值與預算",
        type=['pdf']
    )

    auto_point_val    = None
    official_direct   = None
    official_coworker = None

    if file_alloc is not None:
        pdf_bytes = file_alloc.read()

        # debug：顯示 pdfplumber 原始萃取結果
        try:
            _dbg = ""
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as _pdf:
                for _p in _pdf.pages:
                    _t = _p.extract_text()
                    if _t:
                        _dbg += _t + "\n"
            with st.expander("🔍 pdfplumber 原始萃取文字（debug）", expanded=False):
                st.text(_dbg if _dbg.strip() else "（無法萃取文字，為掃描圖片，將使用 Vision OCR）")
        except Exception:
            pass

        with st.spinner("📖 解析分配表中..."):
            auto_point_val, official_direct, official_coworker, method_msg = \
                parse_alloc_pdf(pdf_bytes)

        if auto_point_val is not None:
            st.success(f"✅ 系統已成功解析分配表！（方式：{method_msg}）")
            cc1, cc2, cc3 = st.columns(3)
            cc1.info(f"💵 **每點獎金:**\n### {auto_point_val}")
            cc2.info(f"👮‍♂️ **直接執行人員獎金:**\n### $ {official_direct:,}"
                     if official_direct else "👮‍♂️ **直接執行人員獎金:**\n### （未取得）")
            cc3.info(f"🤝 **共同作業人員獎金:**\n### $ {official_coworker:,}"
                     if official_coworker else "🤝 **共同作業人員獎金:**\n### （未取得）")
        else:
            st.warning("⚠️ 自動解析失敗，請使用下方手動輸入欄位。")

    # ── 手動輸入 / 覆蓋欄位 ──────────────────────────────────────
    auto_ok = (auto_point_val is not None
               and official_direct is not None
               and official_coworker is not None)

    with st.expander("✏️ 手動輸入 / 覆蓋官方核定額度", expanded=(not auto_ok)):
        mc1, mc2, mc3 = st.columns(3)
        manual_point_val = mc1.number_input(
            "每點獎金",
            value=float(auto_point_val) if auto_point_val is not None else 1.375,
            format="%.3f", step=0.001
        )
        manual_direct = mc2.number_input(
            "直接執行人員獎金",
            value=int(official_direct) if official_direct is not None else 0,
            step=1, min_value=0
        )
        manual_coworker = mc3.number_input(
            "共同作業及配合人員獎金",
            value=int(official_coworker) if official_coworker is not None else 0,
            step=1, min_value=0
        )

    # 手動值永遠覆蓋，防止 AI/OCR 辨識錯誤無法修正
    point_value       = manual_point_val
    official_direct   = manual_direct
    official_coworker = manual_coworker

    # ── 共同作業人員名單 ──────────────────────────────────────────
    st.markdown("##### 👥 共同作業及配合人員名單 (內建完整預設值)")
    st.caption("💡 提示：本表已載入全分局預設名單。可直接在表格內修改異動金額，或在最下方點擊「+」新增人員，勾選左側核取方塊按 Delete 可刪除。")

    default_coworkers_data = [
        {"單位": "龍潭分局",     "職別": "分局長",         "姓名": "施宇峰", "金額": 301, "蓋章": ""},
        {"單位": "龍潭分局",     "職別": "副分局長",       "姓名": "何憶雯", "金額": 100, "蓋章": ""},
        {"單位": "龍潭分局",     "職別": "副分局長",       "姓名": "蔡志明", "金額": 100, "蓋章": ""},
        {"單位": "交通組",       "職別": "業務單位主管",   "姓名": "陳維明", "金額": 298, "蓋章": ""},
        {"單位": "交通組",       "職別": "交通業務承辦人", "姓名": "盧冠仁", "金額": 298, "蓋章": ""},
        {"單位": "交通組",       "職別": "交通業務承辦人", "姓名": "李峯甫", "金額": 298, "蓋章": ""},
        {"單位": "交通組",       "職別": "交通業務承辦人", "姓名": "羅千金", "金額": 298, "蓋章": ""},
        {"單位": "交通組",       "職別": "交通業務承辦人", "姓名": "郭勝隆", "金額": 298, "蓋章": ""},
        {"單位": "交通組",       "職別": "交通業務承辦人", "姓名": "吳享運", "金額": 232, "蓋章": ""},
        {"單位": "交通組",       "職別": "交通業務承辦人", "姓名": "吳沛軒", "金額": 232, "蓋章": ""},
        {"單位": "會計室",       "職別": "主計",           "姓名": "郭貞彣", "金額": 77,  "蓋章": ""},
        {"單位": "會計室",       "職別": "主計",           "姓名": "林玲宜", "金額": 78,  "蓋章": ""},
        {"單位": "秘書室",       "職別": "主任",           "姓名": "陳振貴", "金額": 78,  "蓋章": ""},
        {"單位": "秘書室",       "職別": "出納",           "姓名": "簡啟峯", "金額": 78,  "蓋章": ""},
        {"單位": "人事室",       "職別": "主任",           "姓名": "葉菀容", "金額": 78,  "蓋章": ""},
        {"單位": "人事室",       "職別": "助理員",         "姓名": "王韋翔", "金額": 77,  "蓋章": ""},
        {"單位": "人事室",       "職別": "警務佐",         "姓名": "李福源", "金額": 77,  "蓋章": ""},
        {"單位": "人事室",       "職別": "警員",           "姓名": "陳明祥", "金額": 77,  "蓋章": ""},
        {"單位": "人事室",       "職別": "警員",           "姓名": "黃秀吉", "金額": 77,  "蓋章": ""},
        {"單位": "聖亭派出所",   "職別": "所長",           "姓名": "鄭榮捷", "金額": 195, "蓋章": ""},
        {"單位": "聖亭派出所",   "職別": "副所長",         "姓名": "邱品淳", "金額": 195, "蓋章": ""},
        {"單位": "聖亭派出所",   "職別": "副所長",         "姓名": "曹培翔", "金額": 195, "蓋章": ""},
        {"單位": "聖亭派出所",   "職別": "業務承辦人",     "姓名": "曾建凱", "金額": 90,  "蓋章": ""},
        {"單位": "龍潭派出所",   "職別": "所長",           "姓名": "孫祥愷", "金額": 195, "蓋章": ""},
        {"單位": "龍潭派出所",   "職別": "副所長",         "姓名": "劉重言", "金額": 195, "蓋章": ""},
        {"單位": "龍潭派出所",   "職別": "副所長",         "姓名": "梁順安", "金額": 195, "蓋章": ""},
        {"單位": "龍潭派出所",   "職別": "業務承辦人",     "姓名": "周薇",   "金額": 90,  "蓋章": ""},
        {"單位": "中興派出所",   "職別": "所長",           "姓名": "董亦文", "金額": 195, "蓋章": ""},
        {"單位": "中興派出所",   "職別": "副所長",         "姓名": "何昀融", "金額": 195, "蓋章": ""},
        {"單位": "中興派出所",   "職別": "副所長",         "姓名": "林榮裕", "金額": 195, "蓋章": ""},
        {"單位": "中興派出所",   "職別": "業務承辦人",     "姓名": "鄧雅文", "金額": 90,  "蓋章": ""},
        {"單位": "石門派出所",   "職別": "所長",           "姓名": "林育辰", "金額": 195, "蓋章": ""},
        {"單位": "石門派出所",   "職別": "副所長",         "姓名": "薛德祥", "金額": 195, "蓋章": ""},
        {"單位": "石門派出所",   "職別": "業務承辦人",     "姓名": "陳琦",   "金額": 89,  "蓋章": ""},
        {"單位": "高平派出所",   "職別": "所長",           "姓名": "王梓岳", "金額": 195, "蓋章": ""},
        {"單位": "高平派出所",   "職別": "副所長",         "姓名": "楊勝吉", "金額": 195, "蓋章": ""},
        {"單位": "高平派出所",   "職別": "業務承辦人",     "姓名": "黃丞潁", "金額": 89,  "蓋章": ""},
        {"單位": "三和派出所",   "職別": "所長",           "姓名": "宋開國", "金額": 194, "蓋章": ""},
        {"單位": "三和派出所",   "職別": "副所長",         "姓名": "陳佶汎", "金額": 194, "蓋章": ""},
        {"單位": "三和派出所",   "職別": "業務承辦人",     "姓名": "童霂晟", "金額": 89,  "蓋章": ""},
        {"單位": "龍潭交通分隊", "職別": "分隊長",         "姓名": "卓宜澂", "金額": 195, "蓋章": ""},
        {"單位": "龍潭交通分隊", "職別": "小隊長",         "姓名": "鄭敬思", "金額": 195, "蓋章": ""},
        {"單位": "龍潭交通分隊", "職別": "小隊長",         "姓名": "蔡安龍", "金額": 195, "蓋章": ""},
        {"單位": "龍潭交通分隊", "職別": "業務承辦人",     "姓名": "陳建穎", "金額": 89,  "蓋章": ""},
        {"單位": "勤務中心",     "職別": "主任",           "姓名": "游新枝", "金額": 65,  "蓋章": ""},
        {"單位": "勤務中心",     "職別": "巡佐",           "姓名": "李文章", "金額": 65,  "蓋章": ""},
        {"單位": "勤務中心",     "職別": "巡佐",           "姓名": "余清富", "金額": 65,  "蓋章": ""},
        {"單位": "勤務中心",     "職別": "警務佐",         "姓名": "陳敬霖", "金額": 65,  "蓋章": ""},
        {"單位": "勤務中心",     "職別": "警員",           "姓名": "黃文興", "金額": 65,  "蓋章": ""},
        {"單位": "勤務中心",     "職別": "警員",           "姓名": "王天龍", "金額": 65,  "蓋章": ""},
        {"單位": "勤務中心",     "職別": "警員",           "姓名": "曾嘉偉", "金額": 65,  "蓋章": ""},
        {"單位": "勤務中心",     "職別": "警員",           "姓名": "江文頌", "金額": 64,  "蓋章": ""},
        {"單位": "秘書室",       "職別": "巡官",           "姓名": "陳鵬翔", "金額": 64,  "蓋章": ""},
        {"單位": "督察組",       "職別": "組長",           "姓名": "賴永益", "金額": 64,  "蓋章": ""},
        {"單位": "督察組",       "職別": "督察員",         "姓名": "黃中彥", "金額": 64,  "蓋章": ""},
        {"單位": "督察組",       "職別": "警務員",         "姓名": "陳冠彰", "金額": 64,  "蓋章": ""},
        {"單位": "督察組",       "職別": "巡官",           "姓名": "全楚文", "金額": 64,  "蓋章": ""},
        {"單位": "保安民防組",   "職別": "組長",           "姓名": "蔡奇青", "金額": 64,  "蓋章": ""},
        {"單位": "保安民防組",   "職別": "警務員",         "姓名": "曾盛鉉", "金額": 64,  "蓋章": ""},
        {"單位": "保安民防組",   "職別": "巡官",           "姓名": "李立人", "金額": 64,  "蓋章": ""},
        {"單位": "保安民防組",   "職別": "巡官",           "姓名": "林沛達", "金額": 64,  "蓋章": ""},
        {"單位": "保安民防組",   "職別": "巡官",           "姓名": "吳國棟", "金額": 64,  "蓋章": ""},
        {"單位": "行政組",       "職別": "組長",           "姓名": "周金柱", "金額": 64,  "蓋章": ""},
        {"單位": "行政組",       "職別": "巡官",           "姓名": "蕭凱文", "金額": 64,  "蓋章": ""},
        {"單位": "防治組",       "職別": "組長",           "姓名": "沈鳳漳", "金額": 64,  "蓋章": ""},
        {"單位": "防治組",       "職別": "巡官",           "姓名": "陳冠亘", "金額": 64,  "蓋章": ""},
    ]
    df_coworkers_default = pd.DataFrame(default_coworkers_data)

    edited_df_coworkers = st.data_editor(
        df_coworkers_default,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        height=400,
        column_config={
            "金額": st.column_config.NumberColumn("金額", min_value=0, step=1, format="%d")
        }
    )

    # ── 執行按鈕 ──────────────────────────────────────────────────
    if st.button("🚀 執行彙整、計算獎金與自動寄信", type="primary", use_container_width=True):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種點數統計資料皆已完成上傳！")
            return

        with st.spinner("正在計算數據、產出報表並將雙檔案發送至信箱..."):
            try:
                # ── 讀取事故資料 ──
                df_acc_raw = pd.read_excel(file_acc, header=4)
                df_acc_raw['姓名'] = df_acc_raw['姓名'].astype(str).str.strip()
                dict_acc = df_acc_raw.groupby('姓名')[['A2類', 'A3類']].sum().to_dict(orient='index')

                # ── 讀取交整資料 ──
                df_traf_all = pd.concat([pd.read_excel(f, sheet_name='月彙整總表') for f in file_traf_list])
                df_traf_all['姓名'] = df_traf_all['姓名'].astype(str).str.strip()
                dict_traf = df_traf_all.groupby('姓名')['總計尖峰時數'].sum().to_dict()

                # ── 讀取點數統計表，自動抓年月 ──
                dfs_raw = pd.read_excel(file_template, sheet_name=None, header=None)
                ext_year, ext_month = "115", "4"
                found_date = False
                for _, df_scan in dfs_raw.items():
                    for r in range(min(15, len(df_scan))):
                        for c in range(min(10, len(df_scan.columns))):
                            v = str(df_scan.iloc[r, c])
                            m = re.search(r'開單日期[：:\s]*(\d{3})(\d{2})', v)
                            if m:
                                ext_year, ext_month = m.group(1), str(int(m.group(2)))
                                found_date = True
                                break
                        if found_date:
                            break
                    if found_date:
                        break

                # ── 逐工作表計算點數 ──
                final_sheets     = {}
                summary_rows     = []
                g_cite, g_acc, g_traf, g_all = 0, 0, 0, 0
                direct_exec_list = []

                for sheet_name, df in dfs_raw.items():
                    if '總表' in sheet_name:
                        continue

                    start_r, start_c = None, None
                    for r_idx, row in df.iterrows():
                        row_str = [str(x).strip() for x in row.values]
                        if '員警姓名' in row_str:
                            start_r = r_idx
                            start_c = row_str.index('員警姓名')
                            break

                    if start_r is None:
                        continue

                    df_work = df.iloc[start_r:, start_c:].copy()
                    df_work.reset_index(drop=True, inplace=True)
                    df_work.columns = [str(c).strip() for c in df_work.iloc[0]]
                    df_work = df_work.drop(0).astype(object)
                    col_map = {c: i for i, c in enumerate(df_work.columns)}

                    member_rows = []
                    for r in range(len(df_work)):
                        name_cell = str(df_work.iloc[r, col_map['員警姓名']]).strip()
                        if '小計' in name_cell or '總計' in name_cell or name_cell in ['nan', 'None', '']:
                            continue
                        member_rows.append(r)

                    df_members = df_work.iloc[member_rows].copy()
                    s_cite, s_acc, s_traf = 0, 0, 0

                    for idx, row in df_members.iterrows():
                        name = str(row['員警姓名']).strip()
                        a2   = dict_acc.get(name, {}).get('A2類', 0)
                        a3   = dict_acc.get(name, {}).get('A3類', 0)
                        th   = dict_traf.get(name, 0)
                        ap   = a2 * P_A2 + a3 * P_A3
                        tp   = th * P_TRAF

                        cp = pd.to_numeric(row.get('取締點數', 0), errors='coerce')
                        cp = cp if pd.notna(cp) else 0
                        total_pts = cp + ap + tp

                        if 'A2件數'     in col_map: df_members.at[idx, 'A2件數']     = a2 if a2 > 0 else ""
                        if 'A3件數'     in col_map: df_members.at[idx, 'A3件數']     = a3 if a3 > 0 else ""
                        if '事故點數'   in col_map: df_members.at[idx, '事故點數']   = ap if ap > 0 else ""
                        if '交整時數'   in col_map: df_members.at[idx, '交整時數']   = th if th > 0 else ""
                        if '交整點數'   in col_map: df_members.at[idx, '交整點數']   = tp if tp > 0 else ""
                        if '個人總點數' in col_map: df_members.at[idx, '個人總點數'] = total_pts

                        s_cite += cp; s_acc += ap; s_traf += tp

                        if total_pts > 0:
                            reward = int(np.round(total_pts * point_value))
                            direct_exec_list.append({
                                "單位名稱":   sheet_name,
                                "員警姓名":   name,
                                "取締件數":   row.get('取締件數', ''),
                                "取締點數":   cp if cp > 0 else '',
                                "A2件數":     a2 if a2 > 0 else '',
                                "A3件數":     a3 if a3 > 0 else '',
                                "事故點數":   ap if ap > 0 else '',
                                "交整時數":   th if th > 0 else '',
                                "交整點數":   tp if tp > 0 else '',
                                "個人總點數": total_pts,
                                "每點獎金":   point_value,
                                "實領獎金":   reward,
                                "蓋章":       "",
                            })

                    sub_row_data = {c: "" for c in df_work.columns}
                    sub_row_data['員警姓名'] = '小計'
                    for col_n in df_work.columns:
                        if col_n in ['員警姓名', '蓋章']:
                            continue
                        v_sum = pd.to_numeric(df_members[col_n], errors='coerce').sum()
                        sub_row_data[col_n] = v_sum if v_sum > 0 else 0

                    df_final = pd.concat([df_members, pd.DataFrame([sub_row_data])], ignore_index=True)
                    if '蓋章' in df_final.columns:
                        df_final = df_final.drop(columns=['蓋章'])
                    final_sheets[sheet_name] = df_final

                    summary_rows.append([sheet_name, s_cite, s_acc, s_traf, s_cite + s_acc + s_traf])
                    g_cite += s_cite; g_acc += s_acc; g_traf += s_traf
                    g_all  += (s_cite + s_acc + s_traf)

                # ── 清冊彙整 ──
                df_direct_exec = pd.DataFrame(direct_exec_list)
                df_direct_exec.insert(0, '序號', range(1, len(df_direct_exec) + 1))
                direct_total_money = df_direct_exec['實領獎金'].sum()

                df_coworkers_final = edited_df_coworkers.copy()
                df_coworkers_final.dropna(how='all', inplace=True)
                df_coworkers_final.insert(0, '序號', range(1, len(df_coworkers_final) + 1))
                coworkers_total_money = pd.to_numeric(
                    df_coworkers_final['金額'], errors='coerce').fillna(0).sum()

                # ── 🚨 對帳警告 🚨 ──
                if official_coworker > 0 and coworkers_total_money != official_coworker:
                    st.warning(
                        f"⚠️ **對帳異常 (共同作業)：** 名單金額總和 ${coworkers_total_money:,}，"
                        f"官方分配表 ${official_coworker:,}，"
                        f"相差 {abs(int(coworkers_total_money) - official_coworker):,} 元！"
                    )
                if official_direct > 0 and direct_total_money != official_direct:
                    st.warning(
                        f"⚠️ **對帳異常 (直接執行)：** 系統計算四捨五入總和 ${direct_total_money:,}，"
                        f"官方分配表 ${official_direct:,}，"
                        f"相差 {abs(direct_total_money - official_direct):,} 元。（通常為進位誤差）"
                    )

                # ── 一覽表 ──
                df_payroll_summary = pd.DataFrame([
                    {"項目": "直接執行人員",       "金額": direct_total_money},
                    {"項目": "共同作業及配合人員", "金額": coworkers_total_money},
                    {"項目": "合計",               "金額": direct_total_money + coworkers_total_money},
                    {"項目": "製表人",             "金額": ""},
                ])

                # ── 產出 Excel ──
                pts_output = io.BytesIO()
                df_pts_summary = pd.DataFrame(
                    [['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']]
                    + summary_rows
                    + [['合計', g_cite, g_acc, g_traf, g_all]]
                )
                with pd.ExcelWriter(pts_output, engine='xlsxwriter') as writer:
                    df_pts_summary.to_excel(writer, sheet_name='總表', header=False, index=False)
                    for sn, df_f in final_sheets.items():
                        df_f.to_excel(writer, sheet_name=sn, index=False)
                pts_excel_data = pts_output.getvalue()
                pts_filename   = f"龍潭分局{ext_year}年{ext_month}月份_點數統計表.xlsx"

                payroll_output = io.BytesIO()
                with pd.ExcelWriter(payroll_output, engine='xlsxwriter') as writer:
                    df_direct_exec.to_excel(writer, sheet_name='直接執行人員', index=False)
                    if not df_coworkers_final.empty:
                        df_coworkers_final.to_excel(writer, sheet_name='共同作業及配合人員', index=False)
                    df_payroll_summary.to_excel(writer, sheet_name='獎勵金支領一覽表', index=False)
                    workbook    = writer.book
                    vcenter_fmt = workbook.add_format({'valign': 'vcenter'})
                    for sn, df_s in [('直接執行人員', df_direct_exec),
                                     ('共同作業及配合人員', df_coworkers_final)]:
                        if sn in writer.sheets:
                            ws = writer.sheets[sn]
                            ws.set_row(0, 25, vcenter_fmt)
                            for rn in range(1, len(df_s) + 1):
                                ws.set_row(rn, 45, vcenter_fmt)
                payroll_excel_data = payroll_output.getvalue()
                payroll_filename   = f"龍潭分局{ext_year}年{ext_month}月份_獎勵金印領清冊.xlsx"

                # ── 寄信 ──
                ok, err = send_report_email_auto(
                    [(pts_excel_data, pts_filename), (payroll_excel_data, payroll_filename)],
                    ext_year, ext_month
                )
                if ok:
                    st.success("✅ 雙報表產出成功！已備份至信箱。")
                else:
                    st.warning(f"⚠️ 報表已產出，但郵件發送失敗: {err}")

                c5, c6 = st.columns(2)
                c5.download_button("📥 下載【點數統計表】",
                                   data=pts_excel_data, file_name=pts_filename,
                                   use_container_width=True)
                c6.download_button("📥 下載【印領清冊】",
                                   data=payroll_excel_data, file_name=payroll_filename,
                                   use_container_width=True, type="primary")

            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")


if __name__ == "__main__":
    p18_page()
