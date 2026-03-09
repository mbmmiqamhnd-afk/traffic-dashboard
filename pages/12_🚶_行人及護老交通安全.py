import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib, io, os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# --- 1. 初始化頁面 ---
st.set_page_config(page_title="行人及護老勤務系統", layout="wide", page_icon="🚶")

# --- 2. 核心常數與預設資料 ---
UNIT = "桃園市政府警察局龍潭分局"
DEFAULT_MONTH = "115年3月份"

# 預設任務編組資料
DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "駐區督察", "姓名": "孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫", "任務": "負責規劃本勤務、重點機動督導。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"},
])

# 預設勤務表資料
DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日", "單位": "各派出所", "路段": "校園周邊道路或轄區行人易肇事路口"}
])

# --- 3. 字型載入器 ---
@st.cache_resource
def load_pdf_font():
    """載入標楷體，若失敗則回退至 Helvetica"""
    font_name = "Helvetica"
    # 搜尋環境中可能的字型路徑
    paths = ["kaiu.ttf", "font/kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/kaiu.ttf"]
    for p in paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("標楷體", p))
                return "標楷體"
            except: pass
    return font_name

# --- 4. 核心 PDF 生成函數 (附件與網頁一致的關鍵) ---
def generate_unified_pdf(month, df_cmd, df_sch):
    f_name = load_pdf_font()
    buf = io.BytesIO()
    # 設定 A4 邊界
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_w = 180 * mm # 可用寬度
    elements = []
    
    # 定義樣式
    s_title = ParagraphStyle('T', fontName=f_name, fontSize=16, alignment=1, leading=22, spaceAfter=12)
    s_cell_center = ParagraphStyle('C', fontName=f_name, fontSize=10, alignment=1, leading=14)
    s_cell_left = ParagraphStyle('L', fontName=f_name, fontSize=10, alignment=0, leading=14)
    s_section_head = ParagraphStyle('H', fontName=f_name, fontSize=13, alignment=1, leading=18)

    # 1. 報表主標題
    elements.append(Paragraph(f"<b>{UNIT}{month}執行「行人及護老交通安全」專案勤務規劃表</b>", s_title))

    # 2. 第一部分：任務編組
    data1 = [[Paragraph("<b>任　務　編　組</b>", s_section_head), "", "", ""],
             [Paragraph("<b>職稱</b>", s_cell_center), Paragraph("<b>代號</b>", s_cell_center), 
              Paragraph("<b>姓名</b>", s_cell_center), Paragraph("<b>任務</b>", s_cell_center)]]
    
    for _, r in df_cmd.iterrows():
        name_fmt = str(r.get('姓名','')).replace("、","<br/>").replace(" ","<br/>")
        data1.append([
            Paragraph(f"<b>{r.get('職稱','')}</b>", s_cell_center),
            Paragraph(str(r.get('代號','')), s_cell_center),
            Paragraph(name_fmt, s_cell_center),
            Paragraph(str(r.get('任務','')), s_cell_left)
        ])
    
    t1 = Table(data1, colWidths=[page_w*0.15, page_w*0.1, page_w*0.25, page_w*0.5])
    t1.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.7, colors.black), # 線條厚度統一
        ('SPAN', (0,0), (3,0)), # 標題跨欄
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,1), colors.whitesmoke),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(t1)

    # 3. 關鍵間距：空出一行高度 (10mm)
    elements.append(Spacer(1, 10*mm))

    # 4. 第二部分：警力佈署
    data2 = [[Paragraph("<b>警　力　佈　署</b>", s_section_head), "", ""],
             [Paragraph("<b>日期（時段）</b>", s_cell_center), Paragraph("<b>單位</b>", s_cell_center), Paragraph("<b>路段</b>", s_cell_center)]]
    
    for _, r in df_sch.iterrows():
        road_fmt = str(r.iloc[2]).replace("\n", "<br/>")
        data2.append([
            Paragraph(str(r.iloc[0]), s_cell_center),
            Paragraph(str(r.iloc[1]), s_cell_center),
            Paragraph(road_fmt, s_cell_left)
        ])
    
    t2 = Table(data2, colWidths=[page_w*0.3, page_w*0.2, page_w*0.5])
    
    # 警力佈署表格樣式 (包含自動日期合併邏輯)
    style2 = [
        ('GRID', (0,0), (-1,-1), 0.7, colors.black),
        ('SPAN', (0,0), (2,0)),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,1), colors.whitesmoke),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]
    
    # 自動合併相同日期的儲存格 (PDF SPAN)
    idx_list = [i for i, v in enumerate(df_sch.iloc[:, 0]) if str(v).strip() != ""]
    idx_list.append(len(df_sch))
    for i in range(len(idx_list)-1):
        if idx_list[i+1] - idx_list[i] > 1:
            style2.append(('SPAN', (0, idx_list[i]+2), (0, idx_list[i+1]+1)))

    t2.setStyle(TableStyle(style2))
    elements.append(t2)

    doc.build(elements)
    return buf.getvalue()

# --- 5. 主程式介面 ---
st.title("🚶 行人及護老交通安全勤務規劃系統")

month_input = st.text_input("報表月份名稱", DEFAULT_MONTH)

col_editor_l, col_editor_r = st.columns(2)
with col_editor_l:
    st.subheader("📝 1. 任務編組編輯")
    edit_cmd = st.data_editor(DEFAULT_CMD, num_rows="dynamic", use_container_width=True, key="cmd_ed")

with col_editor_r:
    st.subheader("🚔 2. 警力佈署編輯")
    edit_sch = st.data_editor(DEFAULT_SCHEDULE, num_rows="dynamic", use_container_width=True, key="sch_ed")

st.divider()

# --- 6. 功能執行區 (下載與寄信) ---
btn_col1, btn_col2, btn_col3 = st.columns(3)

try:
    # 預先生成 PDF 二進位資料，確保附件與網頁內容絕對同步
    pdf_bytes = generate_unified_pdf(month_input, edit_cmd, edit_sch)
    
    with btn_col1:
        st.download_button(
            label="📥 下載 PDF 報表檔案",
            data=pdf_bytes,
            file_name=f"Traffic_Plan_{datetime.now().strftime('%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary"
        )

    with btn_col2:
        if st.button("📧 發送郵件附件 (與預覽一致)", use_container_width=True):
            # 這裡填入您的 smtplib 寄信代碼，並將 pdf_bytes 附加進去
            st.success("郵件發送指令已發出！")

    with btn_col3:
        if st.button("☁️ 同步至雲端試算表", use_container_width=True):
            st.info("雲端資料同步中...")

except Exception as e:
    st.error(f"❌ 系統生成 PDF 時發生錯誤: {e}")
