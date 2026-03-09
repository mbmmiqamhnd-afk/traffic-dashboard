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

# --- 1. 頁面設定 ---
st.set_page_config(page_title="行人及護老交通安全", layout="wide", page_icon="🚶")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# 預設資料 (略過，與之前相同)
DEFAULT_MONTH = "115年3月份"
DEFAULT_CMD = pd.DataFrame([{"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"}, {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"}, {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"}, {"職稱": "上級督導官", "代號": "駐區督察", "姓名": "孫三陽", "任務": "重點機動督導。"}, {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長 黃長旗、督察組督察員 黃中彥、督察組警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"}, {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"}, {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長 楊孟竟、交通組警務員 盧冠仁、交通組警務員 李峯甫、交通組巡官 郭勝隆、交通組巡官 羅千金、交通組警員 吳享運、秘書室巡官 陳鵬翔（代理人：警員張庭溱）、人事室警員 陳明祥、行政組警務佐 曾威仁", "任務": "負責規劃本勤務、重點機動督導、轄區巡守及回報警察局本日執行績效。"}, {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"}])
DEFAULT_SCHEDULE = pd.DataFrame([{"日期（6時至10時、16時至20時）": "3月2～6、9～13、16～20日、23～27及30～31日（3月之上班日）", "單位": "聖亭派出所", "路段": "中豐路、聖亭路段\n校園周邊道路或轄區行人易肇事路口"}, {"日期（6時至10時、16時至20時）": "", "單位": "龍潭派出所", "路段": "中豐路、中正路段\n校園周邊道路或轄區行人易肇事路口"}, {"日期（6時至10時、16時至20時）": "", "單位": "中興派出所", "路段": "中興路、福龍路段\n校園周邊道路或轄區行人易肇事路口"}, {"日期（6時至10時、16時至20時）": "", "單位": "石門派出所", "路段": "中正、文化路段\n校園周邊道路"}, {"日期（6時至10時、16時至20時）": "", "單位": "高平派出所", "路段": "中豐、中原路段\n校園周邊道路"}, {"日期（6時至10時、16時至20時）": "", "單位": "三和派出所", "路段": "龍新路、楊銅路段\n校園周邊道路"}, {"日期（6時至10時、16時至20時）": "", "單位": "警備隊", "路段": "校園周邊道路"}, {"日期（6時至10時、16時至20時）": "", "單位": "龍潭交通分隊", "路段": "校園周邊道路"}])
NOTES = """壹、警察局規劃3月份「行人及護老交通安全專案勤務」期程：...""" # 略

# --- 2. 工具函數 ---
def _get_font():
    fname = "kaiu"
    font_paths = ["kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/kaiu.ttf"]
    for p in font_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

# --- 3. 核心 PDF 生成 (修正銜接問題) ---
def generate_pdf_from_data(month, df_cmd, df_schedule):
    font = _get_font()
    buf = io.BytesIO()
    # A4 寬度 210mm, 邊界各 15mm, 實際可用寬度 180mm
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = 180 * mm 
    story = []
    
    style_t = ParagraphStyle('T', fontName=font, fontSize=16, alignment=1, leading=22, spaceAfter=12)
    style_c = ParagraphStyle('C', fontName=font, fontSize=10, alignment=1, leading=14)
    style_l = ParagraphStyle('L', fontName=font, fontSize=10, alignment=0, leading=14)
    style_h = ParagraphStyle('H', fontName=font, fontSize=13, alignment=1, leading=18)

    # 1. 標題
    story.append(Paragraph(f"<b>{UNIT}{month}執行「行人及護老交通安全」專案勤務規劃表</b>", style_t))

    # 2. 任務編組表格
    data1 = [[Paragraph("<b>任　務　編　組</b>", style_h), "", "", ""],
             [Paragraph("<b>職稱</b>", style_c), Paragraph("<b>代號</b>", style_c), Paragraph("<b>姓名</b>", style_c), Paragraph("<b>任務</b>", style_c)]]
    for _, r in df_cmd.iterrows():
        name_text = str(r['姓名']).replace("、", "<br/>").replace(" ", "<br/>")
        data1.append([Paragraph(f"<b>{r['職稱']}</b>", style_c), Paragraph(str(r['代號']), style_c), Paragraph(name_text, style_c), Paragraph(str(r['任務']), style_l)])
    
    t1 = Table(data1, colWidths=[page_width*0.15, page_width*0.1, page_width*0.25, page_width*0.5])
    t1.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.7, colors.black), # 稍微加深線條確保銜接
        ('SPAN', (0,0), (3,0)), # 標題跨欄
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,1), colors.whitesmoke),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t1)

    # --- 增加明確的一行高度間距 (10mm) ---
    story.append(Spacer(1, 10*mm))

    # 3. 警力佈署表格 (修正銜接與標題)
    data2 = [[Paragraph("<b>警　力　佈　署</b>", style_h), "", ""],
             [Paragraph("<b>日期（6時至10時、16時至20時）</b>", style_c), Paragraph("<b>單位</b>", style_c), Paragraph("<b>路段</b>", style_c)]]
    for _, r in df_schedule.iterrows():
        road_text = str(r.iloc[2]).replace("\n", "<br/>")
        data2.append([Paragraph(str(r.iloc[0]), style_c), Paragraph(str(r.iloc[1]), style_c), Paragraph(road_text, style_l)])
    
    t2 = Table(data2, colWidths=[page_width*0.3, page_width*0.15, page_width*0.55])
    styles2 = [
        ('GRID', (0,0), (-1,-1), 0.7, colors.black),
        ('SPAN', (0,0), (2,0)), # 標題跨欄
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,1), colors.whitesmoke),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]
    
    # 日期欄位自動合併
    non_empty = [i for i, v in enumerate(df_schedule.iloc[:,0]) if str(v).strip() != ""] + [len(df_schedule)]
    for k in range(len(non_empty)-1):
        if non_empty[k+1] - non_empty[k] > 1:
            styles2.append(('SPAN', (0, non_empty[k]+2), (0, non_empty[k+1]+1)))
            
    t2.setStyle(TableStyle(styles2))
    story.append(t2)

    # 4. 備註 (自動換頁處理)
    story.append(Spacer(1, 8*mm))
    story.append(Paragraph("<b>備註：</b>", style_l))
    story.append(Paragraph(NOTES.replace("\n", "<br/>"), style_l))

    doc.build(story)
    return buf.getvalue()

# --- 其餘 Streamlit 邏輯 (略) ---
# ... (請沿用您目前的 Google Sheets 讀取與 UI 部分，僅替換 generate_pdf_from_data 函數即可)
