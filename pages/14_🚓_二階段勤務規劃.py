import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib, io, os
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
import re

# --- 1. 頁面設定 ---
st.set_page_config(page_title="二階段勤務規劃系統", layout="wide", page_icon="🚓")

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

DEFAULT_UNIT    = "桃園市政府警察局龍潭分局"
DEFAULT_TIME    = "113年8月1日20至24時"
DEFAULT_PROJ    = "0801全市取締酒後駕車暨防制危險駕車及噪音車輛專案"
DEFAULT_BRIEF   = "於各編組執行前\n地點：現地勤教" 

DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長鄭文銘", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長陳孟欣", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "上級督導官", "代號": "駐區督察", "姓名": "孫三陽", "任務": "重點機動督導。"},
    {"職稱": "督導組", "代號": "隆安6", "姓名": "督察組組長賴永益", "任務": "督導各編組服儀裝備及勤務紀律。"},
    {"職稱": "指導組", "代號": "隆安684", "姓名": "督察組教官郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
    {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "交通組組長蘇志旭", "任務": "規劃本勤務、重點機動督導及回報績效。"},
    {"職稱": "通訊組", "代號": "隆安", "姓名": "主任游新枝", "任務": "指揮、調度及通報本勤務事宜。"},
])

DEFAULT_PTL = pd.DataFrame([
    {"編組": "第一巡邏組", "無線電": "隆安62", "單位": "龍潭所", "服勤人員": "副所長劉重言、警員劉俊德", "任務分工": "於易發生酒駕、危險駕車路段(中豐路等)加強攔檢。"},
    {"編組": "第二巡邏組", "無線電": "隆安52", "單位": "聖亭所", "服勤人員": "副所長曹培翔、警員劉憬霖", "任務分工": "於易發生酒駕、危險駕車路段(聖亭路等)加強攔檢。"},
    {"編組": "第三巡邏組", "無線電": "隆安72", "單位": "中興所", "服勤人員": "副所長何昀融、警員廖佩祺", "任務分工": "於易發生酒駕、危險駕車路段(北龍路等)加強攔檢。"},
    {"編組": "第四巡邏組", "無線電": "隆安83", "單位": "石門所", "服勤人員": "巡佐林偉政、警員陳琦", "任務分工": "於易發生酒駕、危險駕車路段(大昌路二段等)加強攔檢。"},
    {"編組": "第五巡邏組", "無線電": "隆安93", "單位": "高平所", "服勤人員": "警務佐余志誠、警員王天龍", "任務分工": "於易發生酒駕、危險駕車路段(中豐路高平段等)加強攔檢。"},
])

DEFAULT_CHECKPOINT = pd.DataFrame([
    {"編組": "第一路檢組", "無線電": "隆安93", "單位": "高平所、石門所", "服勤人員": "警務佐余志誠、警員王天龍、巡佐林偉政、警員陳琦", "任務分工": "路檢：中正路三坑段與美國路口(攔檢往龍潭市區方向車輛)。"},
    {"編組": "第二路檢組", "無線電": "隆安52", "單位": "聖亭所、三和所", "服勤人員": "副所長曹培翔、警員劉憬霖、警員盧建廷", "任務分工": "路檢：中正路三坑段與美國路口(攔檢往龍源路方向車輛)。"},
    {"編組": "第一機動攔檢組", "無線電": "隆安991", "單位": "交通分隊", "服勤人員": "分隊長卓宜澂、警員陳世傑", "任務分工": "攔截圍捕：於中正路三坑段適當地點，機動攔檢迴轉規避車輛。"},
    {"編組": "第二機動攔檢組", "無線電": "隆安62", "單位": "龍潭所", "服勤人員": "副所長劉重言、警員劉俊德", "任務分工": "路檢：於中正路三坑段適當地點，機動攔檢迴轉規避車輛。"},
    {"編組": "第三機動攔檢組", "無線電": "隆安72", "單位": "中興所", "服勤人員": "副所長何昀融、警員廖佩祺", "任務分工": "攔截圍捕：於美國路文化路來回梭巡，機動攔檢迴轉規避車輛。"}
])

# --- 2. 輔助函數 ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    font_paths = ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/custom/kaiu.ttf"]
    for p in font_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def parse_meeting_time(time_str):
    try:
        match = re.search(r"(\d+)至", time_str)
        if match:
            start_hour = int(match.group(1))
            end_hour = start_hour + 1
            return f"{start_hour}時30分至{end_hour}時00分"
    except:
        pass
    return "19時30分至20時00分"

@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("設定")
        ws_cmd = sh.worksheet("指揮組")
        ws_ptl = sh.worksheet("巡邏組")
        try:
            ws_cp = sh.worksheet("路檢組")
            df_cp = pd.DataFrame(ws_cp.get_all_records())
        except:
            df_cp = None
        return pd.DataFrame(ws_set.get_all_records()), pd.DataFrame(ws_cmd.get_all_records()), pd.DataFrame(ws_ptl.get_all_records()), df_cp, None
    except Exception as e: return None, None, None, None, str(e)

def save_data(unit, time_str, project, briefing, df_cmd, df_ptl, df_cp):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        ws_set = sh.worksheet("設定")
        ws_set.clear()
        ws_set.update([["Key", "Value"], ["unit_name", unit], ["plan_full_time", time_str], ["project_name", project], ["briefing_info", briefing]])
        
        for ws_name, df in [("指揮組", df_cmd), ("巡邏組", df_ptl), ("路檢組", df_cp)]:
            try:
                ws = sh.worksheet(ws_name)
            except:
                ws = sh.add_worksheet(title=ws_name, rows="100", cols="20")
            ws.clear()
            df = df.fillna("")
            ws.update([df.columns.tolist()] + df.values.tolist())
        load_data.clear()
        return True
    except: return False

# --- 3. PDF 生成功能 ---
def generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []
    
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=24, alignment=1, spaceAfter=8)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=10)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=18, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=18, alignment=0)
    
    style_middle_block = ParagraphStyle(
        'MiddleBlock', fontName=font, fontSize=14, leading=22, spaceAfter=2*mm, alignment=TA_LEFT, leftIndent=5*mm, firstLineIndent=0
    )
    style_table_title = ParagraphStyle('TTitle', fontName=font, fontSize=16, alignment=1, leading=22)

    story.append(Paragraph(f"{unit}執行{project}規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    # 清理文字換行用的輔助函數
    def clean(t): return str(t).replace("\n", "<br/>").replace("、", "<br/>")
    def clean_text_only(t): return str(t).replace("\n", "<br/>")

    data_cmd = [[Paragraph("<b>任 務 編 組</b>", style_table_title), '', '', ''],
                [Paragraph(f"<b>{h}</b>", style_cell) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        data_cmd.append([
            Paragraph(f"<b>{clean_text_only(r.get('職稱',''))}</b>", style_cell), 
            Paragraph(clean_text_only(r.get('代號','')), style_cell),
            Paragraph(clean(r.get('姓名','')), style_cell), 
            Paragraph(clean_text_only(r.get('任務','')), style_cell_left)
        ])
    
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.12, page_width*0.28, page_width*0.45])
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('GRID',(0,0),(-1,-1),0.5,colors.black),('SPAN',(0,0),(-1,0)),
                            ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t1)
    
    briefing_clean = clean_text_only(briefing).strip()
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph("<b>📢 勤前教育：</b>", style_middle_block))
    story.append(Paragraph(f"{briefing_clean}", style_middle_block))
    story.append(Spacer(1, 6*mm))

    # 第一階段
    story.append(Paragraph("<b>第一階段：21時至22時30分，機動巡邏</b>", style_middle_block))
    data_ptl = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "代號", "單位", "服勤人員", "任務分工"]]]
    for _, r in df_ptl.iterrows():
        # 將編組及代號等欄位包入 Paragraph 才能自動換行
        data_ptl.append([
            Paragraph(clean_text_only(r.get('編組','')), style_cell), 
            Paragraph(clean_text_only(r.get('無線電','')), style_cell), 
            Paragraph(clean(r.get('單位','')), style_cell), 
            Paragraph(clean(r.get('服勤人員','')), style_cell), 
            Paragraph(clean_text_only(r.get('任務分工','')), style_cell_left)
        ])
    
    t2 = Table(data_ptl, colWidths=[page_width*0.15, page_width*0.12, page_width*0.13, page_width*0.20, page_width*0.40])
    t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('FONTSIZE',(0,0),(-1,-1),14),('GRID',(0,0),(-1,-1),0.5,colors.black),
                            ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t2)

    story.append(Spacer(1, 8*mm))

    # 第二階段
    story.append(Paragraph("<b>第二階段：22時30分至24時，定點路檢及機動攔檢</b>", style_middle_block))
    data_cp = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["編組", "代號", "單位", "服勤人員", "任務分工"]]]
    for _, r in df_cp.iterrows():
        task_text = f"{clean_text_only(r.get('任務分工',''))}<br/><font color='red' size='11'>*雨天備案：轄區治安要點巡邏。</font>"
        data_cp.append([
            Paragraph(clean_text_only(r.get('編組','')), style_cell), 
            Paragraph(clean_text_only(r.get('無線電','')), style_cell), 
            Paragraph(clean(r.get('單位','')), style_cell), 
            Paragraph(clean(r.get('服勤人員','')), style_cell), 
            Paragraph(task_text, style_cell_left)
        ])
    
    t3 = Table(data_cp, colWidths=[page_width*0.15, page_width*0.12, page_width*0.13, page_width*0.20, page_width*0.40])
    t3.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font),('FONTSIZE',(0,0),(-1,-1),14),('GRID',(0,0),(-1,-1),0.5,colors.black),
                            ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e6e6e6')),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(t3)

    doc.build(story)
    return buf.getvalue()

# 簽到表
def generate_attendance_pdf(unit, project, time_str, briefing):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 30*mm
    story = []
    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, leading=22, alignment=1, spaceAfter=8)
    style_top_info = ParagraphStyle('TopInfo', fontName=font, fontSize=12, leading=18, alignment=0)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=14, leading=24, alignment=1)
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=14, leading=24, alignment=0) 
    style_note = ParagraphStyle('Note', fontName=font, fontSize=11, leading=15, alignment=0)
    story.append(Paragraph(f"{unit}執行{project}勤前教育會議人員簽到表", style_title))
    meeting_range = parse_meeting_time(time_str)
    date_part = time_str.split('日')[0] + '日' if '日' in time_str else ""
    story.append(Paragraph(f"時間：{date_part}{meeting_range}", style_top_info))
    loc = str(briefing).strip() if "於" not in str(briefing) else str(briefing).strip().split("於")[1]
    story.append(Paragraph(f"地點：{loc}", style_top_info))
    story.append(Spacer(1, 3*mm))
    table_data = [[Paragraph("分局長：", style_cell_left), "", Paragraph("上級督導：", style_cell_left), ""],
                  [Paragraph("副分局長：", style_cell_left), "", "", ""],
                  [Paragraph("單位", style_cell), Paragraph("參加人員", style_cell), Paragraph("單位", style_cell), Paragraph("參加人員", style_cell)]]
    rows = [("交通組", "中興派出所"), ("勤務指揮中心", "石門派出所"), ("督察組", "高平派出所"), ("聖亭派出所", "三和派出所"), ("龍潭派出所", "龍潭交通分隊")]
    for l, r in rows: table_data.append([Paragraph(l, style_cell), "", Paragraph(r, style_cell), ""])
    t = Table(table_data, colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3], rowHeights=[18*mm, 18*mm, 10*mm] + [26*mm]*len(rows))
    t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), font), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                           ('ALIGN', (0,0), (0,0), 'LEFT'), ('ALIGN', (2,0), (2,0), 'LEFT'), ('ALIGN', (0,1), (0,1), 'LEFT'), ('SPAN', (0,1), (3,1)),
                           ('BACKGROUND', (0,2), (0,2), colors.whitesmoke), ('BACKGROUND', (2,2), (2,2), colors.whitesmoke)]))
    story.append(t)
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph("備註：請將行動電話調整為靜音。", style_note))
    doc.build(story)
    return buf.getvalue()

# 寄信功能
def send_report_email(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"], msg["Subject"] = sender, sender, f"二階段勤務規劃_{datetime.now().strftime('%m%d')}"
        msg.attach(MIMEText("附件為二階段勤務規劃表與人員簽到表 PDF。", "plain", "utf-8"))
        
        pdf1 = generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_ptl, df_cp)
        part1 = MIMEBase("application", "pdf")
        part1.set_payload(pdf1)
        encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote('二階段規劃表.pdf')}")
        msg.attach(part1)
        
        pdf2 = generate_attendance_pdf(unit, project, time_str, briefing)
        part2 = MIMEBase("application", "pdf")
        part2.set_payload(pdf2)
        encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote('簽到表.pdf')}")
        msg.attach(part2)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# 智慧代號計算
def auto_assign_radio_code(df):
    base_prefixes = {"交通分隊": "99", "聖亭": "5", "龍潭": "6", "中興": "7", "石門": "8", "高平": "9", "三和": "3"}
    for idx, row in df.iterrows():
        unit, person = str(row.get('單位', '')), str(row.get('服勤人員', ''))
        if not unit: continue
        first_unit = re.split(r'[\n、 ]', unit.strip())[0]
        base_pfx = next((v for k, v in base_prefixes.items() if k in first_unit), "")
        if base_pfx:
            if "副所長" in person: df.at[idx, '無線電'] = f"隆安{base_pfx}2"
            elif "所長" in person: df.at[idx, '無線電'] = f"隆安{base_pfx}1"
            elif not str(row.get('無線電', '')).startswith(f"隆安{base_pfx}"): df.at[idx, '無線電'] = f"隆安{base_pfx}0"
    return df

# 主介面
df_set, df_cmd, df_ptl, df_cp, err = load_data()
if err or df_set is None:
    u, t, p, b = DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_PROJ, DEFAULT_BRIEF
    ed_cmd, ed_ptl, ed_cp = DEFAULT_CMD.copy(), DEFAULT_PTL.copy(), DEFAULT_CHECKPOINT.copy()
else:
    d = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    u, t, p, b = d.get("unit_name", DEFAULT_UNIT), d.get("plan_full_time", DEFAULT_TIME), d.get("project_name", DEFAULT_PROJ), d.get("briefing_info", DEFAULT_BRIEF)
    ed_cmd, ed_ptl = df_cmd, df_ptl
    ed_cp = df_cp if df_cp is not None and not df_cp.empty else DEFAULT_CHECKPOINT.copy()

st.title("🚓 二階段專案勤務規劃系統")
c1, c2 = st.columns(2)
p_name = c1.text_input("專案名稱", p)
p_time = c2.text_input("勤務時間", t)

st.subheader("1. 指揮編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True)
b_info = st.text_area("📢 勤前教育", b, height=70)

st.subheader("2. 勤務編組 (兩階段)")
tab1, tab2 = st.tabs(["📍 第一階段：機動巡邏", "🚧 第二階段：定點路檢及機動攔檢"])
with tab1:
    res_ptl_raw = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True, key="ptl_editor")
    res_ptl = auto_assign_radio_code(res_ptl_raw.copy())
with tab2:
    res_cp_raw = st.data_editor(ed_cp, num_rows="dynamic", use_container_width=True, key="cp_editor")
    res_cp = auto_assign_radio_code(res_cp_raw.copy())

def get_html():
    style = "<style>body{font-family:'標楷體';padding:10px;} th,td{border:1px solid black;padding:6px;font-size:12pt;text-align:center;} .middle-block{font-size:12pt;margin:15px 0 15px 20px;line-height:1.6; text-align:left;}</style>"
    html = f"<html>{style}<body><h3 style='text-align:center'>{u}<br>{p_name}</h3><div style='text-align:right'><b>時間：{p_time}</b></div><table><tr><th colspan='4'>任 務 編 組</th></tr>"
    for _, r in res_cmd.iterrows():
        html += f"<tr><td><b>{str(r.get('職稱','')).replace('\n', '<br>')}</b></td><td>{str(r.get('代號','')).replace('\n', '<br>')}</td><td>{str(r.get('姓名','')).replace('、','<br>').replace('\n','<br>')}</td><td style='text-align:left'>{str(r.get('任務','')).replace('\n', '<br>')}</td></tr>"
    html += f"</table><div class='middle-block'><b>📢 勤前教育：</b><br>{str(b_info).replace('\n', '<br>')}</div>"
    
    html += "<h4>第一階段：機動巡邏</h4><table><tr><th>編組</th><th>代號</th><th>單位</th><th>人員</th><th>任務</th></tr>"
    for _, r in res_ptl.iterrows():
        html += f"<tr><td>{str(r.get('編組','')).replace('\n', '<br>')}</td><td>{str(r.get('無線電','')).replace('\n', '<br>')}</td><td>{str(r.get('單位','')).replace('、','<br>').replace('\n','<br>')}</td><td>{str(r.get('服勤人員','')).replace('、','<br>').replace('\n','<br>')}</td><td style='text-align:left'>{str(r.get('任務分工','')).replace('\n', '<br>')}</td></tr>"
    html += "</table><h4>第二階段：定點路檢</h4><table><tr><th>編組</th><th>代號</th><th>單位</th><th>人員</th><th>任務</th></tr>"
    for _, r in res_cp.iterrows():
        html += f"<tr><td>{str(r.get('編組','')).replace('\n', '<br>')}</td><td>{str(r.get('無線電','')).replace('\n', '<br>')}</td><td>{str(r.get('單位','')).replace('、','<br>').replace('\n','<br>')}</td><td>{str(r.get('服勤人員','')).replace('、','<br>').replace('\n','<br>')}</td><td style='text-align:left'>{str(r.get('任務分工','')).replace('\n', '<br>')}</td></tr>"
    return html + "</table></body></html>"

st.markdown("---")
with st.expander("點擊展開即時預覽"):
    st.components.v1.html(get_html(), height=600, scrolling=True)

col_dl1, col_dl2 = st.columns(2)
pdf_plan = generate_pdf_from_data(u, p_name, p_time, b_info, res_cmd, res_ptl, res_cp)
col_dl1.download_button("📝 下載 1.勤務規劃表", data=pdf_plan, file_name=f"規劃表_{datetime.now().strftime('%m%d')}.pdf", use_container_width=True)
pdf_attendance = generate_attendance_pdf(u, p_name, p_time, b_info)
col_dl2.download_button("🖋️ 下載 2.人員簽到表", data=pdf_attendance, file_name=f"簽到表_{datetime.now().strftime('%m%d')}.pdf", use_container_width=True)

if st.button("💾 同步雲端並發送備份郵件", use_container_width=True):
    with st.spinner("同步中..."):
        if save_data(u, p_time, p_name, b_info, res_cmd, res_ptl, res_cp):
            ok, mail_err = send_report_email(u, p_name, p_time, b_info, res_cmd, res_ptl, res_cp)
            if ok: st.success("✅ 同步成功並已寄出郵件！")
            else: st.warning(f"⚠️ 雲端已同步，但郵件失敗: {mail_err}")
