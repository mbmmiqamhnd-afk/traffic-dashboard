import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import smtplib, io, os, traceback
import urllib.parse as _ul
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
import re as _re_safe

try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    st.sidebar.warning("找不到 menu.py，跳過側邊欄載入。")

# ==========================================
# 常數與 Google 授權設定 (使用全新 Sheet Tab)
# ==========================================
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

S1_COLS = ["組別", "無線電代號", "派遣單位", "職別", "姓名", "任務分工", "攜行裝備", "機動攔檢區域"]
S2_COLS = ["組別", "無線電代號", "派遣單位", "職別", "姓名", "任務分工", "臨檢目標場所"]
S3_COLS = ["組別", "無線電代號", "派遣單位", "職別", "姓名", "任務分工", "攜行裝備", "定點路檢目標"]

WS_SET_NAME = "三階段_設定"
WS_CMD_NAME = "三階段_指揮組"
WS_S1_NAME  = "三階段_一階機動"
WS_S2_NAME  = "三階段_二階臨檢"
WS_S3_NAME  = "三階段_三階路檢"

DEFAULT_UNIT  = "桃園市政府警察局龍潭分局"
DEFAULT_TIME  = "115年3月25日 18時至22時"
DEFAULT_PROJ  = "0325「雷霆除暴專案」三階段專案"

# 各階段預設勤務時間
DEFAULT_S1_TIME = "18時00分至19時30分"
DEFAULT_S2_TIME = "19時30分至20時30分"
DEFAULT_S3_TIME = "20時30分至22時00分"

DEFAULT_BRIEF = (
    "一、 工作重點任務提示：同仁執行盤查、臨檢及路檢勤務過程中，應強化敵情觀念，提高危機意識，並特別注意人犯戒護。\n"
    "二、 行動要領：除法律另有規定外，警察人員執行場所之臨檢，應限於已發生危害或依客觀合理判斷易生危害之場所。\n"
    "三、 盤查規範：確實依司法院大法官釋字第535號解釋及「警察職權行使法」相關規定，遵守比例原則。\n"
    "四、 全程蒐證：務必全程連續錄音或錄影，以避免因案件招致物議。"
)

DEFAULT_S1_FOCUS = "【第一階段：機動攔檢】採取全面機動巡邏，針對酒駕熱點攔停盤查；攔獲疑似改裝噪音車，立即引導檢驗。"
DEFAULT_S2_FOCUS = "【第二階段：場所臨檢】由各組機動警力，會合偵查隊專案人員，準時進入目標場所執行威力掃蕩。"
DEFAULT_S3_FOCUS = "【第三階段：定點路檢】於各聯外道路、重要路口設立檢測點，加強取締酒後駕車及危險駕車。"

DEFAULT_CMD = pd.DataFrame([{"項目": "指揮官", "通訊代號": "隆安 1 號", "任務目標": "重點機動督導", "負責人員": "分局長 施宇峰", "共同執行人員": "秘書 陳鵬翔"}])
DEFAULT_S1 = pd.DataFrame([{"組別": "第1機動組", "無線電代號": "隆安51", "派遣單位": "龍潭所", "職別": "所長", "姓名": "王小明", "任務分工": "帶班", "攜行裝備": "槍彈、密錄器", "機動攔檢區域": "龍潭市區中正路沿線"}])
DEFAULT_S2 = pd.DataFrame([{"組別": "第1臨檢組", "無線電代號": "隆安51", "派遣單位": "龍潭所", "職別": "所長", "姓名": "王小明", "任務分工": "帶班", "臨檢目標場所": "A. 鉅大撞球館"}])
DEFAULT_S3 = pd.DataFrame([{"組別": "第1路檢組", "無線電代號": "隆安51", "派遣單位": "龍潭所", "職別": "所長", "姓名": "王小明", "任務分工": "帶班兼管制", "攜行裝備": "槍彈、酒測器", "定點路檢目標": "北龍路319號前"}])

# ==========================================
# 工具函數區塊
# ==========================================
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    for p in ["./kaiu.ttf", "kaiu.ttf", "/usr/share/fonts/truetype/custom/kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def safe_str(val):
    if val is None: return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s

def clean_df_to_list(df): return df.astype(str).values.tolist()

def get_commander_name(df_cmd):
    if not df_cmd.empty and "項目" in df_cmd.columns and "負責人員" in df_cmd.columns:
        cmd_row = df_cmd[df_cmd["項目"].str.contains("指揮官", na=False)]
        if not cmd_row.empty: return safe_str(cmd_row.iloc[0]["負責人員"])
    return "分局長"

@st.cache_resource
def get_client():
    try:
        info = dict(st.secrets["gcp_service_account"])
        info["private_key"] = info["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Google 授權失敗：{e}")
        return None

# ==========================================
# 資料存取區塊
# ==========================================
@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, None, None, "權限不足"
        sh = client.open_by_key(SHEET_ID)
        def _get_ws(name, default_df):
            try: return pd.DataFrame(sh.worksheet(name).get_all_records()).fillna("")
            except: return default_df.copy() if default_df is not None else pd.DataFrame()
        return _get_ws(WS_SET_NAME, None), _get_ws(WS_CMD_NAME, DEFAULT_CMD), _get_ws(WS_S1_NAME, DEFAULT_S1), _get_ws(WS_S2_NAME, DEFAULT_S2), _get_ws(WS_S3_NAME, DEFAULT_S3), None
    except Exception as e:
        return None, None, None, None, None, str(e)

def save_data(unit, time_str, project, briefing, df_cmd, df_s1, df_s2, df_s3, stats, t_s1, t_s2, t_s3, f_s1, f_s2, f_s3):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        def _update_ws(name, df, cols=20):
            try: ws = sh.worksheet(name)
            except: ws = sh.add_worksheet(title=name, rows="100", cols=str(cols))
            ws.clear()
            clean_df = df.dropna(how="all").fillna("")
            if not clean_df.empty: ws.update(range_name="A1", values=[clean_df.columns.tolist()] + clean_df_to_list(clean_df))
        
        try: ws_set = sh.worksheet(WS_SET_NAME)
        except: ws_set = sh.add_worksheet(title=WS_SET_NAME, rows="50", cols="5")
        ws_set.clear()
        ws_set.update(range_name="A1", values=[
            ["Key", "Value"], ["unit_name", unit], ["plan_full_time", time_str], ["project_name", project],
            ["briefing_info", briefing], ["stats_cmd", str(stats["cmd"])],
            ["stats_s1", str(stats["s1"])], ["stats_s2", str(stats["s2"])], ["stats_s3", str(stats["s3"])],
            ["stats_inv", str(stats["inv"])], ["stats_civ", str(stats["civ"])],
            ["briefing_time", str(stats["b_time"])], ["briefing_loc", str(stats["b_loc"])],
            ["s1_time", t_s1], ["s2_time", t_s2], ["s3_time", t_s3], 
            ["s1_focus", f_s1], ["s2_focus", f_s2], ["s3_focus", f_s3],
        ])

        _update_ws(WS_CMD_NAME, df_cmd); _update_ws(WS_S1_NAME, df_s1); _update_ws(WS_S2_NAME, df_s2); _update_ws(WS_S3_NAME, df_s3)
        load_data.clear()
        return True
    except Exception as e:
        st.error(f"❌ 同步失敗：{e}")
        return False

# ==========================================
# PDF 產出區塊
# ==========================================
def generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_s1, df_s2, df_s3, stats, t_s1, t_s2, t_s3, f_s1, f_s2, f_s3):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=10*mm, rightMargin=10*mm, topMargin=12*mm, bottomMargin=15*mm)
    page_width = A4[0] - 20*mm
    story = []

    style_title      = ParagraphStyle("Title", fontName=font, fontSize=18, leading=26, alignment=1, spaceAfter=8, wordWrap="CJK")
    style_section    = ParagraphStyle("Section", fontName=font, fontSize=14, leading=20, spaceAfter=2*mm, spaceBefore=4*mm, wordWrap="CJK")
    style_text       = ParagraphStyle("Text", fontName=font, fontSize=14, leading=20, wordWrap="CJK")
    style_cell       = ParagraphStyle("Cell", fontName=font, fontSize=12, leading=17, alignment=1, wordWrap="CJK")
    style_cell_left  = ParagraphStyle("CellLeft", fontName=font, fontSize=12, leading=17, alignment=0, wordWrap="CJK")
    style_target     = ParagraphStyle("Target", fontName=font, fontSize=10, leading=14, alignment=0, wordWrap="CJK")
    style_briefing   = ParagraphStyle("Briefing", fontName=font, fontSize=14, leading=22, leftIndent=22, firstLineIndent=-22, wordWrap="CJK")
    def clean(t): return safe_str(t).replace("\n", "<br/>")

    story.append(Paragraph(f"<b>{unit}執行 {project} 勤務規劃表</b>", style_title))
    story.append(Paragraph("<b>壹、 勤務基本資料</b>", style_section))
    date_str = clean(time_str.split(" ")[0] if " " in time_str else "")
    time_str_only = clean(time_str.split(" ")[1] if " " in time_str else "")
    t_basic = Table([
        [Paragraph(f"<b>{h}</b>", style_cell) for h in ["實施日期", "勤務時間", "指揮官", "勤務編組", "勤前教育"]],
        [Paragraph(date_str, style_cell), Paragraph(time_str_only, style_cell), Paragraph(get_commander_name(df_cmd), style_cell), Paragraph("如任務編組表", style_cell), Paragraph(f"{stats['b_time']}<br/>{stats['b_loc']}", style_cell)]
    ], colWidths=[page_width*0.19, page_width*0.16, page_width*0.19, page_width*0.16, page_width*0.30])
    t_basic.setStyle(TableStyle([("FONTNAME", (0,0), (-1,-1), font), ("GRID", (0,0), (-1,-1), 0.5, colors.black), ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#f2f2f2")), ("VALIGN", (0,0), (-1,-1), "MIDDLE")]))
    story.append(t_basic)

    story.append(Paragraph("<b>貳、 警力統計</b>", style_section))
    total = stats["cmd"] + stats["s1"] + stats["s2"] + stats["s3"] + stats["inv"] + stats["civ"]
    t_stats = Table([
        [Paragraph(f"<b>{h}</b>", style_cell) for h in ["督導組", "一階(機動)", "二階(臨檢)", "三階(路檢)", "偵訊/民力", "總計"]],
        [Paragraph(str(stats["cmd"]), style_cell), Paragraph(str(stats["s1"]), style_cell), Paragraph(str(stats["s2"]), style_cell), Paragraph(str(stats["s3"]), style_cell), Paragraph(f"{stats['inv']} / {stats['civ']}", style_cell), Paragraph(str(total), style_cell)]
    ], colWidths=[page_width/6]*6)
    t_stats.setStyle(TableStyle([("FONTNAME", (0,0), (-1,-1), font), ("GRID", (0,0), (-1,-1), 0.5, colors.black), ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#f2f2f2")), ("VALIGN", (0,0), (-1,-1), "MIDDLE")]))
    story.append(t_stats); story.append(Spacer(1, 4*mm))

    story.append(Paragraph("<b>參、 督導及其他任務編組表</b>", style_section))
    data_cmd = [[Paragraph(f"<b>{h}</b>", style_cell) for h in ["項目","通訊代號","任務目標","負責人員","共同人員"]]]
    for _, r in df_cmd.iterrows(): data_cmd.append([Paragraph(clean(r.get("項目","")), style_cell), Paragraph(clean(r.get("通訊代號","")), style_cell), Paragraph(clean(r.get("任務目標","")), style_cell_left), Paragraph(clean(r.get("負責人員","")), style_cell), Paragraph(clean(r.get("共同執行人員","")), style_cell)])
    t_cmd = Table(data_cmd, colWidths=[page_width*0.13, page_width*0.14, page_width*0.26, page_width*0.25, page_width*0.22])
    t_cmd.setStyle(TableStyle([("FONTNAME", (0,0), (-1,-1), font), ("GRID", (0,0), (-1,-1), 0.5, colors.black), ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#f2f2f2")), ("VALIGN", (0,0), (-1,-1), "MIDDLE")]))
    story.append(t_cmd)

    def build_stage_table(df_stage, target_col, col_widths, bg_color="#f2f2f2"):
        headers = ["組別","無線電\n代號","派遣\n單位","職別","姓名","任務分工"] + target_col
        data = [[Paragraph(f"<b>{h}</b>", style_cell) for h in headers]]
        
        # 【自動防呆處理】複製並改寫 rows，將同組內填寫的地點擴散給全組各列，避免 ReportLab 垂直合併 SPAN 機制吃掉非首列之新地點
        rows = df_stage.reset_index(drop=True).copy()
        for col_name in target_col:
            if col_name in rows.columns:
                rows[col_name] = rows[col_name].replace(r'^\s*$', pd.NA, regex=True)
                rows[col_name] = rows.groupby('組別')[col_name].transform(lambda x: x.ffill().bfill()).fillna("")

        merges, prev_grp, grp_start = [], None, 1
        for i, r in rows.iterrows():
            grp = safe_str(r.get("組別",""))
            if grp != prev_grp:
                if prev_grp is not None: merges.append((grp_start, i))
                prev_grp, grp_start = grp, i + 1
            row_data = [Paragraph(clean(r.get(c,"")), style_cell if c != "任務分工" else style_cell_left) for c in ["組別","無線電代號","派遣單位","職別","姓名","任務分工"]]
            for col_name in target_col: row_data.append(Paragraph(clean(r.get(col_name,"")), style_target if "目標" in col_name or "區域" in col_name else style_cell_left))
            data.append(row_data)
        if prev_grp is not None: merges.append((grp_start, len(rows)))
        t = Table(data, colWidths=col_widths, splitByRow=True)
        ts = [("FONTNAME", (0,0), (-1,-1), font), ("GRID", (0,0), (-1,-1), 0.5, colors.black), ("BACKGROUND", (0,0), (-1,0), colors.HexColor(bg_color)), ("VALIGN", (0,0), (-1,-1), "MIDDLE")]
        for rs, re in merges:
            if re > rs:
                for c in [0, 1, len(headers)-1]: ts.append(("SPAN", (c, rs), (c, re)))
        t.setStyle(TableStyle(ts))
        return t

    # 肆、第一階段 PDF 加上時間
    story.append(Paragraph("<b>肆、【第一階段】機動攔檢任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務時間：</b>{clean(t_s1)}", style_text))
    story.append(Paragraph(f"<b>勤務重點：</b><br/>{clean(f_s1)}", style_text))
    story.append(build_stage_table(df_s1, ["攜行裝備", "機動攔檢區域"], [page_width*0.1, page_width*0.09, page_width*0.09, page_width*0.1, page_width*0.11, page_width*0.12, page_width*0.15, page_width*0.24]))

    # 伍、第二階段 PDF 加上時間
    story.append(Paragraph("<b>伍、【第二階段】場所臨檢任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務時間：</b>{clean(t_s2)}", style_text))
    story.append(Paragraph(f"<b>勤務重點：</b><br/>{clean(f_s2)}", style_text))
    story.append(build_stage_table(df_s2, ["臨檢目標場所"], [page_width*0.1, page_width*0.09, page_width*0.09, page_width*0.1, page_width*0.11, page_width*0.16, page_width*0.35], bg_color="#e6e6e6"))

    # 陸、第三階段 PDF 加上時間
    story.append(Paragraph("<b>陸、【第三階段】定點路檢任務編組</b>", style_section))
    story.append(Paragraph(f"<b>勤務時間：</b>{clean(t_s3)}", style_text))
    story.append(Paragraph(f"<b>勤務重點：</b><br/>{clean(f_s3)}", style_text))
    story.append(build_stage_table(df_s3, ["攜行裝備", "定點路檢目標"], [page_width*0.1, page_width*0.09, page_width*0.09, page_width*0.1, page_width*0.11, page_width*0.12, page_width*0.15, page_width*0.24]))

    story.append(Paragraph("<b>柒、 工作重點與法令宣導</b>", style_section))
    for line in str(briefing).split("\n"):
        if line.strip(): story.append(Paragraph(clean(line), style_briefing))

    def add_footer(canvas, doc):
        canvas.saveState(); canvas.setFont(font, 10); canvas.drawCentredString(A4[0]/2.0, 10*mm, f"-第{canvas.getPageNumber()}頁-"); canvas.restoreState()
    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    return buf.getvalue()

def generate_attendance_pdf(unit, project, time_str, stats, df_cmd):
    font = _get_font()
    buf  = io.BytesIO()
    doc  = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=10*mm, bottomMargin=10*mm)
    page_width = A4[0] - 30*mm
    story = []
    style_title = ParagraphStyle("Title", fontName=font, fontSize=18, leading=26, alignment=1, spaceAfter=8, wordWrap="CJK")
    style_info  = ParagraphStyle("Info",  fontName=font, fontSize=14, leading=22, spaceAfter=1*mm, wordWrap="CJK")
    style_cell  = ParagraphStyle("Cell",  fontName=font, fontSize=14, leading=20, alignment=1, wordWrap="CJK")
    style_sig   = ParagraphStyle("Sig",  fontName=font, fontSize=14, leading=20, alignment=0, wordWrap="CJK") 

    story.append(Paragraph(f"{unit}執行{project}簽到表", style_title))
    date_part = time_str.split(" ")[0] if " " in time_str else "115年3月25日"
    story.append(Paragraph(f"時間：{date_part} {stats['b_time']}", style_info))
    story.append(Paragraph(f"地點：{stats['b_loc']}召開", style_info))
    story.append(Spacer(1, 3*mm))

    commander_title = get_commander_name(df_cmd).split(" ")[0] if " " in get_commander_name(df_cmd) else get_commander_name(df_cmd)
    t_sig = Table([[Paragraph(f"{commander_title}：", style_sig), Paragraph("上級督導：", style_sig)], [Paragraph("副分局長：", style_sig), ""]], colWidths=[page_width/2.0]*2)
    t_sig.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("BOTTOMPADDING", (0,0), (-1,-1), 2)]))
    story.append(t_sig)
    story.append(Spacer(1, 4*mm))

    rows = [("交通組", "聖亭派出所"), ("督察組", "龍潭派出所"), ("行政組", "中興派出所"), ("保安民防組", "石門派出所"), ("勤務指揮中心","高平派出所"), ("偵查隊", "三和派出所"), ("", "龍潭交通分隊")]
    table_data = [[Paragraph("單位", style_cell), Paragraph("參加人員", style_cell), Paragraph("單位", style_cell), Paragraph("參加人員", style_cell)]]
    for l, r in rows: table_data.append([Paragraph(l, style_cell) if l else "", "", Paragraph(r, style_cell) if r else "", ""])
    t = Table(table_data, colWidths=[page_width*0.2, page_width*0.3, page_width*0.2, page_width*0.3], rowHeights=[10*mm] + [20*mm]*len(rows))
    t.setStyle(TableStyle([("FONTNAME", (0,0), (-1,-1), font), ("GRID", (0,0), (-1,-1), 0.5, colors.black), ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("BACKGROUND", (0,0), (3,0), colors.whitesmoke)]))
    story.append(t)
    
    def add_footer(canvas, doc):
        canvas.saveState(); canvas.setFont(font, 10); canvas.drawCentredString(A4[0]/2.0, 10*mm, f"-第{canvas.getPageNumber()}頁-"); canvas.restoreState()
    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    return buf.getvalue()

def send_report_email(unit, project, time_str, briefing, df_cmd, df_s1, df_s2, df_s3, stats, t_s1, t_s2, t_s3, f_s1, f_s2, f_s3):
    try:
        sender = st.secrets["email"]["user"]; pwd = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"] = sender; msg["To"] = sender; msg["Subject"] = f"勤務規劃與簽到表_{datetime.now().strftime('%m%d')} (三階段專案)"
        msg.attach(MIMEText("附件為最新版本「三階段專案」勤務規劃表與簽到表。", "plain"))

        pdf1 = generate_pdf_from_data(unit, project, time_str, briefing, df_cmd, df_s1, df_s2, df_s3, stats, t_s1, t_s2, t_s3, f_s1, f_s2, f_s3)
        part1 = MIMEBase("application", "pdf"); part1.set_payload(pdf1); encoders.encode_base64(part1)
        part1.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(f'{unit}規劃表(三階段).pdf')}"); msg.attach(part1)

        pdf2 = generate_attendance_pdf(unit, project, time_str, stats, df_cmd)
        part2 = MIMEBase("application", "pdf"); part2.set_payload(pdf2); encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(f'{unit}簽到表(三階段).pdf')}"); msg.attach(part2)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd); server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)


# ==========================================
# UI 介面與【自動匯入】邏輯
# ==========================================
df_set, df_cmd, df_s1, df_s2, df_s3, err = load_data()

# 自動匯入二合一舊資料當預設值
if not err and (df_set is None or (isinstance(df_set, pd.DataFrame) and df_set.empty)):
    try:
        client = get_client()
        if client:
            sh = client.open_by_key(SHEET_ID)
            old_set = pd.DataFrame(sh.worksheet("二合一_設定").get_all_records()).fillna("")
            old_cmd = pd.DataFrame(sh.worksheet("二合一_指揮組").get_all_records()).fillna("")
            old_ptl = pd.DataFrame(sh.worksheet("二合一_路檢組").get_all_records()).fillna("")
            old_cp  = pd.DataFrame(sh.worksheet("二合一_擴大臨檢組").get_all_records()).fillna("")
            
            if not old_set.empty:
                st.toast("✨ 偵測到首次啟用，已自動為您匯入舊版『二合一』的人員名單！", icon="📥")
                df_set = old_set
                if not old_cmd.empty: df_cmd = old_cmd
                if not old_ptl.empty:
                    df_s1 = old_ptl.copy()
                    if "臨檢目標" in df_s1.columns: df_s1.rename(columns={"臨檢目標": "機動攔檢區域"}, inplace=True)
                    df_s3 = old_ptl.copy()
                    if "臨檢目標" in df_s3.columns: df_s3.rename(columns={"臨檢目標": "定點路檢目標"}, inplace=True)
                if not old_cp.empty: df_s2 = old_cp.copy()
    except Exception:
        pass 

default_stats = {"cmd": 7, "s1": 10, "s2": 10, "s3": 10, "inv": 3, "civ": 0, "b_time": "18時30分至19時00分", "b_loc": "本分局2樓會議室"}

if err or df_set is None or (isinstance(df_set, pd.DataFrame) and df_set.empty):
    u, t, p = DEFAULT_UNIT, DEFAULT_TIME, DEFAULT_PROJ
    ed_cmd, ed_s1, ed_s2, ed_s3 = DEFAULT_CMD.copy(), DEFAULT_S1.copy(), DEFAULT_S2.copy(), DEFAULT_S3.copy()
    f_s1, f_s2, f_s3 = DEFAULT_S1_FOCUS, DEFAULT_S2_FOCUS, DEFAULT_S3_FOCUS
    t_s1, t_s2, t_s3 = DEFAULT_S1_TIME, DEFAULT_S2_TIME, DEFAULT_S3_TIME 
else:
    d = dict(zip(df_set.iloc[:,0], df_set.iloc[:,1]))
    u, t, p = d.get("unit_name", DEFAULT_UNIT), d.get("plan_full_time", DEFAULT_TIME), d.get("project_name", DEFAULT_PROJ)
    f_s1 = d.get("s1_focus", DEFAULT_S1_FOCUS)
    f_s2 = d.get("s2_focus", DEFAULT_S2_FOCUS)
    f_s3 = d.get("s3_focus", DEFAULT_S3_FOCUS)
    t_s1 = d.get("s1_time", DEFAULT_S1_TIME)
    t_s2 = d.get("s2_time", DEFAULT_S2_TIME)
    t_s3 = d.get("s3_time", DEFAULT_S3_TIME)

    default_stats.update({
        "cmd": int(d.get("stats_cmd", 7)), "s1": int(d.get("stats_s1", 10)), "s2": int(d.get("stats_s2", 10)), "s3": int(d.get("stats_s3", 10)),
        "inv": int(d.get("stats_inv", 3)), "civ": int(d.get("stats_civ", 0)), "b_time": d.get("briefing_time", "18時30分至19時00分"), "b_loc": d.get("briefing_loc", "本分局2樓會議室")
    })
    ed_cmd = df_cmd.astype(str)
    ed_s1 = df_s1[S1_COLS].astype(str) if not df_s1.empty and all(c in df_s1.columns for c in S1_COLS) else df_s1.astype(str)
    ed_s2 = df_s2[S2_COLS].astype(str) if not df_s2.empty and all(c in df_s2.columns for c in S2_COLS) else df_s2.astype(str)
    ed_s3 = df_s3[S3_COLS].astype(str) if not df_s3.empty and all(c in df_s3.columns for c in S3_COLS) else df_s3.astype(str)

st.title("三階段專案勤務規劃系統 🚓")
st.info("💡 系統特色：本功能將勤務分為「一階機動攔檢」、「二階場所臨檢」、「三階定點路檢」進行無縫排班規劃。")

if err: st.warning(f"⚠️ 無法連線 Google Sheets（{err}），顯示預設資料。")

st.subheader("壹、 勤務基本資料")
col_b1, col_b2 = st.columns(2)
with col_b1: p_time = st.text_input("勤務時間", t)
with col_b2: p_input = st.text_input("專案名稱", _re_safe.sub(r'^\d{4}「?', '', p))
col_b3, col_b4 = st.columns(2)
with col_b3: input_b_time = st.text_input("勤前教育時間", default_stats["b_time"])
with col_b4: input_b_loc = st.text_input("勤前教育地點", default_stats["b_loc"])

date_match = _re_safe.search(r'(\d+)年(\d+)月(\d+)日', p_time)
auto_4_digit = f"{int(date_match.group(2)):02d}{int(date_match.group(3)):02d}" if date_match else datetime.now().strftime("%m%d")
p_name = f"{auto_4_digit}「{p_input.replace('「', '').replace('」', '')}」"

st.subheader("參、 指揮編組")
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True).dropna(how="all").fillna("")

st.subheader("勤務執行編組 (三階段)")
tab1, tab2, tab3 = st.tabs(["肆、【第一階段】機動攔檢", "伍、【第二階段】場所臨檢", "陸、【第三階段】定點路檢"])

def create_editor(tab_obj, key, title, time_val, focus_val, ed_df, col_config):
    with tab_obj:
        res_time = st.text_input(f"勤務時間 ({title})", time_val, key=f"{key}_time_in")
        res_focus = st.text_area(f"勤務重點 ({title})", focus_val, height=80, key=f"{key}_focus_in")
        st.caption("💡 同一組的多名人員請填寫相同的「組別」與「無線電代號」，PDF 會自動合併。")
        res_df = st.data_editor(ed_df, num_rows="dynamic", use_container_width=True, key=f"{key}_ed", column_config=col_config).dropna(how="all").fillna("").reset_index(drop=True)
        return res_time, res_focus, res_df

base_col_config = {
    "組別": st.column_config.TextColumn("組別", width="small"),
    "無線電代號": st.column_config.TextColumn("無線電代號", width="small"),
    "派遣單位": st.column_config.TextColumn("派遣單位", width="small"),
    "職別": st.column_config.TextColumn("職別", width="small"),
    "姓名": st.column_config.TextColumn("姓名", width="small"),
    "任務分工": st.column_config.TextColumn("任務分工", width="medium"),
}

s1_config = {**base_col_config, "攜行裝備": st.column_config.TextColumn("攜行裝備", width="medium"), "機動攔檢區域": st.column_config.TextColumn("機動攔檢區域", width="large")}
s2_config = {**base_col_config, "臨檢目標場所": st.column_config.TextColumn("臨檢目標場所", width="large")}
s3_config = {**base_col_config, "攜行裝備": st.column_config.TextColumn("攜行裝備", width="medium"), "定點路檢目標": st.column_config.TextColumn("定點路檢目標", width="large")}

res_s1_time, res_s1_focus, res_s1 = create_editor(tab1, "s1", "第一階段", t_s1, f_s1, ed_s1, s1_config)
res_s2_time, res_s2_focus, res_s2 = create_editor(tab2, "s2", "第二階段", t_s2, f_s2, ed_s2, s2_config)
res_s3_time, res_s3_focus, res_s3 = create_editor(tab3, "s3", "第三階段", t_s3, f_s3, ed_s3, s3_config)

def count_people(df):
    return int(df["姓名"].astype(str).str.strip().loc[lambda x: x!=""].count()) if not df.empty and "姓名" in df.columns else 0

current_stats = {
    "cmd": int(res_cmd["負責人員"].astype(str).str.strip().loc[lambda x: x!=""].count()) if not res_cmd.empty and "負責人員" in res_cmd.columns else 0,
    "s1": count_people(res_s1), "s2": count_people(res_s2), "s3": count_people(res_s3), "b_time": input_b_time, "b_loc": input_b_loc
}

st.subheader("貳、 警力統計")
col_adj1, col_adj2 = st.columns(2)
with col_adj1: current_stats["inv"] = st.number_input("偵訊組人數", value=default_stats["inv"], min_value=0)
with col_adj2: current_stats["civ"] = st.number_input("民力人數", value=default_stats["civ"], min_value=0)
current_stats["total"] = sum([current_stats[k] for k in ["cmd", "s1", "s2", "s3", "inv", "civ"]])

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("總警力", f"{current_stats['total']} 人"); m2.metric("督導組", f"{current_stats['cmd']} 人")
m3.metric("一階機動", f"{current_stats['s1']} 人"); m4.metric("二階臨檢", f"{current_stats['s2']} 人")
m5.metric("三階路檢", f"{current_stats['s3']} 人"); m6.metric("偵/民", f"{current_stats['inv']} / {current_stats['civ']} 人")

st.markdown("---")

if st.button("💾 儲存【三階段專案】規劃並發送郵件", use_container_width=True):
    with st.spinner("同步至 Google Sheets 中..."):
        if save_data(u, p_time, p_name, DEFAULT_BRIEF, res_cmd, res_s1, res_s2, res_s3, current_stats, res_s1_time, res_s2_time, res_s3_time, res_s1_focus, res_s2_focus, res_s3_focus):
            with st.spinner("正在產生 PDF 並寄送郵件..."):
                ok, err = send_report_email(u, p_name, p_time, DEFAULT_BRIEF, res_cmd, res_s1, res_s2, res_s3, current_stats, res_s1_time, res_s2_time, res_s3_time, res_s1_focus, res_s2_focus, res_s3_focus)
                if ok: st.success("✅ 資料已同步，且郵件發送成功！")
                else: st.warning(f"⚠️ 同步成功，但郵件失敗：{err}")
