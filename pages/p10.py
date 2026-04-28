import streamlit as st
# 將 set_page_config 移到最前面
st.set_page_config(page_title="防制危險駕車勤務", layout="wide", page_icon="🚔")

from menu import show_sidebar
show_sidebar() # 確保在 config 之後呼叫

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import smtplib
import io
import os
import urllib.parse as _ul
import re
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

# --- 常數與設定 ---
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT_TITLE = "桃園市政府警察局龍潭分局"

CHECKIN_POINTS = """1. 中油高原交流道站（龍源路2-20號）
2. 萊爾富超商-龍潭石門山店（龍源路大平段262號）
3. 7-11龍潭佳園門市（中正路三坑段776號）
4. 旭日路三坑自然生態公園停車場
5. 旭日路與大溪區交界處"""

NOTES = """一、各編組執行前由帶班人員在駐地實施勤前教育。
二、攔檢、盤查車輛時，應隨時注意自身安全及執勤態度。
三、駕駛巡邏車應開啟警示燈，如發現危險駕車行為「勿追車」，請立即向勤指中心報告攔截圍捕。
四、加強攔查改裝排管、無照駕駛、蛇行、逼車、拆除消音器、毒駕及公共危險罪等事項。"""

# --- 2. 格式化工具 ---
def format_staff_only(val):
    if pd.isna(val) or str(val).strip() in ["None", "nan", ""]: return ""
    # 統一將全形頓號或斜線轉為換行
    s = str(val).replace('\\', '\n').replace('、', '\n').replace('\xa0', ' ')
    s = re.sub(r'(\d{2}[:：]?\d{0,2}\s*-\s*\d{2}[:：]?\d{0,2}[時]?[:：])\s*([^\n\s])', r'\1\n\2', s)
    return '\n'.join([l.strip() for l in s.split('\n') if l.strip()])

# --- 3. 雲端與 PDF 引擎 ---
@st.cache_resource
def get_client():
    if "gcp_service_account" not in st.secrets: return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except:
        return None

@st.cache_data(ttl=10)
def load_data():
    try:
        client = get_client()
        if client is None: return None, None, None, "離線模式"
        sh = client.open_by_key(SHEET_ID)
        df_set = pd.DataFrame(sh.worksheet("危駕_設定").get_all_records())
        df_cmd = pd.DataFrame(sh.worksheet("危駕_指揮組").get_all_records())
        df_ptl = pd.DataFrame(sh.worksheet("危駕_警力佈署").get_all_records())
        return df_set, df_cmd, df_ptl, None
    except Exception as e: return None, None, None, str(e)

def save_data(time_str, commander, df_cmd, df_patrol):
    try:
        client = get_client()
        if client is None: return False
        sh = client.open_by_key(SHEET_ID)
        
        ws_set = sh.worksheet("危駕_設定")
        ws_set.clear()
        ws_set.update(range_name='A1', values=[["Key", "Value"], ["plan_time", time_str], ["commander", commander]])
        
        for ws_name, df in [("危駕_指揮組", df_cmd), ("危駕_警力佈署", df_patrol)]:
            ws = sh.worksheet(ws_name)
            ws.clear()
            # 確保資料為字串且處理空值
            data_to_save = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
            ws.update(range_name='A1', values=data_to_save)
        
        load_data.clear()
        return True
    except:
        return False

def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    for p in ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_pdf_from_data(time_str, commander, df_cmd, df_patrol):
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - 24*mm
    story = []
    
    style_title = ParagraphStyle('T', fontName=font, fontSize=16, alignment=1, spaceAfter=8)
    style_info = ParagraphStyle('I', fontName=font, fontSize=12, alignment=2, spaceAfter=10)
    style_th = ParagraphStyle('H', fontName=font, fontSize=16, alignment=1, leading=22)
    style_cell = ParagraphStyle('C', fontName=font, fontSize=14, leading=20, alignment=1) 
    style_cell_l = ParagraphStyle('L', fontName=font, fontSize=14, leading=20, alignment=0)
    style_note_hanging = ParagraphStyle('NH', fontName=font, fontSize=14, leading=20, alignment=0, leftIndent=28, firstLineIndent=-28)

    story.append(Paragraph(f"<b>{UNIT_TITLE}執行「防制危險駕車專案勤務」規劃表</b>", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def br(txt, bold_time=True):
        if not txt: return ""
        s = str(txt).replace('\n', '<br/>').replace('\xa0', ' ')
        if bold_time:
            s = re.sub(r'(\d{2}[:：]?\d{0,2}-\d{2}[:：]?\d{0,2}[時]?[:：]?)', r'<b>\1</b>', s)
        return s

    data_cmd = [[Paragraph("<b>任　務　編　組</b>", style_th), '', '', ''], 
                [Paragraph(f"<b>{h}</b>", style_th) for h in ["職稱", "代號", "姓名", "任務"]]]
    for _, r in df_cmd.iterrows():
        if all(str(v).strip() == "" for v in r.values): continue
        data_cmd.append([
            Paragraph(br(r.get('職稱','')), style_cell), 
            Paragraph(br(r.get('代號', '')), style_cell), 
            Paragraph(br(r.get('姓名','')), style_cell), 
            Paragraph(br(r.get('任務','')), style_cell_l)
        ])
    
    t1 = Table(data_cmd, colWidths=[page_width*0.15, page_width*0.15, page_width*0.25, page_width*0.45])
    t1.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#f2f2f2'))]))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    data_ptl = [[Paragraph("<b>警　力　佈　署</b>", style_th), '', '', '', ''], 
                [Paragraph(f"<b>交通快打指揮官：</b>{commander}", style_cell_l), '', '', '', ''], 
                [Paragraph(f"<b>{h}</b>", style_th) for h in ["勤務時段", "代號", "編組", "服勤人員", "任務分工"]]]
    for _, r in df_patrol.iterrows():
        if all(str(v).strip() == "" for v in r.values): continue
        data_ptl.append([
            Paragraph(br(r.get('勤務時段','')), style_cell), 
            Paragraph(br(r.get('代號', '')), style_cell), 
            Paragraph(br(r.get('編組','')), style_cell), 
            Paragraph(br(r.get('服勤人員','')), style_cell), 
            Paragraph(br(r.get('任務分工','')), style_cell_l)
        ])

    t2 = Table(data_ptl, colWidths=[page_width*0.22, page_width*0.13, page_width*0.15, page_width*0.23, page_width*0.27])
    t2.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),font), ('GRID',(0,0),(-1,-1),0.5,colors.black), ('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('SPAN',(0,0),(-1,0)), ('SPAN',(0,1),(-1,1)), ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f2f2f2')), ('BACKGROUND',(0,2),(-1,2),colors.HexColor('#f2f2f2'))]))
    story.append(t2)
    
    story.append(Spacer(1, 6*mm)); story.append(Paragraph("<b>📍 巡簽地點：</b>", style_cell_l))
    for l in CHECKIN_POINTS.split('\n'):
        if l.strip(): story.append(Paragraph(l.strip(), style_note_hanging))
    story.append(Spacer(1, 4*mm)); story.append(Paragraph("<b>📝 備註：</b>", style_cell_l))
    for l in NOTES.split('\n'):
        if l.strip(): story.append(Paragraph(l.strip(), style_note_hanging))

    doc.build(story)
    return buf.getvalue()

# --- 4. 寄信功能 ---
def send_report_email(time_str, commander, df_cmd, df_patrol, custom_filename):
    try:
        if "email" not in st.secrets: return False, "未設定 secrets"
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        pdf_bytes = generate_pdf_from_data(time_str, commander, df_cmd, df_patrol)
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender
        msg["Subject"] = custom_filename
        msg.attach(MIMEText(f"附件為 {custom_filename}。", "plain", "utf-8"))
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        filename_encoded = _ul.quote(f"{custom_filename}.pdf")
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{filename_encoded}")
        msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e: return False, str(e)

# --- 5. 介面與邏輯 ---
df_set, df_cmd, df_ptl, err = load_data()

# 🎯 強制防呆與名稱標準化
if df_cmd is not None:
    df_cmd.columns = [str(c).strip() for c in df_cmd.columns]
    if "無線電" in df_cmd.columns: df_cmd = df_cmd.rename(columns={"無線電": "代號"})
if df_ptl is not None:
    df_ptl.columns = [str(c).strip() for c in df_ptl.columns]
    if "無線電" in df_ptl.columns: df_ptl = df_ptl.rename(columns={"無線電": "代號"})

if err or df_set is None:
    t, cmdr_default = "115年3月6日22時至翌日6時", "石門所副所長林榮裕"
    ed_cmd = pd.DataFrame(columns=["職稱", "代號", "姓名", "任務"])
    ed_ptl = pd.DataFrame(columns=["勤務時段", "代號", "編組", "服勤人員", "任務分工"])
else:
    sd = dict(zip(df_set.iloc[:,0].astype(str), df_set.iloc[:,1].astype(str)))
    t = sd.get("plan_time", "115年3月6日22時至翌日6時")
    cmdr_default = sd.get("commander", "石門所副所長林榮裕")
    ed_cmd, ed_ptl = df_cmd, df_ptl

st.title("🚔 防制危險駕車專案勤務規劃表")

p_time = st.text_input("1. 勤務時間", t)
cmdr_input = st.text_input("2. 交通快打指揮官", cmdr_default)

# --- 核心連動邏輯 (已嚴格修正日期判定邏輯) ---
if len(ed_ptl) == 0:
    ed_ptl = pd.DataFrame([["", "", "", "", ""]], columns=["勤務時段", "代號", "編組", "服勤人員", "任務分工"])

ed_ptl = ed_ptl.reset_index(drop=True)

date_match = re.search(r'(?:(\d+)年)?(\d+)月(\d+)日(.*)', p_time)
if date_match and len(ed_ptl) > 0:
    try:
        y_val = date_match.group(1)
        m_val = int(date_match.group(2))
        d_val = int(date_match.group(3))
        # 這裡提取出來的是「22時至翌日6時」這類純時段
        time_part = date_match.group(4).strip() 
        
        y_tw = int(y_val) if y_val else (datetime.now().year - 1911)
        base_dt = datetime(y_tw + 1911, m_val, d_val)
        next_dt = base_dt + timedelta(days=1)
        
        # 專責警力：強制取跨日 (+1天) 的日期
        dedicated_time = f"{next_dt.month}月{next_dt.day}日\n零時至4時"
        
        # 一般警力：強制取當天 (base_dt) 的日期
        normal_time = f"{base_dt.month}月{base_dt.day}日\n{time_part}"
        
        for i in range(len(ed_ptl)):
            current_time_val = str(ed_ptl.at[i, '勤務時段']).strip()
            current_group_val = str(ed_ptl.at[i, '編組']).strip()
            
            # 只有當下該欄位為空值時才進行配發
            if current_time_val in ["", "nan", "None"]:
                # 絕對嚴謹判斷：只有「第 0 列」或是手動打了「專責」的，才給跨日時間
                if i == 0 or "專責" in current_group_val:
                    ed_ptl.at[i, '勤務時段'] = dedicated_time
                else:
                    # 第二列之後的空白列，絕對給當日時間 (如 4月30日)
                    ed_ptl.at[i, '勤務時段'] = normal_time
    except Exception as e: 
        # 防止解析異常導致崩潰
        pass

if len(ed_ptl) > 0:
    # 判斷代號
    if "分隊" in cmdr_input: base = "隆安99"
    elif "石門" in cmdr_input: base = "隆安8"
    elif "高平" in cmdr_input: base = "隆安9"
    elif "聖亭" in cmdr_input: base = "隆安5"
    elif "龍潭" in cmdr_input: base = "隆安6"
    elif "中興" in cmdr_input: base = "隆安7"
    else: base = "隆安"
    
    suffix = "1" if ("所長" in cmdr_input and "副所長" not in cmdr_input) or ("分隊長" in cmdr_input and "副分隊長" not in cmdr_input) else "2"
    
    # 針對第一列(專責)自動填入預設代號與編組
    if str(ed_ptl.at[0, '代號']).strip() in ["", "nan", "None"]:
        ed_ptl.at[0, '代號'] = base + suffix
        
    if str(ed_ptl.at[0, '編組']).strip() in ["", "nan", "None"]:
        unit_match = re.search(r'([\u4e00-\u9fa5]+?(?:派出所|所|分隊|警備隊))', cmdr_input)
        if unit_match:
            pu = re.sub(r'派出所$', '所', unit_match.group(1))
            ed_ptl.at[0, '編組'] = f"專責警力\n（{pu}輪值）"

# 顯示編輯器
st.subheader("3. 任務編組")
if '姓名' in ed_cmd.columns:
    ed_cmd['姓名'] = ed_cmd['姓名'].apply(format_staff_only)
res_cmd = st.data_editor(ed_cmd, num_rows="dynamic", use_container_width=True).fillna("")

st.subheader("4. 警力佈署")
if '服勤人員' in ed_ptl.columns: 
    ed_ptl['服勤人員'] = ed_ptl['服勤人員'].apply(format_staff_only)
res_ptl = st.data_editor(ed_ptl, num_rows="dynamic", use_container_width=True).fillna("")

# 6. 預覽
def get_preview(df_c, df_p, cmdr_n, time_s):
    cmd_rows = ""
    for _, r in df_c.iterrows():
        if all(str(v).strip() == "" for v in r.values): continue
        cmd_rows += f"<tr><td>{str(r.get('職稱','')).replace('\\n','<br>')}</td><td>{r.get('代號','')}</td><td>{str(r.get('姓名','')).replace('\\n','<br>').replace('\n','<br>')}</td><td>{r.get('任務','')}</td></tr>"
    
    ptl_rows = ""
    for _, r in df_p.iterrows():
        if all(str(v).strip() == "" for v in r.values): continue
        ptl_rows += f"<tr><td>{str(r.get('勤務時段','')).replace('\\n','<br>').replace('\n','<br>')}</td><td>{r.get('代號','')}</td><td>{str(r.get('編組','')).replace('\\n','<br>').replace('\n','<br>')}</td><td>{str(r.get('服勤人員','')).replace('\\n','<br>').replace('\n','<br>')}</td><td>{r.get('任務分工','')}</td></tr>"
    
    return f"""<style>table {{ width:100%; border-collapse:collapse; font-family:"標楷體"; }} th,td {{ border:1px solid black; padding:8px; text-align:center; }} th {{ background:#f2f2f2; font-size:16pt; }} td {{ font-size:14pt; }}</style>
    <h2 style='text-align:center;'>{UNIT_TITLE} 規劃表</h2><div style='text-align:right;'>勤務時間：{time_s}</div><br>
    <table><tr><th colspan="4">任 務 編 組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>{cmd_rows}</table><br>
    <table><tr><th colspan="5">警 力 佈 署</th></tr><tr><th colspan="5" style="text-align:left;">交通快打指揮官：{cmdr_n}</th></tr><tr><th>勤務時段</th><th>代號</th><th>編組</th><th>服勤人員</th><th>任務分工</th></tr>{ptl_rows}</table>"""

st.components.v1.html(get_preview(res_cmd, res_ptl, cmdr_input, p_time), height=500, scrolling=True)

# 7. 儲存與下載
if st.button("💾 同步、寄信並下載 PDF", type="primary"):
    date_fn = "".join(re.findall(r'\d+', p_time))[:8] if date_match else datetime.now().strftime('%Y%m%d')
    final_filename = f"防制危險駕車勤務規劃表_{date_fn}"

    # 再次處理換行
    if '姓名' in res_cmd.columns: res_cmd['姓名'] = res_cmd['姓名'].apply(format_staff_only)
    if '服勤人員' in res_ptl.columns: res_ptl['服勤人員'] = res_ptl['服勤人員'].apply(format_staff_only)

    if save_data(p_time, cmdr_input, res_cmd, res_ptl):
        ok, mail_err = send_report_email(p_time, cmdr_input, res_cmd, res_ptl, final_filename)
        if ok: st.success(f"✅ 同步成功，郵件「{final_filename}」已寄送！")
        else: st.error(f"⚠️ 同步成功，但郵件失敗：{mail_err}")
        
        pdf_data = generate_pdf_from_data(p_time, cmdr_input, res_cmd, res_ptl)
        st.download_button("📥 下載 PDF", data=pdf_data, file_name=f"{final_filename}.pdf")
