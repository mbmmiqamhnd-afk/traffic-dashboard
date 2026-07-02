import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import io, os, re, smtplib, calendar
from datetime import datetime, timedelta
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

st.set_page_config(page_title="綜合勤務規劃總署", layout="wide", page_icon="🚓")

# 嘗試載入側邊欄
try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    pass

# ==========================================
# 1. 核心設定與預設資料庫 (DUTY_PROFILES)
# ==========================================
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
UNIT = "桃園市政府警察局龍潭分局"

# --- 預設指揮組 ---
DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"}
])

# --- 綜合勤務配置字典 (新增任何專案只需在這裡加一筆) ---
DUTY_PROFILES = {
    "聯合稽查 (單階段)": {
        "sheet_prefix": "統整_聯合",
        "meta_fields": ["勤務時間", "勤前教育", "環保局臨時檢驗站"],
        "tabs": ["指揮編組", "巡邏編組"],
        "default_dfs": {
            "指揮編組": DEFAULT_CMD.copy(),
            "巡邏編組": pd.DataFrame([
                {"編組": "第一巡邏組", "代號": "隆安52", "單位": "聖亭所", "職別": "副所長", "姓名": "曹培翔", "任務分工": "帶班兼蒐證", "巡邏路段": "於中正路周邊易有噪音車輛滋擾路段機動巡查。"},
                {"編組": "第一巡邏組", "代號": "隆安52", "單位": "聖亭所", "職別": "警員", "姓名": "詹宗澤", "任務分工": "攔檢盤查", "巡邏路段": "於中正路周邊易有噪音車輛滋擾路段機動巡查。"}
            ])
        },
        "business_logic": "auto_radio"
    },
    "防制危險駕車 (時段分佈)": {
        "sheet_prefix": "統整_危駕",
        "meta_fields": ["勤務時間", "交通快打指揮官", "巡簽地點", "備註"],
        "tabs": ["指揮編組", "警力佈署"],
        "default_dfs": {
            "指揮編組": DEFAULT_CMD.copy(),
            "警力佈署": pd.DataFrame([
                {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安80", "編組": "石門", "服勤人員": "線上警力兼任", "巡邏路段": "「區域聯防」勤務，於中正路等路段巡邏"},
                {"勤務時段": "5月22日\n22時至翌日6時", "代號": "隆安90", "編組": "高平", "服勤人員": "線上警力兼任", "巡邏路段": "「區域聯防」勤務，於中豐路等路段巡邏"},
            ])
        },
        "business_logic": "shift_date"
    },
    "行人及護老專案 (月份排班)": {
        "sheet_prefix": "統整_護老",
        "meta_fields": ["月份", "放假日", "備註"],
        "tabs": ["指揮編組", "警力佈署"],
        "default_dfs": {
            "指揮編組": DEFAULT_CMD.copy(),
            "警力佈署": pd.DataFrame([
                {"日期": "", "單位": "聖亭派出所", "路段": "中豐路、聖亭路段\n轄區行人易肇事路口"},
                {"日期": "", "單位": "石門派出所", "路段": "中正、文化路段\n轄區行人易肇事路口"}
            ])
        },
        "business_logic": "monthly_calendar"
    },
    "三階段專案 (多模組)": {
        "sheet_prefix": "統整_三階",
        "meta_fields": ["勤務時間", "勤前教育"],
        "tabs": ["指揮編組", "一階機動", "二階臨檢", "三階路檢"],
        "default_dfs": {
            "指揮編組": DEFAULT_CMD.copy(),
            "一階機動": pd.DataFrame([{"組別": "第1機動組", "代號": "隆安51", "單位": "龍潭所", "職別": "所長", "姓名": "", "任務分工": "帶班", "攜行裝備": "槍彈", "機動攔檢區域": "龍潭市區"}]),
            "二階臨檢": pd.DataFrame([{"組別": "第1臨檢組", "代號": "隆安51", "單位": "龍潭所", "職別": "所長", "姓名": "", "任務分工": "帶班", "臨檢目標場所": "A. 鉅大撞球館"}]),
            "三階路檢": pd.DataFrame([{"組別": "第1路檢組", "代號": "隆安51", "單位": "龍潭所", "職別": "所長", "姓名": "", "任務分工": "管制", "攜行裝備": "酒測器", "定點路檢目標": "北龍路319號前"}])
        },
        "business_logic": "none"
    }
}

# ==========================================
# 2. 泛用型字體與 PDF 引擎 (Universal Engine)
# ==========================================
@st.cache_resource
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    for p in ["./kaiu.ttf", "kaiu.ttf", "/usr/share/fonts/truetype/custom/kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_universal_pdf(duty_name, project_name, meta_dict, dfs_dict):
    """
    終極 PDF 產出引擎：無論傳入幾個 DataFrame、長什麼形狀，都會自動計算欄寬、自動繪製表格，並自動合併相同屬性的列。
    """
    font = _get_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=12*mm, rightMargin=12*mm, topMargin=12*mm, bottomMargin=15*mm)
    W = A4[0] - 24*mm
    story = []

    s_title = ParagraphStyle("title", fontName=font, fontSize=16, alignment=1, leading=24, spaceAfter=8)
    s_sec   = ParagraphStyle("sec", fontName=font, fontSize=14, leading=20, spaceAfter=4, spaceBefore=8)
    s_txt   = ParagraphStyle("txt", fontName=font, fontSize=12, leading=18)
    s_cell  = ParagraphStyle("cell", fontName=font, fontSize=12, alignment=1, leading=16)
    s_left  = ParagraphStyle("left", fontName=font, fontSize=12, alignment=0, leading=16)

    def c(txt, style=s_cell): return Paragraph(str(txt).replace("\n", "<br/>"), style)

    # --- 檔頭資訊 ---
    story.append(Paragraph(f"<b>{UNIT}執行「{project_name}」勤務規劃表</b>", s_title))
    
    # 輸出上半部 Meta 資訊 (時間、勤教等)
    for k, v in meta_dict.items():
        if k not in ["巡簽地點", "備註"] and str(v).strip():
            story.append(Paragraph(f"<b>{k}：</b>", s_sec))
            story.append(Paragraph(str(v).replace("\n", "<br/>"), s_txt))
            
    # --- 動態渲染各階段表格 ---
    for tab_name, df in dfs_dict.items():
        clean_df = df.dropna(how="all").fillna("")
        if clean_df.empty: continue
            
        story.append(Paragraph(f"<b>【{tab_name}】</b>", s_sec))
        headers = clean_df.columns.tolist()
        
        # 智慧計算欄寬：依據欄位名稱特徵賦予權重
        col_widths = []
        for h in headers:
            if any(x in h for x in ["姓名", "人員", "裝備"]): col_widths.append(W * 0.12)
            elif "代號" in h: col_widths.append(W * 0.09)
            elif any(x in h for x in ["單位", "職別", "職稱"]): col_widths.append(W * 0.10)
            elif any(x in h for x in ["時間", "時段", "日期"]): col_widths.append(W * 0.18)
            elif "組別" in h or "編組" in h: col_widths.append(W * 0.11)
            else: col_widths.append(0) # 剩下的長文字欄位平分剩餘空間
            
        rem = W - sum(col_widths)
        zeros = col_widths.count(0)
        if zeros > 0:
            for i in range(len(col_widths)):
                if col_widths[i] == 0: col_widths[i] = rem / zeros
                
        # 組裝表格資料
        data = [[c(f"<b>{h}</b>") for h in headers]]
        for _, row in clean_df.iterrows():
            row_data = []
            for h in headers:
                val = str(row.get(h, ""))
                style = s_left if any(x in h for x in ["任務", "路段", "目標", "區域", "人員"]) else s_cell
                row_data.append(c(val, style))
            data.append(row_data)
            
        # 表格樣式與動態合併邏輯 (Span)
        ts = [("FONTNAME", (0,0), (-1,-1), font), ("GRID", (0,0), (-1,-1), 0.5, colors.black),
              ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#f2f2f2")), ("VALIGN", (0,0), (-1,-1), "MIDDLE")]
        
        # 若第一欄是識別性的群組欄位（如組別、時段），進行相鄰相同值的垂直合併
        if headers[0] in ["組別", "編組", "勤務時段", "日期"]:
            i = 1
            while i <= len(clean_df):
                j = i + 1
                while j <= len(clean_df) and str(clean_df.iloc[i-1, 0]) == str(clean_df.iloc[j-1, 0]) and str(clean_df.iloc[i-1, 0]).strip() != "":
                    j += 1
                if j - i > 1:
                    ts.append(("SPAN", (0, i), (0, j-1)))
                    # 順便合併代號欄位 (若存在)
                    if len(headers) > 1 and headers[1] in ["代號", "無線電代號", "單位"]:
                        ts.append(("SPAN", (1, i), (1, j-1)))
                i = j
                
        t = Table(data, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle(ts))
        story.append(t)
        story.append(Spacer(1, 4*mm))

    # 輸出下半部 Meta 資訊 (備註等)
    for k, v in meta_dict.items():
        if k in ["巡簽地點", "備註"] and str(v).strip():
            story.append(Paragraph(f"<b>{k}：</b>", s_sec))
            for line in str(v).split("\n"):
                if line.strip(): story.append(Paragraph(line, s_txt))

    def add_page_number(canvas, doc):
    canvas.saveState()
    canvas.setFont(font, 10)
    canvas.drawCentredString(
        A4[0] / 2,
        8 * mm,
        f"- 第 {canvas.getPageNumber()} 頁 -"
    )
    canvas.restoreState()

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    buf.seek(0)
    return buf.getvalue()


# ==========================================
# 3. 特定業務邏輯輔助函數
# ==========================================
def parse_monthly_workdays(month_str, holiday_str):
    """護老專案專用：產生上班日清單文字"""
    year_match = re.search(r"(\d+)年", month_str)
    mon_match  = re.search(r"(\d+)月", month_str)
    if not year_match or not mon_match: return "請確認月份格式 (例: 115年3月份)"
    
    ce_year = int(year_match.group(1)) + 1911
    month = int(mon_match.group(1))
    
    holidays = set()
    for part in holiday_str.replace("，", ",").split(","):
        m = re.match(r"(\d+)[/月](\d+)", part.strip())
        if m:
            try: holidays.add(datetime(ce_year, int(m.group(1)), int(m.group(2))).date())
            except ValueError: pass
            
    _, days_in_month = calendar.monthrange(ce_year, month)
    workdays = []
    for day in range(1, days_in_month + 1):
        d = datetime(ce_year, month, day)
        if d.weekday() < 5 and d.date() not in holidays:
            workdays.append(d)
            
    if not workdays: return "（本月無上班日）"
    
    parts = []
    for i, d in enumerate(workdays):
        wd = ["一", "二", "三", "四", "五", "六", "日"][d.weekday()]
        if i == 0: parts.append(f"{month}月{d.day}日(星期{wd})")
        else: parts.append(f"{d.day}日({wd})")
    return "、".join(parts)


# ==========================================
# 4. 主介面 (Dynamic Router)
# ==========================================
st.title("🚓 綜合勤務規劃系統")

# --- 左側選單：切換勤務模組 ---
selected_duty = st.sidebar.selectbox("📌 選擇專案勤務類型", list(DUTY_PROFILES.keys()))
profile = DUTY_PROFILES[selected_duty]
biz_logic = profile["business_logic"]

st.info(f"💡 目前載入模組：**{selected_duty}**。下方欄位與表格已自動變換為該專案專屬架構。")

col1, col2 = st.columns([1, 2])
with col1:
    project_name = st.text_input("專案名稱", value=f"預設{selected_duty.split(' ')[0]}專案")

# --- 動態渲染文字欄位 (Meta Fields) ---
meta_inputs = {}
for field in profile["meta_fields"]:
    if field in ["備註", "巡簽地點", "勤前教育"]:
        meta_inputs[field] = st.text_area(field, height=80)
    else:
        meta_inputs[field] = st.text_input(field, value=f"預設{field}")

# --- 注入特定業務邏輯 (Business Logics) ---
if biz_logic == "monthly_calendar":
    c_month = meta_inputs.get("月份", "115年3月份")
    c_holi = meta_inputs.get("放假日", "3/3, 3/10")
    workday_label = parse_monthly_workdays(c_month, c_holi)
    st.caption(f"📅 系統自動推算上班日：{workday_label}")
    # 動態覆寫護老專案的第一個日期欄位
    for k in profile["default_dfs"].keys():
        if "佈署" in k and "日期" in profile["default_dfs"][k].columns:
            profile["default_dfs"][k]["日期"] = ""
            profile["default_dfs"][k].loc[0, "日期"] = workday_label

if biz_logic == "shift_date":
    fast_cmd = meta_inputs.get("交通快打指揮官", "")
    time_str = meta_inputs.get("勤務時間", "")
    st.caption(f"⚙️ 系統將依據指揮官單位及時間，自動微調下方【專責警力】之排班時段與代號。")

# --- 動態渲染資料表格 (Tabs) ---
st.markdown("---")
st.subheader("勤務編組與佈署設定")

tabs = st.tabs(profile["tabs"])
result_dfs = {}

for tab_obj, tab_name in zip(tabs, profile["tabs"]):
    with tab_obj:
        st.caption(f"編輯【{tab_name}】資料：")
        default_df = profile["default_dfs"][tab_name]
        
        # 建立 Data Editor
        edited_df = st.data_editor(
            default_df, 
            num_rows="dynamic", 
            use_container_width=True,
            key=f"editor_{selected_duty}_{tab_name}"
        )
        result_dfs[tab_name] = edited_df

# --- 產出 PDF 報表 ---
st.markdown("---")
if st.button("📄 產生綜合勤務規劃 PDF", type="primary", use_container_width=True):
    with st.spinner("啟動泛用型 PDF 引擎進行排版計算中..."):
        try:
            pdf_bytes = generate_universal_pdf(selected_duty, project_name, meta_inputs, result_dfs)
            
            st.success("✅ PDF 產生成功！")
            st.download_button(
                label="⬇️ 點擊下載 PDF 報表",
                data=pdf_bytes,
                file_name=f"{project_name}規劃表.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"產出 PDF 時發生錯誤: {str(e)}")
