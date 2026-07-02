import streamlit as st
import pandas as pd
import io
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# 必須是第一個 Streamlit 指令
st.set_page_config(page_title="綜合勤務規劃系統", layout="wide", page_icon="🚓")

# ==========================================
# 1. 泛用型 PDF 輸出引擎
# ==========================================
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames(): return fname
    font_paths = ["kaiu.ttf", "./kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf", "/usr/share/fonts/truetype/custom/kaiu.ttf"]
    for p in font_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fname, p))
            return fname
    return "Helvetica"

def generate_dynamic_pdf(duty_name, focus_text, df):
    font = _get_font()
    buf = io.BytesIO()
    
    margin_lr = 12 * mm
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=margin_lr, rightMargin=margin_lr, topMargin=15*mm, bottomMargin=15*mm)
    page_width = A4[0] - (2 * margin_lr)
    story = []

    # 定義字體樣式
    style_title = ParagraphStyle('Title', fontName=font, fontSize=18, leading=26, alignment=1, spaceAfter=12, wordWrap='CJK')
    style_section = ParagraphStyle('Section', fontName=font, fontSize=14, leading=20, spaceAfter=8, wordWrap='CJK')
    style_text = ParagraphStyle('Text', fontName=font, fontSize=12, leading=18, wordWrap='CJK')
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=12, leading=16, alignment=1, wordWrap='CJK')
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=12, leading=16, alignment=0, wordWrap='CJK')

    # 標題與重點
    story.append(Paragraph(f"<b>桃園市政府警察局龍潭分局 {duty_name} 勤務規劃表</b>", style_title))
    story.append(Paragraph("<b>壹、勤務重點：</b>", style_section))
    story.append(Paragraph(focus_text.replace("\n", "<br/>"), style_text))
    story.append(Spacer(1, 6*mm))

    # 動態表格生成
    story.append(Paragraph("<b>貳、勤務執行編組：</b>", style_section))
    
    clean_df = df.dropna(how="all").fillna("")
    headers = clean_df.columns.tolist()
    
    # 計算欄寬邏輯
    fixed_widths = {
        "組別": page_width * 0.08,
        "無線電代號": page_width * 0.12,
        "派遣單位": page_width * 0.12,
        "姓名": page_width * 0.12,
        "任務分工": page_width * 0.12,
    }
    
    col_widths = []
    used_width = 0
    dynamic_cols_count = 0
    
    for h in headers:
        if h in fixed_widths:
            used_width += fixed_widths[h]
        else:
            dynamic_cols_count += 1
            
    remaining_width = page_width - used_width
    dynamic_width = remaining_width / dynamic_cols_count if dynamic_cols_count > 0 else 0
    
    for h in headers:
        col_widths.append(fixed_widths.get(h, dynamic_width))

    # 組裝表格資料
    table_data = [[Paragraph(f"<b>{h}</b>", style_cell) for h in headers]]
    
    for _, row in clean_df.iterrows():
        row_data = []
        for h in headers:
            val = str(row.get(h, "")).strip().replace("\n", "<br/>")
            cell_style = style_cell_left if h not in ["組別", "無線電代號", "派遣單位", "姓名"] else style_cell
            row_data.append(Paragraph(val, cell_style))
        table_data.append(row_data)

    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f2f2f2')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t)

    def add_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 10)
        canvas.drawCentredString(A4[0] / 2.0, 10 * mm, f"- 第 {canvas.getPageNumber()} 頁 -")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    return buf.getvalue()

# ==========================================
# 2. 勤務版型設定檔 (Configuration Dictionary)
# ==========================================
DUTY_PROFILES = {
    "一般機動巡邏": {
        "extra_cols": ["攜行裝備", "巡邏路段"],
        "default_focus": "採取全面機動巡邏，針對重點路段加強攔檢；發現異常立即通報，注意自身三安。",
        "col_config": {
            "巡邏路段": st.column_config.TextColumn("巡邏路段", width="large"),
            "攜行裝備": st.column_config.TextColumn("攜行裝備", width="medium")
        }
    },
    "定點場所臨檢": {
        "extra_cols": ["臨檢目標場所"],
        "default_focus": "會合專案人員，準時進入目標場所執行威力掃蕩，全程開啟密錄器蒐證。",
        "col_config": {
            "臨檢目標場所": st.column_config.TextColumn("臨檢目標場所", width="large")
        }
    },
    "定點路檢": {
        "extra_cols": ["攜行裝備", "定點路檢目標"],
        "default_focus": "於各聯外道路、重要路口設立檢測點，加強取締違規。",
        "col_config": {
            "定點路檢目標": st.column_config.TextColumn("定點路檢目標", width="large"),
            "攜行裝備": st.column_config.TextColumn("攜行裝備", width="medium")
        }
    }
}

# ==========================================
# 3. 主畫面 UI 與動態表單
# ==========================================
st.title("🚓 綜合勤務規劃系統 (動態模組化測試)")
st.info("💡 只要切換下方的勤務類型，系統的「勤務重點」、「表格欄位」以及最終的「PDF排版」都會自動跟著變形。")

# 選擇器
selected_duty = st.selectbox("📌 請選擇本次規劃的勤務類型", list(DUTY_PROFILES.keys()))
profile = DUTY_PROFILES[selected_duty]

st.markdown("---")
st.subheader(f"【{selected_duty}】編組設定")

# 動態勤務重點
focus_text = st.text_area("📢 勤務重點", value=profile["default_focus"], height=80)

# 準備基礎與動態欄位
base_cols = ["組別", "無線電代號", "派遣單位", "姓名", "任務分工"]
dynamic_cols = base_cols + profile["extra_cols"]

# 產生預設假資料
default_units = ["石門", "中興", "聖亭", "龍潭", "高平", "三和", "交通分隊"]
default_data = []
for i, unit in enumerate(default_units[:3]): 
    row = {
        "組別": f"第{i+1}組",
        "無線電代號": f"隆安{i+1}1" if i == 0 else f"隆安{i+1}0",
        "派遣單位": unit,
        "姓名": "",
        "任務分工": "帶班" if i == 0 else "攔檢盤查"
    }
    for col in profile["extra_cols"]:
        row[col] = ""
    default_data.append(row)

df_template = pd.DataFrame(default_data, columns=dynamic_cols)

# 欄位寬度設定合併
base_config = {
    "組別": st.column_config.TextColumn("組別", width="small"),
    "無線電代號": st.column_config.TextColumn("無線電代號", width="small"),
    "派遣單位": st.column_config.TextColumn("派遣單位", width="small"),
    "姓名": st.column_config.TextColumn("姓名", width="small"),
    "任務分工": st.column_config.TextColumn("任務分工", width="medium"),
}
final_col_config = {**base_config, **profile["col_config"]}

# 渲染動態 Data Editor
st.caption("💡 下方表格已自動變換欄位，編輯完畢後請點擊下方按鈕產出 PDF：")
res_df = st.data_editor(
    df_template,
    num_rows="dynamic",
    use_container_width=True,
    column_config=final_col_config,
    key=f"editor_{selected_duty}" 
)

# ==========================================
# 4. 觸發按鈕與 PDF 下載
# ==========================================
st.markdown("---")
if st.button("📄 生成動態 PDF 報表", type="primary"):
    with st.spinner("正在計算欄寬並產生 PDF..."):
        
        pdf_bytes = generate_dynamic_pdf(selected_duty, focus_text, res_df)
        
        st.success("✅ PDF 產生成功！請點擊下方按鈕下載。")
        st.download_button(
            label="⬇️ 點擊下載 PDF 規劃表",
            data=pdf_bytes,
            file_name=f"{selected_duty}規劃表.pdf",
            mime="application/pdf"
        )
