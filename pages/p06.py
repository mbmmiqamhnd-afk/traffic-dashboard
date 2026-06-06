import streamlit as st
try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    pass

import io
import pdfplumber
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# 設置頁面
st.set_page_config(page_title="PDF 原生轉 PPTX 工具", page_icon="📂")
st.header("📂 智慧型 PDF 轉檔中心 (可編輯版)")
st.caption("💡 說明：本工具使用內建的 pdfplumber 解析技術，轉換後的文字與表格皆可在 PowerPoint 中直接點擊修改。")

uploaded_file = st.file_uploader("請上傳 PDF 檔案", type=["pdf"])

if uploaded_file:
    if st.button("🚀 開始智慧轉換為可編輯 PPTX"):
        with st.spinner("正在深度解析 PDF 欄位結構並重建原生投影片..."):
            try:
                # 1. 讀取上傳檔案的 Bytes 並用 pdfplumber 開啟
                pdf_bytes = uploaded_file.read()
                
                # 2. 初始化 python-pptx 簡報物件
                prs = Presentation()
                # 設定標準 16:9 寬螢幕簡報尺寸
                prs.slide_width = Inches(13.333)
                prs.slide_height = Inches(7.5)
                
                # 3. 逐頁解析 PDF
                with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                    for page_num, page in enumerate(pdf.pages):
                        
                        # --- 核心邏輯 A：檢查是否有表格 ---
                        tables = page.extract_tables()
                        
                        if tables:
                            for table in tables:
                                # 建立空白投影片 (版型 6 為全空白)
                                slide = prs.slides.add_slide(prs.slide_layouts[6])
                                
                                # 過濾掉完全為 None 的空列或空欄
                                cleaned_table = [[str(cell) if cell is not None else "" for cell in row] for row in table]
                                
                                num_rows = len(cleaned_table)
                                num_cols = len(cleaned_table[0]) if num_rows > 0 else 0
                                
                                if num_rows == 0 or num_cols == 0:
                                    continue
                                    
                                # 在 PPT 中建立真正的「原生可編輯表格物件」
                                left = Inches(0.5)
                                top = Inches(1.0)
                                width = Inches(12.333)
                                height = Inches(5.5)
                                
                                ppt_table_shape = slide.shapes.add_table(num_rows, num_cols, left, top, width, height)
                                ppt_table = ppt_table_shape.table
                                
                                # 將 PDF 表格數據填入 PPT 原生表格中
                                for r_idx, row in enumerate(cleaned_table):
                                    for c_idx, cell_value in enumerate(row):
                                        ppt_cell = ppt_table.cell(r_idx, c_idx)
                                        ppt_cell.text = cell_value
                                        
                                        # 表格樣式微調（防止公務報表字型太大爆格）
                                        for paragraph in ppt_cell.text_frame.paragraphs:
                                            paragraph.font.size = Pt(11)
                                            paragraph.font.name = "Microsoft JhengHei"  # 微軟正黑體
                                            paragraph.alignment = PP_ALIGN.CENTER       # 文字置中
                        
                        # --- 核心邏輯 B：檢查是否有純文字段落 ---
                        text = page.extract_text()
                        if text and not tables:  # 如果這頁是純文字報告（無表格）
                            slide = prs.slides.add_slide(prs.slide_layouts[5])  # 使用帶標題版型
                            slide.shapes.title.text = f"資料報告 - 第 {page_num + 1} 頁"
                            
                            # 建立一個大文字方塊
                            txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11.333), Inches(4.5))
                            tf = txBox.text_frame
                            tf.word_wrap = True  # 自動換行
                            
                            # 填入文字
                            tf.text = text
                            
                            # 微調文字方塊內的所有段落字體
                            for paragraph in tf.paragraphs:
                                paragraph.font.size = Pt(14)
                                paragraph.font.name = "Microsoft JhengHei"
                
                # 4. 如果整份 PDF 什麼都沒抓到（例如純掃描圖片 PDF）
                if len(prs.slides) == 0:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    txBox = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(11.333), Inches(1))
                    txBox.text_frame.text = "【系統提示】未偵測到內嵌的文字或結構化表格。若此 PDF 是紙本掃描檔或純照片，請改用 OCR 功能處理。"

                # 5. 將 PPTX 儲存至記憶體提供下載
                pptx_out = io.BytesIO()
                prs.save(pptx_out)
                
                st.success(f"🎉 智慧結構重組成功！共生成 {len(prs.slides)} 頁原生投影片。")
                st.download_button(
                    label="📥 點擊下載【完全可編輯】PPTX",
                    data=pptx_out.getvalue(),
                    file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_可編輯版.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary"
                )
                
            except Exception as e:
                st.error(f"❌ 智慧解析失敗：{str(e)}")
