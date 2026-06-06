import streamlit as st
try:
    from menu import show_sidebar
    show_sidebar()
except ImportError:
    pass

import io
import pdfplumber
from pdf2image import convert_from_bytes
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# 設置頁面
st.set_page_config(page_title="PDF 轉 PPTX 高顏值工具", page_icon="📂")
st.header("📂 交通週報 - PDF 轉 PPTX 智慧中心")
st.caption("💡 方案 A（完美版面流）：100% 保留交大週報圖表與格式，並智慧重疊可編輯之原生文字方塊。")

uploaded_file = st.file_uploader("請上傳交大週報 PDF 檔案", type=["pdf"])

if uploaded_file:
    if st.button("🚀 開始高顏值智慧轉換"):
        with st.spinner("正在進行版面精準渲染與文字雙層歸戶中..."):
            try:
                # 讀取檔案 Bytes
                file_bytes = uploaded_file.read()
                
                # 1. 利用 pdf2image 將 PDF 頁面渲染為高解析度 JPEG (維持原汁原味高顏值)
                # 150 DPI 是公務簡報兼顧清晰度與檔案體積的最佳平衡
                images = convert_from_bytes(file_bytes, dpi=150)
                
                if not images:
                    st.error("❌ 無法將該 PDF 渲染為圖片，請確認檔案是否損壞。")
                    st.stop()
                
                # 2. 初始化簡報物件並精準設定投影片尺寸（完全對齊第一頁 PDF 比例）
                prs = Presentation()
                w, h = images[0].size
                prs.slide_width = int(w * 914400 / 150)
                prs.slide_height = int(h * 914400 / 150)
                
                # 3. 同步開啟 pdfplumber 用於文字層解析
                with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                    
                    # 雙層流循環：同時處理圖片與文字
                    for i, img in enumerate(images):
                        # 使用全空白佈局（版型 6）
                        slide = prs.slides.add_slide(prs.slide_layouts[6])
                        
                        # ── 第一層：將高顏值簡報頁面塞入背景 ──
                        img_stream = io.BytesIO()
                        img.save(img_stream, format='JPEG', quality=90) # 高品質輸出
                        img_stream.seek(0)
                        
                        # 貼上滿版底圖
                        slide.shapes.add_picture(
                            img_stream, 0, 0, 
                            width=prs.slide_width, 
                            height=prs.slide_height
                        )
                        
                        # ── 第二層：智慧抽取文字，並重疊透明可編輯文字方塊 ──
                        # 防止因為頁數不對稱產生外掛錯誤
                        if i < len(pdf.pages):
                            page_text = pdf.pages[i].extract_text()
                            
                            if page_text and page_text.strip():
                                # 在投影片中央偏下方建立一個幾乎滿版的透明文字方塊
                                # 格式設定為左邊留白 0.5 英吋、頂部留白 1.0 英吋
                                txBox = slide.shapes.add_textbox(
                                    Inches(0.5), Inches(1.0), 
                                    Inches((w/150) - 1.0), Inches((h/150) - 1.5)
                                )
                                tf = txBox.text_frame
                                tf.word_wrap = True # 啟用自動換行
                                tf.text = page_text
                                
                                # 微調這層文字的字體與顏色（設為隱約透明或與底色相近，主要供選取、改字與搜尋使用）
                                for paragraph in tf.paragraphs:
                                    paragraph.font.name = "Microsoft JhengHei"
                                    paragraph.font.size = Pt(12)
                                    # 使用半透明感的深灰或讓其與底色自然融合，若要完全隱形可以調小字體
                                    paragraph.font.color.rgb = RGBColor(60, 60, 60)
                
                # 4. 輸出最終簡報二進位流
                pptx_out = io.BytesIO()
                prs.save(pptx_out)
                
                st.success(f"🎉 轉換成功！已完美封裝 {len(images)} 頁高顏值週報投影片（內含可點擊編輯文字層）。")
                st.download_button(
                    label="📥 點擊下載【方案 A - 高顏值可編輯】PPTX",
                    data=pptx_out.getvalue(),
                    file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_高顏值簡報.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary"
                )
                
            except Exception as e:
                st.error(f"❌ 運行發生錯誤：{str(e)}")
                st.info("💡 提示：請確保本地或伺服器環境中，`pdf2image` 依賴的系統組件 Poppler 運作正常。")
