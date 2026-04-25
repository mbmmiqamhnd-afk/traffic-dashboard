import streamlit as st
from menu import show_sidebar
show_sidebar()
import io
from pdf2image import convert_from_bytes
from pptx import Presentation
from pptx.util import Inches

# 設置頁面
st.set_page_config(page_title="PDF 轉 PPTX 工具", page_icon="📂")
st.header("📂 PDF 轉檔中心")

uploaded_file = st.file_uploader("請上傳 PDF 檔案", type=["pdf"])

if uploaded_file:
    if st.button("🚀 開始轉換為 PPTX"):
        with st.spinner("正在解析 PDF 並生成投影片..."):
            try:
                # 1. 將 PDF 轉為圖片 (調整 DPI 兼顧清晰度與檔案大小)
                # 這裡不需要 read() 兩次，直接傳入 bytes
                file_bytes = uploaded_file.read()
                images = convert_from_bytes(file_bytes, dpi=150)
                
                # 2. 建立 PPTX
                prs = Presentation()
                
                if images:
                    # 根據第一張圖設定 PPT 頁面尺寸
                    w, h = images[0].size
                    # 914400 是 PPT 內部單位 (EMU)，1 英吋 = 914400 EMU
                    prs.slide_width = int(w * 914400 / 150)
                    prs.slide_height = int(h * 914400 / 150)
                    
                    for img in images:
                        slide = prs.slides.add_slide(prs.slide_layouts[6]) # 使用空白佈局
                        
                        img_stream = io.BytesIO()
                        img.save(img_stream, format='JPEG', quality=85) # 優化品質
                        img_stream.seek(0)
                        
                        slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
                    
                    # 3. 輸出檔案
                    pptx_out = io.BytesIO()
                    prs.save(pptx_out)
                    
                    st.success(f"✅ 轉換成功！共處理 {len(images)} 頁")
                    st.download_button(
                        label="📥 點擊下載 PPTX",
                        data=pptx_out.getvalue(),
                        file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                
            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")
                st.info("提示：如果是部署在 Linux 環境，請確認是否已安裝 poppler-utils。")
