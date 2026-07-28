import streamlit as st
from menu import show_sidebar
show_sidebar()
import io
import zipfile
from pdf2image import convert_from_bytes
from pptx import Presentation

# 設置頁面
st.set_page_config(page_title="PDF 轉檔中心", page_icon="📂")
st.header("📂 PDF 轉檔中心")

uploaded_file = st.file_uploader("請上傳 PDF 檔案", type=["pdf"])

if uploaded_file:
    # 新增：讓使用者選擇要轉換的格式
    option = st.radio("請選擇轉換格式：", ("轉成 PPTX", "轉成圖片 (ZIP壓縮檔)"))

    if st.button(f"🚀 開始{option}"):
        with st.spinner("正在解析 PDF 並處理中..."):
            try:
                # 1. 將 PDF 轉為圖片 (共通步驟)
                file_bytes = uploaded_file.read()
                images = convert_from_bytes(file_bytes, dpi=150)
                
                if images:
                    if option == "轉成 PPTX":
                        # 2A. 建立 PPTX
                        prs = Presentation()
                        
                        # 根據第一張圖設定 PPT 頁面尺寸
                        w, h = images[0].size
                        prs.slide_width = int(w * 914400 / 150)
                        prs.slide_height = int(h * 914400 / 150)
                        
                        for img in images:
                            slide = prs.slides.add_slide(prs.slide_layouts[6]) # 使用空白佈局
                            
                            img_stream = io.BytesIO()
                            img.save(img_stream, format='JPEG', quality=85) # 優化品質
                            img_stream.seek(0)
                            
                            slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
                        
                        # 3A. 輸出 PPTX 檔案
                        pptx_out = io.BytesIO()
                        prs.save(pptx_out)
                        
                        st.success(f"✅ PPTX 轉換成功！共處理 {len(images)} 頁")
                        st.download_button(
                            label="📥 點擊下載 PPTX",
                            data=pptx_out.getvalue(),
                            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                        
                    elif option == "轉成圖片 (ZIP壓縮檔)":
                        # 2B. 建立 ZIP 壓縮檔來存放多張圖片
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                            for i, img in enumerate(images):
                                img_stream = io.BytesIO()
                                img.save(img_stream, format='JPEG', quality=100) # 圖片模式下保留較高畫質
                                # 將每一頁圖片加入壓縮檔中，命名為 page_1.jpg, page_2.jpg...
                                zip_file.writestr(f"page_{i+1}.jpg", img_stream.getvalue())
                        
                        # 3B. 輸出 ZIP 檔案
                        st.success(f"✅ 圖片轉換成功！共打包 {len(images)} 張圖片")
                        st.download_button(
                            label="📥 點擊下載圖片 ZIP",
                            data=zip_buffer.getvalue(),
                            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}_images.zip",
                            mime="application/zip"
                        )
                else:
                    st.warning("⚠️ 此 PDF 沒有可提取的頁面。")
                    
            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")
                st.info("提示：如果是部署在 Linux 環境，請確認是否已安裝 poppler-utils。")
