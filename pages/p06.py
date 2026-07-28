import streamlit as st
from menu import show_sidebar
# show_sidebar() # 假設 menu.py 會處理側邊欄，此處先註解或保留依你的架構而定
import io
import zipfile
from pdf2image import convert_from_bytes
from pptx import Presentation
from PIL import Image

# 設置頁面 (更新標題以涵蓋新功能)
st.set_page_config(page_title="綜合檔案轉檔中心", page_icon="🗂️")

# 這裡呼叫你的側邊欄函數
show_sidebar()

st.header("🗂️ 綜合檔案轉檔中心")
st.write("支援 PDF 轉 PPTX/圖片，以及圖片合併轉 PDF")

# 使用 Tabs 來區分不同的功能模組，讓介面更乾淨
tab1, tab2 = st.tabs(["📄 PDF 轉其他格式", "🖼️ 圖片轉 PDF"])

# ==========================================
# Tab 1: PDF 轉 PPTX 或 圖片
# ==========================================
with tab1:
    st.subheader("從 PDF 轉換")
    uploaded_pdf = st.file_uploader("請上傳 PDF 檔案", type=["pdf"], key="pdf_uploader")

    if uploaded_pdf:
        option = st.radio("請選擇轉換格式：", ("轉成 PPTX", "轉成圖片 (ZIP壓縮檔)"))

        if st.button(f"🚀 開始{option}"):
            with st.spinner("正在解析 PDF 並處理中..."):
                try:
                    # 1. 將 PDF 轉為圖片
                    file_bytes = uploaded_pdf.read()
                    images = convert_from_bytes(file_bytes, dpi=150)
                    
                    if images:
                        if option == "轉成 PPTX":
                            # 2A. 建立 PPTX
                            prs = Presentation()
                            
                            w, h = images[0].size
                            prs.slide_width = int(w * 914400 / 150)
                            prs.slide_height = int(h * 914400 / 150)
                            
                            for img in images:
                                slide = prs.slides.add_slide(prs.slide_layouts[6])
                                
                                img_stream = io.BytesIO()
                                img.save(img_stream, format='JPEG', quality=85)
                                img_stream.seek(0)
                                
                                slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
                            
                            # 3A. 輸出 PPTX
                            pptx_out = io.BytesIO()
                            prs.save(pptx_out)
                            
                            st.success(f"✅ PPTX 轉換成功！共處理 {len(images)} 頁")
                            st.download_button(
                                label="📥 點擊下載 PPTX",
                                data=pptx_out.getvalue(),
                                file_name=f"{uploaded_pdf.name.rsplit('.', 1)[0]}.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                            
                        elif option == "轉成圖片 (ZIP壓縮檔)":
                            # 2B. 建立 ZIP
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                                for i, img in enumerate(images):
                                    img_stream = io.BytesIO()
                                    img.save(img_stream, format='JPEG', quality=100)
                                    zip_file.writestr(f"page_{i+1}.jpg", img_stream.getvalue())
                            
                            # 3B. 輸出 ZIP
                            st.success(f"✅ 圖片轉換成功！共打包 {len(images)} 張圖片")
                            st.download_button(
                                label="📥 點擊下載圖片 ZIP",
                                data=zip_buffer.getvalue(),
                                file_name=f"{uploaded_pdf.name.rsplit('.', 1)[0]}_images.zip",
                                mime="application/zip"
                            )
                    else:
                        st.warning("⚠️ 此 PDF 沒有可提取的頁面。")
                        
                except Exception as e:
                    st.error(f"❌ 發生錯誤：{str(e)}")
                    st.info("提示：如果是部署在 Linux 環境，請確認是否已安裝 poppler-utils。")


# ==========================================
# Tab 2: 圖片轉 PDF
# ==========================================
with tab2:
    st.subheader("將多張圖片合併為單一 PDF")
    # 允許上傳多個檔案，並限定圖片格式
    uploaded_images = st.file_uploader(
        "請上傳圖片 (支援多選，按上傳順序排列)", 
        type=["jpg", "jpeg", "png", "bmp"], 
        accept_multiple_files=True,
        key="img_uploader"
    )

    if uploaded_images:
        st.info(f"已選擇 {len(uploaded_images)} 張圖片。")
        
        if st.button("🚀 開始合併為 PDF"):
            with st.spinner("正在處理圖片並建立 PDF..."):
                try:
                    image_list = []
                    
                    # 讀取並處理所有上傳的圖片
                    for uploaded_img in uploaded_images:
                        # 使用 PIL 開啟圖片
                        img = Image.open(uploaded_img)
                        # 確保圖片轉換為 RGB 模式 (PDF 不支援帶透明度的 RGBA)
                        if img.mode != 'RGB':
                            img = img.convert('RGB')
                        image_list.append(img)
                        
                    if image_list:
                        pdf_buffer = io.BytesIO()
                        # 取第一張圖片作為基礎，將後續圖片 append 進去
                        first_image = image_list[0]
                        remaining_images = image_list[1:] if len(image_list) > 1 else []
                        
                        # 儲存為 PDF
                        first_image.save(
                            pdf_buffer, 
                            format="PDF", 
                            save_all=True, 
                            append_images=remaining_images,
                            resolution=100.0
                        )
                        
                        st.success(f"✅ PDF 建立成功！共包含 {len(image_list)} 頁")
                        st.download_button(
                            label="📥 點擊下載合併後的 PDF",
                            data=pdf_buffer.getvalue(),
                            file_name="merged_images.pdf",
                            mime="application/pdf"
                        )
                except Exception as e:
                    st.error(f"❌ 發生錯誤：{str(e)}")
