import streamlit as st
from menu import show_sidebar
show_sidebar()
import io
import os
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import white, black
from PIL import Image, ImageDraw, ImageFont

# --- 設定網頁標題 ---
st.set_page_config(page_title="商標頁碼工具", page_icon="📝")
st.header("📝 檔案商標遮蓋與頁碼工具")

# --- 自動偵測字型 ---
def get_font_path():
    possible_paths = ["kaiu.ttf", "font.ttf", "pages/kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]
    for p in possible_paths:
        if os.path.exists(p):
            return p
    return None

font_path = get_font_path()

# --- PDF 遮蓋邏輯 ---
def create_pdf_overlay(page_width, page_height, page_num, current_font):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    text = f"交通組製 - 第 {page_num} 頁"
    box_width, box_height = 130, 20
    rect_x, rect_y = page_width - box_width, 0
    
    c.setFillColor(white)
    c.rect(rect_x, rect_y, box_width, box_height, fill=1, stroke=0)
    c.setFillColor(black)
    c.setFont(current_font, 14)
    c.drawRightString(page_width - 4, 4, text)
    c.save()
    packet.seek(0)
    return packet

# --- 圖片處理邏輯 ---
def process_image(image_file, font_p):
    img = Image.open(image_file).convert("RGB")
    draw = ImageDraw.Draw(img)
    width, height = img.size
    
    # 定義遮蓋框大小 (圖片像素通常較多，這裡設為寬度的 20%)
    box_w, box_h = 250, 50 
    rect_x0, rect_y0 = width - box_w, height - box_h
    
    # 畫白色遮蓋矩形
    draw.rectangle([rect_x0, rect_y0, width, height], fill="white")
    
    # 載入字型
    try:
        font = ImageFont.truetype(font_p, 24) if font_p else ImageFont.load_default()
    except:
        font = ImageFont.load_default()
        
    text = "交通組製"
    # 簡單置中文字在白框內
    draw.text((rect_x0 + 10, rect_y0 + 10), text, fill="black", font=font)
    
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='JPEG')
    return img_byte_arr.getvalue()

# --- 主介面 ---
uploaded_file = st.file_uploader("上傳 PDF 或 圖片 (JPG/PNG)", type=["pdf", "jpg", "jpeg", "png"])

if uploaded_file and st.button("開始加工"):
    file_ext = uploaded_file.name.split('.')[-1].lower()
    
    try:
        # 處理 PDF
        if file_ext == "pdf":
            reader = PdfReader(uploaded_file)
            writer = PdfWriter()
            
            # 在處理 PDF 頁面「前」，先註冊好字型
            pdf_font = "Helvetica" # 預設字型
            if font_path:
                try:
                    pdfmetrics.registerFont(TTFont('CustomFont', font_path))
                    pdf_font = 'CustomFont'
                except Exception as font_e:
                    st.warning(f"字體載入失敗，改用預設字體 (錯誤: {font_e})")

            # 開始處理每一頁
            for i, page in enumerate(reader.pages):
                w, h = float(page.mediabox.width), float(page.mediabox.height)
                overlay = create_pdf_overlay(w, h, i+1, pdf_font) 
                
                page.merge_page(PdfReader(overlay).pages[0])
                writer.add_page(page)
            
            out = io.BytesIO()
            writer.write(out)
            st.success("🎉 PDF 加工完成！")
            st.download_button("📥 下載加工版 PDF", out.getvalue(), "加工版.pdf", "application/pdf")

        # 處理圖片
        else:
            result_img = process_image(uploaded_file, font_path)
            st.image(result_img, caption="預覽加工後的圖片")
            st.success("🎉 圖片加工完成！")
            st.download_button("📥 下載加工版圖片", result_img, f"processed_{uploaded_file.name}", f"image/{file_ext}")

    except Exception as e:
        st.error(f"發生錯誤: {e}")
