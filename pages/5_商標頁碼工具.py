import streamlit as st
import io
import os
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import white, black
from PIL import Image, ImageDraw, ImageFont  # 新增 PIL 用於處理圖片

# --- 設定網頁標題 ---
st.set_page_config(page_title="商標頁碼工具", page_icon="📝")
st.header("📝 PDF & 圖片 商標遮蓋與頁碼工具")

# --- 自動偵測字型 (kaiu.ttf) ---
def get_font_path():
    possible_paths = [
        "kaiu.ttf", "font.ttf", 
        "pages/kaiu.ttf", "pages/font.ttf", 
        "../kaiu.ttf", "../font.ttf"
    ]
    for p in possible_paths:
        if os.path.exists(p):
            return p
    return None

# --- 字型載入 ---
font_path = get_font_path()
if font_path:
    try:
        pdfmetrics.registerFont(TTFont('CustomFont', font_path))
        font_name = 'CustomFont'
        st.success(f"✅ 字型載入成功 ({os.path.basename(font_path)})")
    except:
        font_name = "Helvetica"
        st.error("❌ 字型載入失敗")
else:
    font_name = "Helvetica"
    st.warning("⚠️ 未偵測到中文字型 (kaiu.ttf)，PDF 文字可能顯示為方塊。")

# --- PDF 遮蓋層建立 ---
def create_pdf_overlay(page_width, page_height, page_num, current_font):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    
    text = f"交通組製 - 第 {page_num} 頁"
    
    box_width = 112
    box_height = 20
    
    # PDF 座標原點在左下角
    rect_x = page_width - box_width
    rect_y = 0
    
    # 畫白框 (遮蓋層)
    c.setFillColor(white)
    c.setStrokeColor(white)
    c.rect(rect_x, rect_y, box_width, box_height, fill=1, stroke=1)
    
    # 寫字 (黑色)
    c.setFillColor(black)
    c.setFont(current_font, 14) 
    
    text_end_x = page_width - 4
    text_y = 4
    
    c.drawRightString(text_end_x, text_y, text)
    c.save()
    packet.seek(0)
    return packet

# --- 圖片處理邏輯 ---
def process_image(uploaded_img_file, f_path):
    # 開啟圖片並轉換為 RGB (避免 PNG 透明背景問題)
    img = Image.open(uploaded_img_file).convert("RGBA")
    file_ext = uploaded_img_file.name.split('.')[-1].lower()
    
    if file_ext in ['jpg', 'jpeg']:
        img = img.convert("RGB")

    draw = ImageDraw.Draw(img)
    width, height = img.size

    box_width = 112
    box_height = 20
    
    # PIL 座標原點在左上角，所以 Y 座標為 高度 - 框高
    rect_x = width - box_width
    rect_y = height - box_height

    # 畫白框
    draw.rectangle([rect_x, rect_y, width, height], fill="white")

    # 設定字型
    try:
        if f_path:
            pil_font = ImageFont.truetype(f_path, 14)
        else:
            pil_font = ImageFont.load_default()
    except:
        pil_font = ImageFont.load_default()

    text = "交通組製 - 第 1 頁" # 圖片通常只有單頁

    # 計算文字大小以進行對齊
    text_bbox = draw.textbbox((0, 0), text, font=pil_font)
    text_w = text_bbox[2] - text_bbox[0]
    text_h = text_bbox[3] - text_bbox[1]

    # 文字位置微調 (靠右 4 px，垂直置中)
    text_x = width - text_w - 4
    # PIL 的文字 Y 座標微調通常需要考慮字體本身的 offset
    text_y = rect_y + (box_height - text_h) / 2 - 2 

    # 繪製文字
    draw.text((text_x, text_y), text, fill="black", font=pil_font)

    # 輸出為 Bytes
    out_io = io.BytesIO()
    out_format = 'PNG' if file_ext == 'png' else 'JPEG'
    img.save(out_io, format=out_format)
    
    return out_io.getvalue(), file_ext, out_format

# --- 主處理邏輯 ---
# 更新 uploader 支援圖片格式
uploaded_file = st.file_uploader("上傳原始 PDF 或圖片", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file and st.button("開始加工"):
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    try:
        # 處理 PDF 邏輯
        if file_extension == 'pdf':
            reader = PdfReader(uploaded_file)
            writer = PdfWriter()
            
            progress_bar = st.progress(0)
            total = len(reader.pages)
            
            for i, page in enumerate(reader.pages):
                w = float(page.mediabox.width)
                h = float(page.mediabox.height)
                
                overlay = create_pdf_overlay(w, h, i+1, font_name)
                page.merge_page(PdfReader(overlay).pages[0])
                writer.add_page(page)
                progress_bar.progress((i + 1) / total)
                
            out = io.BytesIO()
            writer.write(out)
            st.success("🎉 PDF 加工完成！")
            st.download_button("📥 下載加工版 PDF", out.getvalue(), "交通組_加工版.pdf", "application/pdf")

        # 處理圖片邏輯
        elif file_extension in ['png', 'jpg', 'jpeg']:
            img_bytes, ext, out_format = process_image(uploaded_file, font_path)
            
            st.success("🎉 圖片加工完成！")
            st.image(img_bytes, caption="預覽加工結果")
            st.download_button(
                label="📥 下載加工版圖片", 
                data=img_bytes, 
                file_name=f"交通組_加工版.{ext}", 
                mime=f"image/{out_format.lower()}"
            )
            
    except Exception as e:
        st.error(f"錯誤發生: {e}")
