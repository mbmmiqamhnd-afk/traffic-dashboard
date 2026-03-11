import streamlit as st
import io
import os
# 處理 PDF 需要的套件
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import white, black
# 處理 圖片 需要的套件
from PIL import Image, ImageDraw, ImageFont

# --- 設定網頁標題 ---
st.set_page_config(page_title="商標遮蓋工具", page_icon="📝")
st.header("📝 PDF & 圖片 商標遮蓋工具")
st.markdown("---")

# ==============================
#  共用工具函數：字型偵測與載入
# ==============================

def get_font_path():
    """搜尋系統中可能的字型路徑"""
    possible_paths = [
        "kaiu.ttf", "font.ttf", 
        "pages/kaiu.ttf", "pages/font.ttf", 
        "../kaiu.ttf", "../font.ttf",
        "C:\\Windows\\Fonts\\kaiu.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc" 
    ]
    for p in possible_paths:
        if os.path.exists(p):
            return p
    return None

# --- 初始化字型 ---
font_path = get_font_path()
pdf_font_name = "Helvetica" 
pil_font_path = None        

if font_path:
    try:
        pdfmetrics.registerFont(TTFont('CustomFont', font_path))
        pdf_font_name = 'CustomFont'
        pil_font_path = font_path 
        st.sidebar.success(f"✅ 已載入中文字型：{os.path.basename(font_path)}")
    except Exception as e:
        st.sidebar.error(f"❌ 字型載入失敗: {e}")
else:
    st.sidebar.warning("⚠️ 未偵測到 kaiu.ttf。中文字元可能無法正確顯示。")


# ==============================
#  核心邏輯 1：處理 PDF (保持原樣，含頁碼)
# ==============================

def create_pdf_overlay(page_width, page_height, page_num, current_font):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    
    text = f"交通組製 - 第 {page_num} 頁"
    
    box_width = 112
    box_height = 20
    
    rect_x = page_width - box_width
    rect_y = 0
    
    c.setFillColor(white)
    c.setStrokeColor(white)
    c.rect(rect_x, rect_y, box_width, box_height, fill=1, stroke=1)
    
    c.setFillColor(black)
    c.setFont(current_font, 14) 
    
    text_end_x = page_width - 4
    text_y = 4
    
    c.drawRightString(text_end_x, text_y, text)
    c.save()
    packet.seek(0)
    return packet

def process_pdf(uploaded_pdf_file):
    reader = PdfReader(uploaded_pdf_file)
    writer = PdfWriter()
    
    progress_text = st.empty()
    progress_bar = st.progress(0)
    total_pages = len(reader.pages)
    
    for i, page in enumerate(reader.pages):
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        
        overlay = create_pdf_overlay(w, h, i+1, pdf_font_name)
        page.merge_page(PdfReader(overlay).pages[0])
        writer.add_page(page)
        
        current_progress = (i + 1) / total_pages
        progress_bar.progress(current_progress)
        progress_text.text(f"正在處理第 {i+1}/{total_pages} 頁...")
        
    out_io = io.BytesIO()
    writer.write(out_io)
    return out_io.getvalue()

# ==============================
#  核心邏輯 2：處理 圖片 (放大字體與完美置中)
# ==============================

def process_image(uploaded_img_file, f_path):
    img = Image.open(uploaded_img_file).convert("RGBA")
    file_ext = uploaded_img_file.name.split('.')[-1].lower()

    draw = ImageDraw.Draw(img)
    width, height = img.size

    # --- 遮蓋範圍擴大一倍 ---
    box_width = 224  
    box_height = 40  
    
    rect_x = width - box_width
    rect_y = height - box_height

    # 畫巨大白框
    draw.rectangle([rect_x, rect_y, width, height], fill="white")

    # --- 【修改點 1】將字體大小設定為 28 (配合高度 40 的框) ---
    font_size = 28
    try:
        if f_path:
            pil_font = ImageFont.truetype(f_path, font_size)
        else:
            pil_font = ImageFont.load_default()
    except:
        pil_font = ImageFont.load_default()

    # --- 【修改點 2】移除頁碼，僅保留固定文字 ---
    text = "交通組製"

    # 取得文字實際的寬高
    text_bbox = draw.textbbox((0, 0), text, font=pil_font)
    text_w = text_bbox[2] - text_bbox[0]
    text_h = text_bbox[3] - text_bbox[1]

    # --- 【修改點 3】計算置中座標 ---
    # 水平：(白框寬度 - 文字寬度) / 2，加上白框起始 X 座標
    text_x = rect_x + (box_width - text_w) / 2
    # 垂直：(白框高度 - 文字高度) / 2，加上白框起始 Y 座標 (微調 -4 避免偏下)
    text_y = rect_y + (box_height - text_h) / 2 - 4 

    # 繪製文字 (黑色)
    draw.text((text_x, text_y), text, fill="black", font=pil_font)

    # 輸出轉換
    out_io = io.BytesIO()
    if file_ext in ['jpg', 'jpeg']:
        img = img.convert("RGB")
        out_format = 'JPEG'
    else:
        out_format = 'PNG'
        
    img.save(out_io, format=out_format)
    return out_io.getvalue(), file_ext, out_format

# ==============================
#  主介面邏輯
# ==============================

st.write("請上傳檔案進行處理 (支援多檔上傳)")

uploaded_files = st.file_uploader(
    "選擇 PDF 或 圖片檔案 (支援 PDF, PNG, JPG, JPEG)", 
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True 
)

if uploaded_files and st.button("🚀 開始批次加工"):
    st.write("---")
    st.subheader("處理結果列表")
    
    results_container = st.container()

    for i, uploaded_file in enumerate(uploaded_files):
        file_name = uploaded_file.name
        file_ext = file_name.split('.')[-1].lower()
        base_name = os.path.splitext(file_name)[0]
        new_file_name = f"{base_name}_加工版.{file_ext}"
        
        with results_container.expander(f"📄 檔案 {i+1}: {file_name}", expanded=True):
            try:
                if file_ext == 'pdf':
                    with st.spinner(f"正在處理 PDF: {file_name}..."):
                        processed_pdf_bytes = process_pdf(uploaded_file)
                        st.success("PDF 加工完成！")
                        st.download_button(
                            label=f"📥 下載 {new_file_name}",
                            data=processed_pdf_bytes,
                            file_name=new_file_name,
                            mime="application/pdf",
                            key=f"dl_pdf_{i}"
                        )
                        
                elif file_ext in ['png', 'jpg', 'jpeg']:
                    with st.spinner(f"正在處理圖片: {file_name}..."):
                        img_bytes, ext, out_fmt = process_image(uploaded_file, pil_font_path)
                        st.success("圖片 加工完成！")
                        
                        col1, col2 = st.columns([3, 2])
                        with col1:
                            st.image(img_bytes, caption="加工後預覽 (字體已放大並置中)", use_column_width=True)
                        with col2:
                            st.write("") 
                            st.write("")
                            st.download_button(
                                label=f"📥 下載 {new_file_name}",
                                data=img_bytes,
                                file_name=new_file_name,
                                mime=f"image/{out_fmt.lower()}",
                                key=f"dl_img_{i}"
                            )
                            
            except Exception as e:
                st.error(f"處理 {file_name} 時發生錯誤: {e}")

    st.write("---")
    st.success("✅ 所有檔案處理完畢！")
