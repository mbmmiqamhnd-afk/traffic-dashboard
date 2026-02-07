import streamlit as st
import io
import os
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import white, black

# --- 設定網頁標題 ---
st.set_page_config(page_title="商標頁碼工具", page_icon="📝")
st.header("📝 PDF 商標遮蓋與頁碼工具")

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
    st.warning("⚠️ 未偵測到中文字型 (kaiu.ttf)，文字將顯示為方塊。")

# --- 核心修改：精確尺寸調整 ---
def create_overlay(page_width, page_height, page_num, current_font):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    
    text = f"交通組製 - 第 {page_num} 頁"
    
    # --- 【修改點】尺寸計算 ---
    # 寬度：135 * (5/6) ≈ 112
    # 高度：維持 20 (輕薄貼紙風格)
    box_width = 112
    box_height = 20
    
    # 貼齊右下角 (無邊距)
    rect_x = page_width - box_width
    rect_y = 0
    
    # 畫白框 (遮蓋層)
    c.setFillColor(white)
    c.setStrokeColor(white)
    c.rect(rect_x, rect_y, box_width, box_height, fill=1, stroke=1)
    
    # 寫字 (黑色)
    c.setFillColor(black)
    
    # 字體大小設為 10 (配合 112 的寬度)
    c.setFont(current_font, 10) 
    
    # 文字位置微調
    # 水平：靠右對齊，留 4 點邊距
    text_end_x = page_width - 4
    # 垂直：高度20，字高約10，放在 y=6 左右最置中
    text_y = 6
    
    c.drawRightString(text_end_x, text_y, text)
    
    c.save()
    packet.seek(0)
    return packet

# --- 主處理邏輯 ---
uploaded_file = st.file_uploader("上傳原始 PDF", type=["pdf"])

if uploaded_file and st.button("開始加工"):
    try:
        reader = PdfReader(uploaded_file)
        writer = PdfWriter()
        
        progress_bar = st.progress(0)
        total = len(reader.pages)
        
        for i, page in enumerate(reader.pages):
            w = float(page.mediabox.width)
            h = float(page.mediabox.height)
            
            overlay = create_overlay(w, h, i+1, font_name)
            page.merge_page(PdfReader(overlay).pages[0])
            writer.add_page(page)
            progress_bar.progress((i + 1) / total)
            
        out = io.BytesIO()
        writer.write(out)
        st.success("🎉 加工完成！")
        st.download_button("📥 下載加工版 PDF", out.getvalue(), "交通組_加工版.pdf", "application/pdf")
        
    except Exception as e:
        st.error(f"錯誤: {e}")
