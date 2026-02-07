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
st.info("功能：自動遮蓋右下角舊商標，並加上「交通組製」落款與頁碼。")

# --- 1. 字型註冊函式 ---
def register_font():
    # 自動搜尋 font.ttf (無論在根目錄還是 pages 都能找到)
    paths = ["font.ttf", "../font.ttf", "pages/font.ttf"]
    for p in paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont('CustomFont', p))
                return 'CustomFont'
            except:
                pass
    return "Helvetica"

# --- 2. 製作浮水印圖層 ---
def create_overlay(page_width, page_height, page_num, font_name):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    
    # 設定顯示文字
    text = f"交通組製 - 第 {page_num} 頁"
    
    # 設定遮罩與文字位置 (右下角)
    box_width = 200   # 遮罩寬度 (白色貼紙大小)
    box_height = 30   # 遮罩高度
    margin_right = 20
    margin_bottom = 10
    
    # 計算位置
    rect_x = page_width - box_width - margin_right
    rect_y = margin_bottom
    
    # A. 畫白色遮罩 (像立可白一樣蓋掉舊 Logo)
    c.setFillColor(white)
    c.setStrokeColor(white)
    c.rect(rect_x, rect_y, box_width, box_height, fill=1, stroke=1)
    
    # B. 寫上新文字
    c.setFillColor(black)
    c.setFont(font_name, 12)
    
    # 文字靠右對齊計算
    text_end_x = page_width - margin_right - 10 
    text_y = rect_y + 8 # 垂直微調
    
    c.drawRightString(text_end_x, text_y, text)
    
    c.save()
    packet.seek(0)
    return packet

# --- 3. 主處理邏輯 ---
uploaded_file = st.file_uploader("上傳原始 PDF", type=["pdf"])

if uploaded_file and st.button("開始加工"):
    font_name = register_font()
    reader = PdfReader(uploaded_file)
    writer = PdfWriter()
    
    # 進度條
    progress_bar = st.progress(0)
    total_pages = len(reader.pages)
    
    for i, page in enumerate(reader.pages):
        # 取得頁面尺寸
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        
        # 製作每一頁的浮水印
        overlay = create_overlay(w, h, i+1, font_name)
        overlay_page = PdfReader(overlay).pages[0]
        
        # 合併
        page.merge_page(overlay_page)
        writer.add_page(page)
        
        # 更新進度
        progress_bar.progress((i + 1) / total_pages)
        
    # 輸出
    out = io.BytesIO()
    writer.write(out)
    st.success("完成！")
    st.download_button(
        label="📥 下載加工後的 PDF",
        data=out.getvalue(),
        file_name="交通組_加工版.pdf",
        mime="application/pdf"
    )
