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

# --- 自動偵測字型 (這裡是關鍵修改) ---
def get_font_path():
    # 程式會依序尋找這些檔案，直到找到為止
    possible_paths = [
        "kaiu.ttf",         # 您的檔名 (根目錄)
        "font.ttf",         # 備用檔名
        "pages/kaiu.ttf",   # 您的檔名 (pages目錄)
        "pages/font.ttf",   # 備用檔名
        "../kaiu.ttf",      # 上一層
        "../font.ttf"
    ]
    
    for p in possible_paths:
        if os.path.exists(p):
            return p
    return None

# --- 除錯與註冊 ---
font_path = get_font_path()
if font_path:
    st.success(f"✅ 成功載入字型檔：{font_path}")
    try:
        pdfmetrics.registerFont(TTFont('CustomFont', font_path))
        font_name = 'CustomFont'
    except Exception as e:
        st.error(f"❌ 字型載入失敗，檔案可能損毀：{e}")
        font_name = "Helvetica"
else:
    st.error("❌ 找不到 kaiu.ttf！請確認檔案已上傳到 GitHub。")
    font_name = "Helvetica" # 暫時用英文型，避免程式崩潰，但會顯示方塊

# --- 製作浮水印圖層 ---
def create_overlay(page_width, page_height, page_num, current_font):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    
    text = f"交通組製 - 第 {page_num} 頁"
    
    # 設定位置
    box_width = 200
    box_height = 30
    rect_x = page_width - box_width - 20
    rect_y = 10
    
    # 畫白框
    c.setFillColor(white)
    c.setStrokeColor(white)
    c.rect(rect_x, rect_y, box_width, box_height, fill=1, stroke=1)
    
    # 寫字
    c.setFillColor(black)
    c.setFont(current_font, 12)
    c.drawRightString(page_width - 30, rect_y + 8, text)
    
    c.save()
    packet.seek(0)
    return packet

# --- 主處理邏輯 ---
uploaded_file = st.file_uploader("上傳原始 PDF", type=["pdf"])

if uploaded_file and st.button("開始加工"):
    if font_name == "Helvetica":
        st.warning("⚠️ 警告：目前使用預設字型，中文可能會顯示為方塊。請先解決上方的紅色錯誤。")

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
        st.error(f"處理過程發生錯誤: {e}")
