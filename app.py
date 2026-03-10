import streamlit as st
import pandas as pd
import io
import platform
import shutil
from pdf2image import convert_from_bytes
from pptx import Presentation

# ==========================================
# 1. 基本設定 (這行必須在最上面)
# ==========================================
st.set_page_config(page_title="龍潭分局交通戰情室", page_icon="🚓", layout="wide")

# ==========================================
# 2. 側邊欄導覽選單
# ==========================================
with st.sidebar:
    st.title("🚓 龍潭分局選單")
    app_mode = st.selectbox(
        "選擇功能模組",
        ["🏠 戰情室首頁", "📊 交通事故統計", "📂 PDF 轉 PPTX 工具"]
    )
    
    st.markdown("---")
    # 環境檢查 (僅在 PDF 轉檔模式顯示)
    if app_mode == "📂 PDF 轉 PPTX 工具":
        poppler_path = shutil.which("pdftoppm")
        if poppler_path:
            st.success("✅ 系統環境已就緒")
        else:
            st.error("❌ 缺少轉檔引擎 (Poppler)")

# ==========================================
# 3. 功能模組：🏠 戰情室首頁
# ==========================================
if app_mode == "🏠 戰情室首頁":
    st.title("🚓 桃園市政府警察局龍潭分局 - 交通戰情室")
    st.markdown("---")
    st.info("👈 請從左側選單選擇您要使用的功能模組。")
    st.write("本系統提供龍潭轄區交通數據統計及相關行政文書工具。")

# ==========================================
# 4. 功能模組：📊 交通事故統計 (原功能位置)
# ==========================================
elif app_mode == "📊 交通事故統計":
    st.header("📊 交通事故統計模組")
    # --- 這裡放入您原本的 CSV 處理或統計代碼 ---
    uploaded_csv = st.file_uploader("上傳交通事故報表 (CSV/Excel)", type=["csv", "xlsx"])
    if uploaded_csv:
        st.write("檔案已上傳，這裡是您的統計邏輯顯示區。")
        # df = pd.read_csv(uploaded_csv)
        # st.dataframe(df.head())

# ==========================================
# 5. 功能模組：📂 PDF 轉 PPTX 工具
# ==========================================
elif app_mode == "📂 PDF 轉 PPTX 工具":
    st.header("📂 PDF 轉檔中心 (生成 PPTX)")
    st.info("將 PDF 每一頁轉換為投影片圖片，保持原始排版不跑版。")
    
    uploaded_pdf = st.file_uploader("上傳 PDF 檔案 (建議用加工後的版本)", type=["pdf"])
    
    # 使用 placeholder 容器防止 DOM 渲染錯誤 (removeChild error)
    placeholder = st.empty()
    
    if uploaded_pdf:
        if st.button("🚀 開始轉換為 PPTX", use_container_width=True):
            with st.spinner("正在解析 PDF 並生成投影片..."):
                try:
                    # 1. 讀取 PDF
                    pdf_bytes = uploaded_pdf.read()
                    
                    # 2. 轉為圖片
                    images = convert_from_bytes(pdf_bytes, dpi=150)
                    
                    if not images:
                        st.error("無法解析該 PDF，請確認檔案。")
                    else:
                        # 3. 建立 PPT
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
                        
                        # 4. 準備下載
                        pptx_out = io.BytesIO()
                        prs.save(pptx_out)
                        
                        # 5. 顯示成功訊息與按鈕
                        placeholder.empty()
                        with placeholder.container():
                            st.success(f"✅ 轉換成功！共 {len(images)} 頁")
                            st.download_button(
                                label="📥 下載轉換後的 PPTX",
                                data=pptx_out.getvalue(),
                                file_name=f"{uploaded_pdf.name.split('.')[0]}.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True
                            )
                            
                except Exception as e:
                    st.error(f"轉換過程中發生錯誤：{e}")
    else:
        placeholder.warning("尚未上傳任何 PDF 檔案。")
