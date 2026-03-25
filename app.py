import streamlit as st
import io
from pdf2image import convert_from_bytes
from pptx import Presentation

# ==========================================
# 0. 系統初始化與全局設定
# ==========================================
st.set_page_config(page_title="龍潭分局交通智慧戰情室", page_icon="🚓", layout="wide")

# ==========================================
# 🏰 導覽選單
# ==========================================
with st.sidebar:
    st.title("🚓 龍潭分局戰情室")
    app_mode = st.selectbox("功能模組", ["🏠 智慧上傳中心 (首頁)", "📂 PDF 轉 PPTX 工具"])
    st.divider()
    st.info("💡 10秒流程：將所有報表全部拖入首頁，系統會自動分類並引導您前往對應功能。")

# ==========================================
# 🏠 核心：智慧上傳中心 (總機路由模式)
# ==========================================
if app_mode == "🏠 智慧上傳中心 (首頁)":
    st.header("📈 交通數據智慧分析中心")
    st.markdown("請將您從警政系統匯出的報表（無論是科技執法、超載、交通事故等）**一次全部拖入下方**。")
    
    uploads = st.file_uploader("📂 拖入所有報表檔案，系統將自動分類處理", type=["xlsx", "csv", "xls"], accept_multiple_files=True)

    if uploads:
        # --- 1. 建立分類收納盒 ---
        category_files = {"科技執法": [], "重大違規": [], "超載統計": [], "強化專案": [], "交通事故": [], "未知分類": []}

        # --- 2. 智慧分配：依據檔名特徵分類 ---
        for f in uploads:
            name = f.name.lower()
            
            # [科技執法]
            if "list" in name or "地點" in name or "科技" in name:
                category_files["科技執法"].append(f)
                
            # [超載統計]
            elif "stone" in name or "超載" in name:
                category_files["超載統計"].append(f)
                
            # [重大違規]
            elif "重大" in name:
                category_files["重大違規"].append(f)
                
            # [強化專案] (包含大型車違規)
            elif "強化" in name or "專案" in name or "砂石車" in name or "r17" in name:
                category_files["強化專案"].append(f)
                
            # [交通事故] (包含案件統計表)
            elif "a1" in name or "a2" in name or "事故" in name or "案件統計" in name:
                category_files["交通事故"].append(f)
                
            # [未知分類]
            else:
                category_files["未知分類"].append(f)

        # --- 3. 顯示分類結果與跳轉按鈕 ---
        st.divider()
        st.subheader("🎯 檔案分類結果與處理通道")
        
        # 建立多欄位排版，讓畫面更整齊
        cols = st.columns(2)
        col_idx = 0

        # [超載統計] 路由
        if category_files["超載統計"]:
            with cols[col_idx % 2]:
                st.info(f"🚛 **超載統計**：已識別 {len(category_files['超載統計'])} 份檔案")
                st.session_state["auto_files_overload"] = category_files["超載統計"]
                if st.button("🚀 前往處理【超載統計】", key="btn_overload", use_container_width=True):
                    st.switch_page("pages/3_超載統計.py")
            col_idx += 1

        # [科技執法] 路由
        if category_files["科技執法"]:
            with cols[col_idx % 2]:
                st.info(f"📸 **科技執法**：已識別 {len(category_files['科技執法'])} 份檔案")
                st.session_state["auto_files_tech"] = category_files["科技執法"]
                if st.button("🚀 前往處理【科技執法】", key="btn_tech", use_container_width=True):
                    st.switch_page("pages/7_📸_科技執法成效.py")
            col_idx += 1

        # [重大違規] 路由
        if category_files["重大違規"]:
            with cols[col_idx % 2]:
                st.info(f"🚨 **重大違規**：已識別 {len(category_files['重大違規'])} 份檔案")
                st.session_state["auto_files_major"] = category_files["重大違規"]
                if st.button("🚀 前往處理【重大交通違規】", key="btn_major", use_container_width=True):
                    st.switch_page("pages/2_取締重大交通違規統計.py")
            col_idx += 1

        # [強化專案] 路由
        if category_files["強化專案"]:
            with cols[col_idx % 2]:
                st.info(f"🔥 **強化專案 (含砂石車)**：已識別 {len(category_files['強化專案'])} 份檔案")
                st.session_state["auto_files_project"] = category_files["強化專案"]
                if st.button("🚀 前往處理【強化專案勤務】", key="btn_project", use_container_width=True):
                    st.switch_page("pages/8_強化交通安全執法專案勤務取締件數統計表.py")
            col_idx += 1

        # [交通事故] 路由
        if category_files["交通事故"]:
            with cols[col_idx % 2]:
                st.info(f"🚑 **交通事故**：已識別 {len(category_files['交通事故'])} 份檔案")
                st.session_state["auto_files_accident"] = category_files["交通事故"]
                if st.button("🚀 前往處理【交通事故統計】", key="btn_accident", use_container_width=True):
                    st.switch_page("pages/1_交通事故統計.py")
            col_idx += 1

        # [未知分類防呆]
        if category_files["未知分類"]:
            st.divider()
            for f in category_files["未知分類"]:
                st.warning(f"⚠️ 無法識別此檔案，請確認檔名是否正確或手動前往各分頁上傳：{f.name}")

# ==========================================
# 📂 模式二：PDF 轉 PPTX (保留原功能)
# ==========================================
elif app_mode == "📂 PDF 轉 PPTX 工具":
    st.header("📂 PDF 行政文書轉 PPTX 簡報")
    st.markdown("快速將 PDF 報表、公文每一頁轉換成 PowerPoint 投影片。")
    
    pdf_file = st.file_uploader("上傳 PDF 檔案", type=["pdf"])
    
    if pdf_file:
        if st.button("🚀 開始轉換"):
            with st.spinner("正在將 PDF 轉換為圖片，並合成簡報..."):
                try:
                    pdf_bytes = pdf_file.read()
                    images = convert_from_bytes(pdf_bytes, dpi=200)
                    
                    prs = Presentation()
                    blank_slide_layout = prs.slide_layouts[6]
                    
                    for img in images:
                        slide = prs.slides.add_slide(blank_slide_layout)
                        img_byte_arr = io.BytesIO()
                        img.save(img_byte_arr, format='PNG')
                        img_byte_arr.seek(0)
                        slide.shapes.add_picture(img_byte_arr, 0, 0, width=prs.slide_width, height=prs.slide_height)
                    
                    pptx_io = io.BytesIO()
                    prs.save(pptx_io)
                    pptx_io.seek(0)
                    
                    st.success("✅ 轉換完成！")
                    st.download_button(
                        label="📥 下載 PPTX 檔案",
                        data=pptx_io,
                        file_name=f"{pdf_file.name.replace('.pdf', '')}_轉換.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                except Exception as e:
                    st.error(f"轉換過程發生錯誤。請確認系統環境是否已安裝 Poppler。錯誤訊息：{e}")
