import streamlit as st
import pandas as pd
import io
import re
import gspread
import smtplib
from datetime import datetime
from pdf2image import convert_from_bytes
from pptx import Presentation
from pptx.util import Inches
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ==========================================
# 0. 系統初始化與全局設定
# ==========================================
st.set_page_config(page_title="龍潭分局交通智慧戰情室", page_icon="🚓", layout="wide")

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"
TO_EMAIL = "mbmmiqamhnd@gmail.com"

# 目標值設定區
TARGETS_MAJOR = {'科技執法': 6006, '聖亭所': 1941, '龍潭所': 2588, '中興所': 1941, '石門所': 1479, '高平所': 1294, '三和所': 339, '交通分隊': 2526}
TARGETS_OVERLOAD = {'聖亭所': 20, '龍潭所': 27, '中興所': 20, '石門所': 16, '高平所': 14, '三和所': 8, '警備隊': 0, '交通分隊': 22}

# ==========================================
# 🛠️ 通用工具箱 (Helper Functions)
# ==========================================
def get_std_unit(n):
    """將五花八門的單位名稱標準化"""
    if pd.isna(n): return None
    n = str(n).strip()
    if '分隊' in n: return '交通分隊'
    if '科技' in n or '交通組' in n: return '科技執法'
    if '警備' in n: return '警備隊'
    for k in ['聖亭', '龍潭', '中興', '石門', '高平', '三和']:
        if k in n: return k + '所'
    return None

def sync_gsheet_batch(ws, title, df_data, font_size=16):
    """通用雲端同步：含藍紅標題格式化"""
    ws.clear()
    # 填補 NaN 避免 JSON 序列化錯誤
    df_data = df_data.fillna("")
    data = [[title]] + [df_data.columns.tolist()] + df_data.values.tolist()
    ws.update(values=data)
    
    blue = {"red": 0, "green": 0, "blue": 1}
    red = {"red": 1, "green": 0, "blue": 0}
    split_idx = title.find("(") if "(" in title else len(title)
    
    reqs = [{
        "updateCells": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1},
            "rows": [{"values": [{"userEnteredValue": {"stringValue": title},
                "textFormatRuns": [
                    {"startIndex": 0, "format": {"foregroundColor": blue, "bold": True, "fontSize": font_size}},
                    {"startIndex": split_idx, "format": {"foregroundColor": red, "bold": True, "fontSize": font_size}}
                ]}]}],
            "fields": "userEnteredValue,textFormatRuns"
        }
    }]
    ws.spreadsheet.batch_update({"requests": reqs})

def send_email_alert(subject, body, df_attachment=None):
    """發送自動化電子郵件"""
    try:
        # 需在 st.secrets 中設定 email_sender 與 email_password
        sender = st.secrets.get("email_sender", "")
        password = st.secrets.get("email_password", "")
        if not sender or not password:
            st.warning("⚠️ 尚未設定 Email 帳密，跳過郵件發送。")
            return
            
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = TO_EMAIL
        msg.attach(MIMEText(body, 'html'))
        
        if df_attachment is not None:
            csv_data = df_attachment.to_csv(index=False).encode('utf-8-sig')
            part = MIMEApplication(csv_data, Name="報表附件.csv")
            part['Content-Disposition'] = 'attachment; filename="報表附件.csv"'
            msg.attach(part)
            
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender, password)
            server.send_message(msg)
        st.success("📧 通知信件已成功發送！")
    except Exception as e:
        st.error(f"郵件發送失敗: {e}")

def load_data(file):
    """自動判斷副檔名並讀取資料"""
    if file.name.endswith('.csv'):
        return pd.read_csv(file, encoding='big5', on_bad_lines='skip') # 台灣公部門常為 big5
    else:
        return pd.read_excel(file)

# ==========================================
# 🏰 導覽選單
# ==========================================
with st.sidebar:
    st.title("🚓 龍潭分局戰情室")
    app_mode = st.selectbox("功能模組", ["🏠 智慧上傳中心", "📂 PDF 轉 PPTX 工具"])
    st.divider()
    st.info("💡 10秒流程：首頁直接拖入報表即可分析。")

# ==========================================
# 🏠 核心：智慧上傳中心 (特徵分配版)
# ==========================================
if app_mode == "🏠 智慧上傳中心":
    st.header("📈 交通數據智慧分析中心")
    uploads = st.file_uploader("📂 一次拖入所有報表檔案，系統將自動分類處理", type=["xlsx", "csv", "xls"], accept_multiple_files=True)

    if uploads:
        try:
            gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
            sh = gc.open_by_url(GOOGLE_SHEET_URL)
        except Exception as e:
            st.error("連線 Google Sheets 失敗，請確認 st.secrets 設定。")
            st.stop()
            
        category_files = {"科技執法": [], "重大違規": [], "超載統計": [], "強化專案": [], "交通事故": [], "未知分類": []}

        # --- 2. 智慧分配：依據檔名特徵分類 ---
        for f in uploads:
            name = f.name.lower() # 轉小寫方便比對英文
            
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

        # --- 3. 執行統計邏輯 ---
        with st.spinner("資料運算與雲端同步中..."):
            
            # [科技執法] - Worksheet 4
            if category_files["科技執法"]:
                st.success(f"📸 識別「科技執法」：共 {len(category_files['科技執法'])} 份")
                try:
                    df_list = [load_data(f) for f in category_files["科技執法"]]
                    df = pd.concat(df_list, ignore_index=True)
                    # 假設有欄位包含「單位」
                    unit_col = next((col for col in df.columns if '單位' in col), None)
                    if unit_col:
                        df['標準單位'] = df[unit_col].apply(get_std_unit)
                        summary = df.groupby('標準單位').size().reset_index(name='件數')
                        sync_gsheet_batch(sh.get_worksheet(4), f"科技執法統計({datetime.now().strftime('%Y-%m-%d')})", summary)
                except Exception as e:
                    st.error(f"科技執法處理錯誤: {e}")

            # [重大交通違規] - Worksheet 0
            if category_files["重大違規"]:
                st.success(f"🚨 識別「重大交通違規」：共 {len(category_files['重大違規'])} 份")
                try:
                    df_list = [load_data(f) for f in category_files["重大違規"]]
                    df = pd.concat(df_list, ignore_index=True)
                    sync_gsheet_batch(sh.get_worksheet(0), f"重大交通違規({datetime.now().strftime('%Y-%m-%d')})", df.head(50))
                except Exception as e:
                    st.error(f"重大違規處理錯誤: {e}")

            # [超載統計] - Worksheet 1
            if category_files["超載統計"]:
                st.success(f"🚛 識別「超載違規」：共 {len(category_files['超載統計'])} 份")
                try:
                    df_list = [load_data(f) for f in category_files["超載統計"]]
                    df = pd.concat(df_list, ignore_index=True)
                    unit_col = next((col for col in df.columns if '單位' in col), None)
                    if unit_col:
                        df['標準單位'] = df[unit_col].apply(get_std_unit)
                        summary = df.groupby('標準單位').size().reset_index(name='本月件數')
                        target_df = pd.DataFrame(list(TARGETS_OVERLOAD.items()), columns=['標準單位', '目標值'])
                        final_df = pd.merge(target_df, summary, on='標準單位', how='left').fillna(0)
                        final_df['達成率'] = (final_df['本月件數'] / final_df['目標值'] * 100).round(2).astype(str) + '%'
                        
                        sync_gsheet_batch(sh.get_worksheet(1), f"超載統計({datetime.now().strftime('%Y-%m-%d')})", final_df)
                        send_email_alert("龍潭分局-超載違規自動統計完成", "長官您好，本期超載違規數據已結算完畢，請參閱附件。", final_df)
                except Exception as e:
                    st.error(f"超載統計處理錯誤: {e}")

            # [強化專案] - Worksheet 5
            if category_files["強化專案"]:
                st.success(f"🔥 識別「強化交通安全專案 (含砂石車)」：共 {len(category_files['強化專案'])} 份")
                try:
                    df_list = [load_data(f) for f in category_files["強化專案"]]
                    df = pd.concat(df_list, ignore_index=True)
                    # TODO: 這裡可以填入針對強化專案/砂石車的具體 Pandas 計算邏輯
                    # 暫時先直接上傳前50筆作為示意
                    sync_gsheet_batch(sh.get_worksheet(5), f"強化專案統計({datetime.now().strftime('%Y-%m-%d')})", df.head(50))
                except Exception as e:
                    st.error(f"強化專案處理錯誤: {e}")

            # [交通事故] - Worksheet 2, 3
            if category_files["交通事故"]:
                st.success(f"🚑 識別「交通事故 (含案件統計)」：共 {len(category_files['交通事故'])} 份")
                try:
                    df_list = [load_data(f) for f in category_files["交通事故"]]
                    df = pd.concat(df_list, ignore_index=True)
                    # TODO: 這裡可以填入交通事故 A1/A2 的具體 Pandas 計算邏輯
                    # 暫時先上傳至 Worksheet 2 作為示意
                    sync_gsheet_batch(sh.get_worksheet(2), f"交通事故統計({datetime.now().strftime('%Y-%m-%d')})", df.head(50))
                except Exception as e:
                    st.error(f"交通事故處理錯誤: {e}")

            # [未知防呆]
            if category_files["未知分類"]:
                for f in category_files["未知分類"]:
                    st.warning(f"⚠️ 無法自動識別此檔案，請確認檔名包含對應關鍵字：{f.name}")

# ==========================================
# 📂 模式二：PDF 轉 PPTX
# ==========================================
elif app_mode == "📂 PDF 轉 PPTX 工具":
    st.header("📂 PDF 行政文書轉 PPTX 簡報")
    st.markdown("快速將 PDF 報表、公文每一頁轉換成 PowerPoint 投影片。")
    
    pdf_file = st.file_uploader("上傳 PDF 檔案", type=["pdf"])
    
    if pdf_file:
        if st.button("🚀 開始轉換"):
            with st.spinner("正在將 PDF 轉換為圖片，並合成簡報..."):
                try:
                    # 1. 將 PDF 轉為圖片清單 (需確認伺服器有安裝 poppler-utils)
                    pdf_bytes = pdf_file.read()
                    images = convert_from_bytes(pdf_bytes, dpi=200)
                    
                    # 2. 建立 PPTX
                    prs = Presentation()
                    blank_slide_layout = prs.slide_layouts[6] # 空白排版
                    
                    # 3. 將圖片貼上投影片
                    for img in images:
                        slide = prs.slides.add_slide(blank_slide_layout)
                        img_byte_arr = io.BytesIO()
                        img.save(img_byte_arr, format='PNG')
                        img_byte_arr.seek(0)
                        
                        # 調整尺寸以填滿投影片 (預設比例)
                        slide.shapes.add_picture(img_byte_arr, 0, 0, width=prs.slide_width, height=prs.slide_height)
                    
                    # 4. 輸出檔案供下載
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
