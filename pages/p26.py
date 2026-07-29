import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Font
import io
import os
import smtplib
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ReportLab 相關匯入 (用於生成正式交辦單 PDF)
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 優先載入專案資料夾中的 kaiu.ttf (標楷體)
font_path = 'kaiu.ttf'
if os.path.exists(font_path):
    pdfmetrics.registerFont(TTFont('KaiTi', font_path))
    FONT_NAME = 'KaiTi'
else:
    # 若找不到則嘗試 Linux 系統備援路徑或內建字型
    fallback_path = '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf'
    if os.path.exists(fallback_path):
        pdfmetrics.registerFont(TTFont('KaiTi', fallback_path))
        FONT_NAME = 'KaiTi'
    else:
        FONT_NAME = 'Helvetica'

# 匯入系統原本的側邊欄設定
try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        pass

def send_csv_email(file_bytes, file_name):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        
        msg["Subject"] = f"龍潭分局_行人及護老專案自動產出結果_{file_name}"
        
        body_text = (
            f"您好，\n\n"
            f"系統已自動產出「115年上半年行人及護老專案」相關檔案，附件為對應的 Excel 統計檔案。\n\n"
            f"本信件由交通執法自動化分析引擎發送。"
        )
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(file_bytes.getvalue())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(file_name)}")
        msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
            
        return True, None
    except Exception as e:
        return False, str(e)

def generate_all_slips_pdf():
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, right_margin=30, left_margin=30, top_margin=15, bottom_margin=15)
    story = []
    
    title_style = ParagraphStyle(
        'TitleStyle', fontName=FONT_NAME, fontSize=17, leading=22, alignment=1
    )
    
    header_style = ParagraphStyle(
        'HeaderStyle', fontName=FONT_NAME, fontSize=13, leading=16, alignment=1
    )

    body_style = ParagraphStyle(
        'BodyStyle', fontName=FONT_NAME, fontSize=13, leading=18, alignment=4
    )
    
    # 針對 13pt 字體設定懸掛縮排
    char_w = 13
    body_l1 = ParagraphStyle('BodyL1', parent=body_style, leftIndent=char_w*2, firstLineIndent=-char_w*2)
    body_l2 = ParagraphStyle('BodyL2', parent=body_style, leftIndent=char_w*4, firstLineIndent=-char_w*2)
    body_l3 = ParagraphStyle('BodyL3', parent=body_style, leftIndent=char_w*6, firstLineIndent=-char_w*3)
    body_step = ParagraphStyle('BodyStep', parent=body_style, leftIndent=char_w*2, firstLineIndent=0)

    # 所有受文單位清單
    units = ["龍潭所", "聖亭所", "中興所", "石門所", "高平所", "勤務指揮中心", "龍潭交通分隊"]

    for idx, unit_name in enumerate(units):
        story.append(Paragraph("桃園市政府警察局龍潭分局交通組交辦單", title_style))
        story.append(Spacer(1, 8)) 
        
        notice_paragraphs = [
            Paragraph("一、為辦理本分局115年1至6月各單位執行「行人及護老交通安全實施計畫」工作出力人員敘獎案，請統計所屬執勤時數並彙整敘獎人員名冊。", body_l1),
            Paragraph("二、獎勵規則：", body_l1),
            Paragraph("1、行人及護老交通安全實施計畫：", body_l2),
            Paragraph("(一)本案專責勤務人員執行成效良好，每半年執勤時數累計達40小時以上者核予嘉獎一次、80小時以上者核予嘉獎二次。", body_l3),
            Paragraph("(二)略.......", body_l3),
            Paragraph("(三)本案專責勤務及督導人員每人每半年獎勵額度以嘉獎二次為限，其中上半年未達獎勵額度之時數(勤務人員未達40小時或督導人員未達60小時)，得累計至當年度下半年計算。", body_l3),
            Paragraph("三、請至網路硬碟/交通組/巡官郭勝隆的資料夾/☆115年1至6月「行人及護老交通安全勤務實施計畫」工作出力人員敘獎區☆，下載115年1至6月「行人及護老交通安全勤務實施計畫」工作出力人員獎勵清冊，再依「115年辦公日曆表」，「1月至6月非假日之每日06-10時段及16-20時段，凡勤務分配表有顯示「護老專案」之服勤時數均得計算，再請於人事資訊整合管理系統登錄獎懲資料。", body_l1),
            Paragraph("四、115年1至6月「行人及護老交通安全勤務實施計畫」工作出力人員獎勵清冊由主管核章後附交辦單，逕送本組。", body_l1),
            Paragraph("五、登錄獎懲資料的流程：", body_l1),
            Paragraph("新增 > 主旨事由: 115年1至6月執行「行人及護老交通安全實施計畫」出力人員獎勵案 > 受理單位代碼: 交通組 > 受理人員代碼: 郭勝隆 > 再按新增 > 加入獎懲人員 > 獎懲事由: 115年1至6月執行「行人及護老交通安全」專責勤務達幾小時。", body_step),
            Paragraph("六、請不用寫辛勞得力及備極辛勞，系統會自動帶入。", body_l1),
            Spacer(1, 5),
            Paragraph("辦理期限：115年7月28日前辦理完畢連同原件具報。", body_style)
        ]

        data = [
            [
                Paragraph("受文者", header_style), Paragraph(unit_name, header_style), 
                Paragraph("交辦日期", header_style), Paragraph("115年7月21日", header_style), 
                Paragraph("承辦人", header_style), Paragraph("", header_style), 
                Paragraph("單位主管", header_style), Paragraph("組長楊孟竟", header_style)
            ],
            [
                Paragraph("交<br/><br/>辦<br/><br/>事<br/><br/>由", header_style), 
                notice_paragraphs, '', '', '', '', '', ''
            ],
            [
                Paragraph("承辦內容", header_style), 
                Paragraph("承辦人：<br/><br/>所(隊)長：<br/><br/>年 月 日", body_style), '', '', '', '', '', ''
            ]
        ]
        
        # 【精算寬度調整】：
        # 總寬維持 535。
        # 受文單位設為 45 (3個字)
        # 交辦日期設為 56 (精準容納 4 個中文字寬: 13*4 + 4 padding)
        # 將多出的空間補給右側日期值欄位 (拓寬至 101)
        t = Table(data, colWidths=[56, 45, 56, 101, 45, 102, 56, 74])
        
        # 加入左右微小內縮 (PADDING=2) 確保字體緊湊服貼
        t.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('SPAN', (1,1), (7,1)),
            ('SPAN', (1,2), (7,2)),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('LEFTPADDING', (0,0), (-1,-1), 2),
            ('RIGHTPADDING', (0,0), (-1,-1), 2),
        ]))
        
        story.append(t)
        
        # 除最後一個單位外，其他單位後面都加入換頁符號
        if idx < len(units) - 1:
            story.append(PageBreak())

    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_excel_file(combined_date_str, total_hours):
    file_path = '376431843C_1150087037_ATTACH4.xlsx'
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到範本檔案 {file_path}")

    wb = openpyxl.load_workbook(file_path)
    ws = wb['警力統計']

    for row in range(3, 7):
        cell_c = ws.cell(row=row, column=3)
        cell_c.value = combined_date_str
        cell_c.alignment = Alignment(wrapText=True, vertical='center', horizontal='left')
        
        cell_d = ws.cell(row=row, column=4)
        cell_d.value = total_hours
        cell_d.alignment = Alignment(vertical='center', horizontal='center')
        cell_d.font = Font(name='微軟正黑體', size=12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    show_sidebar()

    st.title("👵 行人及護老專案交辦單與時數統計系統")
    st.info("本系統整合了「全單位交辦單 PDF 批量產出」與「護老專案時數統計與報表匯出」功能。")

    tab1, tab2 = st.tabs(["📄 1. 各單位交辦單 PDF 產生器", "📊 2. 護老專案時數統計與匯出"])

    # ==========================================
    # TAB 1: 交辦單 PDF 產生器
    # ==========================================
    with tab1:
        st.subheader("📋 桃市警龍潭分局交通組交辦單 (PDF 全單位一次匯出)")
        st.write("已將「交辦日期」欄寬精準限制在容納「4個中文字寬」，並完美分配表格其餘空間。點擊下方按鈕即可一鍵產出包含所有單位的 PDF（共 7 頁）。")
        
        # 產生包含所有單位的 PDF
        pdf_data = generate_all_slips_pdf()
        
        st.download_button(
            label="📥 一鍵下載【全單位】正式交辦單 (PDF)",
            data=pdf_data,
            file_name="115年上半年行人及護老專案交辦單_全單位.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary"
        )

    # ==========================================
    # TAB 2: 時數統計與匯出
    # ==========================================
    with tab2:
        st.subheader("📝 勤務日期與時段設定")
        
        default_dates = []
        for m in range(1, 7):
            for d in [5, 12, 19, 26]:
                default_dates.append(f"{m}/{d}(06-10)")
                default_dates.append(f"{m}/{d}(16-20)")
                
        default_str = "、".join(default_dates)

        combined_date_str = st.text_area(
            "請確認或修改督勤日期清單 (各時段請以頓號「、」分隔)：", 
            value=default_str, 
            height=200
        )

        hours_per_shift = 4

        if combined_date_str.strip() == "":
            total_shifts = 0
        else:
            total_shifts = len(combined_date_str.split("、"))
            
        total_hours = total_shifts * hours_per_shift
        total_days = total_shifts // 2

        st.subheader("📋 運算結果預覽")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric(label="合計排定天數", value=f"{total_days} 天")
        with col2:
            st.metric(label="合計督勤班次", value=f"{total_shifts} 班")
        with col3:
            st.metric(label="總計時數 (D欄輸出值)", value=f"{total_hours} 小時")

        st.divider()

        st.subheader("📥 匯出與寄送督勤表")
        
        if total_shifts > 0:
            try:
                excel_data = generate_excel_file(combined_date_str, total_hours)
                file_name = "115年上半年督導行人及護老交通安全勤務表.xlsx"

                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        label="📥 下載督勤日期表 (Excel)",
                        data=excel_data,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col2:
                    if st.button("📧 將此報表一鍵寄至我的信箱", use_container_width=True):
                        with st.spinner("信件發送中，請稍候…"):
                            ok, mail_err = send_csv_email(excel_data, file_name)
                            if ok:
                                st.success("✅ 信件發送成功！報表已隨信夾帶至您的信箱。")
                            else:
                                st.error(f"❌ 發信失敗: {mail_err}")

            except FileNotFoundError as fnf_err:
                st.error(f"⚠️ 錯誤：{fnf_err}。請確認 `376431843C_1150087037_ATTACH4.xlsx` 檔案是否存在於主目錄中。")
            except Exception as e:
                st.error(f"⚠️ 發生未知的錯誤：{e}")
        else:
            st.warning("請先於上方輸入督勤日期，方可匯出表單。")

if __name__ == "__main__":
    main()
