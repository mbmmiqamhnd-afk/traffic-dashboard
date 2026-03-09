# --- 字型 & PDF & 寄信函數 (全新改良版) ---
def _get_font():
    fname = "kaiu"
    if fname in pdfmetrics.getRegisteredFontNames():
        return fname
    # 嘗試載入字型
    font_paths = ["kaiu.ttf", "./kaiu.ttf", "font/kaiu.ttf", "C:/Windows/Fonts/kaiu.ttf"]
    font_path = None
    for p in font_paths:
        if os.path.exists(p):
            font_path = p
            break   
    if font_path:
        try:
            pdfmetrics.registerFont(TTFont(fname, font_path))
            return fname
        except Exception:
            pass
    return "Helvetica" # 若無中文字型則回退，避免當機

def generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    """
    直接根據資料生成 PDF，確保格式與 HTML 完全一致
    """
    font = _get_font()
    buf = io.BytesIO()
    
    # 頁面邊距設定 (與 HTML padding 類似)
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=15*mm, bottomMargin=15*mm,
        title=f"{unit}執行{project}規劃表")
        
    # 計算可用寬度
    page_width = A4[0] - 30*mm
    
    story = []
    
    # --- 定義樣式 ---
    # 標題 (H2)
    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, leading=22, spaceAfter=6)
    # 右上角時間 (Info)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=12)
    # 表格內文字 (一般)
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=10, leading=13, alignment=1) # 置中
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=10, leading=13, alignment=0) # 靠左
    # 勤教與檢驗站 (Note)
    style_note = ParagraphStyle('Note', fontName=font, fontSize=11, leading=16, spaceAfter=4)

    # 1. 標題與時間
    story.append(Paragraph(f"{unit}執行{project}規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    # helper: 處理文字換行與格式
    def clean_text(txt):
        return str(txt).replace("\n", "<br/>").replace("、", "<br/>")

    # 2. 指揮組表格
    # 定義欄位標題
    headers_cmd = ["職稱", "代號", "姓名", "任務"]
    # 欄寬比例: 15%, 10%, 25%, 50%
    col_widths_cmd = [page_width * 0.15, page_width * 0.10, page_width * 0.25, page_width * 0.50]
    
    data_cmd = []
    # 表頭 (跨欄標題)
    title_row = [Paragraph("<b>任　務　編　組</b>", style_cell)]
    # 欄位名稱
    header_row = [Paragraph(f"<b>{h}</b>", style_cell) for h in headers_cmd]
    
    data_cmd.append(header_row) # 注意：這裡我們分開處理 Title Row，ReportLab Table 比較難做 colspan，這裡簡化直接放 Header
    
    for _, row in df_cmd.iterrows():
        job = Paragraph(f"<b>{row.get('職稱','')}</b>", style_cell)
        code = Paragraph(str(row.get('代號','')), style_cell)
        name = Paragraph(clean_text(row.get('姓名','')), style_cell)
        task = Paragraph(str(row.get('任務','')), style_cell_left) # 任務靠左
        data_cmd.append([job, code, name, task])

    # 建立表格
    t1 = Table(data_cmd, colWidths=col_widths_cmd, repeatRows=1)
    t1.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#f2f2f2')), # 表頭灰底
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    
    # 模擬 HTML 的 "任務編組" 大標題 (用另一個表格或直接文字)
    story.append(Paragraph("<b>任　務　編　組</b>", ParagraphStyle('THead', fontName=font, fontSize=12, alignment=1, spaceAfter=2)))
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    # 3. 勤教與檢驗站資訊
    story.append(Paragraph(f"<b>📢 勤前教育：</b>{briefing}", style_note))
    # 處理檢驗站換行
    st_text = str(station).replace("\n", "<br/>")
    story.append(Paragraph(f"<b>🚧 檢驗站資訊：</b><br/>{st_text}", style_note))
    story.append(Spacer(1, 6*mm))

    # 4. 巡邏組表格
    # 定義欄位
    headers_ptl = ["編組", "代號", "單位", "服勤人員", "任務分工"]
    # 欄寬比例: 10%, 8%, 12%, 18%, 52%
    col_widths_ptl = [page_width * 0.10, page_width * 0.08, page_width * 0.12, page_width * 0.18, page_width * 0.52]
    
    data_ptl = []
    data_ptl.append([Paragraph(f"<b>{h}</b>", style_cell) for h in headers_ptl])
    
    for _, row in df_ptl.iterrows():
        group = Paragraph(str(row.get('編組','')), style_cell)
        code = Paragraph(str(row.get('無線電','')), style_cell)
        unit_p = Paragraph(clean_text(row.get('單位','')), style_cell)
        ppl = Paragraph(clean_text(row.get('服勤人員','')), style_cell)
        
        # 處理任務與雨備 (藍色字體)
        task_text = str(row.get('任務分工',''))
        # 在 ReportLab 中使用 <font color='blue'> 標籤
        full_task = f"{task_text}<br/><font color='blue' size='9'>*雨備方案：轄區治安要點巡邏。</font>"
        task_cell = Paragraph(full_task, style_cell_left)
        
        data_ptl.append([group, code, unit_p, ppl, task_cell])

    t2 = Table(data_ptl, colWidths=col_widths_ptl, repeatRows=1)
    t2.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#f2f2f2')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    
    story.append(t2)

    try:
        doc.build(story)
        return buf.getvalue()
    except Exception as e:
        print(f"PDF Build Error: {e}")
        return None

def send_report_email(html_content, subject, unit, time_str, project, briefing, station, df_cmd, df_ptl):
    """
    修改後的寄信函數，接收原始資料以生成高品質 PDF
    """
    import urllib.parse as _ul
    try:
        sender   = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        receiver = sender 
        
        # 改用新的 PDF 生成函數
        pdf_bytes = generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl)
        
        if pdf_bytes is None:
            return False, "PDF 生成失敗"

        msg = MIMEMultipart()
        msg["From"]    = sender
        msg["To"]      = receiver
        msg["Subject"] = subject
        msg.attach(MIMEText("請見附件 PDF 報表。\n\n本郵件由雲端勤務系統自動發送。", "plain", "utf-8"))
        
        part = MIMEBase("application", "pdf")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        encoded_name = _ul.quote(f"{subject}.pdf", safe='')
        part.add_header(
            "Content-Disposition",
            f"attachment; filename=\"report.pdf\"; filename*=UTF-8''{encoded_name}"
        )
        msg.attach(part)
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, receiver, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)
