def generate_pdf_from_data(unit, project, time_str, briefing, station, df_cmd, df_ptl):
    """
    修正版：將「任務編組」整合進表格第一列，並設定跨欄、邊框與底色
    """
    font = _get_font()
    buf = io.BytesIO()
    
    # 設定頁面邊距
    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=15*mm, bottomMargin=15*mm,
        title=f"{unit}執行{project}規劃表")
        
    page_width = A4[0] - 30*mm
    story = []
    
    # --- 樣式定義 ---
    style_title = ParagraphStyle('Title', fontName=font, fontSize=16, leading=22, spaceAfter=6)
    style_info = ParagraphStyle('Info', fontName=font, fontSize=12, alignment=2, spaceAfter=12) # 靠右
    style_cell = ParagraphStyle('Cell', fontName=font, fontSize=10, leading=13, alignment=1) # 置中
    style_cell_left = ParagraphStyle('CellLeft', fontName=font, fontSize=10, leading=13, alignment=0) # 靠左
    style_note = ParagraphStyle('Note', fontName=font, fontSize=11, leading=16, spaceAfter=4)
    # 表格大標題樣式 (任務編組用)
    style_table_title = ParagraphStyle('TableTitle', fontName=font, fontSize=14, alignment=1, leading=18) 

    # 1. 標題與時間
    story.append(Paragraph(f"{unit}執行{project}規劃表", style_title))
    story.append(Paragraph(f"勤務時間：{time_str}", style_info))
    
    def clean_text(txt):
        return str(txt).replace("\n", "<br/>").replace("、", "<br/>")

    # ====================
    # 2. 指揮組表格 (包含「任務編組」大標題列)
    # ====================
    col_widths_cmd = [page_width * 0.15, page_width * 0.10, page_width * 0.25, page_width * 0.50]
    headers_cmd = ["職稱", "代號", "姓名", "任務"]
    
    data_cmd = []
    
    # [Row 0] 加入「任務編組」作為表格的第一列 (需要跨 4 欄)
    # 技巧：第一個放內容，後面放空字串，稍後用 SPAN 合併
    title_cell = Paragraph("<b>任　務　編　組</b>", style_table_title)
    data_cmd.append([title_cell, '', '', '']) 
    
    # [Row 1] 欄位名稱 (職稱、代號...)
    header_row = [Paragraph(f"<b>{h}</b>", style_cell) for h in headers_cmd]
    data_cmd.append(header_row)
    
    # [Row 2+] 資料內容
    for _, row in df_cmd.iterrows():
        job = Paragraph(f"<b>{row.get('職稱','')}</b>", style_cell)
        code = Paragraph(str(row.get('代號','')), style_cell)
        name = Paragraph(clean_text(row.get('姓名','')), style_cell)
        task = Paragraph(str(row.get('任務','')), style_cell_left)
        data_cmd.append([job, code, name, task])

    # 建立表格並設定樣式
    t1 = Table(data_cmd, colWidths=col_widths_cmd, repeatRows=2) # repeatRows=2 表示換頁時重複前兩列(標題+欄位名)
    t1.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),                 # 全部有框線
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),                        # 垂直置中
        
        # --- 關鍵修正：任務編組列設定 ---
        ('SPAN', (0,0), (-1,0)),                                     # 合併第一列 (Row 0) 所有欄位
        ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#f2f2f2')),  # 第一列灰底
        ('ALIGN', (0,0), (-1,0), 'CENTER'),                          # 第一列置中
        
        # --- 欄位名稱列設定 ---
        ('BACKGROUND', (0,1), (-1, 1), colors.HexColor('#f2f2f2')),  # 第二列 (Header) 灰底
        
        # --- 調整 Padding ---
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    
    story.append(t1)
    story.append(Spacer(1, 6*mm))

    # ====================
    # 3. 勤教與檢驗站資訊
    # ====================
    story.append(Paragraph(f"<b>📢 勤前教育：</b>{briefing}", style_note))
    st_text = str(station).replace("\n", "<br/>")
    story.append(Paragraph(f"<b>🚧 檢驗站資訊：</b><br/>{st_text}", style_note))
    story.append(Spacer(1, 6*mm))

    # ====================
    # 4. 巡邏組表格
    # ====================
    col_widths_ptl = [page_width * 0.10, page_width * 0.08, page_width * 0.12, page_width * 0.18, page_width * 0.52]
    headers_ptl = ["編組", "代號", "單位", "服勤人員", "任務分工"]
    
    data_ptl = []
    # 欄位名稱
    data_ptl.append([Paragraph(f"<b>{h}</b>", style_cell) for h in headers_ptl])
    
    for _, row in df_ptl.iterrows():
        group = Paragraph(str(row.get('編組','')), style_cell)
        code = Paragraph(str(row.get('無線電','')), style_cell)
        unit_p = Paragraph(clean_text(row.get('單位','')), style_cell)
        ppl = Paragraph(clean_text(row.get('服勤人員','')), style_cell)
        
        task_text = str(row.get('任務分工',''))
        # 藍色雨備方案
        full_task = f"{task_text}<br/><font color='blue' size='9'>*雨備方案：轄區治安要點巡邏。</font>"
        task_cell = Paragraph(full_task, style_cell_left)
        
        data_ptl.append([group, code, unit_p, ppl, task_cell])

    t2 = Table(data_ptl, colWidths=col_widths_ptl, repeatRows=1)
    t2.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1, 0), colors.HexColor('#f2f2f2')), # 標題列灰底
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
