# --- 自動化流程 ---
        if st.session_state.get("processed_hash") != file_hash:
            with st.status("🚀 執行雲端同步與自動寄信...") as s:
                try:
                    # ==========================================
                    # 1. Google Sheets 同步 (修改版：新增標題列)
                    # ==========================================
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    ws = sh.get_worksheet(1) # 假設是第2個工作表
                    
                    clean_cols = ['統計期間', raw_wk, raw_yt, raw_ly, '本年與去年同期比較', '目標值', '達成率']
                    
                    # A. 寫入資料
                    # A1 寫入標題
                    ws.update(range_name='A1', values=[['取締超載違規件數統計表']])
                    # A2 開始寫入欄位名稱與資料
                    ws.update(range_name='A2', values=[clean_cols] + df_final.values.tolist())
                    
                    # B. 格式化請求 (標題 + 內文紅字)
                    reqs = []
                    
                    # (1) 標題格式：合併 A1:G1、藍色粗體、置中、字型加大
                    reqs.append({
                        "mergeCells": {
                            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 7},
                            "mergeType": "MERGE_ALL"
                        }
                    })
                    reqs.append({
                        "updateCells": {
                            "rows": [{"values": [{"userEnteredFormat": {
                                "horizontalAlignment": "CENTER",
                                "verticalAlignment": "MIDDLE",
                                "textFormat": {"foregroundColor": {"blue": 1.0}, "fontSize": 18, "bold": True}
                            }}]}],
                            "fields": "userEnteredFormat",
                            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1}
                        }
                    })

                    # (2) 內文紅字邏輯 (注意 row_idx 要 +1，因為多了一列標題)
                    # 欄位標題 (現在在第 2 列)
                    for i, t in enumerate(clean_cols[1:4], 2):
                        reqs.append(get_header_num_red_req(ws.id, 2, i, t))
                    
                    # 底部說明文字 (現在在 資料長度 + 2(標題列) + 1(緩衝) + 1(index修正) = len + 4)
                    footer_row_idx = 2 + len(df_final) + 1
                    reqs.append(get_footer_precise_red_req(ws.id, footer_row_idx, 1, f_plain))
                    
                    sh.batch_update({"requests": reqs})
                    st.write("✅ 試算表同步與格式化完成 (含標題)")

                    # ==========================================
                    # 2. 自動寄信 (修改版：Excel 增加標題)
                    # ==========================================
                    st.write("📧 正在準備郵件附件並寄信...")
                    df_sync = df_final.copy()
                    df_sync.columns = clean_cols
                    
                    df_excel_buffer = io.BytesIO()
                    
                    # 使用 ExcelWriter 引擎來製作漂亮的標題
                    with pd.ExcelWriter(df_excel_buffer, engine='xlsxwriter') as writer:
                        # 資料從第 2 列開始寫 (startrow=1，Excel index 從 0 開始算)
                        df_sync.to_excel(writer, index=False, startrow=1, sheet_name='Sheet1')
                        
                        workbook = writer.book
                        worksheet = writer.sheets['Sheet1']
                        
                        # 定義標題格式
                        title_format = workbook.add_format({
                            'bold': True,
                            'font_size': 18,
                            'align': 'center',
                            'valign': 'vcenter',
                            'font_color': 'blue'
                        })
                        
                        # 合併 A1:G1 並寫入標題
                        worksheet.merge_range('A1:G1', '取締超載違規件數統計表', title_format)
                        
                        # (選用) 調整欄寬讓它好看一點
                        worksheet.set_column('A:A', 15)
                        worksheet.set_column('B:G', 12)

                    mail_res = send_report_email(df_excel_buffer.getvalue(), f"🚛 超載報表 - {e_yt} ({prog_str})")
                    
                    if mail_res == "成功":
                        st.write("✅ 電子郵件自動寄送成功")
                    else:
                        st.error(f"❌ 郵件自動寄送失敗\n{mail_res}")

                    st.session_state["processed_hash"] = file_hash
                    st.balloons()
                    s.update(label="自動化流程處理完畢", state="complete")
                    
                except Exception as ex:
                    st.error(f"❌ 自動化流程中斷: {ex}")
                    st.write(traceback.format_exc()) # 印出詳細錯誤以便除錯
