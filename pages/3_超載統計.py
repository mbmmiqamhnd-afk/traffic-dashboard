# --- 自動化流程 ---
        if st.session_state.get("processed_hash") != file_hash:
            with st.status("🚀 執行雲端同步與自動寄信...") as s:
                try:
                    # ==========================================
                    # 1. Google Sheets 同步 (修改版：僅寫入資料，不改格式)
                    # ==========================================
                    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
                    sh = gc.open_by_url(GOOGLE_SHEET_URL)
                    ws = sh.get_worksheet(1) # 假設是第2個工作表
                    
                    clean_cols = ['統計期間', raw_wk, raw_yt, raw_ly, '本年與去年同期比較', '目標值', '達成率']
                    
                    # 計算底部說明文字的位置 (標題佔1列 + 欄位名佔1列 + 資料列數 + 1列緩衝)
                    footer_row_idx = 1 + 1 + len(df_final) + 1
                    
                    # 準備要寫入的資料範圍
                    # A1: 標題
                    # A2: 表格內容
                    # Footer: 底部說明
                    
                    # 批次寫入資料以提升效能 (注意：這裡只更新值，不會動格式)
                    ws.update(range_name='A1', values=[['取締超載違規件數統計表']])
                    ws.update(range_name='A2', values=[clean_cols] + df_final.values.tolist())
                    ws.update(range_name=f'A{footer_row_idx}', values=[[f_plain]])
                    
                    st.write("✅ 試算表數據已更新 (保留原格式)")

                    # ==========================================
                    # 2. 自動寄信 (維持不變，Excel 附件仍保持美觀格式)
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
                    st.write(traceback.format_exc())
