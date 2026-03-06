def sync_to_gsheet(df_a1, df_a2):
    try:
        gc = gspread.service_account_from_dict(GCP_CREDS)
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        
        # --- 更新 A1 類 (維持原樣) ---
        ws_a1 = sh.get_worksheet(2) # 索引 2，第 3 個分頁
        ws_a1.batch_clear(["A3:Z100"])
        a1_data = [[int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in row] for row in df_a1.values.tolist()]
        ws_a1.update('A3', a1_data)

        # --- 更新 A2 類 (修正：確保受傷人數在 C 欄) ---
        ws_a2 = sh.get_worksheet(3) # 索引 3，第 4 個分頁
        ws_a2.batch_clear(["A3:Z100"])
        
        # 重新整理 df_a2 的順序，確保輸出格式符合：
        # A 欄: 統計期間, B 欄: 本期件數(假設), C 欄: 本期受傷人數
        # 這裡我們根據您的需求，將 A2_Injuries_wk 強制排在第三位 (Index 2)
        a2_list = []
        for _, row in df_a2.iterrows():
            # row[0] 是單位, row[1] 是本期數值
            # 我們構造一個列表：[單位名稱, "-", row['本期(日期)']] 
            # 這樣 row['本期'] 就會落在試算表的 C 欄
            line = [
                row[0],  # A 欄: 單位
                "-",     # B 欄: 佔位
                row[1],  # C 欄: A2類受傷人數 (來自本期資料)
                row[3],  # D 欄: 本年累計
                row[4],  # E 欄: 去年累計
                row[5],  # F 欄: 比較
                row[6]   # G 欄: 增減比例
            ]
            # 數值轉換
            line = [int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in line]
            a2_list.append(line)
            
        ws_a2.update('A3', a2_list)
        
        # 套用紅字格式標題
        reqs = [get_gsheet_rich_text_req(ws_a2.id, 1, i, col) for i, col in enumerate(df_a2.columns)]
        sh.batch_update({"requests": reqs})
        
        return True, "✅ 雲端試算表同步成功 (A2受傷人數已輸出至C欄)"
    except Exception as e:
        return False, f"❌ 試算表同步失敗: {e}"
