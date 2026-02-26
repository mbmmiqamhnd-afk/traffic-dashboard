# --- Google Sheets 同步函式 (針對指定網址優化) ---
def sync_to_gsheet(df_a1, df_a2):
    try:
        # 使用 Secrets 中的 GCP 憑證
        gc = gspread.service_account_from_dict(GCP_CREDS)
        # 開啟指定的試算表
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        
        def update_sheet_values(ws_index, df, title_text):
            # ws_index 2 為 A1類統計, 3 為 A2類統計
            ws = sh.get_worksheet(ws_index)
            
            # 1. 清除舊資料 (從 A3 開始)
            ws.batch_clear(["A3:Z100"])
            
            # 2. 更新標題 (A1 儲存格)
            ws.update_acell('A1', title_text)
            
            # 3. 準備寫入資料 (將 DataFrame 轉為列表，處理數值型態)
            data_rows = [[int(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else x for x in row] for row in df.values.tolist()]
            
            # 4. 寫入資料 (從 A3 開始，保留 A2 的原始欄位標題)
            if data_rows:
                ws.update('A3', data_rows)
            
            # 5. 套用紅字格式 (呼叫原有的 Rich Text 請求)
            reqs = [get_gsheet_rich_text_req(ws.id, 1, col_idx, col_name) for col_idx, col_name in enumerate(df.columns)]
            if reqs:
                sh.batch_update({"requests": reqs})
            
            return True

        # 執行同步 (分頁索引需與您原有的版本一致)
        update_sheet_values(2, df_a1, "A1類交通事故死亡人數統計表")
        update_sheet_values(3, df_a2, "A2類交通事故受傷人數統計表")
        
        return True, "✅ 指定雲端試算表同步成功 (格式已保留)"
    except Exception as e:
        return False, f"❌ 試算表同步失敗: {e}"
