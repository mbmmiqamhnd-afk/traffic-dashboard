# --- 工具函數 3: 同步 Google Sheets (改用名稱對應，更穩定) ---
def sync_to_gsheet_tech(df_loc, df_hour):
    try:
        if "gcp_service_account" not in st.secrets:
            return False, "❌ Secrets 遺失 GCP 設定"
        gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        sh = gc.open_by_url(GOOGLE_SHEET_URL)
        
        # --- 同步路段統計 ---
        try:
            # 建議：在您的試算表中新增一個分頁命名為「科技執法路段」
            # 或者改回 index (例如原本成功的 index 4)
            ws_loc = sh.worksheet("科技執法路段") 
        except:
            # 如果找不到該名稱的分頁，就改用 index 4 (第5張)
            ws_loc = sh.get_worksheet(4) 
            
        ws_loc.clear()
        ws_loc.update([df_loc.columns.values.tolist()] + df_loc.values.tolist())
        
        # --- 同步時段統計 ---
        try:
            # 建議：在您的試算表中新增一個分頁命名為「科技執法時段」
            ws_hour = sh.worksheet("科技執法時段")
        except:
            # 如果還是想用 index，但 index 5 不存在，我們先暫時改用 index 0 測試
            # 或是請您務必在 Google 試算表按下「+」新增一個分頁
            st.warning("⚠️ 找不到第 6 個分頁，請在試算表中按『+』新增分頁，或檢查分頁名稱。")
            return False, "❌ 同步失敗: 試算表分頁數量不足 (找不到 index 5)"
            
        ws_hour.clear()
        ws_hour.update([df_hour.columns.values.tolist()] + df_hour.values.tolist())
        
        return True, "✅ Google 試算表同步成功"
    except Exception as e:
        return False, f"❌ 同步失敗: {e}"
