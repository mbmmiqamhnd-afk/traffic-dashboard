raw_text = re.sub(r'\n```$', '', raw_text)
                
            if raw_text and raw_text != "[]" and raw_text != "{}":
                parsed = json.loads(raw_text)
                if isinstance(parsed, list):
                    results.extend(parsed)
                elif isinstance(parsed, dict):
                    results.append(parsed)
                    
        except Exception as e:
            st.warning(f"單位 {unit_idx+1} 刑案單第 {i+1} 頁辨識失敗或無資料: {e}")
            
    return results

# ==========================================
# 6. 主介面 UI 路由與封裝
# ==========================================
def p16_page():
    st.header("📋 勤務督導報告自動生成系統")
    insp_date = st.date_input("選擇督導日期", datetime.now(), key="insp_d")
    num_units = st.number_input("待督導單位數量", 1, 8, 3, key="num_u")
    u_tabs = st.tabs([f"🏢 單位 {i+1}" for i in range(num_units)] + ["📄 總匯整報告"])

    for i in range(num_units):
        with u_tabs[i]:
            u_time = st.time_input("抵達時間", datetime.now().time(), key=f"ut_{i}")
            
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1: u_duty = st.file_uploader(f"單位 {i+1} 勤務表 (XLSX)", type=['xlsx'], key=f"ud_{i}")
            with col_f2: u_eq = st.file_uploader(f"單位 {i+1} 交接簿 (XLSX)", type=['xlsx'], key=f"ue_{i}")
            with col_f3: u_pdf = st.file_uploader(f"單位 {i+1} 刑案單優蹟 (PDF選填)", type=['pdf'], key=f"up_{i}")
            
            # 強制加回按鈕觸發，按下一口氣完成，避免上傳中途 Reactive 機制卡死
            if st.button(f"🚀 開始產生 單位 {i+1} 督導報告", key=f"btn_run_{i}"):
                if not (u_duty and u_eq):
                    st.error("⚠️ 勤務表與交接簿為必備檔案，請確認已上傳！")
                else:
                    dr = extract_duty_v2(u_duty, u_time.hour)
                    er = extract_equip_v2(u_eq)
                    
                    if not er:
                        er = {'gi':0, 'go':0, 'bi':0, 'bo':0, 'ri':0, 'ro':0, 'vi':0, 'vo':0}
                        
                    t, loc = dr['term'], dr['loc_term']
                    d_e = insp_date - timedelta(days=1)
                    d_5, d_3 = (insp_date - timedelta(days=5)), (insp_date - timedelta(days=3))
                    
                    if dr['v_name'] == "該時段無值班人員":
                        line_1 = f"{u_time.strftime('%H%M')}，{t}該時段無值班人員。"
                    else:
                        line_1 = f"{u_time.strftime('%H%M')}，{t}值班{dr['v_name']}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。"
                    
                    lns = [
                        line_1,
                        f"{t}{'駐地監錄設備及天羅地網系統' if dr['has_skyline'] else '駐地監錄設備'}均運作正常，無故障，{d_5.strftime('%m月%d日')}至{d_e.strftime('%m月%d日')}有逐日檢測2次以上紀錄。",
                        f"{t}{d_3.strftime('%m月%d日')}至{d_e.strftime('%m月%d日')}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均點閱。",
                        f"{t}環境內務擺設整齊清潔，符合規定。",
                        f"{t}手槍出勤 {er['go']} 把、在{loc} {er['gi']} 把，子彈出勤 {er['bo']} 顆、在{loc} {er['bi']} 顆，無線電出勤 {er['ro']} 臺、在{loc} {er['ri']} 臺；防彈背心出勤 {er['vo']} 件、在{loc} {er['vi']} 件，幹部對械彈每日檢查管制良好，符合規定。",
                        f"本日{dr['cadre_status']}",
                        f"{t}酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。"
                    ]
                    
                    if dr['is_guard_unit']:
                        lns.append(f"拘留室值班警員{dr['detention_name']}，對人犯監控良好，無異常狀況發生。" if dr['detention_name'] else "拘留室目前無人犯。")
                    
                    # 只有在上傳 PDF 時才載入 AI 模組，並整合成第 8 點以後
                    if u_pdf:
                        try:
                            with st.spinner(f"單位 {i+1} 刑案單優蹟影像全速分析中..."):
                                cases = parse_crime_pdf_gemini(u_pdf, dr.get('roster', []), i)
                            if cases:
                                for case in cases:
                                    officers = case.get('查獲員警', '')
                                    if isinstance(officers, list):
                                        officers = "、".join(officers)
                                    case_time = case.get('查獲時間', '')
                                    case_loc = case.get('查獲地點', '')
                                    suspect = case.get('嫌疑人', '')
                                    crime = case.get('觸犯法條', '')
                                    
                                    lns.append(f"優蹟紀錄：{dr['unit_name']}同仁 {officers} 於 {case_time} 在 {case_loc} 查獲 {suspect} 涉嫌 {crime} 案。")
                            else:
                                st.warning("⚠️ 刑案單已上傳，但 AI 未能提取出任何有效資料。")
                        except Exception as ai_err:
                            st.error(f"AI 辨識發生預期外錯誤: {ai_err}")

                    final_lines = []
                    for idx, line in enumerate(lns):
                        final_lines.append(f"{idx+1}、{line}")
                        
                    final_text = "\n".join(final_lines)
                    st.session_state.unit_reports[i] = f"【{dr['unit_name']} 督導報告】\n{final_text}"
                    
                    if "中斷" in dr['cadre_status'] or "失敗" in dr['v_name']:
                        st.error(f"⚠️ {dr['unit_name']} 解析可能不完全：{dr['cadre_status']}")
                    else:
                        st.success(f"✅ {dr['unit_name']} 報告輸出完成")

            # 確保切換分頁時，已產生的預覽報告能透過 Session State 留存
            if i in st.session_state.unit_reports:
                st.text_area("預覽報告", st.session_state.unit_reports[i], height=350, key=f"preview_{i}")

    # ==========================================
    # 5. 總匯整報告分頁與寄信
    # ==========================================
    with u_tabs[-1]:
        reports_list = [st.session_state.unit_reports[k] for k in sorted(st.session_state.unit_reports.keys()) if k < num_units]
        if reports_list:
            full_text = ("\n\n" + "─" * 40 + "\n\n").join(reports_list)
            st.subheader("📋 匯整結果")
            st.text_area("匯整文本", full_text, height=600)
            target_mail = st.text_input("收件信箱", "mbmmiqamhnd@gmail.com")
            if st.button("🚀 立即寄送郵件"):
                if send_gmail(f"勤務督導報告匯整_{insp_date.strftime('%Y%m%d')}", full_text, target_mail):
                    st.success(f"✅ 郵件發送成功")
        else:
            st.warning("請先於各單位分頁上傳檔案並點擊「產生督導報告」。")

if __name__ == "__main__":
    p16_page()
