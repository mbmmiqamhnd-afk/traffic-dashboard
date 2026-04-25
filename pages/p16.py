# --- 幹部動態精準偵測邏輯：積極偵測模式 ---
for code in ["A", "B", "C"]:
    full_name = f_map.get(code, titles[code])
    found, is_off, d_names = False, False, set()
    
    for r in range(vr_idx, len(df)):
        dt_area = "".join([str(x) for x in df.iloc[r, :2]]) # A欄、B欄
        cell_val = str(df.iloc[r, t_col_idx]) if t_col_idx != -1 else ""
        cell_codes = [d_normalize_code(x) for x in re.findall(r'[A-Za-z0-9]{1,2}', cell_val)]
        
        if code in cell_codes:
            found = True
            # 💡 邏輯 A：只要格子裡有你的代號，先預設「你在上班」
            currently_working = True
            
            # 💡 邏輯 B：如果 A 欄寫休假，且該列從頭到尾都沒有任何勤務關鍵字，才判定休假
            row_all_text = "".join([str(x) for x in df.iloc[r, :]])
            is_keyword_present = any(k in row_all_text for k in ["巡","守","望","臨","交","路","督","勤","備"])
            
            if any(k in dt_area for k in ["休", "假", "輪", "補"]) and not is_keyword_present:
                is_off = True
            else:
                is_off = False # 強制判定為在勤
                
            # 💡 邏輯 C：嘗試從該列或時段格抓取勤務名稱 (擴大詞庫)
            keywords = {
                "巡": "巡邏", "守": "守望", "望": "守望", "臨": "臨檢", 
                "交": "交整", "路": "路檢", "督": "督導", "備": "備勤",
                "辦": "辦公", "內": "內勤", "專": "專案"
            }
            for k, kn in keywords.items():
                if k in dt_area or k in cell_val: 
                    d_names.add(kn)
    
    if not found or is_off: 
        c_notes.append(f"{full_name}休假")
    else:
        # 如果有代號但沒抓到關鍵字，給予保底描述「在所督勤」
        if d_names:
            sh_slot, eh_slot = t_cols.get(t_col_idx, (0,0))
            e_str = "24" if eh_slot in (24, 0) else f"{eh_slot % 24:02d}"
            c_notes.append(f"{full_name}在所督勤，編排{sh_slot:02d}至{e_str}時段{'、'.join(sorted(d_names))}勤務")
        else: 
            c_notes.append(f"{full_name}在所督勤")
