def d_extract_duty(d_file, hour):
    res = {'v_name': '解析失敗', 'cadre_status': '無幹部資料', 'unit_name': '未偵測單位', 'term': '該所'}
    try:
        df = pd.read_excel(d_file, header=None, dtype=str).fillna("")
        
        # 1. 偵測單位與稱謂
        for r in range(5):
            rt = "".join([str(x) for x in df.iloc[r].values])
            m = re.search(r'([\u4e00-\u9fa5]+(分局|派出所|分隊|局))', rt)
            if m: res['unit_name'] = m.group(1); break
        
        is_traffic_unit = "分隊" in res['unit_name']
        res['term'] = "該分隊" if is_traffic_unit else "該所"
        
        # 2. 人員雷達 (建立 f_map)
        full = " ".join([str(x).strip() for x in df.values.flatten() if x])
        p = r'([A-Z]|[0-9]{1,2})\s*(所長|副所長|分隊長|小隊長|巡官|巡佐|警員|警員兼副所長|實習)[\s\n]*([\u4e00-\u9fa5]{2,4})'
        matches = re.findall(p, full)
        n_map, f_map = {}, {}
        for m in matches:
            code_id, title, name = d_normalize_code(m[0]), m[1].strip(), m[2]
            if len(name) >= 2: n_map[code_id] = name; f_map[code_id] = f"{title}{name}"
            
        # 3. 偵測時間欄位與目前時段
        tr_idx, t_cols, t_col_idx = 2, {}, -1
        for r in range(6):
            tmp = {c: d_parse_time(df.iloc[r, c]) for c in range(len(df.columns)) if d_parse_time(df.iloc[r, c])[0] is not None}
            if len(tmp) > len(t_cols): tr_idx, t_cols = r, tmp
        for c, (sh, eh) in t_cols.items():
            ce = eh if eh > sh else eh + 24
            ch = hour if hour >= 6 or ce <= 24 else hour + 24
            if sh <= ch < ce: t_col_idx = c
            
        # 4. 偵測值班列 (擴大掃描)
        vr_idx = tr_idx + 1
        for r in range(tr_idx + 1, min(tr_idx + 8, len(df))):
            if "值" in "".join([str(x) for x in df.iloc[r, :4]]): vr_idx = r; break
        
        if t_col_idx != -1:
            raw = str(df.iloc[vr_idx, t_col_idx]).strip()
            mc = re.search(r'[A-Za-z0-9]{1,2}', raw)
            code_v = d_normalize_code(mc.group(0)) if mc else ""
            if code_v in f_map: res['v_name'] = f_map[code_v]
            else:
                for cid, nm in n_map.items():
                    if nm in raw: res['v_name'] = f_map[cid]; break
                    
        # 5. 🌟 積極偵測幹部動態 (解決休假列有代號的問題)
        titles_dict = {"A": "分隊長" if is_traffic_unit else "所長", 
                       "B": "小隊長" if is_traffic_unit else "副所長", 
                       "C": "幹部"}
        c_notes = []
        
        for code_c in ["A", "B", "C"]:
            full_name = f_map.get(code_c, titles_dict[code_c])
            found, is_actually_off, d_names = False, False, set()
            
            for r in range(vr_idx, len(df)):
                # 取得該時段格內容
                cell_val = str(df.iloc[r, t_col_idx]) if t_col_idx != -1 else ""
                cell_codes = [d_normalize_code(x) for x in re.findall(r'[A-Za-z0-9]{1,2}', cell_val)]
                
                # 💡 如果這格出現幹部代號
                if code_c in cell_codes:
                    found = True
                    # 檢查 A/B 欄是否有休假字眼
                    dt_area = "".join([str(x) for x in df.iloc[r, :2]])
                    row_all_text = "".join([str(x) for x in df.iloc[r, :]])
                    
                    # 判斷是否為「真休假」：寫了休假且整列都沒有勤務關鍵字
                    has_work_kw = any(kw in row_all_text for kw in ["巡","守","望","臨","交","路","督","勤","備","辦","內","專"])
                    if any(k in dt_area for k in ["休", "假", "輪", "補"]) and not has_work_kw:
                        is_actually_off = True
                    else:
                        is_actually_off = False # 強制視為在勤
                    
                    # 抓取勤務名稱
                    kw_map = {"巡":"巡邏", "守":"守望", "望":"守望", "臨":"臨檢", "交":"交整", "路":"路檢", "督":"督導", "備":"備勤", "辦":"辦公", "內":"內勤", "專":"專案"}
                    for k, kn in kw_map.items():
                        if k in row_all_text: d_names.add(kn)
            
            if not found or is_actually_off: 
                c_notes.append(f"{full_name}休假")
            else:
                if d_names:
                    sh_slot, eh_slot = t_cols.get(t_col_idx, (0, 0))
                    e_str = "24" if eh_slot in (24, 0) else f"{eh_slot % 24:02d}"
                    c_notes.append(f"{full_name}在所督勤，編排{sh_slot:02d}至{e_str}時段{'、'.join(sorted(list(d_names)))}勤務")
                else: 
                    c_notes.append(f"{full_name}在所督勤")
                    
        res['cadre_status'] = "；".join(c_notes) + "。"
    except Exception:
        res['cadre_status'] = "幹部資料解析錯誤"
    return res
