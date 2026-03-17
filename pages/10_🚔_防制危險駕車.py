# ====== (中間省略，保持原有的 import 與基礎設定) ======

def auto_format_personnel(val):
    if pd.isna(val) or str(val).strip() in ["None", "nan", ""]: 
        return ""
    s = str(val).replace('：', ':').replace('、', '\n')
    # 強制時段後換行，並加粗時段
    s = re.sub(r'(\d{2}-\d{2}時)[:\s]*', r'<b>\1</b>\n', s)
    lines = [line.strip() for line in s.split('\n') if line.strip()]
    return '\n'.join(lines)

# --- 主介面邏輯 ---
st.title("🚔 防制危險駕車專案勤務規劃表")
p_time = st.text_input("勤務時間", t)
cmdr_input = st.text_input("交通快打指揮官", cmdr)

# ====== 魔法連動：單位、代號與垂直排版同步 ======
if len(ed_ptl) > 0:
    # 抓取單位
    m_unit = re.search(r'([\u4e00-\u9fa5]+(?:所|分隊|分局))', cmdr_input)
    if m_unit:
        unit_name = m_unit.group(1)
        title_name = cmdr_input.replace(unit_name, "").strip() # 剩餘職稱姓名
        
        # 同步編組單位
        ed_ptl.loc[0, '編組'] = f"專責警力\n（{unit_name}輪值）"
        
        # 同步無線電代號
        unit_map = {"石門": "隆安8", "高平": "隆安9", "聖亭": "隆安5", "龍潭": "隆安6", "中興": "隆安7", "分隊": "隆安99"}
        for k, v in unit_map.items():
            if k in unit_name:
                suffix = "1" if any(x in title_name for x in ["所長", "分隊長"]) else "2"
                ed_ptl.loc[0, '無線電'] = v + suffix
                break
        
        # 同步服勤人員垂直排版
        current_ppl = str(ed_ptl.loc[0, '服勤人員'])
        time_slots = re.findall(r'(\d{2}-\d{2}時)', current_ppl)
        if time_slots and title_name:
            new_val = ""
            for ts in time_slots:
                new_val += f"{ts}\n{title_name}\n" # 這裡會被下方的 auto_format 處理成加粗
            ed_ptl.loc[0, '服勤人員'] = new_val.strip()

# 強制套用排版引擎
if '服勤人員' in ed_ptl.columns:
    ed_ptl['服勤人員'] = ed_ptl['服勤人員'].apply(auto_format_personnel)

# ====== (後續 st.data_editor 與 PDF 生成邏輯保持一致) ======
