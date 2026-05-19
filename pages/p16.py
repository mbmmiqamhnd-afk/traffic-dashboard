import streamlit as st
import openpyxl
import re
import json
import google.generativeai as genai
from datetime import datetime
from pdf2image import convert_from_bytes

# ==========================================
# 0. 設定與權限 (付費解鎖版：全速 2.5-flash)
# ==========================================
st.set_page_config(page_title="勤務督導報告系統", layout="wide")

try:
    api_key = st.secrets["api"]["GOOGLE_API_KEY"]
    genai.configure(api_key=api_key)
    
    model = genai.GenerativeModel('gemini-2.5-flash')
    st.sidebar.success("✅ 目前就緒: gemini-2.5-flash (付費全速模式)")

except Exception as e:
    st.error(f"系統初始化失敗: {e}")
    st.stop()

# ==========================================
# 1. 勤務與交接解析
# ==========================================
def extract_duty_v2(file, current_hour: int) -> dict:
    try:
        wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
        title_cell = str(all_rows[0][0]) if all_rows[0][0] else ''
        m_term = re.search(r'龍潭分局\s*([\u4e00-\u9fa5]+派出所|[\u4e00-\u9fa5]+隊)', title_cell)
        term = m_term.group(1) if m_term else '本所'
        code_map = {}
        for row in all_rows[43:48]:
            for grp in range(6):
                base = 1 + grp * 6
                if base + 1 >= len(row): break
                code, name_cell = row[base], row[base + 1]
                if code and name_cell and isinstance(name_cell, str):
                    names = re.findall(r'[\u4e00-\u9fa5]{2,4}', name_cell)
                    if names: code_map[str(code).strip()] = names[-1]
        
        time_headers = list(all_rows[2][13:])
        col = next((13 + i for i, h in enumerate(time_headers) if h and str(h).split('\n')[0].strip() == str(current_hour)), None)
        v_code = str(all_rows[3][col]).strip() if col is not None and all_rows[3][col] else ''
        return {'term': term, 'v_name': code_map.get(v_code, f'番號{v_code}'), 'roster': list(code_map.values())}
    except Exception as e:
        return {'term': '本所', 'v_name': '（解析失敗）', 'roster': [], '_error': str(e)}

# ==========================================
# 2. Gemini Vision (精準提取職稱)
# ==========================================
def parse_crime_pdf_gemini(pdf_file, roster: list) -> list:
    pdf_file.seek(0)
    images = convert_from_bytes(pdf_file.read(), dpi=150)
    results = []
    roster_str = "、".join(roster)
    
    prompt = f"請提取：嫌疑人, 查獲時間, 查獲地點, 觸犯法條, 查獲員警(請完整提取「職稱+姓名」，例如「警員蕭漢祥」、「巡佐董倢亨」)。名冊供比對錯別字參考：{roster_str}。僅回傳標準 JSON。"
    
    total_pages = len(images)
    
    for i, img in enumerate(images):
        try:
            st.info(f"AI 正在全速辨識第 {i+1}/{total_pages} 頁...")
            
            response = model.generate_content([prompt, img])
            raw_text = response.text.replace("```json", "").replace("```", "").strip()
            
            if raw_text and raw_text != "{}":
                results.append(json.loads(raw_text))
                
        except Exception as e:
            st.error(f"第 {i+1} 頁辨識失敗: {e}")
            
    return results

# ==========================================
# 3. UI 介面與完整報告框架
# ==========================================
st.header("📋 勤務督導報告自動生成系統")

# 增加日期與人員的輸入欄位，讓報告更完整
col1, col2, col3 = st.columns(3)
with col1:
    u_date = st.date_input("督導日期", datetime.now())
with col2:
    u_time = st.time_input("抵達時間", datetime.now().time())
with col3:
    u_inspector = st.text_input("督考人員", "交通組 ")

u_duty = st.file_uploader("勤務表 (XLSX)", type=['xlsx'])
u_pdf = st.file_uploader("刑案單 (PDF)", type=['pdf'])

if u_duty and u_pdf:
    if st.button("開始 AI 辨識"):
        dr = extract_duty_v2(u_duty, u_time.hour)
        
        with st.spinner("AI 影像全速分析中..."):
            cases = parse_crime_pdf_gemini(u_pdf, dr.get('roster', []))
        
        # 轉換為民國年
        tw_year = u_date.year - 1911
        date_str = f"{tw_year}年{u_date.month:02d}月{u_date.day:02d}日"
        time_str = u_time.strftime('%H時%M分')
        
        # 建立完整督導報告框架
        lns = [
            "【勤務督考報告】",
            f"一、受考單位：{dr['term']}",
            f"二、督考時間：{date_str} {time_str}",
            f"三、督考人員：{u_inspector}",
            "四、督考情形：",
            f"（一）{u_time.strftime('%H%M')}，{dr['term']}值班{dr['v_name']}，員警服裝儀容整齊，應對進退得宜，駐地內外環境整潔，各項勤務均能按表操課。",
            "（二）優蹟紀錄："
        ]
        
        if cases:
            # 依序填入 AI 辨識出的優蹟紀錄
            for idx, case in enumerate(cases):
                officers = case.get('查獲員警', '')
                if isinstance(officers, list):
                    officers = "、".join(officers)
                    
                case_time = case.get('查獲時間', '')
                case_loc = case.get('查獲地點', '')
                suspect = case.get('嫌疑人', '')
                crime = case.get('觸犯法條', '')
                
                lns.append(f"      {idx+1}. {dr['term']}同仁 {officers} 於 {case_time} 在 {case_loc} 查獲 {suspect} 涉嫌 {crime} 案。")
        else:
            lns.append("      無特殊優蹟紀錄。")
        
        st.text_area("報告預覽", "\n".join(lns), height=400)
