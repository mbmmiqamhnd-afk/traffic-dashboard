import streamlit as st
import openpyxl
import re
import json
import google.generativeai as genai
from datetime import datetime
from pdf2image import convert_from_bytes

# ==========================================
# 0. 設定與權限
# ==========================================
st.set_page_config(page_title="勤務督導報告系統", layout="wide")

try:
    api_key = st.secrets["api"]["GOOGLE_API_KEY"]
    genai.configure(api_key=api_key)
    
    # 自動抓取帳號下可用的模型
    all_models = [m.name for m in genai.list_models()]
    target_model = next((m for m in all_models if 'gemini-1.5-flash' in m), None)
    if not target_model:
        target_model = next((m for m in all_models if 'gemini-1.5-pro' in m), all_models[0])
    
    model = genai.GenerativeModel(target_model)
    st.sidebar.info(f"系統已連線: {target_model}")
except Exception as e:
    st.error(f"系統初始化錯誤: {e}")
    st.stop()

# ==========================================
# 1. 勤務與交接解析函式
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
# 2. Gemini Vision 刑案單解析 (加入偵錯輸出)
# ==========================================
def parse_crime_pdf_gemini(pdf_file, roster: list) -> list:
    try:
        pdf_file.seek(0)
        images = convert_from_bytes(pdf_file.read(), dpi=200)
        results = []
        roster_str = "、".join(roster)
        prompt = f"請分析刑案呈報單圖片，提取：嫌疑人, 查獲時間, 查獲地點, 觸犯法條, 查獲員警。員警名冊：{roster_str}。僅回傳乾淨的 JSON，不要 Markdown。"
        
        for img in images:
            response = model.generate_content([prompt, img])
            raw_text = response.text.replace("```json", "").replace("```", "").strip()
            
            try:
                data = json.loads(raw_text)
                results.append(data)
            except:
                st.warning(f"AI 未回傳有效 JSON。原始輸出: {raw_text}")
    except Exception as e:
        st.error(f"解析失敗: {e}")
    return results

# ==========================================
# 3. UI 介面
# ==========================================
st.header("📋 勤務督導報告系統")
num_units = st.number_input("待督導單位數量", 1, 8, 3)
u_tabs = st.tabs([f"單位 {i+1}" for i in range(num_units)] + ["匯整"])

for i in range(num_units):
    with u_tabs[i]:
        u_time = st.time_input("抵達時間", datetime.now().time(), key=f"ut_{i}")
        c1, c2, c3 = st.columns(3)
        u_duty = c1.file_uploader("勤務表", type=['xlsx'], key=f"ud_{i}")
        u_eq = c2.file_uploader("交接簿", type=['xlsx'], key=f"ue_{i}")
        u_pdf = c3.file_uploader("刑案單", type=['pdf'], accept_multiple_files=True, key=f"updf_{i}")
        
        if u_duty:
            dr = extract_duty_v2(u_duty, u_time.hour)
            lns = [f"{u_time.strftime('%H%M')}，{dr['term']}值班{dr['v_name']}。"]
            if u_pdf:
                with st.spinner("AI 正在分析刑案單..."):
                    for f in u_pdf:
                        for case in parse_crime_pdf_gemini(f, dr.get('roster', [])):
                            lns.append(f"優蹟紀錄：{dr['term']}同仁 {case.get('查獲員警','')} 於 {case.get('查獲地點','')} 查獲 {case.get('嫌疑人','')}。")
            
            st.text_area("預覽", "\n".join(lns), height=200, key=f"prev_{i}")
