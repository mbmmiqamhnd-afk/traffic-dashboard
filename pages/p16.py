import streamlit as st
import openpyxl
import re
import io
import json
import google.generativeai as genai
from datetime import datetime
from pdf2image import convert_from_bytes

# ==========================================
# 0. 設定與權限
# ==========================================
st.set_page_config(page_title="勤務督導報告系統", layout="wide")

# 初始化 Gemini API (請確保在 Streamlit Secret 中設定 GOOGLE_API_KEY)
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-1.5-pro')

# ==========================================
# 1. 勤務與交接解析函式 (保留您原本的邏輯)
# ==========================================
def extract_duty_v2(file, current_hour: int) -> dict:
    # ... (請貼上您原本的 extract_duty_v2 程式碼) ...
    return res

def extract_equip_v2(file) -> dict:
    # ... (請貼上您原本的 extract_equip_v2 程式碼) ...
    return equip_data

# ==========================================
# 2. Gemini Vision 刑案單解析 (核心更新)
# ==========================================
def parse_crime_pdf_gemini(pdf_file, roster: list) -> list:
    """使用 Gemini 1.5 Pro 解析刑案單"""
    try:
        pdf_file.seek(0)
        images = convert_from_bytes(pdf_file.read(), dpi=150)
        results = []
        
        roster_str = "、".join(roster)
        prompt = f"""請分析這張「刑案呈報單」影像，並以純 JSON 格式回傳以下欄位：
        {{
            "嫌疑人": "姓名",
            "查獲時間": "年月日時間",
            "查獲地點": "地址",
            "觸犯法條": "法條",
            "查獲員警": ["姓名1", "姓名2"]
        }}
        若姓名不清楚，請對照名冊比對：{roster_str}。如果看不清楚填「不詳」。"""

        for img in images:
            response = model.generate_content([prompt, img])
            # 清理 AI 可能產生的 markdown 格式
            raw = response.text.replace("```json", "").replace("```", "").strip()
            data = json.loads(raw)
            results.append(data)
    except Exception as e:
        st.error(f"Gemini 解析錯誤: {e}")
        return []
    return results

# ==========================================
# 3. UI 介面
# ==========================================
st.header("📋 勤務督導報告自動生成系統 (Gemini AI 版)")

# ... (其餘 UI 邏輯與您原本結構保持一致) ...
# 在解析 PDF 的地方，改為呼叫：
# cases = parse_crime_pdf_gemini(pdf_file, dr.get('roster', []))
