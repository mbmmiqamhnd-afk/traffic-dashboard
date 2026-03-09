import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import os
import io

# --- 1. 基礎設定 ---
st.set_page_config(page_title="行人及護老交通安全", layout="wide")
st.title("🚶 行人及護老交通安全專案勤務規劃表")

UNIT = "桃園市政府警察局龍潭分局"
SHEET_ID = "1dOrFjewsdpTGy0JyBJXmuBhr8p_LSpSb6Lp2gC39KK0"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# 預設資料 (供測試用)
DEFAULT_MONTH = "115年3月份"
DEFAULT_CMD = pd.DataFrame([
    {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
    {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
])

# 這裡刻意讓日期相同，以測試合併效果
DEFAULT_SCHEDULE = pd.DataFrame([
    {"日期（6時至10時、16時至20時）": "3月2～31日(上班日)", "單位": "聖亭派出所", "路段": "中豐路、聖亭路段"},
    {"日期（6時至10時、16時至20時）": "3月2～31日(上班日)", "單位": "龍潭派出所", "路段": "中豐路、中正路段"},
    {"日期（6時至10時、16時至20時）": "3月2～31日(上班日)", "單位": "中興派出所", "路段": "中興路、福龍路段"},
    {"日期（6時至10時、16時至20時）": "其他日期", "單位": "警備隊", "路段": "轄區易肇事路口"}
])

# --- 2. HTML 生成函式 (核心：自動計算 rowspan) ---
def generate_html(month, df_c, df_s):
    font_family = "'標楷體', 'DFKai-SB', 'BiauKai', 'KaiTi', serif"
    style = f"""
    <style>
        body {{ font-family: {font_family}; color: #000; padding: 20px; }}
        .container {{ max-width: 900px; margin: auto; border: 1px solid #000; padding: 30px; }}
        h2 {{ text-align: center; font-weight: bold; font-size: 24px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
        th, td {{ border: 1px solid black; padding: 10px; text-align: center; font-size: 16px; }}
        th {{ background-color: #f2f2f2; font-weight: bold; }}
        .left {{ text-align: left; padding-left: 10px; }}
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'><h2>{UNIT}<br>{month}執行「行人及護老交通安全」專案勤務規劃表</h2>"

    # --- 任務編組表 ---
    html += "<table><tr><th colspan='4'>任 務 編 組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>"
    for _, r in df_c.iterrows():
        html += f"<tr><td><b>{r['職稱']}</b></td><td>{r['代號']}</td><td>{str(r['姓名']).replace('、','<br>')}</td><td class='left'>{r['任務']}</td></tr>"
    html += "</table><br>"

    # --- 警力佈署表 (關鍵日期合併邏輯) ---
    html += "<table><tr><th colspan='3'>警 力 佈 署</th></tr><tr><th width='35%'>日期（時段）</th><th width='20%'>單位</th><th width='45%'>路段</th></tr>"
    
    # 建立日期清單
    date_col = df_s.columns[0]
    dates = df_s[date_col].astype(str).tolist()
    
    i = 0
    while i < len(dates):
        current_date = dates[i]
        # 計算此日期連續出現幾次
        span = 1
        for j in range(i + 1, len(dates)):
            if dates[j] == current_date and current_date != "" and current_date != "nan":
                span += 1
            else:
                break
        
        # 產生對應的行
        for k in range(span):
            row = df_s.iloc[i + k]
            html += "<tr>"
            # 只有同一組的第一列需要加上 rowspan
            if k == 0:
                html += f"<td rowspan='{span}'>{current_date}</td>"
            
            html += f"<td>{row['單位']}</td><td class='left'>{str(row['路段']).replace('\\n','<br>').replace('\n','<br>')}</td>"
            html += "</tr>"
        i += span # 跳到下一組日期

    html += "</table></div></body></html>"
    return html

# --- 3. 介面與下載 ---
cur_month = st.text_input("月份標題", DEFAULT_MONTH)
st.subheader("編輯警力佈署 (第一欄相同日期會自動合併)")
ed_sch = st.data_editor(DEFAULT_SCHEDULE, num_rows="dynamic", use_container_width=True)

final_html = generate_html(cur_month, DEFAULT_CMD, ed_sch)

st.markdown("---")
st.components.v1.html(final_html, height=500, scrolling=True)

st.download_button(
    label="💾 下載標楷體 HTML (日期已合併)",
    data=final_html,
    file_name=f"勤務規劃表_{cur_month}.html",
    mime="text/html",
    use_container_width=True
)
