import streamlit as st
import pandas as pd
import os

# --- 1. 頁面基礎設定 ---
st.set_page_config(page_title="取締砂石車專案勤務", layout="wide")
st.title("🚛 取締砂石（大型貨）車重點違規專案勤務規劃表")

UNIT = "桃園市政府警察局龍潭分局"
NOTES = "※ 加強取締砂石（大型貨）車重點違規事項..."

# --- 2. 核心：優化後的 HTML 生成 (確保不漏掉任何單位) ---
def generate_html(month, briefing, df_cmd, df_schedule):
    font_family = "'標楷體', 'DFKai-SB', 'BiauKai', 'KaiTi', serif"
    style = f"""
    <style>
        body {{ font-family: {font_family}; color: #000; font-size: 15px; padding: 20px; }}
        .container {{ max-width: 850px; margin: auto; border: 1px solid #000; padding: 30px; line-height: 1.5; }}
        h2 {{ text-align: center; font-weight: bold; margin-bottom: 20px; }}
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; }}
        th, td {{ border: 1px solid black; padding: 8px; text-align: center; vertical-align: middle; }}
        th {{ background-color: #f2f2f2; font-weight: bold; }}
        .left {{ text-align: left; }}
        @media print {{ body {{ padding: 0; }} .container {{ border: none; }} }}
    </style>
    """
    
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    html += f"<h2>{UNIT}<br>執行{month}「取締砂石（大型貨）車重點違規」專案勤務規劃表</h2>"
    
    # 任務編組表格
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr><tr><th>職稱</th><th>代號</th><th>姓名</th><th>任務</th></tr>"
    for _, r in df_cmd.iterrows():
        html += f"<tr><td><b>{r['職稱']}</b></td><td>{r['代號']}</td><td>{str(r['姓名']).replace('、','<br>')}</td><td class='left'>{r['任務']}</td></tr>"
    html += "</table>"

    html += f"<p><b>📢 勤前教育：</b><br>{briefing.replace('\n','<br>')}</p>"

    # --- 警力佈署：修正後的 Rowspan 合併邏輯 ---
    html += "<table><tr><th colspan='4'>警　力　佈　署</th></tr>"
    html += "<tr><th width='30%'>日期</th><th width='20%'>執行單位</th><th width='10%'>人數</th><th width='40%'>執行路段</th></tr>"
    
    df_s = df_schedule.fillna("") # 確保沒有 NaN
    data = df_s.values.tolist()
    total_rows = len(data)
    
    idx = 0
    while idx < total_rows:
        date_text = str(data[idx][0]).strip()
        
        # 尋找合併範圍：只有當 date_text 有內容時才計算 rowspan
        span = 1
        if date_text != "":
            for j in range(idx + 1, total_rows):
                next_date = str(data[j][0]).strip()
                if next_date == "":
                    span += 1
                else:
                    break
        
        # 輸出這組資料
        for k in range(span):
            curr_idx = idx + k
            html += "<tr>"
            # 只有在這一組的第一列且有內容時顯示日期單元格
            if k == 0:
                html += f"<td rowspan='{span}'>{date_text if date_text != '' else '（續上項）'}</td>"
            
            html += f"<td>{data[curr_idx][1]}</td>" # 單位
            html += f"<td>{data[curr_idx][2]}</td>" # 人數
            html += f"<td class='left'>{str(data[curr_idx][3]).replace('\n','<br>')}</td>" # 路段
            html += "</tr>"
            
        idx += span # 跳轉到下一組

    html += f"</table><p><b>備註：</b><br>{NOTES}</p></div></body></html>"
    return html

# --- 3. UI 測試資料 ---
if 'cmd_data' not in st.session_state:
    st.session_state.cmd_data = pd.DataFrame([{"職稱":"指揮官","代號":"隆安1","姓名":"施宇峰","任務":"督導"}])
if 'sch_data' not in st.session_state:
    st.session_state.sch_data = pd.DataFrame([
        {"日期": "3月13日", "執行單位": "聖亭所", "執行人數": "4", "執行路段": "中豐路"},
        {"日期": "", "執行單位": "龍潭所", "執行人數": "2", "執行路段": "大昌路"},
        {"日期": "3月27日", "執行單位": "石門所", "執行人數": "3", "執行路段": "中正路"},
        {"日期": "", "執行單位": "高平所", "執行人數": "2", "執行路段": "龍源路"},
    ])

# 編輯介面
month = st.text_input("月份", "115年3月份")
brief = st.text_area("勤前教育", "地點：分局圓環")
ed_sch = st.data_editor(st.session_state.sch_data, num_rows="dynamic", use_container_width=True)

# 產生 HTML
final_html = generate_html(month, brief, st.session_state.cmd_data, ed_sch)

st.markdown("---")
st.components.v1.html(final_html, height=500, scrolling=True)

st.download_button(
    label="💾 下載報表 (確認所有單位皆顯示)",
    data=final_html.encode("utf-8"),
    file_name="勤務規劃表.html",
    mime="text/html"
)
