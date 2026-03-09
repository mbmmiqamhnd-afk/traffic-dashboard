# --- 產生 HTML 預覽字串 (修正變數名稱錯誤) ---
def generate_html_content():
    style = """
    <style>
        body { font-family: 'DFKai-SB', 'BiauKai', '標楷體', serif; color: #000; }
        .container { width: 100%; max-width: 800px; margin: 0 auto; padding: 20px; }
        h2 { text-align: left; margin-bottom: 5px; letter-spacing: 2px; }
        .info { text-align: right; font-weight: bold; margin-bottom: 15px; font-size: 14px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid black; padding: 5px; text-align: center; font-size: 14px; vertical-align: middle; }
        th { background-color: #f2f2f2; }
        .left-align { text-align: left; }
    </style>
    """
    html = f"<html><head><meta charset='utf-8'>{style}</head><body><div class='container'>"
    # 使用正確的全域變數名稱
    html += f"<h2>{current_unit}執行{project_name}規劃表</h2>"
    html += f"<div class='info'>勤務時間：{plan_time}</div>"
    
    html += "<table><tr><th colspan='4'>任　務　編　組</th></tr>"
    html += "<tr><th width='15%'>職稱</th><th width='10%'>代號</th><th width='25%'>姓名</th><th width='50%'>任務</th></tr>"
    for _, row in edited_cmd.iterrows():
        name = str(row.get('姓名', '')).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        html += f"<tr><td><b>{row.get('職稱','')}</b></td><td>{row.get('代號','')}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務','')}</td></tr>"
    html += "</table>"
    
    html += f"<div class='left-align' style='margin-bottom:20px;line-height:1.6'>"
    # 這裡原本是 {briefing}，改成 {brief_info}
    html += f"<div><b>📢 勤前教育：</b>{brief_info}</div>"
    # 這裡原本是 {station}，改成 {check_st}
    html += f"<div style='white-space:pre-wrap'><b>🚧 {check_st}</b></div></div>"
    
    html += "<table><tr><th width='10%'>編組</th><th width='8%'>代號</th><th width='12%'>單位</th><th width='18%'>服勤人員</th><th width='52%'>任務分工</th></tr>"
    for _, row in edited_ptl.iterrows():
        name = str(row.get('服勤人員', '')).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        unit_cell = str(row.get("單位","")).replace("、","<br>").replace(",","<br>")
        html += f"<tr><td>{row.get('編組','')}</td><td>{row.get('無線電','')}</td><td style='line-height:1.4'>{unit_cell}</td><td style='line-height:1.4'>{name}</td><td class='left-align'>{row.get('任務分工','')}<br><span style='color:blue;font-size:0.9em'>*雨備方案：轄區治安要點巡邏。</span></td></tr>"
    html += "</table></div></body></html>"
    return html
