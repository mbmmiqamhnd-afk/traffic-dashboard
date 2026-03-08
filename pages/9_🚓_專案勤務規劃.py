import streamlit as st
import pandas as pd
from datetime import datetime

# 設定頁面資訊
st.set_page_config(page_title="專案勤務規劃", layout="wide")

st.title("🚓 專案勤務規劃表產生器")
st.caption("即時編輯勤務名單，並輸出標準格式報表")

# --- 側邊欄：全域設定 ---
with st.sidebar:
    st.header("⚙️ 勤務基本設定")
    plan_date = st.text_input("勤務日期", value="115年2月26日")
    plan_time = st.text_input("勤務時間", value="19至23時")
    unit_name = st.text_input("執行單位", value="桃園市政府警察局龍潭分局")
    project_name = st.text_input("專案名稱", value="0226「取締改裝(噪音)車輛專案監、警、環聯合稽查勤務」")

# --- 核心資料區 ---

# 1. 指揮與幕僚編組
st.subheader("1. 指揮與幕僚編組")
st.info("💡 提示：姓名若有多人，請用「、」分隔，報表輸出時會自動變為「上下並列」。")

with st.expander("📝 點此編輯【指揮官與幕僚】名單", expanded=True):
    # 預設資料：姓名使用頓號分隔
    command_data = [
        {"職稱": "指揮官", "代號": "隆安1", "姓名": "分局長 施宇峰", "任務": "核定本勤務執行並重點機動督導。"},
        {"職稱": "副指揮官", "代號": "隆安2", "姓名": "副分局長 何憶雯", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
        {"職稱": "副指揮官", "代號": "隆安3", "姓名": "副分局長 蔡志明", "任務": "襄助指揮官執行本勤務並重點機動督導。"},
        {"職稱": "上級督導官", "代號": "駐區督察", "姓名": "孫三陽", "任務": "重點機動督導。"},
        {"職稱": "督導組", "代號": "隆安6", "姓名": "組長 黃長旗、督察員 黃中彥、警務員 陳冠彰", "任務": "督導各編組服儀裝備及勤務紀律。"},
        {"職稱": "指導組", "代號": "隆安684", "姓名": "教官 郭文義", "任務": "指導各編組勤務執行及狀況處置。"},
        {"職稱": "作業及督巡組", "代號": "隆安13", "姓名": "組長 楊孟竟、警務員 盧冠仁、警務員 李峯甫、巡官 郭勝隆", "任務": "規劃勤務、督導、回報績效。"},
        {"職稱": "通訊組", "代號": "隆安", "姓名": "主任 蔡奇青、執勤官 李文章、執勤員 黃文興", "任務": "指揮、調度及通報本勤務事宜。"}
    ]
    
    df_command = st.data_editor(
        pd.DataFrame(command_data), 
        num_rows="dynamic", 
        use_container_width=True,
        key="command_editor"
    )

# 2. 勤務細節
col1, col2 = st.columns(2)
with col1:
    briefing_info = st.text_area("📢 勤前教育", value="19時30分於分局二樓會議室召開", height=100)
with col2:
    check_station = st.text_area("🚧 環保局臨時檢驗站", value="時間：20時至23時\n地點：桃園市龍潭區大昌路一段277號（龍潭區警政聯合辦公大樓）廣場", height=100)

# 3. 巡邏編組
st.subheader("2. 執行勤務編組 (巡邏組)")
st.caption("👇 直接在表格中修改人員或地點，下方會即時更新")

patrol_data = [
    {"編組": "第一巡邏組", "無線電": "隆安54", "單位": "聖亭所", "服勤人員": "巡佐傅錫城、警員曾建凱", "任務分工": "於大昌路一段周邊易有噪音車輛滋擾、聚集路段機動巡查。", "雨備": "轄區治安要點巡邏。"},
    {"編組": "第二巡邏組", "無線電": "隆安62", "單位": "龍潭所", "服勤人員": "副所長全楚文、警員龔品璇", "任務分工": "於大昌路二段周邊易有噪音車輛滋擾、聚集路段機動巡查。", "雨備": "轄區治安要點巡邏。"},
    {"編組": "第三巡邏組", "無線電": "隆安72", "單位": "中興所", "服勤人員": "副所長薛德祥、警員冷柔萱", "任務分工": "於中興路周邊易有噪音車輛滋擾、聚集路段機動巡查。", "雨備": "轄區治安要點巡邏。"},
    {"編組": "第四巡邏組", "無線電": "隆安83", "單位": "石門所", "服勤人員": "巡佐林偉政、警員盧瑾瑤", "任務分工": "於北龍路周邊易有噪音車輛滋擾、聚集路段機動巡查。", "雨備": "轄區治安要點巡邏。"},
    {"編組": "第五巡邏組", "無線電": "隆安33", "單位": "三和所/高平所", "服勤人員": "警員唐銘聰、警員張湃柏", "任務分工": "於大昌路一、二段、北龍路及中興路周邊機動巡查。", "雨備": "轄區治安要點巡邏。"},
    {"編組": "第六巡邏組", "無線電": "隆安994", "單位": "龍潭交通分隊", "服勤人員": "小隊長林振生、警員吳沛軒", "任務分工": "於大昌路一、二段、北龍路及中興路周邊機動巡查。", "雨備": "轄區治安要點巡邏。"}
]

df_patrol = st.data_editor(
    pd.DataFrame(patrol_data), 
    num_rows="dynamic", 
    use_container_width=True,
    key="patrol_editor"
)

# --- 輸出邏輯 (HTML 生成) ---
def generate_html(unit, project, date, time, briefing, station, df_cmd, df_ptl):
    style = """
    <style>
        body { font-family: 'DFKai-SB', 'BiauKai', '標楷體', serif; color: #000; }
        .container { width: 100%; max-width: 800px; margin: 0 auto; padding: 20px; }
        h2 { text-align: center; margin-bottom: 5px; }
        .info { text-align: center; font-weight: bold; margin-bottom: 15px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid black; padding: 8px; text-align: center; font-size: 14px; vertical-align: middle; }
        th { background-color: #f2f2f2; }
        .left-align { text-align: left; }
        .name-col { white-space: nowrap; } /* 避免名字被過度擠壓 */
        .rain-plan { color: blue; font-size: 0.9em; display: block; margin-top: 4px; }
    </style>
    """
    
    html = f"""
    <html>
    <head>{style}</head>
    <body>
    <div class="container">
        <h2>{unit}執行{project}規劃表</h2>
        <div class="info">勤務時間：{date} {time}</div>
    """
    
    # 第一個表格：任務編組
    html += """
        <table>
            <tr><th colspan="4">任　務　編　組</th></tr>
            <tr>
                <th width="15%">職稱</th>
                <th width="10%">代號</th>
                <th width="25%">姓名</th>
                <th width="50%">任務</th>
            </tr>
    """
    for _, row in df_cmd.iterrows():
        # 【關鍵修改】將頓號(、)、逗號(,) 取代為換行標籤 <br>
        formatted_name = str(row['姓名']).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        
        html += f"""
            <tr>
                <td><b>{row['職稱']}</b></td>
                <td>{row['代號']}</td>
                <td style="line-height: 1.5;">{formatted_name}</td>
                <td class="left-align">{row['任務']}</td>
            </tr>
        """
    html += "</table>"
    
    # 勤前教育與檢驗站
    html += f"""
        <div class="left-align" style="margin-bottom: 20px;">
            <div><b>📢 勤前教育：</b>{briefing}</div>
            <div style="margin-top:5px;"><b>🚧 {station.replace(chr(10), '<br>')}</b></div>
        </div>
    """
    
    # 第二個表格：巡邏編組
    html += """
        <table>
            <tr>
                <th width="12%">編組</th>
                <th width="10%">代號</th>
                <th width="15%">單位</th>
                <th width="20%">服勤人員</th>
                <th width="43%">任務分工 / 雨備方案</th>
            </tr>
    """
    for _, row in df_ptl.iterrows():
        # 巡邏組的人員也可能需要換行，這裡一併處理
        formatted_ptl_name = str(row['服勤人員']).replace("、", "<br>").replace(",", "<br>").replace("\n", "<br>")
        rain_text = f"<span class='rain-plan'>*雨備：{row['雨備']}</span>" if row['雨備'] else ""
        
        html += f"""
            <tr>
                <td>{row['編組']}</td>
                <td>{row['無線電']}</td>
                <td>{row['單位']}</td>
                <td style="line-height: 1.5;">{formatted_ptl_name}</td>
                <td class="left-align">
                    {row['任務分工']}
                    {rain_text}
                </td>
            </tr>
        """
    html += """
        </table>
    </div>
    </body>
    </html>
    """
    return html

# 產生 HTML
html_content = generate_html(unit_name, project_name, plan_date, plan_time, briefing_info, check_station, df_command, df_patrol)

# --- 預覽與下載區 ---
st.markdown("---")
col_preview, col_download = st.columns([3, 1])

with col_preview:
    st.subheader("📄 即時預覽")
    st.components.v1.html(html_content, height=800, scrolling=True)

with col_download:
    st.subheader("📥 輸出報表")
    st.write("點擊下方按鈕下載完整表格。")
    
    st.download_button(
        label="下載列印用報表 (.html)",
        data=html_content,
        file_name=f"勤務規劃表_{datetime.now().strftime('%Y%m%d')}.html",
        mime="text/html"
    )
    st.info("💡 姓名欄位已設定自動換行。\n(編輯時用「、」分隔即可)")
