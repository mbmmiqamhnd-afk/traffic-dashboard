import streamlit as st
import pandas as pd
import io
import sys
import os
import re
import smtplib
import numpy as np
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# 自動將上層目錄加入路徑
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
try:
    from app import show_sidebar
except ImportError:
    def show_sidebar():
        pass


def send_report_email_auto(files, year, month):
    try:
        if "email" not in st.secrets:
            return False, "找不到 st.secrets 中的 email 設定"
          
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
      
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = sender
        msg['Subject'] = f"【系統備份】龍潭分局 {year}年{month}月 獎勵金點數統計表暨印領清冊"
      
        body = f"郭同仁您好：\n\n系統已自動完成 {year}年{month}月份的獎勵金點數彙整與印領清冊產出。\n本次附件包含「點數統計表」與「印領清冊」共兩份 Excel 檔案，請查收。"
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
      
        for file_data, filename in files:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file_data)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(filename)}")
            msg.attach(part)
      
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)


def sort_coworkers(df):
    df = df.copy()
    df['姓名'] = df['姓名'].fillna("").astype(str).str.strip()
    df['單位'] = df['單位'].fillna("").astype(str).str.strip()
    df['職別'] = df['職別'].fillna("").astype(str).str.strip()
    df['分配類別'] = df['分配類別'].fillna("").astype(str).str.strip()
  
    if '排序調整' in df.columns:
        df['排序調整'] = pd.to_numeric(df['排序調整'], errors='coerce').fillna(999).astype(int)
    else:
        df.insert(0, '排序調整', range(100, 100 + len(df)))
  
    cat_order = ["負責管考(72%)", "勤務督導(20%)", "其他配合(8%)", ""]
    df['分配類別'] = pd.Categorical(df['分配類別'], categories=cat_order, ordered=True)
  
    unit_order = ["交通組", "會計室", "秘書室", "人事室", "龍潭分局", "勤務中心", "督察組",
                  "保安民防組", "行政組", "防治組", "聖亭派出所", "龍潭派出所", "中興派出所",
                  "石門派出所", "高平派出所", "三和派出所", "龍潭交通分隊", ""]
  
    for u in df['單位'].unique():
        if u not in unit_order:
            unit_order.append(u)
          
    df['單位'] = pd.Categorical(df['單位'], categories=unit_order, ordered=True)
  
    def get_rank_weight(title):
        title = str(title)
        if title == '分局長': return 1
        if title == '副分局長': return 2
        if any(x in title for x in ['副所長', '小隊長']): return 4
        if any(x in title for x in ['主管', '組長', '主任', '所長', '分隊長', '主計']): return 3
        if any(x in title for x in ['巡佐', '督察員', '警務員']): return 5
        if '巡官' in title: return 6
        return 7
  
    df['職級權重'] = df['職別'].apply(get_rank_weight)
  
    df.sort_values(by=['排序調整', '分配類別', '單位', '職級權重', '姓名'],
                    ascending=[True, True, True, True, True], inplace=True)
  
    df.drop(columns=['職級權重'], inplace=True, errors='ignore')
    df.reset_index(drop=True, inplace=True)
    return df


def on_data_edited():
    changes = st.session_state.co_editor
    df = st.session_state.current_roster.copy()
  
    for row_idx, updated_cols in changes.get("edited_rows", {}).items():
        for col_name, val in updated_cols.items():
            df.at[row_idx, col_name] = val
          
    for new_row in changes.get("added_rows", []):
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
      
    deleted_indices = changes.get("deleted_rows", [])
    if deleted_indices:
        df.drop(index=deleted_indices, inplace=True)
      
    st.session_state.current_roster = sort_coworkers(df)


def p18_page():
    show_sidebar()
    st.title("💰 龍潭分局 - 獎勵金點數統計表暨印領清冊產生器")
    st.info("💡 挪移順序教學：表格最左側「排序調整」欄位可手動調整順序")
    
    P_A2, P_A3, P_TRAF = 10.0, 5.0, 5.0

    # 1. 檔案上傳
    st.subheader("📂 1. 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)",
                                      type=['xlsx'], accept_multiple_files=True)
  
    # 2. 設定
    st.subheader("📝 2. 印領清冊與獎金分配設定")
    point_value = st.number_input("💵 直接執行人員 - 每點獎金金額", value=1.905, format="%.3f", step=0.001)
    target_direct_budget = st.number_input("🎯 警察局核撥【直接執行人員】總獎金目標 (元) *若為 0 則不啟動自動平帳", value=0, step=1)
  
    st.markdown("##### 👥 共同作業及配合人員 - 分配模式")
    alloc_mode = st.radio(
        "請選擇「共同作業及配合人員」的獎金計算方式：",
        ["🤖 系統自動按比例分配 (72%差異化)", "✍️ 手動輸入固定金額 (僅於總表分類顯示)"]
    )
 
    if "系統自動" in alloc_mode:
        st.info("💡 負責管考(72%)：正副主官固定8%（分局長60%、副分局長40%），其餘按56%：26%：10%比例分配。")
        budget_type = st.selectbox("請選擇預算輸入方式：", [
            "A. 直接輸入【共同作業人員】的總分配預算",
            "B. 輸入【全分局】本月核撥總預算"
        ])
        if "A" in budget_type:
            budget_input = st.number_input("💰 輸入【共同作業人員】總預算 (元)", value=10000, step=100)
        else:
            budget_input = st.number_input("💰 輸入【全分局】核撥總預算 (元)", value=50000, step=100)
    else:
        st.info("💡 手動模式：請於表格內自行填寫實際金額。")
        budget_input = 0
        budget_type = ""

    st.markdown("**共同作業名單**")
    roster_file = 'coworkers_roster.csv'

    if 'current_roster' not in st.session_state:
        if os.path.exists(roster_file):
            df_init = pd.read_csv(roster_file)
        else:
            default_coworkers_data = [
                {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "業務單位主管", "姓名": "陳維明"},
                {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "盧冠仁"},
                {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "李峯甫"},
                {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "羅千金"},
                {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "郭勝隆"},
                {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "吳享運"},
                {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "吳沛軒"},
                {"分配類別": "其他配合(8%)", "單位": "會計室", "職別": "主任", "姓名": ""},
                {"分配類別": "其他配合(8%)", "單位": "會計室", "職別": "主計", "姓名": "郭貞彣"},
                {"分配類別": "其他配合(8%)", "單位": "會計室", "職別": "主計", "姓名": "林玲宜"},
                {"分配類別": "其他配合(8%)", "單位": "秘書室", "職別": "主任", "姓名": "陳振貴"},
                {"分配類別": "其他配合(8%)", "單位": "秘書室", "職別": "出納", "姓名": "簡啟峯"},
                {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "主任", "姓名": "葉菀容"},
                {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "助理員", "姓名": "王韋翔"},
                {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "警務佐", "姓名": "李福源"},
                {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "警員", "姓名": "陳明祥"},
                {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "警員", "姓名": "黃秀吉"},
                {"分配類別": "勤務督導(20%)", "單位": "秘書室", "職別": "巡官", "姓名": "陳鵬翔"},
                {"分配類別": "勤務督導(20%)", "單位": "龍潭分局", "職別": "分局長", "姓名": "施宇峰"},
                {"分配類別": "勤務督導(20%)", "單位": "龍潭分局", "職別": "副分局長", "姓名": "何憶雯"},
                {"分配類別": "勤務督導(20%)", "單位": "龍潭分局", "職別": "副分局長", "姓名": "蔡志明"},
                {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "主任", "姓名": "游新枝"},
                {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "巡佐", "姓名": "李文章"},
                {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "巡佐", "姓名": "余清富"},
                {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "警務佐", "姓名": "陳敬霖"},
                {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "警員", "姓名": "黃文興"},
                {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "警員", "姓名": "王天龍"},
                {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "警員", "姓名": "曾嘉偉"},
                {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "警員", "姓名": "江文頌"},
                {"分配類別": "勤務督導(20%)", "單位": "督察組", "職別": "組長", "姓名": "賴永益"},
                {"分配類別": "勤務督導(20%)", "單位": "督察組", "職別": "督察員", "姓名": "黃中彥"},
                {"分配類別": "勤務督導(20%)", "單位": "督察組", "職別": "警務員", "姓名": "陳冠彰"},
                {"分配類別": "勤務督導(20%)", "單位": "督察組", "職別": "巡官", "姓名": "全楚文"},
                {"分配類別": "勤務督導(20%)", "單位": "保安民防組", "職別": "組長", "姓名": "蔡奇青"},
                {"分配類別": "勤務督導(20%)", "單位": "保安民防組", "職別": "警務員", "姓名": "曾盛鉉"},
                {"分配類別": "勤務督導(20%)", "單位": "保安民防組", "職別": "巡官", "姓名": "李立人"},
                {"分配類別": "勤務督導(20%)", "單位": "保安民防組", "職別": "巡官", "姓名": "林沛達"},
                {"分配類別": "勤務督導(20%)", "單位": "保安民防組", "職別": "巡官", "姓名": "吳國棟"},
                {"分配類別": "勤務督導(20%)", "單位": "行政組", "職別": "組長", "姓名": "周金柱"},
                {"分配類別": "勤務督導(20%)", "單位": "行政組", "職別": "巡官", "姓名": "蕭凱文"},
                {"分配類別": "勤務督導(20%)", "單位": "防治組", "職別": "組長", "姓名": "沈鳳漳"},
                {"分配類別": "勤務督導(20%)", "單位": "防治組", "職別": "巡官", "姓名": "陳冠亘"},
                {"分配類別": "勤務督導(20%)", "單位": "聖亭派出所", "職別": "所長", "姓名": "鄭榮捷"},
                {"分配類別": "勤務督導(20%)", "單位": "聖亭派出所", "職別": "副所長", "姓名": "邱品淳"},
                {"分配類別": "勤務督導(20%)", "單位": "聖亭派出所", "職別": "副所長", "姓名": "曹培翔"},
                {"分配類別": "勤務督導(20%)", "單位": "聖亭派出所", "職別": "業務承辦人", "姓名": "曾建凱"},
                {"分配類別": "勤務督導(20%)", "單位": "龍潭派出所", "職別": "所長", "姓名": "孫祥愷"},
                {"分配類別": "勤務督導(20%)", "單位": "龍潭派出所", "職別": "副所長", "姓名": "劉重言"},
                {"分配類別": "勤務督導(20%)", "單位": "龍潭派出所", "職別": "副所長", "姓名": "梁順安"},
                {"分配類別": "勤務督導(20%)", "單位": "龍潭派出所", "職別": "業務承辦人", "姓名": "周薇"},
                {"分配類別": "勤務督導(20%)", "單位": "中興派出所", "職別": "所長", "姓名": "董亦文"},
                {"分配類別": "勤務督導(20%)", "單位": "中興派出所", "職別": "副所長", "姓名": "何昀融"},
                {"分配類別": "勤務督導(20%)", "單位": "中興派出所", "職別": "副所長", "姓名": "林榮裕"},
                {"分配類別": "勤務督導(20%)", "單位": "中興派出所", "職別": "業務承辦人", "姓名": "鄧雅文"},
                {"分配類別": "勤務督導(20%)", "單位": "石門派出所", "職別": "所長", "姓名": "林育辰"},
                {"分配類別": "勤務督導(20%)", "單位": "石門派出所", "職別": "副所長", "姓名": "薛德祥"},
                {"分配類別": "勤務督導(20%)", "單位": "石門派出所", "職別": "業務承辦人", "姓名": "陳琦"},
                {"分配類別": "勤務督導(20%)", "單位": "高平派出所", "職別": "所長", "姓名": "王梓岳"},
                {"分配類別": "勤務督導(20%)", "單位": "高平派出所", "職別": "副所長", "姓名": "楊勝吉"},
                {"分配類別": "勤務督導(20%)", "單位": "高平派出所", "職別": "業務承辦人", "姓名": "黃丞潁"},
                {"分配類別": "勤務督導(20%)", "單位": "三和派出所", "職別": "所長", "姓名": "宋開國"},
                {"分配類別": "勤務督導(20%)", "單位": "三和派出所", "職別": "副所長", "姓名": "陳佶汎"},
                {"分配類別": "勤務督導(20%)", "單位": "三和派出所", "職別": "業務承辦人", "姓名": "童霂晟"},
                {"分配類別": "勤務督導(20%)", "單位": "龍潭交通分隊", "職別": "分隊長", "姓名": "卓宜澂"},
                {"分配類別": "勤務督導(20%)", "單位": "龍潭交通分隊", "職別": "小隊長", "姓名": "鄭敬思"},
                {"分配類別": "勤務督導(20%)", "單位": "龍潭交通分隊", "職別": "小隊長", "姓名": "蔡安龍"},
                {"分配類別": "勤務督導(20%)", "單位": "龍潭交通分隊", "職別": "業務承辦人", "姓名": "陳建穎"}
            ]
            df_init = pd.DataFrame(default_coworkers_data)
        st.session_state.current_roster = sort_coworkers(df_init)

    df_display = st.session_state.current_roster.copy()
  
    if "系統自動" not in alloc_mode:
        if '金額' not in df_display.columns:
            df_display.insert(5, '金額', 0)
        col_cfg = {
            "排序調整": st.column_config.NumberColumn("排序調整 🔢", help="修改數字可調整順序", min_value=1, format="%d"),
            "分配類別": st.column_config.SelectboxColumn("分配類別", options=["負責管考(72%)", "勤務督導(20%)", "其他配合(8%)"], required=True),
            "金額": st.column_config.NumberColumn("金額", min_value=0, step=1, format="%d")
        }
    else:
        if '金額' in df_display.columns:
            df_display = df_display.drop(columns=['金額'])
        col_cfg = {
            "排序調整": st.column_config.NumberColumn("排序調整 🔢", help="修改數字可調整順序", min_value=1, format="%d"),
            "分配類別": st.column_config.SelectboxColumn("分配類別", options=["負責管考(72%)", "勤務督導(20%)", "其他配合(8%)"], required=True)
        }

    st.data_editor(df_display, num_rows="dynamic", use_container_width=True, hide_index=True, height=500, column_config=col_cfg, key="co_editor", on_change=on_data_edited)

    if st.button("💾 儲存最新名單為預設值", use_container_width=True, type="secondary"):
        st.session_state.current_roster.to_csv(roster_file, index=False, encoding='utf-8-sig')
        st.success("✅ 名單已永久儲存！")
        st.rerun()

    st.markdown("---")

    if st.button("🚀 執行彙整、計算獎金與發送報表", type="primary", use_container_width=True):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種檔案皆已上傳！")
            return

        with st.spinner("正在精算比例與發放金額..."):
            try:
                # A. 資料讀取
                df_acc_raw = pd.read_excel(file_acc, header=4)
                df_acc_raw['姓名'] = df_acc_raw['姓名'].astype(str).str.strip()
                dict_acc = df_acc_raw.groupby('姓名')[['A2類', 'A3類']].sum().to_dict(orient='index')

                df_traf_all = pd.concat([pd.read_excel(f, sheet_name='月彙整總表') for f in file_traf_list])
                df_traf_all['姓名'] = df_traf_all['姓名'].astype(str).str.strip()
                dict_traf = df_traf_all.groupby('姓名')['總計尖峰時數'].sum().to_dict()

                # B. 日期偵測
                dfs_raw = pd.read_excel(file_template, sheet_name=None, header=None)
                ext_year, ext_month = "115", "4"
                found_date = False
                for _, df_scan in dfs_raw.items():
                    for r in range(min(15, len(df_scan))):
                        for c in range(min(10, len(df_scan.columns))):
                            v = str(df_scan.iloc[r, c])
                            m = re.search(r'開單日期[：:\s]*(\d{3})(\d{2})', v)
                            if m:
                                ext_year, ext_month = m.group(1), str(int(m.group(2)))
                                found_date = True
                                break
                        if found_date: break
                    if found_date: break

                # C. 直接執行人員計算
                final_sheets = {}
                summary_rows = []
                g_cite, g_acc, g_traf, g_all = 0, 0, 0, 0
                direct_exec_list = []

                for sheet_name, df in dfs_raw.items():
                    if '總表' in sheet_name: continue
                   
                    start_r, start_c = None, None
                    for r_idx, row in df.iterrows():
                        row_str = [str(x).strip() for x in row.values]
                        if '員警姓名' in row_str:
                            start_r, start_c = r_idx, row_str.index('員警姓名')
                            break
                   
                    if start_r is not None:
                        df_work = df.iloc[start_r:, start_c:].copy()
                        df_work.reset_index(drop=True, inplace=True)
                        df_work.columns = [str(c).strip() for c in df_work.iloc[0]]
                        df_work = df_work.drop(0).astype(object)
                      
                        member_rows = []
                        for r in range(len(df_work)):
                            name_cell = str(df_work.iloc[r, 0]).strip()
                            if '小計' in name_cell or '總計' in name_cell or name_cell in ['nan', 'None', '']:
                                continue
                            member_rows.append(r)
                       
                        df_members = df_work.iloc[member_rows].copy()
                        s_cite, s_acc, s_traf = 0, 0, 0
                      
                        for idx, row in df_members.iterrows():
                            name = str(row.get('員警姓名', '')).strip()
                            a2 = dict_acc.get(name, {}).get('A2類', 0)
                            a3 = dict_acc.get(name, {}).get('A3類', 0)
                            th = dict_traf.get(name, 0)
                            ap = a2 * P_A2 + a3 * P_A3
                            tp = th * P_TRAF
                            cp = pd.to_numeric(row.get('取締點數', 0), errors='coerce') or 0
                            total_pts = cp + ap + tp
                           
                            if 'A2件數' in df_members.columns: df_members.at[idx, 'A2件數'] = a2 if a2 > 0 else ""
                            if 'A3件數' in df_members.columns: df_members.at[idx, 'A3件數'] = a3 if a3 > 0 else ""
                            if '事故點數' in df_members.columns: df_members.at[idx, '事故點數'] = ap if ap > 0 else ""
                            if '交整時數' in df_members.columns: df_members.at[idx, '交整時數'] = th if th > 0 else ""
                            if '交整點數' in df_members.columns: df_members.at[idx, '交整點數'] = tp if tp > 0 else ""
                            if '個人總點數' in df_members.columns: df_members.at[idx, '個人總點數'] = total_pts
                           
                            s_cite += cp; s_acc += ap; s_traf += tp
                           
                            if total_pts > 0:
                                reward = int(np.round(total_pts * point_value))
                                direct_exec_list.append({
                                    "單位名稱": sheet_name, "員警姓名": name,
                                    "取締件數": row.get('取締件數', ''), "取締點數": cp if cp > 0 else '',
                                    "A2件數": a2 if a2 > 0 else '', "A3件數": a3 if a3 > 0 else '',
                                    "事故點數": ap if ap > 0 else '', "交整時數": th if th > 0 else '',
                                    "交整點數": tp if tp > 0 else '', "個人總點數": total_pts,
                                    "每點獎金": point_value, "實領獎金": reward, "蓋章": ""
                                })
                       
                        sub_row_data = {c: "" for c in df_work.columns}
                        sub_row_data['員警姓名'] = '小計'
                        for col_n in df_work.columns:
                            if col_n in ['員警姓名', '蓋章']: continue
                            v_sum = pd.to_numeric(df_members[col_n], errors='coerce').sum()
                            sub_row_data[col_n] = v_sum if v_sum > 0 else 0
                      
                        df_final = pd.concat([df_members, pd.DataFrame([sub_row_data])], ignore_index=True)
                        if '蓋章' in df_final.columns: df_final = df_final.drop(columns=['蓋章'])
                        final_sheets[sheet_name] = df_final
                      
                        summary_rows.append([sheet_name, s_cite, s_acc, s_traf, s_cite + s_acc + s_traf])
                        g_cite += s_cite; g_acc += s_acc; g_traf += s_traf; g_all += (s_cite + s_acc + s_traf)

                df_direct_exec = pd.DataFrame(direct_exec_list)
                if not df_direct_exec.empty:
                    df_direct_exec.insert(0, '序號', range(1, len(df_direct_exec) + 1))

                # 金額誤差處理
                direct_total_money = df_direct_exec['實領獎金'].sum() if not df_direct_exec.empty else 0
                if not df_direct_exec.empty and target_direct_budget > 0:
                    current_sum = direct_total_money
                    diff = target_direct_budget - current_sum
                    if diff != 0:
                        st.info(f"🎯 目標金額：**{target_direct_budget:,}** 元 | 目前計算：{current_sum:,} 元 | **差額 {diff:+,} 元**")
                        if abs(diff) <= 5:
                            df_direct_exec.at[0, '實領獎金'] += diff
                            st.success(f"✅ 已自動調整第 1 筆資料 {diff:+,} 元")
                        else:
                            n = len(df_direct_exec)
                            base = diff // n
                            rem = abs(diff) % n
                            sign = 1 if diff > 0 else -1
                            df_direct_exec['實領獎金'] += base
                            if rem > 0:
                                df_direct_exec.iloc[:rem, df_direct_exec.columns.get_loc('實領獎金')] += sign
                            st.success(f"✅ 已將差額 {diff:+,} 元平均分散調整")
                        direct_total_money = df_direct_exec['實領獎金'].sum()

                # D. 共同作業人員處理
                df_coworkers_work = st.session_state.current_roster.copy()
                df_coworkers_work.dropna(how='all', inplace=True)
                df_coworkers_work = sort_coworkers(df_coworkers_work)

                if "系統自動" in alloc_mode:
                    if "A" in budget_type:
                        coworker_pool = int(budget_input)
                    else:
                        coworker_pool = int(budget_input) - direct_total_money
                        if coworker_pool < 0:
                            st.error(f"❌ 全分局預算 ({budget_input}) 不足支付直接執行人員 ({direct_total_money})")
                            return
                    pool_72 = int(np.round(coworker_pool * 0.72))
                    pool_20 = int(np.round(coworker_pool * 0.20))
                    pool_08 = coworker_pool - pool_72 - pool_20
                    df_coworkers_work['核發金額'] = 0
                    mask_72 = df_coworkers_work['分配類別'] == "負責管考(72%)"
                    df_72 = df_coworkers_work[mask_72].copy()
                    df_72['核發金額'] = 0
                    if not df_72.empty and pool_72 > 0:
                        main_officers = df_72[df_72['職別'].str.contains('分局長|副分局長', na=False)].copy()
                        if not main_officers.empty:
                            pool_main = int(np.round(pool_72 * 0.08))
                            for idx, row in main_officers.iterrows():
                                title = str(row['職別'])
                                if '分局長' in title:
                                    amount = int(np.round(pool_main * 0.60))
                                elif '副分局長' in title:
                                    amount = pool_main - int(np.round(pool_main * 0.60))
                                else:
                                    amount = 0
                                df_72.at[idx, '核發金額'] = amount
                        remaining_pool = pool_72 - df_72['核發金額'].sum()
                        other_mask = ~df_72['職別'].str.contains('分局長|副分局長', na=False)
                        df_other = df_72[other_mask].copy()
                        if not df_other.empty and remaining_pool > 0:
                            def get_sub_weight(row):
                                title = str(row['職別'])
                                unit = str(row['單位'])
                                if '交通組' in unit: return 26
                                elif any(x in title for x in ['所長', '分隊長', '副所長', '小隊長']) and any(x in unit for x in ['派出所', '交通分隊']): return 56
                                elif any(x in title for x in ['承辦', '業務']) and any(x in unit for x in ['派出所', '交通分隊']): return 10
                                return 1
                            df_other['weight'] = df_other.apply(get_sub_weight, axis=1)
                            total_weight = df_other['weight'].sum()
                            if total_weight > 0:
                                exact_amounts = (df_other['weight'] / total_weight) * remaining_pool
                                int_amounts = np.floor(exact_amounts).astype(int)
                                diff_rem = int(remaining_pool - int_amounts.sum())
                                if diff_rem > 0:
                                    remainders = exact_amounts - int_amounts
                                    top_indices = remainders.nlargest(diff_rem).index
                                    int_amounts.loc[top_indices] += 1
                                df_other['核發金額'] = int_amounts
                                df_72.loc[other_mask, '核發金額'] = df_other['核發金額']
                        df_coworkers_work.loc[mask_72, '核發金額'] = df_72['核發金額']
                    
                    for cat, pool in [("勤務督導(20%)", pool_20), ("其他配合(8%)", pool_08)]:
                        cat_mask = df_coworkers_work['分配類別'] == cat
                        count = cat_mask.sum()
                        if count > 0 and pool > 0:
                            int_amount = int(np.floor(pool / count))
                            amounts = np.full(count, int_amount)
                            diff_rem = pool - amounts.sum()
                            if diff_rem > 0:
                                amounts[:diff_rem] += 1
                            df_coworkers_work.loc[cat_mask, '核發金額'] = amounts
                  
                    df_coworkers_output = df_coworkers_work.rename(columns={'核發金額': '金額'})
                else:
                    df_coworkers_output = df_coworkers_work.copy()

                if '金額' not in df_coworkers_output.columns:
                    df_coworkers_output['金額'] = 0

                # 【重點】交通組兼領人員金額合併到管考類別（顯示用）
                df_coworkers_output['顯示類別'] = df_coworkers_output['分配類別']
                df_coworkers_output.loc[df_coworkers_output['單位'] == "交通組", '顯示類別'] = "負責管考(72%)"

                # 移除原本分配類別，改用顯示類別
                df_coworkers_output = df_coworkers_output.drop(columns=['分配類別']).rename(columns={'顯示類別': '分配類別'})

                df_coworkers_output = sort_coworkers(df_coworkers_output)
                if '排序調整' in df_coworkers_output.columns:
                    df_coworkers_output = df_coworkers_output.drop(columns=['排序調整'])
                
                df_coworkers_output.insert(0, '序號', range(1, len(df_coworkers_output) + 1))
                df_coworkers_output['蓋章'] = ""

                # 小計（使用顯示後的類別）
                sub_72 = df_coworkers_output[df_coworkers_output['分配類別'] == "負責管考(72%)"]['金額'].sum()
                sub_20 = df_coworkers_output[df_coworkers_output['分配類別'] == "勤務督導(20%)"]['金額'].sum()
                sub_08 = df_coworkers_output[df_coworkers_output['分配類別'] == "其他配合(8%)"]['金額'].sum()
                coworkers_total_money = sub_72 + sub_20 + sub_08

                summary_data = [
                    {"項目": "一、直接執行人員", "金額": direct_total_money},
                    {"項目": "二、共同作業-負責管考(72%)", "金額": sub_72},
                    {"項目": "二、共同作業-勤務督導(20%)", "金額": sub_20},
                    {"項目": "二、共同作業-其他配合(8%)", "金額": sub_08},
                    {"項目": "共同作業人員小計", "金額": coworkers_total_money},
                    {"項目": "本月合計應發放", "金額": direct_total_money + coworkers_total_money},
                    {"項目": "製表人", "金額": ""}
                ]
                df_payroll_summary = pd.DataFrame(summary_data)

                # E. Excel 輸出
                pts_output = io.BytesIO()
                df_pts_summary = pd.DataFrame([['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']] + summary_rows + [['合計', g_cite, g_acc, g_traf, g_all]])
                with pd.ExcelWriter(pts_output, engine='xlsxwriter') as writer:
                    df_pts_summary.to_excel(writer, sheet_name='總表', header=False, index=False)
                    for sn, df_f in final_sheets.items():
                        df_f.to_excel(writer, sheet_name=sn, index=False)
                pts_excel_data = pts_output.getvalue()
                pts_filename = f"龍潭分局{ext_year}年{ext_month}月份_點數統計表.xlsx"

                # 印領清冊 - 只要有資料的列，整列都有格線
                payroll_output = io.BytesIO()
                with pd.ExcelWriter(payroll_output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    border_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})

                    # 直接執行人員
                    df_direct_exec.to_excel(writer, sheet_name='直接執行人員', index=False)
                    ws1 = writer.sheets['直接執行人員']
                    stamp_col = df_direct_exec.columns.get_loc('蓋章')
                    ws1.set_column(stamp_col, stamp_col, 22)
                    for r in range(len(df_direct_exec) + 1):
                        ws1.set_row(r, 38 if r > 0 else 25)
                        for c in range(len(df_direct_exec.columns)):
                            value = df_direct_exec.iloc[r-1, c] if r > 0 else df_direct_exec.columns[c]
                            ws1.write(r, c, value, border_format)

                    # 共同作業人員
                    if not df_coworkers_output.empty:
                        df_coworkers_output.to_excel(writer, sheet_name='共同作業及配合人員', index=False)
                        ws2 = writer.sheets['共同作業及配合人員']
                        stamp_col2 = df_coworkers_output.columns.get_loc('蓋章')
                        ws2.set_column(stamp_col2, stamp_col2, 22)
                        for r in range(len(df_coworkers_output) + 1):
                            ws2.set_row(r, 38 if r > 0 else 25)
                            for c in range(len(df_coworkers_output.columns)):
                                value = df_coworkers_output.iloc[r-1, c] if r > 0 else df_coworkers_output.columns[c]
                                ws2.write(r, c, value, border_format)

                    df_payroll_summary.to_excel(writer, sheet_name='獎勵金支領一覽表', index=False)

                payroll_excel_data = payroll_output.getvalue()
                payroll_filename = f"龍潭分局{ext_year}年{ext_month}月份_獎勵金印領清冊.xlsx"

                files_to_attach = [(pts_excel_data, pts_filename), (payroll_excel_data, payroll_filename)]
                ok, err = send_report_email_auto(files_to_attach, ext_year, ext_month)
              
                if ok:
                    st.success("✅ 報表產出成功！已自動寄送至信箱。")
                else:
                    st.warning(f"⚠️ 報表已產出，但郵件發送失敗: {err}")

                c5, c6 = st.columns(2)
                c5.download_button("📥 下載【點數統計表】", pts_excel_data, pts_filename, use_container_width=True)
                c6.download_button("📥 下載【印領清冊】", payroll_excel_data, payroll_filename, use_container_width=True, type="primary")

            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")


if __name__ == "__main__":
    p18_page()
