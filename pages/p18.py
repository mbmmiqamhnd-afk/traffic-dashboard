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

def p18_page():
    show_sidebar()

    st.title("💰 龍潭分局 - 獎勵金點數統計表暨印領清冊產生器")
    st.info("權重已固定 (A2:10, A3:5, 交整:5)。系統支援【管考72% / 督導20% / 其他8%】依照金額比例自動精算！")

    P_A2, P_A3, P_TRAF = 10.0, 5.0, 5.0

    # 1. 檔案上傳區
    st.subheader("📂 1. 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)", type=['xlsx'], accept_multiple_files=True)
    
    # 2. 印領清冊參數與名單設定
    st.subheader("📝 2. 印領清冊與獎金分配設定")
    point_value = st.number_input("💵 直接執行人員 - 每點獎金金額", value=1.905, format="%.3f", step=0.001)

    st.markdown("##### 👥 共同作業及配合人員 - 分配模式")
    alloc_mode = st.radio(
        "請選擇「共同作業及配合人員」的獎金計算方式：",
        ["🤖 系統自動依比例分配 (管考72%、督導20%、其他8%)", "✍️ 手動輸入固定金額 (僅於總表分類顯示)"]
    )
    
    # 根據選擇模式切換顯示與設定
    if "系統自動" in alloc_mode:
        st.info("💡 系統會依您設定的總預算切成三塊 (72/20/8)，再按照下方名單內的「金額」做為比例基礎，全自動平均拆分。")
        budget_type = st.selectbox("請選擇預算輸入方式：", [
            "A. 直接輸入【共同作業人員】的總分配預算", 
            "B. 輸入【全分局】本月核撥總預算 (系統會自動先扣掉直接執行人員的總獎金)"
        ])
        
        if "A" in budget_type:
            budget_input = st.number_input("💰 輸入【共同作業人員】總預算 (元)", value=10000, step=100)
        else:
            budget_input = st.number_input("💰 輸入【全分局】核撥總預算 (元)", value=50000, step=100)
    else:
        st.info("💡 系統將直接使用您在下方表格填寫的實際金額進行發放。")
        budget_input = 0
        budget_type = ""

    st.markdown(f"**共同作業名單 (已切換為純金額模式)**")
    
    # 完整名單 (將原來的「設定數值」改回「金額」)
    default_coworkers_data = [
        # --- 負責管考 (72%) ---
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "業務單位主管", "姓名": "陳維明", "金額": 298, "蓋章": ""},
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "盧冠仁", "金額": 298, "蓋章": ""},
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "李峯甫", "金額": 298, "蓋章": ""},
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "羅千金", "金額": 298, "蓋章": ""},
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "郭勝隆", "金額": 298, "蓋章": ""},
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "吳享運", "金額": 232, "蓋章": ""},
        {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "吳沛軒", "金額": 232, "蓋章": ""},
        
        # --- 其他配合 (8%)：會計室、人事室，以及秘書室的主任與出納 ---
        {"分配類別": "其他配合(8%)", "單位": "會計室", "職別": "主計", "姓名": "郭貞彣", "金額": 77, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "會計室", "職別": "主計", "姓名": "林玲宜", "金額": 78, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "秘書室", "職別": "主任", "姓名": "陳振貴", "金額": 78, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "秘書室", "職別": "出納", "姓名": "簡啟峯", "金額": 78, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "主任", "姓名": "葉菀容", "金額": 78, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "助理員", "姓名": "王韋翔", "金額": 77, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "警務佐", "姓名": "李福源", "金額": 77, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "警員", "姓名": "陳明祥", "金額": 77, "蓋章": ""},
        {"分配類別": "其他配合(8%)", "單位": "人事室", "職別": "警員", "姓名": "黃秀吉", "金額": 77, "蓋章": ""},

        # --- 勤務督導 (20%)：包含秘書室巡官與其他各單位 ---
        {"分配類別": "勤務督導(20%)", "單位": "秘書室", "職別": "巡官", "姓名": "陳鵬翔", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭分局", "職別": "分局長", "姓名": "施宇峰", "金額": 301, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭分局", "職別": "副分局長", "姓名": "何憶雯", "金額": 100, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭分局", "職別": "副分局長", "姓名": "蔡志明", "金額": 100, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "主任", "姓名": "游新枝", "金額": 65, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "巡佐", "姓名": "李文章", "金額": 65, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "巡佐", "姓名": "余清富", "金額": 65, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "警務佐", "姓名": "陳敬霖", "金額": 65, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "警員", "姓名": "黃文興", "金額": 65, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "警員", "姓名": "王天龍", "金額": 65, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "警員", "姓名": "曾嘉偉", "金額": 65, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "警員", "姓名": "江文頌", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "督察組", "職別": "組長", "姓名": "賴永益", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "督察組", "職別": "督察員", "姓名": "黃中彥", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "督察組", "職別": "警務員", "姓名": "陳冠彰", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "督察組", "職別": "巡官", "姓名": "全楚文", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "保安民防組", "職別": "組長", "姓名": "蔡奇青", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "保安民防組", "職別": "警務員", "姓名": "曾盛鉉", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "保安民防組", "職別": "巡官", "姓名": "李立人", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "保安民防組", "職別": "巡官", "姓名": "林沛達", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "保安民防組", "職別": "巡官", "姓名": "吳國棟", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "行政組", "職別": "組長", "姓名": "周金柱", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "行政組", "職別": "巡官", "姓名": "蕭凱文", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "防治組", "職別": "組長", "姓名": "沈鳳漳", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "防治組", "職別": "巡官", "姓名": "陳冠亘", "金額": 64, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "聖亭派出所", "職別": "所長", "姓名": "鄭榮捷", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "聖亭派出所", "職別": "副所長", "姓名": "邱品淳", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "聖亭派出所", "職別": "副所長", "姓名": "曹培翔", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "聖亭派出所", "職別": "業務承辦人", "姓名": "曾建凱", "金額": 90, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭派出所", "職別": "所長", "姓名": "孫祥愷", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭派出所", "職別": "副所長", "姓名": "劉重言", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭派出所", "職別": "副所長", "姓名": "梁順安", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭派出所", "職別": "業務承辦人", "姓名": "周薇", "金額": 90, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "中興派出所", "職別": "所長", "姓名": "董亦文", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "中興派出所", "職別": "副所長", "姓名": "何昀融", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "中興派出所", "職別": "副所長", "姓名": "林榮裕", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "中興派出所", "職別": "業務承辦人", "姓名": "鄧雅文", "金額": 90, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "石門派出所", "職別": "所長", "姓名": "林育辰", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "石門派出所", "職別": "副所長", "姓名": "薛德祥", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "石門派出所", "職別": "業務承辦人", "姓名": "陳琦", "金額": 89, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "高平派出所", "職別": "所長", "姓名": "王梓岳", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "高平派出所", "職別": "副所長", "姓名": "楊勝吉", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "高平派出所", "職別": "業務承辦人", "姓名": "黃丞潁", "金額": 89, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "三和派出所", "職別": "所長", "姓名": "宋開國", "金額": 194, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "三和派出所", "職別": "副所長", "姓名": "陳佶汎", "金額": 194, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "三和派出所", "職別": "業務承辦人", "姓名": "童霂晟", "金額": 89, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭交通分隊", "職別": "分隊長", "姓名": "卓宜澂", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭交通分隊", "職別": "小隊長", "姓名": "鄭敬思", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭交通分隊", "職別": "小隊長", "姓名": "蔡安龍", "金額": 195, "蓋章": ""},
        {"分配類別": "勤務督導(20%)", "單位": "龍潭交通分隊", "職別": "業務承辦人", "姓名": "陳建穎", "金額": 89, "蓋章": ""}
    ]
    df_coworkers_default = pd.DataFrame(default_coworkers_data)

    edited_df_coworkers = st.data_editor(
        df_coworkers_default,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        height=400,
        column_config={
            "分配類別": st.column_config.SelectboxColumn("分配類別", options=["負責管考(72%)", "勤務督導(20%)", "其他配合(8%)"], required=True),
            "金額": st.column_config.NumberColumn("金額", min_value=0, step=1, format="%d")
        }
    )

    if st.button("🚀 執行彙整、計算獎金與發送報表", type="primary", use_container_width=True):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種點數統計資料皆已完成上傳！")
            return

        with st.spinner("正在精算比例與發放金額..."):
            try:
                # --- A. 點數數據預處理 ---
                df_acc_raw = pd.read_excel(file_acc, header=4)
                df_acc_raw['姓名'] = df_acc_raw['姓名'].astype(str).str.strip()
                dict_acc = df_acc_raw.groupby('姓名')[['A2類', 'A3類']].sum().to_dict(orient='index')

                df_traf_all = pd.concat([pd.read_excel(f, sheet_name='月彙整總表') for f in file_traf_list])
                df_traf_all['姓名'] = df_traf_all['姓名'].astype(str).str.strip()
                dict_traf = df_traf_all.groupby('姓名')['總計尖峰時數'].sum().to_dict()

                # --- B. 讀取點數範本與日期偵測 ---
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
                                found_date = True; break
                        if found_date: break
                    if found_date: break

                # --- C. 點數表裁剪與重新計算直接執行人員 ---
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
                        
                        col_map = {c: i for i, c in enumerate(df_work.columns)}
                        member_rows = []
                        for r in range(len(df_work)):
                            name_cell = str(df_work.iloc[r, col_map['員警姓名']]).strip()
                            if '小計' in name_cell or '總計' in name_cell or name_cell in ['nan', 'None', '']:
                                continue
                            member_rows.append(r)

                        df_members = df_work.iloc[member_rows].copy()
                        s_cite, s_acc, s_traf = 0, 0, 0
                        
                        for idx, row in df_members.iterrows():
                            name = str(row['員警姓名']).strip()
                            a2 = dict_acc.get(name, {}).get('A2類', 0)
                            a3 = dict_acc.get(name, {}).get('A3類', 0)
                            th = dict_traf.get(name, 0)
                            ap, tp = a2 * P_A2 + a3 * P_A3, th * P_TRAF
                            
                            cp = pd.to_numeric(row.get('取締點數', 0), errors='coerce')
                            cp = cp if pd.notna(cp) else 0
                            
                            total_pts = cp + ap + tp
                            
                            if 'A2件數' in col_map: df_members.at[idx, 'A2件數'] = a2 if a2 > 0 else ""
                            if 'A3件數' in col_map: df_members.at[idx, 'A3件數'] = a3 if a3 > 0 else ""
                            if '事故點數' in col_map: df_members.at[idx, '事故點數'] = ap if ap > 0 else ""
                            if '交整時數' in col_map: df_members.at[idx, '交整時數'] = th if th > 0 else ""
                            if '交整點數' in col_map: df_members.at[idx, '交整點數'] = tp if tp > 0 else ""
                            if '個人總點數' in col_map: df_members.at[idx, '個人總點數'] = total_pts
                            
                            s_cite += cp; s_acc += ap; s_traf += tp
                            
                            # 直接執行人員精算
                            if total_pts > 0:
                                reward = int(np.round(total_pts * point_value))
                                direct_exec_list.append({
                                    "單位名稱": sheet_name,
                                    "員警姓名": name,
                                    "取締件數": row.get('取締件數', ''),
                                    "取締點數": cp if cp > 0 else '',
                                    "A2件數": a2 if a2 > 0 else '',
                                    "A3件數": a3 if a3 > 0 else '',
                                    "事故點數": ap if ap > 0 else '',
                                    "交整時數": th if th > 0 else '',
                                    "交整點數": tp if tp > 0 else '',
                                    "個人總點數": total_pts,
                                    "每點獎金": point_value,
                                    "實領獎金": reward,
                                    "蓋章": ""
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
                df_direct_exec.insert(0, '序號', range(1, len(df_direct_exec) + 1))
                direct_total_money = df_direct_exec['實領獎金'].sum()
                
                # --- D. 處理共同作業人員 (照金額比例分配 - 最大餘數法) ---
                df_coworkers_work = edited_df_coworkers.copy()
                df_coworkers_work.dropna(subset=['姓名'], inplace=True)
                
                if "系統自動" in alloc_mode:
                    if "A" in budget_type:
                        coworker_pool = int(budget_input)
                    else:
                        coworker_pool = int(budget_input) - direct_total_money
                        if coworker_pool < 0:
                            st.error(f"❌ 全分局預算 ({budget_input}) 不足支付直接執行人員 ({direct_total_money})，請確認預算金額。")
                            return
                            
                    # 切割三大塊 72% / 20% / 8%
                    pool_72 = int(np.round(coworker_pool * 0.72))
                    pool_20 = int(np.round(coworker_pool * 0.20))
                    pool_08 = coworker_pool - pool_72 - pool_20
                    
                    df_coworkers_work['核發金額'] = 0
                    
                    # 針對每一類別根據「金額」比例進行分配 (引入最大餘數法)
                    for cat, pool in [("負責管考(72%)", pool_72), ("勤務督導(20%)", pool_20), ("其他配合(8%)", pool_08)]:
                        cat_mask = df_coworkers_work['分配類別'] == cat
                        sum_w = df_coworkers_work.loc[cat_mask, '金額'].sum()
                        
                        if sum_w > 0 and pool > 0:
                            # 算出精確包含小數點的分配金額
                            exact_amounts = (df_coworkers_work.loc[cat_mask, '金額'] / sum_w) * pool
                            # 全部先無條件捨去取整數
                            int_amounts = np.floor(exact_amounts).astype(int)
                            
                            # 看看還剩下幾塊錢發不出去
                            diff = int(pool - int_amounts.sum())
                            # 計算每個人被捨去了多少小數點
                            remainders = exact_amounts - int_amounts
                            
                            # 將剩下的零錢，優先補給小數點餘數最大的那幾個人 (最均勻平分的演算法)
                            if diff > 0:
                                top_indices = remainders.nlargest(diff).index
                                int_amounts.loc[top_indices] += 1
                                
                            df_coworkers_work.loc[cat_mask, '核發金額'] = int_amounts
                            
                    # 用計算完真實核發的金額，取代掉原本作為比例的輸入金額
                    df_coworkers_output = df_coworkers_work.drop(columns=['金額']).rename(columns={'核發金額': '金額'})
                else:
                    # 手動模式：原本輸入多少錢就發多少錢
                    df_coworkers_output = df_coworkers_work.copy()
                
                # 自動排序並插入序號
                df_coworkers_output.sort_values(by='分配類別', ascending=False, inplace=True)
                df_coworkers_output.insert(0, '序號', range(1, len(df_coworkers_output) + 1))
                
                # 分類加總供支領一覽表使用
                sub_72 = df_coworkers_output.loc[df_coworkers_output['分配類別'] == "負責管考(72%)", '金額'].sum()
                sub_20 = df_coworkers_output.loc[df_coworkers_output['分配類別'] == "勤務督導(20%)", '金額'].sum()
                sub_08 = df_coworkers_output.loc[df_coworkers_output['分配類別'] == "其他配合(8%)", '金額'].sum()
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

                # --- E. 封裝成兩個 Excel 檔案 ---
                pts_output = io.BytesIO()
                df_pts_summary = pd.DataFrame([['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']] + summary_rows + [['合計', g_cite, g_acc, g_traf, g_all]])
                with pd.ExcelWriter(pts_output, engine='xlsxwriter') as writer:
                    df_pts_summary.to_excel(writer, sheet_name='總表', header=False, index=False)
                    for sn, df_f in final_sheets.items():
                        df_f.to_excel(writer, sheet_name=sn, index=False)
                pts_excel_data = pts_output.getvalue()
                pts_filename = f"龍潭分局{ext_year}年{ext_month}月份_點數統計表.xlsx"

                payroll_output = io.BytesIO()
                with pd.ExcelWriter(payroll_output, engine='xlsxwriter') as writer:
                    df_direct_exec.to_excel(writer, sheet_name='直接執行人員', index=False)
                    if not df_coworkers_output.empty:
                        df_coworkers_output.to_excel(writer, sheet_name='共同作業及配合人員', index=False)
                    df_payroll_summary.to_excel(writer, sheet_name='獎勵金支領一覽表', index=False)
                payroll_excel_data = payroll_output.getvalue()
                payroll_filename = f"龍潭分局{ext_year}年{ext_month}月份_獎勵金印領清冊.xlsx"

                # --- F. 自動發送郵件 (同時夾帶兩份附件) ---
                files_to_attach = [
                    (pts_excel_data, pts_filename),
                    (payroll_excel_data, payroll_filename)
                ]
                ok, err = send_report_email_auto(files_to_attach, ext_year, ext_month)
                
                if ok:
                    st.success(f"✅ 雙報表產出成功！已使用最新公平比例精算法，檔案已自動備份至您的信箱。")
                else:
                    st.warning(f"⚠️ 報表已產出，但郵件發送失敗: {err}")

                c5, c6 = st.columns(2)
                c5.download_button(label="📥 下載【點數統計表】", data=pts_excel_data, file_name=pts_filename, use_container_width=True)
                c6.download_button(label="📥 下載【印領清冊】", data=payroll_excel_data, file_name=payroll_filename, use_container_width=True, type="primary")

            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")

if __name__ == "__main__":
    p18_page()
