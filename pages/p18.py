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
        
        body = f"郭同仁您好：\n\n系統已自動完成 {year}年{month}月份的獎勵金點數彙整與印領清冊產出。\n本次附件包含「點數統計表」與「印領清冊」共包含兩份 Excel 檔案，請查收。"
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


# --- 排序函數 ---
def sort_coworkers(df):
    df = df.copy()
    for col in ['分配類別', '單位', '職別', '姓名']:
        if col in df.columns:
            df[col] = df[col].astype(str)
            
    df['姓名'] = df['姓名'].replace(['nan', 'None'], '').fillna("").str.strip()
    df['單位'] = df['單位'].replace(['nan', 'None'], '').fillna("").str.strip()
    df['職別'] = df['職別'].replace(['nan', 'None'], '').fillna("").str.strip()
    df['分配類別'] = df['分配類別'].replace(['nan', 'None'], '').fillna("").str.strip()
    
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
    df.drop(columns=['職級權重'], inplace=True)
    
    for col in ['分配類別', '單位']:
        df[col] = df[col].astype(str)
        
    df.reset_index(drop=True, inplace=True)
    return df


def on_data_edited():
    changes = st.session_state.co_editor
    df = st.session_state.current_roster.copy()
    
    for col in ['分配類別', '單位', '職別', '姓名']:
        if col in df.columns:
            df[col] = df[col].astype(str)
            
    for row_idx, updated_cols in changes.get("edited_rows", {}).items():
        for col_name, val in updated_cols.items():
            df.at[row_idx, col_name] = val
            
    for new_row in changes.get("added_rows", {}):
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        
    deleted_indices = changes.get("deleted_rows", [])
    if deleted_indices:
        df.drop(index=deleted_indices, inplace=True)
        
    st.session_state.current_roster = sort_coworkers(df)


def p18_page():
    show_sidebar()
    st.title("💰 龍潭分局 - 獎勵金點數統計表暨印領清冊產生器")

    P_A2, P_A3, P_TRAF = 10.0, 5.0, 5.0

    st.subheader("📂 1. 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)", 
                                      type=['xlsx'], accept_multiple_files=True)
    
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
        st.info("💡 手動模式：請於表格內自行新增「金額」欄位並填寫實際發放金額。")
        budget_input = 0
        budget_type = ""

    st.markdown("**共同作業名單**")
    roster_file = 'coworkers_roster.csv'
   
    if 'current_roster' not in st.session_state:
        if os.path.exists(roster_file):
            df_init = pd.read_csv(roster_file)
        else:
            default_coworkers_data = [
                {"分配類別": "負責管考(72%)", "單位": "龍潭分局", "職別": "分局長", "姓名": "施宇峰"},
                {"分配類別": "負責管考(72%)", "單位": "龍潭分局", "職別": "副分局長", "姓名": "何憶雯"},
                {"分配類別": "負責管考(72%)", "單位": "龍潭分局", "職別": "副分局長", "姓名": "蔡志明"},
                {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "業務單位主管", "姓名": "陳維明"},
                {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "盧冠仁"},
                {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "李峯甫"},
                {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承辦人", "姓名": "羅千金"},
                {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "交通業務承承辦人", "姓名": "郭勝隆"},
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
                {"分配類別": "勤務督導(20%)", "單位": "勤務中心", "職別": "巡視員", "姓名": "林榮裕"},
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
            "排序調整": st.column_config.NumberColumn("排序調整 🔢", min_value=1, format="%d"),
            "分配類別": st.column_config.SelectboxColumn("分配類別", options=["負責管考(72%)", "勤務督導(20%)", "其他配合(8%)"], required=True),
            "金額": st.column_config.NumberColumn("金額", min_value=0, step=1, format="%d")
        }
    else:
        if '金額' in df_display.columns:
            df_display = df_display.drop(columns=['金額'])
        col_cfg = {
            "排序調整": st.column_config.NumberColumn("排序調整 🔢", min_value=1, format="%d"),
            "分配類別": st.column_config.SelectboxColumn("分配類別", options=["負責管考(72%)", "勤務督導(20%)", "其他配合(8%)"], required=True)
        }

    edited_df_coworkers = st.data_editor(
        df_display, num_rows="dynamic", use_container_width=True, hide_index=True, height=500, column_config=col_cfg, key="co_editor", on_change=on_data_edited
    )

    if st.button("💾 儲存最新名單為預設值", use_container_width=True, type="secondary"):
        st.session_state.current_roster.to_csv(roster_file, index=False, encoding='utf-8-sig')
        st.success("✅ 名單順序已永久記憶！")
        st.rerun()

    st.markdown("---")

    if st.button("🚀 執行彙整、計算獎金與發送報表", type="primary", use_container_width=True):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種點數統計資料皆已完成上傳！")
            return

        with st.spinner("正在精算比例與發放金額..."):
            try:
                # A. 點數數據預處理
                df_acc_raw = pd.read_excel(file_acc, header=4)
                df_acc_raw['姓名'] = df_acc_raw['姓名'].astype(str).str.strip()
                dict_acc = df_acc_raw.groupby('姓名')[['A2類', 'A3類']].sum().to_dict(orient='index')

                df_traf_all = pd.concat([pd.read_excel(f, sheet_name='月彙整總表') for f in file_traf_list])
                df_traf_all['姓名'] = df_traf_all['姓名'].astype(str).str.strip()
                dict_traf = df_traf_all.groupby('姓名')['總計尖峰時數'].sum().to_dict()

                # B. 讀取點數範本與日期偵測
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

                # C. 直接執行人員原始金額計算
                final_sheets = {}
                summary_rows = []
                g_cite, g_acc, g_traf, g_all = 0, 0, 0, 0
                direct_exec_list = []
                sheet_reward_diff = {} 

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
                       
                        for idx, row in df_members.iterrows():
                            name = str(row.get('員警姓名', '')).strip()
                            a2 = dict_acc.get(name, {}).get('A2類', 0)
                            a3 = dict_acc.get(name, {}).get('A3類', 0)
                            th = dict_traf.get(name, 0)
                            ap = a2 * P_A2 + a3 * P_A3
                            tp = th * P_TRAF
                            cp = pd.to_numeric(row.get('取締點數', 0), errors='coerce') or 0
                            total_pts = cp + ap + tp
                            
                            reward = int(np.round(total_pts * point_value))
                            if total_pts > 0:
                                direct_exec_list.append({
                                    "單位名稱": sheet_name, "員警姓名": name,
                                    "取締件數": row.get('取締件數', ''), "取締點數": cp if cp > 0 else '',
                                    "A2件數": a2 if a2 > 0 else '', "A3件數": a3 if a3 > 0 else '',
                                    "事故點數": ap if ap > 0 else '', "交整時數": th if th > 0 else '',
                                    "交整點數": tp if tp > 0 else '', "個人總點數": total_pts,
                                    "每點獎金": point_value, "實領獎金": reward, "蓋章": "",
                                    "_orig_df_idx": idx
                                })
                        final_sheets[sheet_name] = {"df_members": df_members, "df_work": df_work}

                # === 直接執行人員 1元 自動平帳調頂區塊 ===
                df_direct_exec = pd.DataFrame(direct_exec_list)
                if not df_direct_exec.empty and target_direct_budget > 0:
                    current_sum = df_direct_exec['實領獎金'].sum()
                    diff = target_direct_budget - current_sum
                    
                    if diff != 0 and abs(diff) <= 5: 
                        df_direct_exec.at[0, '實領獎金'] += diff
                        target_sheet = df_direct_exec.at[0, '單位名稱']
                        sheet_reward_diff[target_sheet] = diff
                        st.info(f"⚖️ 直接執行人員完成 ±1 元強制平帳補齊。")

                # 重新回填與建立各單位的 Excel 點數分頁
                for sheet_name, f_data in final_sheets.items():
                    df_members = f_data["df_members"]
                    df_work = f_data["df_work"]
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
                    
                    sub_row_data = {c: "" for c in df_work.columns}
                    sub_row_data['員警姓名'] = '小計'
                    for col_n in df_work.columns:
                        if col_n in ['員警姓名', '蓋章']: continue
                        v_sum = pd.to_numeric(df_members[col_n], errors='coerce').sum()
                        sub_row_data[col_n] = v_sum if v_sum > 0 else 0
                    
                    if '實領獎金' in df_members.columns:
                        df_members['實領獎金'] = (pd.to_numeric(df_members['個人總點數'], errors='coerce') * point_value).round().fillna(0).astype(int)
                        if sheet_name in sheet_reward_diff:
                            first_idx = df_members.index[0]
                            df_members.at[first_idx, '實領獎金'] += sheet_reward_diff[sheet_name]
                        sub_row_data['實領獎金'] = df_members['實領獎金'].sum()

                    df_final = pd.concat([df_members, pd.DataFrame([sub_row_data])], ignore_index=True)
                    if '蓋章' in df_final.columns: df_final = df_final.drop(columns=['蓋章'])
                    final_sheets[sheet_name] = df_final
                    
                    summary_rows.append([sheet_name, s_cite, s_acc, s_traf, s_cite + s_acc + s_traf])
                    g_cite += s_cite; g_acc += s_acc; g_traf += s_traf; g_all += (s_cite + s_acc + s_traf)

                if not df_direct_exec.empty:
                    if '_orig_df_idx' in df_direct_exec.columns:
                        df_direct_exec = df_direct_exec.drop(columns=['_orig_df_idx'])
                    df_direct_exec.insert(0, '序號', range(1, len(df_direct_exec) + 1))
                
                direct_total_money = df_direct_exec['實領獎金'].sum() if not df_direct_exec.empty else 0

                # D. 共同作業人員處理
                df_coworkers_work = st.session_state.current_roster.copy()
                df_coworkers_work.dropna(how='all', inplace=True)
                df_coworkers_work = sort_coworkers(df_coworkers_work)

                pool_72, pool_20, pool_08 = 0, 0, 0

                if "系統自動" in alloc_mode:
                    if "A" in budget_type:
                        coworker_pool = int(budget_input)
                    else:
                        coworker_pool = int(budget_input) - direct_total_money
                        if coworker_pool < 0:
                            st.error(f"❌ 全分局預算不足支付直接執行人員")
                            return
                           
                    pool_72 = int(np.round(coworker_pool * 0.72))
                    pool_20 = int(np.round(coworker_pool * 0.20))
                    pool_08 = coworker_pool - pool_72 - pool_20
                    
                    df_coworkers_work['核發金額'] = 0
                    
                    # === 負責管考(72%) ===
                    mask_72 = df_coworkers_work['分配類別'] == "負責管考(72%)"
                    df_72 = df_coworkers_work[mask_72].copy()
                    df_72['核發金額'] = 0

                    if not df_72.empty and pool_72 > 0:
                        main_officers_mask = df_72['職別'].str.contains('分局長|副分局長', na=False)
                        main_officers = df_72[main_officers_mask].copy()
                        
                        if not main_officers.empty:
                            pool_main = int(np.round(pool_72 * 0.08))
                            num_deputies = main_officers['職別'].str.contains('副分局長').sum()
                            
                            main_exact_list = []
                            for idx, row in main_officers.iterrows():
                                t = str(row['職別'])
                                if '分局長' in t and '副' not in t:
                                    w = 0.60
                                elif '副分局長' in t:
                                    w = 0.40 / num_deputies if num_deputies > 0 else 0.40
                                else:
                                    w = 0
                                main_exact_list.append((idx, w * pool_main))
                                
                            tot_w_main = sum(x[1] for x in main_exact_list)
                            if tot_w_main > 0:
                                main_amounts = np.floor([x[1] for x in main_exact_list]).astype(int)
                                diff_m = pool_main - main_amounts.sum()
                                if diff_m > 0:
                                    main_amounts[:diff_m] += 1
                                for idx_m, (orig_idx, _) in enumerate(main_exact_list):
                                    df_72.at[orig_idx, '核發金額'] = main_amounts[idx_m]

                        remaining_pool = pool_72 - df_72['核發金額'].sum()
                        other_mask = ~df_72['職別'].str.contains('分局長|副分局長', na=False)
                        df_other = df_72[other_mask].copy()

                        if not df_other.empty and remaining_pool > 0:
                            def get_sub_weight(row):
                                title = str(row['職別'])
                                unit = str(row['單位'])
                                if '交通組' in unit:
                                    return 26
                                elif any(x in title for x in ['所長', '分隊長', '副所長', '小隊長']) and any(x in unit for x in ['派出所', '交通分隊', '警備隊']):
                                    return 56
                                elif any(x in title for x in ['承辦', '業務', '同仁', '警員']) and any(x in unit for x in ['派出所', '交通分隊', '警備隊']):
                                    return 10
                                return 1

                            df_other['weight'] = df_other.apply(get_sub_weight, axis=1)
                            total_weight = df_other['weight'].sum()
                            if total_weight > 0:
                                exact_amounts = (df_other['weight'] / total_weight) * remaining_pool
                                int_amounts = np.floor(exact_amounts).astype(int)
                                
                                diff = int(remaining_pool - int_amounts.sum())
                                if diff > 0:
                                    remainders = exact_amounts - int_amounts
                                    top_indices = remainders.nlargest(diff).index
                                    int_amounts.loc[top_indices] += 1
                                elif diff < 0:
                                    top_indices = int_amounts.nlargest(abs(diff)).index
                                    int_amounts.loc[top_indices] -= 1
                                    
                                df_other['核發金額'] = int_amounts
                                df_72.loc[other_mask, '核發金額'] = df_other['核發金額']

                        df_coworkers_work.loc[mask_72, '核發金額'] = df_72['核發金額']

                    for cat, pool in [("勤務督導(20%)", pool_20), ("其他配合(8%)", pool_08)]:
                        cat_mask = df_coworkers_work['分配類別'] == cat
                        count = cat_mask.sum()
                        if count > 0 and pool > 0:
                            int_amount = int(np.floor(pool / count))
                            amounts = np.full(count, int_amount)
                            diff = pool - amounts.sum()
                            if diff > 0:
                                amounts[:diff] += 1
                            df_coworkers_work.loc[cat_mask, '核發金額'] = amounts
                    
                    df_coworkers_output = df_coworkers_work.rename(columns={'核發金額': '金額'})
                    
                    # 跨類別金額歸併
                    mgr_lookup = {}
                    for idx, row in df_coworkers_output[df_coworkers_output['分配類別'] == '負責管考(72%)'].iterrows():
                        key = (str(row['單位']).strip(), str(row['姓名']).strip())
                        if key[1] != "":
                            mgr_lookup[key] = idx
                    
                    indices_to_drop = []
                    for idx, row in df_coworkers_output[df_coworkers_output['分配類別'] == '勤務督導(20%)'].iterrows():
                        key = (str(row['單位']).strip(), str(row['姓名']).strip())
                        if key in mgr_lookup:
                            target_mgr_idx = mgr_lookup[key]
                            df_coworkers_output.at[target_mgr_idx, '金額'] += row['金額']
                            indices_to_drop.append(idx)
                    
                    if indices_to_drop:
                        df_coworkers_output.drop(index=indices_to_drop, inplace=True)
                else:
                    df_coworkers_output = df_coworkers_work.copy()

                if '金額' not in df_coworkers_output.columns:
                    df_coworkers_output['金額'] = 0
                
                df_coworkers_output = sort_coworkers(df_coworkers_output)
                if '排序調整' in df_coworkers_output.columns:
                    df_coworkers_output = df_coworkers_output.drop(columns=['排序調整'])
                
                df_coworkers_output.insert(0, '序號', range(1, len(df_coworkers_output) + 1))
                df_coworkers_output['蓋章'] = ""

                # 支領一覽表摘要
                if "系統自動" in alloc_mode:
                    sub_72 = pool_72
                    sub_20 = pool_20
                    sub_08 = pool_08
                    coworkers_total_money = pool_72 + pool_20 + pool_08
                else:
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

                # E. Excel 雙報表美化與邊界鎖定輸出
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
                    # 1. 輸出：直接執行人員
                    df_direct_exec.to_excel(writer, sheet_name='直接執行人員', index=False)
                    workbook  = writer.book
                    worksheet1 = writer.sheets['直接執行人員']
                    
                    # 建立安全不報錯的樣式物件
                    cell_format = workbook.add_format()
                    cell_format.set_align('center')   
                    cell_format.set_align('vcenter')  
                    cell_format.set_border(1)         
                    cell_format.set_font_name('微軟正黑體')
                    cell_format.set_font_size(11)
                    
                    # 💡 【重大更新】：限制蓋章範圍並重設其餘空白列，鎖定外圍防線
                    worksheet1.set_default_row(18)  # 其餘沒有資料的地方一律維持預設列高
                    num_direct_rows = len(df_direct_exec)
                    
                    # 精準只在「有資料的行數」加高列高
                    for r_idx in range(1, num_direct_rows + 1):
                        worksheet1.set_row(r_idx, 35, cell_format)
                    
                    # 精準加寬 L 欄（蓋章欄）
                    worksheet1.set_column('L:L', 16, cell_format)
                    
                    # 2. 輸出：共同作業及配合人員
                    if not df_coworkers_output.empty:
                        df_coworkers_output.to_excel(writer, sheet_name='共同作業及配合人員', index=False)
                        worksheet2 = writer.sheets['共同作業及配合人員']
                        
                        worksheet2.set_default_row(18)  # 其餘空白列鎖定
                        num_co_rows = len(df_coworkers_output)
                        
                        for r_idx in range(1, num_co_rows + 1):
                            worksheet2.set_row(r_idx, 35, cell_format)
                        
                        stamp_col_idx = df_coworkers_output.columns.get_loc('蓋章')
                        stamp_col_letter = chr(65 + stamp_col_idx)
                        worksheet2.set_column(f'{stamp_col_letter}:{stamp_col_letter}', 16, cell_format)
                            
                    # 3. 輸出：一覽表摘要
                    df_payroll_summary.to_excel(writer, sheet_name='獎勵金支領一覽表', index=False)
                    worksheet3 = writer.sheets['獎勵金支領一覽表']
                    worksheet3.set_default_row(18)
                    worksheet3.set_column('A:B', 25, cell_format)
                    for r_idx in range(1, len(df_payroll_summary) + 1):
                        worksheet3.set_row(r_idx, 22, cell_format)
                        
                payroll_excel_data = payroll_output.getvalue()
                payroll_filename = f"龍潭分局{ext_year}年{ext_month}月份_獎勵金印領清冊.xlsx"

                # F. 發送與下載
                files_to_attach = [(pts_excel_data, pts_filename), (payroll_excel_data, payroll_filename)]
                ok, err = send_report_email_auto(files_to_attach, ext_year, ext_month)
                
                if ok:
                    st.success("✅ 雙報表產出成功！Excel 右方及下方已切齊邊界，無任何多餘贅欄列框。")
                else:
                    st.warning(f"⚠️ 報表已產出，但郵件發送失敗: {err}")

                c5, c6 = st.columns(2)
                c5.download_button("📥 下載【點數統計表】", pts_excel_data, pts_filename, use_container_width=True)
                c6.download_button("📥 下載【印領清冊】", payroll_excel_data, payroll_filename, use_container_width=True, type="primary")

            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")


if __name__ == "__main__":
    p18_page()
