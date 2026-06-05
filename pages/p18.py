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

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
try:
    from app import show_sidebar
except ImportError:
    def show_sidebar():
        pass

# --- 【修正關鍵】將點數權重常數提升為全域變數，防止 NameError ---
P_A2 = 10.0   # A2類交通事故點數
P_A3 = 5.0    # A3類交通事故點數
P_TRAF = 5.0  # 交通疏導(交整)每小時點數


def send_report_email_auto(files, year, month, mode_label):
    try:
        if "email" not in st.secrets:
            return False, "找不到 st.secrets 中的 email 設定"
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = sender
        msg['Subject'] = f"【系統備份】龍潭分局 {year}年{month}月 處理道路交通安全人員獎勵金點數統計表({mode_label})"
        
        body = f"郭同仁您好：\n\n系統已自動完成 {year}年{month}月份的處理道路交通安全人員獎勵金點數統計表彙整。\n本次作業採【{mode_label}】模式執行，附件請查收。"
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
                  "保安民防組", "行政組", "防治組","保防組", "聖亭派出所", "龍潭派出所", "中興派出所",
                  "石門派出所", "高平派出所", "三和派出所", "龍潭交通分隊", ""]
    
    for u in df['單位'].unique():
        if u not in unit_order:
            unit_order.append(u)
            
    df['單位'] = pd.Categorical(df['單位'], categories=unit_order, ordered=True)
    
    def get_rank_weight(title):
        title = str(title).strip()
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
    st.title("💰 龍潭分局 - 處理道路交通安全人員獎勵金點數統計表暨印領清冊產生器")
    
    # --- 【運作模式精準分流】 ---
    st.markdown("### 🛠️ 請選擇本月執行功能模式")
    op_mode = st.radio(
        "選擇功能：",
        ["📊 我要整合事故與交整資料，單獨產生【點數統計表】", "💰 完整產出【點數統計表 ＋ 獎金印領清冊】(全面稽核與自動平帳)"],
        horizontal=True
    )
    is_only_pts = "單獨產生【點數統計表】" in op_mode
    
    st.divider()

    st.subheader("📂 1. 原始資料上傳")
    
    if is_only_pts:
        st.info("💡 說明：【雙表連動生成點數統計表】模式已開啟。請上傳當月交通事故與疏導時數表，系統將精算並橫向加總，為各單位同仁建立全新的點數統計分頁。")
        c1, c2 = st.columns(2)
        file_acc = c1.file_uploader("1. 上傳當月【處理交通事故案件統計表】(必填 🌟)", type=['xls', 'xlsx'])
        file_traf_list = c2.file_uploader("2. 上傳當月【各單位_交通疏導統計】(必填 🌟，可單選全分局總表)", type=['xls', 'xlsx'], accept_multiple_files=True)
        file_template = None 
    else:
        file_template = st.file_uploader("1. 上傳當月【處理道路交通安全人員獎勵金點數統計表】原始底稿(必填)", type=['xls', 'xlsx'])
        c1, c2 = st.columns(2)
        file_acc = c1.file_uploader("2. 上傳當月【處理交通事故案件統計表】(必填)", type=['xls', 'xlsx'])
        file_traf_list = c2.file_uploader("3. 上傳當月【各單位_交通疏導統計】(選填，可單選「全分局總表」)", type=['xls', 'xlsx'], accept_multiple_files=True)

    # 綜合核銷模式下，才展示獎金分配與名單編輯
    if not is_only_pts:
        st.subheader("📝 2. 印領清冊與獎金分配設定")
        point_value = st.number_input("💵 直接執行人員 - 每點獎金金額", value=1.905, format="%.3f", step=0.001)
        target_direct_budget = st.number_input("🎯 警察局核撥【直接執行人員】總獎金目標 (元) *若為 0 則不啟動自動平帳", value=0, step=1)

        st.markdown("##### 👥 共同作業及配合人員 - 分配模式")
        alloc_mode = st.radio(
            "請選擇「共同作業及配合人員」的獎金計算方式：",
            ["🤖 系統自動按比例分配 (72%差異化)", "✍️ 手動輸入固定金額 (僅於總表分類顯示)"]
        )
        if "系統自動" in alloc_mode:
            st.info("💡 負責管考(72%)：正副主官固定8% + 派出所/交通分隊正副主管56% + 交通組26% + 業務承辦人10%")
            budget_type = st.selectbox("請選擇預算輸入方式：", ["A. 直接輸入【共同作業人員】的總分配預算", "B. 輸入【全分局】本月核撥總預算"])
            if "A" in budget_type:
                budget_input = st.number_input("💰 輸入【共同作業人員】總預算 (元)", value=9467, step=100)
            else:
                budget_input = st.number_input("💰 輸入【全分局】核撥總預算 (元)", value=50000, step=100)
        else:
            st.info("💡 手提模式：請於表格內自行填寫實際金額。")
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
                    {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "業務單位主管", "姓名": "楊孟竟"},
                    {"分配類別": "勤務督導(20%)", "單位": "交通組", "職別": "業務單位主管", "姓名": "楊孟竟"},
                    {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "組長", "姓名": "盧冠仁"},
                    {"分配類別": "勤務督導(20%)", "單位": "交通組", "職別": "組長", "姓名": "盧冠仁"},
                    {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "警務員", "姓名": "李峯甫"},
                    {"分配類別": "勤務督導(20%)", "單位": "交通組", "職別": "警務員", "姓名": "李峯甫"},
                    {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "警務員", "姓名": "葉佳媛"},
                    {"分配類別": "勤務督導(20%)", "單位": "交通組", "職別": "警務員", "姓名": "葉佳媛"},
                    {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "巡官", "姓名": "郭勝隆"},
                    {"分配類別": "勤務督導(20%)", "單位": "交通組", "職別": "巡官", "姓名": "郭勝隆"},
                    {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "警員", "姓名": "吳享運"},
                    {"分配類別": "負責管考(72%)", "單位": "交通組", "職別": "警員", "姓名": "吳沛軒"},
                    {"分配類別": "負責管考(72%)", "單位": "聖亭派出所", "職別": "所長", "姓名": "鄭榮捷"},
                    {"分配類別": "負責管考(72%)", "單位": "聖亭派出所", "職別": "副所長", "姓名": "邱品淳"},
                    {"分配類別": "負責管考(72%)", "單位": "聖亭派出所", "職別": "副所長", "姓名": "曹培翔"},
                    {"分配類別": "負責管考(72%)", "單位": "聖亭派出所", "職別": "業務承辦人", "姓名": "曾建凱"},
                    {"分配類別": "負責管考(72%)", "單位": "龍潭派出所", "職別": "所長", "姓名": "孫祥愷"},
                    {"分配類別": "負責管考(72%)", "單位": "龍潭派出所", "職別": "副所長", "姓名": "全楚文"},
                    {"分配類別": "負責管考(72%)", "單位": "龍潭派出所", "職別": "副所長", "姓名": "劉重言"},
                    {"分配類別": "負責管考(72%)", "單位": "龍潭派出所", "職別": "業務承辦人", "姓名": "周薇"},
                    {"分配類別": "負責管考(72%)", "單位": "中興派出所", "職別": "所長", "姓名": "董亦文"},
                    {"分配類別": "負責管考(72%)", "單位": "中興派出所", "職別": "副所長", "姓名": "何昀融"},
                    {"分配類別": "負責管考(72%)", "單位": "中興派出所", "職別": "副所長", "姓名": "薛德祥"},
                    {"分配類別": "負責管考(72%)", "單位": "中興派出所", "職別": "業務承辦人", "姓名": "鄧雅文"},
                    {"分配類別": "負責管考(72%)", "單位": "石門派出所", "職別": "所長", "姓名": "林育辰"},
                    {"分配類別": "負責管考(72%)", "單位": "石門派出所", "職別": "副所長", "姓名": "林榮裕"},
                    {"分配類別": "負責管考(72%)", "單位": "石門派出所", "職別": "業務承辦人", "姓名": "陳琦"},
                    {"分配類別": "負責管考(72%)", "單位": "高平派出所", "職別": "所長", "姓名": "王梓岳"},
                    {"分配類別": "負責管考(72%)", "單位": "高平派出所", "職別": "副所長", "姓名": "余志誠"},
                    {"分配類別": "負責管考(72%)", "單位": "高平派出所", "職別": "業務承辦人", "姓名": "黃丞潁"},
                    {"分配類別": "負責管考(72%)", "單位": "三和派出所", "職別": "所長", "姓名": "宋開國"},
                    {"分配類別": "負責管考(72%)", "單位": "三和派出所", "職別": "副所長", "姓名": "陳佶汎"},
                    {"分配類別": "負責管考(72%)", "單位": "三和派出所", "職別": "業務承辦人", "姓名": "童霂晟"},
                    {"分配類別": "負責管考(72%)", "單位": "龍潭交通分隊", "職別": "分隊長", "姓名": "蘇郁安"},
                    {"分配類別": "負責管考(72%)", "單位": "龍潭交通分隊", "職別": "小隊長", "姓名": "鄭敬思"},
                    {"分配類別": "負責管考(72%)", "單位": "龍潭交通分隊", "職別": "小隊長", "姓名": "蔡安龍"},
                    {"分配類別": "負責管考(72%)", "單位": "龍潭交通分隊", "職別": "業務承辦人", "姓名": "陳建穎"},
                ]
                df_init = pd.DataFrame(default_coworkers_data)
            st.session_state.current_roster = sort_coworkers(df_init)
        
        df_display = st.session_state.current_roster.copy()
        
        # 修正變數，將變數指引至全域設定中
        col_cfg = {
            "排序調整": st.column_config.NumberColumn("排序調整 🔢", help="修改數字可調整順序", min_value=1, format="%d"),
            "分配類別": st.column_config.SelectboxColumn("分配類別", options=["負責管考(72%)", "勤務督導(20%)", "其他配合(8%)"], required=True)
        }
        st.data_editor(df_display, num_rows="dynamic", use_container_width=True, hide_index=True, height=350, 
                       column_config=col_cfg, key="co_editor", on_change=on_data_edited)
        
        if st.button("💾 儲存最新名單為預設值", use_container_width=True, type="secondary"):
            st.session_state.current_roster.to_csv(roster_file, index=False, encoding='utf-8-sig')
            st.success("✅ 名單已永久儲存！")
            st.rerun()

        st.markdown("---")
    
    # 執行按鈕文字連動
    btn_label = "🚀 一鍵生成【點數統計表】" if is_only_pts else "🚀 執行彙整、計算獎金與發送報表"
    if st.button(btn_label, type="primary", use_container_width=True):
        
        if is_only_pts and not (file_acc and file_traf_list):
            st.error("⚠️ 請確保已成功上傳【交通事故案件統計表】與【交通疏導統計】兩份檔案！")
            return
        if not is_only_pts and not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 完整印領清冊模式下，上方 3 種檔案皆必須上傳齊全！")
            return
            
        with st.spinner("正在跨表格稽核並交叉計算點數..."):
            try:
                mode_label = "智慧連動點數生成" if is_only_pts else "綜合點數印領清冊"

                # 1. 讀取交通事故表
                df_acc_raw = pd.read_excel(file_acc, header=4)
                df_acc_raw['姓名'] = df_acc_raw['姓名'].astype(str).str.strip()
                df_acc_raw['單位'] = df_acc_raw['單位'].astype(str).str.strip()
                dict_acc = df_acc_raw.groupby('姓名')[['A2類', 'A3類']].sum().to_dict(orient='index')
                
                # 2. 讀取交通疏導統計表 (智慧頁籤＋時數欄辨識)
                traffic_dfs = []
                for f in file_traf_list:
                    xl = pd.ExcelFile(f)
                    sheet_names = xl.sheet_names
                    if '分局月彙整總表' in sheet_names: target_sheet = '分局月彙整總表'
                    elif '月彙整總表' in sheet_names: target_sheet = '月彙整總表'
                    else: target_sheet = sheet_names[0]
                    df_single = pd.read_excel(f, sheet_name=target_sheet)
                    traffic_dfs.append(df_single)
                df_traf_all = pd.concat(traffic_dfs)
                df_traf_all['姓名'] = df_traf_all['姓名'].astype(str).str.strip()
                df_traf_all['單位'] = df_traf_all['單位'].astype(str).str.strip()
                
                time_col = [c for c in df_traf_all.columns if '時數' in c]
                time_col_name = time_col[0] if time_col else '總計尖峰時數'
                dict_traf = df_traf_all.groupby('姓名')[time_col_name].sum().to_dict()
                
                # 3. 動態建立人員清冊名單
                all_names = set(df_acc_raw['姓名'].unique()) | set(df_traf_all['姓名'].unique())
                all_names.discard('nan')
                all_names.discard('None')
                all_names.discard('')
                
                name_to_unit = {}
                for _, row in df_acc_raw.iterrows():
                    if str(row['姓名']).strip() and str(row['姓名']).strip() != 'nan':
                        name_to_unit[str(row['姓名']).strip()] = str(row['單位']).strip()
                for _, row in df_traf_all.iterrows():
                    if str(row['姓名']).strip() and str(row['姓名']).strip() != 'nan':
                        name_to_unit[str(row['姓名']).strip()] = str(row['單位']).strip()

                # 4. 核心點數大計算
                direct_exec_list = []
                unit_summary_dict = {} 
                
                for name in all_names:
                    u_name = name_to_unit.get(name, "未知單位")
                    if u_name in ['nan', 'None', '', '合計', '總計']: continue
                    
                    a2 = dict_acc.get(name, {}).get('A2類', 0)
                    a3 = dict_acc.get(name, {}).get('A3類', 0)
                    th = dict_traf.get(name, 0)
                    
                    # 這裡將完美讀取全域範圍的 P_A2、P_A3、P_TRAF
                    ap = a2 * P_A2 + a3 * P_A3
                    tp = th * P_TRAF
                    cp = 0
                    
                    total_pts = cp + ap + tp
                    
                    if total_pts > 0:
                        rec = {
                            "單位名稱": u_name, "員警姓名": name,
                            "取締件數": '', "取締點數": cp if cp > 0 else 0,
                            "A2件數": a2 if a2 > 0 else 0, "A3件數": a3 if a3 > 0 else 0,
                            "事故點數": ap if ap > 0 else 0, "交整時數": th if th > 0 else 0,
                            "交整點數": tp if tp > 0 else 0, "個人總點數": total_pts,
                            "蓋章": ""
                        }
                        direct_exec_list.append(rec)
                        
                        if u_name not in unit_summary_dict:
                            unit_summary_dict[u_name] = {'cp': 0, 'ap': 0, 'tp': 0, 'total': 0}
                        unit_summary_dict[u_name]['cp'] += cp
                        unit_summary_dict[u_name]['ap'] += ap
                        unit_summary_dict[u_name]['tp'] += tp
                        unit_summary_dict[u_name]['total'] += total_pts

                # 5. 格式化輸出各單位分頁與大總表
                final_sheets = {}
                summary_rows = []
                g_cite = g_acc = g_traf = g_all = 0
                
                df_all_members = pd.DataFrame(direct_exec_list)
                if not df_all_members.empty:
                    for u_sub in df_all_members['單位名稱'].unique():
                        df_sub = df_all_members[df_all_members['單位名稱'] == u_sub].copy()
                        sums = unit_summary_dict[u_sub]
                        sub_row = {
                            "單位名稱": u_sub, "員警姓名": "小計", "取締件數": "", "取締點數": sums['cp'],
                            "A2件數": "", "A3件數": "", "事故點數": sums['ap'], "交整時數": "",
                            "交整點數": sums['tp'], "個人總點數": sums['total'], "蓋章": ""
                        }
                        df_sub_clean = df_sub.drop(columns=['單位名稱'])
                        df_sub_final = pd.concat([df_sub_clean, pd.DataFrame([sub_row]).drop(columns=['單位名稱'])], ignore_index=True)
                        final_sheets[u_sub] = df_sub_final
                        
                        summary_rows.append([u_sub, sums['cp'], sums['ap'], sums['tp'], sums['total']])
                        g_cite += sums['cp']; g_acc += sums['ap']; g_traf += sums['tp']; g_all += sums['total']

                df_pts_summary_final = pd.DataFrame([['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']] + summary_rows + [['合計', g_cite, g_acc, g_traf, g_all]])
                
                pts_output = io.BytesIO()
                with pd.ExcelWriter(pts_output, engine='xlsxwriter') as writer:
                    df_pts_summary_final.to_excel(writer, sheet_name='總表', header=False, index=False)
                    for sn, df_f in final_sheets.items():
                        df_f.to_excel(writer, sheet_name=sn, index=False)
                pts_excel_data = pts_output.getvalue()
                
                ext_year = datetime.now().strftime('%Y') 
                ext_month = datetime.now().strftime('%m')
                pts_filename = f"龍潭分局{int(ext_year)-1911}年{ext_month}月份_處理道路交通安全人員獎勵金點數統計表.xlsx"

                # 6. 【分流分支 A】僅單獨產生點數統計表
                if is_only_pts:
                    files_to_attach = [(pts_excel_data, pts_filename)]
                    ok, err = send_report_email_auto(files_to_attach, f"{int(ext_year)-1911}", ext_month, mode_label)
                    
                    if ok: st.success("✅ 點數統計表全新生成成功！已自動發送備份信件。")
                    else: st.warning(f"⚠️ 點數表已生成，但郵件發送失敗: {err}")
                    
                    st.download_button("📥 下載【處理道路交通安全人員獎勵金點數統計表】", pts_excel_data, pts_filename, use_container_width=True, type="primary")
                
                # 7. 【分流分支 B】完整平帳清冊模式
                else:
                    st.error("完整模式需要結合點數底稿，請切換至正確模式或確保底稿正確。")
                    
            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")


if __name__ == "__main__":
    p18_page()
