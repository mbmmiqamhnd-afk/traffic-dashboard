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
from datetime import datetime

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
try:
    from app import show_sidebar
except ImportError:
    def show_sidebar():
        pass

# --- 點數權重全域常數 ---
P_A2 = 10.0   # A2類交通事故點數
P_A3 = 5.0    # A3類交通事故點數
P_TRAF = 5.0  # 交通疏導(交整)每小時點數


def send_report_email_auto(files, year, month, msg_subject, body_text):
    try:
        if "email" not in st.secrets:
            return False, "未偵測到 email 設定"
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = sender
        msg['Subject'] = msg_subject
        
        msg.attach(MIMEText(body_text, 'plain', 'utf-8'))
        
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
    st.title("💰 龍潭分局 - 處理道路交通安全人員獎勵金核銷自動化產生器")
    st.markdown("---")
    
    # --- 使用 Tab 元件，完美將工作階段徹底分流 ---
    tab_pts, tab_pay = st.tabs(["📊 階段一：自動填報對帳（產生點數統計表）", "💰 階段二：月底獎金請款（產生獎金印領清冊）"])
    
    # ==========================================================================
    # 📊 階段一：自動填報對帳面板
    # ==========================================================================
    with tab_pts:
        st.subheader("📂 請上傳對帳所需的三種原始資料")
        st.info("💡 說明：本區塊專用於月初對帳。系統會保留原始底稿的開單數據，自動填入事故與交整點數。")
        
        file_template = st.file_uploader("① 上傳當月【處理道路交通安全人員獎勵金點數統計表】原始取締底稿", type=['xls', 'xlsx'], key="pts_tpl")
        c1, c2 = st.columns(2)
        file_acc = c1.file_uploader("② 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'], key="pts_acc")
        file_traf_list = c2.file_uploader("③ 上傳當月【各單位_交通疏導統計】(可多選)", type=['xls', 'xlsx'], accept_multiple_files=True, key="pts_traf")
        
        st.markdown("---")
        if st.button("📊 一鍵自動填入並生成【點數統計表】", type="primary", use_container_width=True, key="btn_pts"):
            if not (file_template and file_acc and file_traf_list):
                st.error("⚠️ 請確保上述三項必填檔案（底稿、事故統計、疏導統計）皆已成功選取並上傳！")
            else:
                with st.spinner("正在讀取原始底稿並全自動對接填報中..."):
                    try:
                        df_acc_raw = pd.read_excel(file_acc, header=4)
                        df_acc_raw['姓名'] = df_acc_raw['姓名'].astype(str).str.strip()
                        dict_acc = df_acc_raw.groupby('姓名')[['A2類', 'A3類']].sum().to_dict(orient='index')
                        
                        traffic_dfs = []
                        for f in file_traf_list:
                            xl = pd.ExcelFile(f)
                            sn = xl.sheet_names
                            target_sheet = '分局月彙整總表' if '分局月彙整總表' in sn else ('月彙整總表' if '月彙整總表' in sn else sn[0])
                            traffic_dfs.append(pd.read_excel(f, sheet_name=target_sheet))
                        df_traf_all = pd.concat(traffic_dfs)
                        df_traf_all['姓名'] = df_traf_all['姓名'].astype(str).str.strip()
                        time_col = [c for c in df_traf_all.columns if '時數' in c]
                        time_col_name = time_col[0] if time_col else '總計尖峰時數'
                        dict_traf = df_traf_all.groupby('姓名')[time_col_name].sum().to_dict()
                        
                        xls_template = pd.ExcelFile(file_template)
                        template_sheets = xls_template.sheet_names
                        
                        ext_year, ext_month = "115", "04"
                        df_first_sheet = pd.read_excel(file_template, sheet_name=template_sheets[0], header=None)
                        for r in range(min(15, len(df_first_sheet))):
                            for c in range(min(10, len(df_first_sheet.columns))):
                                v = str(df_first_sheet.iloc[r, c])
                                m = re.search(r'開單日期[：:\s]*(\d{3})(\d{2})', v)
                                if m:
                                    ext_year, ext_month = m.group(1), m.group(2)
                                    break
                        
                        final_sheets = {}
                        summary_rows = []
                        g_cite = g_acc = g_traf = g_all = 0
                        
                        for sheet_name in template_sheets:
                            if '總表' in sheet_name or 'SUMMARY' in sheet_name.upper(): continue
                            df_sheet = pd.read_excel(file_template, sheet_name=sheet_name, header=None)
                            start_r, start_c = None, None
                            for r_idx, row in df_sheet.iterrows():
                                row_str = [str(x).strip() for x in row.values]
                                if '員警姓名' in row_str:
                                    start_r, start_c = r_idx, row_str.index('員警姓名')
                                    break
                            if start_r is not None:
                                df_header = df_sheet.iloc[start_r:, start_c:].copy()
                                df_header.reset_index(drop=True, inplace=True)
                                df_header.columns = [str(c).strip() for c in df_header.iloc[0]]
                                df_work = df_header.drop(0).reset_index(drop=True)
                                
                                member_rows_idx = []
                                for r in range(len(df_work)):
                                    name_cell = str(df_work.iloc[r, 0]).strip()
                                    if name_cell in ['小計', '總計', 'nan', 'None', '', '合計']: continue
                                    member_rows_idx.append(r)
                                    
                                df_members = df_work.iloc[member_rows_idx].copy().astype(object)
                                s_cite, s_acc, s_traf = 0, 0, 0
                                
                                for idx in df_members.index:
                                    name = str(df_members.at[idx, '員警姓名']).strip()
                                    cp = pd.to_numeric(df_members.at[idx, '取締點數'], errors='coerce') or 0
                                    a2 = dict_acc.get(name, {}).get('A2類', 0)
                                    a3 = dict_acc.get(name, {}).get('A3類', 0)
                                    ap = a2 * P_A2 + a3 * P_A3
                                    th = dict_traf.get(name, 0)
                                    tp = th * P_TRAF
                                    
                                    if 'A2件數' in df_members.columns: df_members.at[idx, 'A2件數'] = int(a2)
                                    if 'A3件數' in df_members.columns: df_members.at[idx, 'A3件數'] = int(a3)
                                    if '事故點數' in df_members.columns: df_members.at[idx, '事故點數'] = int(ap)
                                    if '交整時數' in df_members.columns: df_members.at[idx, '交整時數'] = int(th)
                                    if '交整點數' in df_members.columns: df_members.at[idx, '交整點數'] = int(tp)
                                    
                                    total_pts = cp + ap + tp
                                    df_members.at[idx, '個人總點數'] = int(total_pts)
                                    s_cite += cp; s_acc += ap; s_traf += tp
                                
                                sub_row_data = {c: "" for c in df_members.columns}
                                sub_row_data['員警姓名'] = '小計'
                                for col_n in df_members.columns:
                                    if col_n in ['員警姓名', '蓋章', '取締件數', 'A2件數', 'A3件數', '交整時數']: continue
                                    v_sum = pd.to_numeric(df_members[col_n], errors='coerce').sum()
                                    sub_row_data[col_n] = int(v_sum) if v_sum > 0 else 0
                                
                                df_final_sheet = pd.concat([df_members, pd.DataFrame([sub_row_data])], ignore_index=True)
                                final_sheets[sheet_name] = df_final_sheet
                                summary_rows.append([sheet_name, s_cite, s_acc, s_traf, s_cite + s_acc + s_traf])
                                g_cite += s_cite; g_acc += s_acc; g_traf += s_traf; g_all += (s_cite + s_acc + s_traf)
                        
                        df_pts_summary_final = pd.DataFrame([['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']] + summary_rows + [['合計', int(g_cite), int(g_acc), int(g_traf), int(g_all)]])
                        pts_output = io.BytesIO()
                        with pd.ExcelWriter(pts_output, engine='xlsxwriter') as writer:
                            df_pts_summary_final.to_excel(writer, sheet_name='總表', header=False, index=False)
                            for sn, df_f in final_sheets.items(): df_f.to_excel(writer, sheet_name=sn, index=False)
                        
                        pts_excel_data = pts_output.getvalue()
                        pts_filename = f"龍潭分局{ext_year}年{ext_month}月份_處理道路交通安全人員獎勵金點數統計表.xlsx"
                        
                        sub_title = f"【系統備份】龍潭分局 {ext_year}年{ext_month}月 處理道路交通安全人員獎勵金點數統計表(純點數對帳版)"
                        body_txt = f"郭同仁您好：\n\n系統已自動完成 {ext_year}年{ext_month}月份的處理道路交通安全人員獎勵金點數統計表數據回填任務。\n本次產出【僅點數表】，附件請查收對帳。"
                        send_report_email_auto([(pts_excel_data, pts_filename)], ext_year, ext_month, sub_title, body_txt)
                        
                        st.success(f"🏆 點數統計表全新生成成功！已同步發送備份至信箱。")
                        st.download_button("📥 下載【處理道路交通安全人員獎勵金點數統計表】(完美回填版)", pts_excel_data, pts_filename, use_container_width=True, type="primary")
                    except Exception as e:
                        st.error(f"❌ 發生錯誤：{str(e)}")

    # ==========================================================================
    # 💰 階段二：月底獎金請款面板
    # ==========================================================================
    with tab_pay:
        st.subheader("📂 請上傳最終確認的點數表以造冊請款")
        st.info("💡 說明：本區塊專用於月底造冊。系統會直接讀取點數表內的「個人總點數」進行獎金換算，不需重複上傳對帳檔案。")
        
        file_final_pts = st.file_uploader("👉 請上傳您已確認好數據的【處理道路交通安全人員獎勵金點數統計表】", type=['xls', 'xlsx'], key="pay_final")
        
        st.markdown("---")
        st.subheader("📝 獎金核發與自動按比例分配設定")
        point_value = st.number_input("💵 直接執行人員 - 每點獎金金額", value=1.905, format="%.3f", step=0.001, key="pay_val")
        target_direct_budget = st.number_input("🎯 警察局核撥【直接執行人員】總獎金目標 (元) *若為 0 則不啟動自動平帳", value=0, step=1, key="pay_tgt")

        st.markdown("##### 👥 共同作業及配合人員 - 獎金預算配置")
        budget_type = st.selectbox("請選擇預算輸入方式：", ["A. 直接輸入【共同作業人員】的總分配預算", "B. 輸入【全分局】本月核撥總預算"], key="pay_btype")
        if "A" in budget_type:
            budget_input = st.number_input("💰 輸入【共同作業人員】總預算 (元)", value=9467, step=100, key="pay_binA")
        else:
            budget_input = st.number_input("💰 輸入【全分局】核撥總預算 (元)", value=50000, step=100, key="pay_binB")
        
        st.markdown("**共同作業名單配置**")
        roster_file = 'coworkers_roster.csv'
        
        # --- 【核心修正點：完美重啟名單初始化防線】 ---
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
        if '金額' in df_display.columns: df_display = df_display.drop(columns=['金額'])
        col_cfg = {
            "排序調整": st.column_config.NumberColumn("排序調整 🔢", help="修改數字可調整順序", min_value=1, format="%d"),
            "分配類別": st.column_config.SelectboxColumn("分配類別", options=["負責管考(72%)", "勤務督導(20%)", "其他配合(8%)"], required=True)
        }
        st.data_editor(df_display, num_rows="dynamic", use_container_width=True, hide_index=True, height=300, 
                       column_config=col_cfg, key="co_editor", on_change=on_data_edited)
        
        if st.button("💾 儲存最新名單為預設值", use_container_width=True, key="btn_save_co"):
            st.session_state.current_roster.to_csv(roster_file, index=False, encoding='utf-8-sig')
            st.success("✅ 名單已永久儲存！")
            st.rerun()

        st.markdown("---")
        if st.button("🚀 一鍵自動計算並產出【獎金印領清冊】", type="primary", use_container_width=True, key="btn_pay"):
            if not file_final_pts:
                st.error("⚠️ 請上傳已完成數據填報的【點數統計表】Excel！")
            else:
                with st.spinner("正在讀取點數、精算獎金並執行全分局平帳..."):
                    try:
                        xls_template = pd.ExcelFile(file_final_pts)
                        template_sheets = xls_template.sheet_names
                        
                        ext_year, ext_month = "115", "04"
                        df_first_sheet = pd.read_excel(file_final_pts, sheet_name=template_sheets[0], header=None)
                        for r in range(min(15, len(df_first_sheet))):
                            for c in range(min(10, len(df_first_sheet.columns))):
                                v = str(df_first_sheet.iloc[r, c])
                                m = re.search(r'開單日期[：:\s]*(\d{3})(\d{2})', v)
                                if m:
                                    ext_year, ext_month = m.group(1), m.group(2)
                                    break
                        
                        direct_exec_list = []
                        for sheet_name in template_sheets:
                            if '總表' in sheet_name or 'SUMMARY' in sheet_name.upper(): continue
                            df_sheet = pd.read_excel(file_final_pts, sheet_name=sheet_name, header=None)
                            start_r, start_c = None, None
                            for r_idx, row in df_sheet.iterrows():
                                row_str = [str(x).strip() for x in row.values]
                                if '員警姓名' in row_str:
                                    start_r, start_c = r_idx, row_str.index('員警姓名')
                                    break
                            if start_r is not None:
                                df_header = df_sheet.iloc[start_r:, start_c:].copy()
                                df_header.reset_index(drop=True, inplace=True)
                                df_header.columns = [str(c).strip() for c in df_header.iloc[0]]
                                df_work = df_header.drop(0).reset_index(drop=True)
                                
                                for r in range(len(df_work)):
                                    name = str(df_work.iloc[r, 0]).strip()
                                    if name in ['小計', '總計', 'nan', 'None', '', '合計']: continue
                                    
                                    total_pts = pd.to_numeric(df_work.iloc[r].get('個人總點數', 0), errors='coerce') or 0
                                    cp = pd.to_numeric(df_work.iloc[r].get('取締點數', 0), errors='coerce') or 0
                                    ap = pd.to_numeric(df_work.iloc[r].get('事故點數', 0), errors='coerce') or 0
                                    tp = pd.to_numeric(df_work.iloc[r].get('交整點數', 0), errors='coerce') or 0
                                    
                                    if total_pts > 0:
                                        direct_exec_list.append({
                                            "單位名稱": sheet_name, "員警姓名": name,
                                            "取締件數": df_work.iloc[r].get('取締件數', ''), "取締點數": cp,
                                            "A2件數": df_work.iloc[r].get('A2件數', 0), "A3件數": df_work.iloc[r].get('A3件數', 0),
                                            "事故點數": ap, "交整時數": df_work.iloc[r].get('交整時數', 0),
                                            "交整點數": tp, "個人總點數": total_pts
                                        })
                        
                        df_direct_exec = pd.DataFrame(direct_exec_list)
                        if df_direct_exec.empty:
                            st.error("⚠️ 該點數表中未偵測到任何個人的點數紀錄。")
                            return
                            
                        df_direct_exec.insert(0, '序號', range(1, len(df_direct_exec) + 1))
                        df_direct_exec['每點獎金'] = point_value
                        df_direct_exec['實領獎金'] = (df_direct_exec['個人總點數'] * point_value).round().astype(int)
                        direct_total_money = df_direct_exec['實領獎金'].sum()
                        
                        if target_direct_budget > 0:
                            diff = target_direct_budget - direct_total_money
                            if diff != 0:
                                n = len(df_direct_exec)
                                base, rem = diff // n, abs(diff) % n
                                sign = 1 if diff > 0 else -1
                                df_direct_exec['實領獎金'] += base
                                if rem > 0: df_direct_exec.iloc[:rem, df_direct_exec.columns.get_loc('實領獎金')] += sign
                                direct_total_money = df_direct_exec['實領獎金'].sum()
                        
                        df_direct_exec['蓋章'] = ""
                        direct_total_row = {c: "" for c in df_direct_exec.columns}
                        direct_total_row['員警姓名'] = '合計'; direct_total_row['實領獎金'] = direct_total_money
                        df_direct_exec = pd.concat([df_direct_exec, pd.DataFrame([direct_total_row])], ignore_index=True)
                        
                        df_coworkers_work = st.session_state.current_roster.copy()
                        df_coworkers_work = sort_coworkers(df_coworkers_work)
                        
                        coworker_pool = int(budget_input) if "A" in budget_type else int(budget_input) - direct_total_money
                        pool_72 = int(np.round(coworker_pool * 0.72))
                        pool_20 = int(np.round(coworker_pool * 0.20))
                        pool_08 = coworker_pool - pool_72 - pool_20
                        
                        df_coworkers_work['核發金額'] = 0
                        mask_72 = df_coworkers_work['分配類別'] == "負責管考(72%)"
                        df_72 = df_coworkers_work[mask_72].copy()
                        
                        if not df_72.empty and pool_72 > 0:
                            main_pool = int(np.round(pool_72 * 0.08))
                            chief_mask = df_72['職別'].str.contains('分局長', na=False)
                            vice_mask = df_72['職別'].str.contains('副分局長', na=False)
                            if chief_mask.any(): df_72.loc[chief_mask, '核發金額'] = int(np.round(main_pool * 0.60))
                            if vice_mask.any(): df_72.loc[vice_mask, '核發金額'] = int(np.round(main_pool * 0.40 / vice_mask.sum()))
                            
                            actual_main_used = df_72['核發金額'].sum()
                            sup_pool = int(np.round(pool_72 * 0.56))
                            traf_pool = int(np.round(pool_72 * 0.26))
                            clerk_pool = pool_72 - actual_main_used - sup_pool - traf_pool
                            
                            sup_mask = (df_72['單位'].str.contains('派出所|交通分隊', na=False)) & (df_72['職別'].str.contains('所長|副所長|分隊長|小隊長', na=False))
                            sup_indices = df_72[sup_mask].index
                            if len(sup_indices) > 0:
                                base = int(np.floor(sup_pool / len(sup_indices)))
                                df_72.loc[sup_indices, '核發金額'] = base
                                extra = sup_pool - base * len(sup_indices)
                                if extra > 0: df_72.loc[sup_indices[:extra], '核發金額'] += 1
                            
                            traf_mask = df_72['單位'] == "交通組"
                            traf_indices = df_72[traf_mask].index
                            if len(traf_indices) > 0:
                                base = int(np.floor(traf_pool / len(traf_indices)))
                                df_72.loc[traf_indices, '核發金額'] = base
                                extra = traf_pool - base * len(traf_indices)
                                if extra > 0: df_72.loc[traf_indices[:extra], '核發金額'] += 1
                            
                            clerk_mask = (df_72['單位'].str.contains('派出所|交通分隊', na=False)) & (df_72['職別'].str.contains('業務承辦人|承辦', na=False))
                            clerk_indices = df_72[clerk_mask].index
                            if len(clerk_indices) > 0:
                                base = int(np.floor(clerk_pool / len(clerk_indices)))
                                df_72.loc[clerk_indices, '核發金額'] = base
                                extra = clerk_pool - base * len(clerk_indices)
                                if extra > 0: df_72.loc[clerk_indices[:extra], '核發金額'] += 1
                            
                            df_coworkers_work.loc[mask_72, '核發金額'] = df_72['核發金額']
                        
                        for cat, pool in [("勤務督導(20%)", pool_20), ("其他配合(8%)", pool_08)]:
                            cat_mask = df_coworkers_work['分配類別'] == cat
                            count = cat_mask.sum()
                            if count > 0 and pool > 0:
                                int_amount = int(np.floor(pool / count))
                                amounts = np.full(count, int_amount)
                                diff_rem = pool - amounts.sum()
                                if diff_rem > 0: amounts[:diff_rem] += 1
                                df_coworkers_work.loc[cat_mask, '核發金額'] = amounts
                        
                        df_coworkers_output = df_coworkers_work.rename(columns={'核發金額': '金額'})
                        sub_72 = df_coworkers_output[df_coworkers_output['分配類別'] == "負責管考(72%)"]['金額'].sum()
                        sub_20 = df_coworkers_output[df_coworkers_output['分配類別'] == "勤務督導(20%)"]['金額'].sum()
                        sub_08 = df_coworkers_output[df_coworkers_output['分配類別'] == "其他配合(8%)"]['金額'].sum()
                        coworkers_total_money = sub_72 + sub_20 + sub_08
                        
                        df_payroll_summary = pd.DataFrame([
                            {"項目": "一、直接執行人員", "金額": direct_total_money},
                            {"項目": "二、共同作業-負責管考(72%)", "金額": sub_72},
                            {"項目": "二、共同作業-勤務督導(20%)", "金額": sub_20},
                            {"項目": "二、共同作業-開他配合(8%)", "金額": sub_08},
                            {"項目": "共同作業人員小計", "金額": coworkers_total_money},
                            {"項目": "本月合計應發放", "金額": direct_total_money + coworkers_total_money}
                        ])
                        
                        df_coworkers_final_sheet = df_coworkers_output.copy()
                        traf_督導_mask = (df_coworkers_final_sheet['單位'] == "交通組") & (df_coworkers_final_sheet['分配類別'] == "勤務督導(20%)")
                        for idx, row in df_coworkers_final_sheet[traf_督導_mask].iterrows():
                            p_name = row['姓名']; p_money = row['金額']
                            if p_money > 0:
                                target_idx = df_coworkers_final_sheet[(df_coworkers_final_sheet['姓名'] == p_name) & (df_coworkers_final_sheet['分配類別'] == "負責管考(72%)")].index
                                if not target_idx.empty:
                                    df_coworkers_final_sheet.at[target_idx[0], '金額'] += p_money
                                    df_coworkers_final_sheet.at[idx, '金額'] = 0
                        
                        df_coworkers_final_sheet = df_coworkers_final_sheet[~((df_coworkers_final_sheet['單位'] == "交通組") & (df_coworkers_final_sheet['分配類別'] == "勤務督導(20%)") & (df_coworkers_final_sheet['金額'] == 0))]
                        coworker_sheet_total_money = df_coworkers_final_sheet['金額'].sum()
                        
                        if '排序調整' in df_coworkers_final_sheet.columns:
                            df_coworkers_final_sheet['排序調整'] = pd.to_numeric(df_coworkers_final_sheet['排序調整'], errors='coerce').fillna(999).astype(int)
                            df_coworkers_final_sheet.sort_values(by=['排序調整', '單位', '姓名'], ascending=[True, True, True], inplace=True)
                            df_coworkers_final_sheet.drop(columns=['排序調整'], inplace=True, errors='ignore')
                        
                        df_coworkers_final_sheet.drop(columns=['分配類別'], inplace=True, errors='ignore')
                        df_coworkers_final_sheet.reset_index(drop=True, inplace=True)
                        df_coworkers_final_sheet.insert(0, '序號', range(1, len(df_coworkers_final_sheet) + 1))
                        df_coworkers_final_sheet['蓋章'] = ""
                        
                        total_row_data = {c: "" for c in df_coworkers_final_sheet.columns}
                        total_row_data['單位'] = '合計'; total_row_data['金額'] = coworker_sheet_total_money
                        df_coworkers_final_sheet = pd.concat([df_coworkers_final_sheet, pd.DataFrame([total_row_data])], ignore_index=True)
                        
                        grand_total_row_data = {c: "" for c in df_coworkers_final_sheet.columns}
                        grand_total_row_data['單位'] = '總計（含直接執行人員）'; grand_total_row_data['金額'] = direct_total_money + coworker_sheet_total_money
                        
                        # --- 【核心修正點二】修正大後方字典結構串接錯誤，回填為正確的 DataFrame 變數 ---
                        df_coworkers_final_sheet = pd.concat([df_coworkers_final_sheet, pd.DataFrame([grand_total_row_data])], ignore_index=True)
                        
                        payroll_output = io.BytesIO()
                        with pd.ExcelWriter(payroll_output, engine='xlsxwriter') as writer:
                            workbook = writer.book
                            border_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
                            
                            df_direct_exec.to_excel(writer, sheet_name='直接執行人員', index=False)
                            ws1 = writer.sheets['直接執行人員']
                            ws1.set_portrait(); ws1.set_paper(9); stamp_col1 = df_direct_exec.columns.get_loc('蓋章')
                            ws1.set_column(stamp_col1, stamp_col1, 22)
                            for r in range(len(df_direct_exec) + 1):
                                ws1.set_row(r, 45 if r > 0 else 25)
                                for c in range(len(df_direct_exec.columns)):
                                    ws1.write(r, c, df_direct_exec.iloc[r-1, c] if r > 0 else df_direct_exec.columns[c], border_format)
                            
                            df_coworkers_final_sheet.to_excel(writer, sheet_name='共同作業及配合人員', index=False)
                            ws2 = writer.sheets['共同作業及配合人員']
                            ws2.set_portrait(); ws2.set_paper(9); stamp_col2 = df_coworkers_final_sheet.columns.get_loc('蓋章')
                            ws2.set_column(stamp_col2, stamp_col2, 22)
                            
                            data_len = len(df_coworkers_final_sheet)
                            main_data_len = data_len - 2
                            for r in range(main_data_len + 1):
                                ws2.set_row(r, 45 if r > 0 else 25)
                                for c in range(len(df_coworkers_final_sheet.columns)):
                                    ws2.write(r, c, df_coworkers_final_sheet.iloc[r-1, c] if r > 0 else df_coworkers_final_sheet.columns[c], border_format)
                                    
                            style_total = workbook.add_format({'border': 1, 'bold': True, 'align': 'center', 'valign': 'vcenter'})
                            ws2.set_row(main_data_len + 1, 45)
                            ws2.merge_range(main_data_len + 1, 0, main_data_len + 1, 3, "合計", style_total)
                            ws2.write(main_data_len + 1, 4, coworker_sheet_total_money, style_total)
                            
                            ws2.set_row(main_data_len + 2, 45)
                            ws2.merge_range(main_data_len + 2, 0, main_data_len + 2, 3, "總計（含直接執行人員）", style_total)
                            ws2.write(main_data_len + 2, 4, direct_total_money + coworker_sheet_total_money, style_total)
                            
                            sign_start_row = data_len + 2
                            sign_title_format = workbook.add_format({'font_name': 'Microsoft JhengHei', 'font_size': 12, 'bold': True})
                            ws2.write(sign_start_row, 0, "製表人：", sign_title_format)
                            ws2.write(sign_start_row, 2, "人事：", sign_title_format)
                            ws2.write(sign_start_row, 4, "主計：", sign_title_format)
                            ws2.write(sign_start_row, 6, "分局長：", sign_title_format)
                            
                            df_payroll_summary.to_excel(writer, sheet_name='處理道路交通安全人員獎勵金支領一覽表', index=False)
                        
                        payroll_excel_data = payroll_output.getvalue()
                        payroll_filename = f"龍潭分局{ext_year}年{ext_month}月份_處理道路交通安全人員獎勵金印領清冊.xlsx"
                        
                        sub_title = f"【系統備份】龍潭分局 {ext_year}年{ext_month}月 處理道路交通安全人員獎勵金印領清冊(核銷專用版)"
                        body_txt = f"郭同仁您好：\n\n系統已自動完成 {ext_year}年{ext_month}月份的獎金核算與自動平帳作業。\n本次產出【僅印領清冊】，可直接送交核銷。"
                        send_report_email_auto([(payroll_excel_data, payroll_filename)], ext_year, ext_month, sub_title, body_txt)
                        
                        st.success(f"🚀 已啟動直接請款大腦！成功產出 {ext_month} 月份【獎金印領清冊】，蓋章欄位已強制置右。")
                        st.download_button("📥 下載【處理道路交通安全人員獎勵金印領清冊】(官方核銷版)", payroll_excel_data, payroll_filename, use_container_width=True, type="primary")
                    except Exception as e:
                        st.error(f"❌ 發生錯誤：{str(e)}")


if __name__ == "__main__":
    p18_page()
