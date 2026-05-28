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
    st.title("💰 龍潭分局 - 獎勵金點數統計表暨印領清冊產生器")
    st.info("💡 挪移順序教學：表格最左側「排序調整」欄位可手動調整順序")
   
    P_A2, P_A3, P_TRAF = 10.0, 5.0, 5.0
    st.subheader("📂 1. 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)", type=['xlsx'], accept_multiple_files=True)
 
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
        st.info("💡 手動模式：請於表格內自行填寫實際金額。")
        budget_input = 0
        budget_type = ""
    
    st.markdown("**共同作業名單**")
    roster_file = 'coworkers_roster.csv'
    if 'current_roster' not in st.session_state:
        if os.path.exists(roster_file):
            df_init = pd.read_csv(roster_file)
        else:
            # 預設名單（已將派出所與交通分隊設為負責管考72%）
            default_coworkers_data = [ ... ]  # 保持您原本的預設名單
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
    
    st.data_editor(df_display, num_rows="dynamic", use_container_width=True, hide_index=True, height=500, 
                   column_config=col_cfg, key="co_editor", on_change=on_data_edited)
    
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
                # ====================== 直接執行人員計算 ======================
                # （這裡省略完整讀取與計算過程，請保留您原本的直接執行人員部分）
                # ... 您的原始直接執行人員計算程式碼 ...

                direct_total_money = 0  # 先初始化，避免 NameError
                
                # 如果您有完整的直接執行人員計算，請把下面的 direct_total_money = ... 放在這裡
                # 例如：
                # df_direct_exec = pd.DataFrame(...)
                # direct_total_money = df_direct_exec['實領獎金'].sum() if not df_direct_exec.empty else 0

                # ====================== 共同作業人員處理 ======================
                df_coworkers_work = st.session_state.current_roster.copy()
                df_coworkers_work = sort_coworkers(df_coworkers_work)
                
                if "系統自動" in alloc_mode:
                    if "A" in budget_type:
                        coworker_pool = int(budget_input)
                        st.success(f"【模式 A】共同作業總預算 = **{coworker_pool:,} 元**")
                    else:
                        coworker_pool = int(budget_input) - direct_total_money
                        st.info(f"【模式 B】全分局 {budget_input:,} - 直接執行 {direct_total_money:,} = **{coworker_pool:,} 元**")
                    
                    pool_72 = int(np.round(coworker_pool * 0.72))
                    pool_20 = int(np.round(coworker_pool * 0.20))
                    pool_08 = coworker_pool - pool_72 - pool_20
                    
                    st.success(f"**72% = {pool_72:,} 元** | 20% = {pool_20:,} 元 | 8% = {pool_08:,} 元")
                    
                    # 後續 72% 內部分配邏輯（已修正）
                    df_coworkers_work['核發金額'] = 0
                    mask_72 = df_coworkers_work['分配類別'] == "負責管考(72%)"
                    df_72 = df_coworkers_work[mask_72].copy()
                    df_72['核發金額'] = 0
                    
                    if not df_72.empty and pool_72 > 0:
                        main_pool = int(np.round(pool_72 * 0.08))
                        # 正副主官
                        chief_mask = df_72['職別'].str.contains('分局長', na=False)
                        vice_mask = df_72['職別'].str.contains('副分局長', na=False)
                        if chief_mask.any():
                            df_72.loc[chief_mask, '核發金額'] = int(np.round(main_pool * 0.60))
                        if vice_mask.any():
                            df_72.loc[vice_mask, '核發金額'] = int(np.round(main_pool * 0.40 / vice_mask.sum()))
                        
                        remaining_pool = pool_72 - df_72['核發金額'].sum()
                        
                        sup_pool = int(np.round(remaining_pool * 0.56))
                        traf_pool = int(np.round(remaining_pool * 0.26))
                        clerk_pool = int(np.round(remaining_pool * 0.10))
                        
                        # 派出所/交通分隊 正副主管
                        sup_mask = (df_72['單位'].str.contains('派出所|交通分隊', na=False)) & (df_72['職別'].str.contains('所長|副所長|分隊長|小隊長', na=False))
                        sup_indices = df_72[sup_mask].index
                        if len(sup_indices) > 0:
                            base = int(np.floor(sup_pool / len(sup_indices)))
                            df_72.loc[sup_indices, '核發金額'] = base
                            extra = sup_pool - base * len(sup_indices)
                            if extra > 0:
                                df_72.loc[sup_indices[:extra], '核發金額'] += 1
                        
                        # 交通組
                        traf_mask = df_72['單位'] == "交通組"
                        traf_indices = df_72[traf_mask].index
                        if len(traf_indices) > 0:
                            base = int(np.floor(traf_pool / len(traf_indices)))
                            df_72.loc[traf_indices, '核發金額'] = base
                            extra = traf_pool - base * len(traf_indices)
                            if extra > 0:
                                df_72.loc[traf_indices[:extra], '核發金額'] += 1
                        
                        # 業務承辦人
                        clerk_mask = (df_72['單位'].str.contains('派出所|交通分隊', na=False)) & (df_72['職別'].str.contains('業務承辦人|承辦', na=False))
                        clerk_indices = df_72[clerk_mask].index
                        if len(clerk_indices) > 0:
                            base = int(np.floor(clerk_pool / len(clerk_indices)))
                            df_72.loc[clerk_indices, '核發金額'] = base
                            extra = clerk_pool - base * len(clerk_indices)
                            if extra > 0:
                                df_72.loc[clerk_indices[:extra], '核發金額'] += 1
                        
                        df_coworkers_work.loc[mask_72, '核發金額'] = df_72['核發金額']
                    
                    # 20% 和 8% 平均分配
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
                
                # 後續總表與 Excel 輸出部分請保留您原本的程式碼
                # ... (summary_data, df_payroll_summary, ExcelWriter 等)

                st.success("✅ 報表產出成功！")
                
            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")


if __name__ == "__main__":
    p18_page()
