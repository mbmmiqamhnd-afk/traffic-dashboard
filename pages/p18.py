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

def send_report_email_auto(file_data, filename, year, month):
    """
    自動從 st.secrets 讀取設定並寄信給自己
    """
    try:
        if "email" not in st.secrets:
            return False, "找不到 st.secrets 中的 email 設定"
            
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = sender
        msg['Subject'] = f"【系統備份】龍潭分局 {year}年{month}月 獎勵金點數統計表暨印領清冊"
        
        body = f"郭同仁您好：\n\n系統已自動完成 {year}年{month}月份的獎勵金點數彙整與印領清冊產出。\n附件為最新產出的統計報表，請查收。"
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
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
    st.info("權重已固定 (A2:10, A3:5, 交整:5)。系統將自動裁剪格式、計算獎金、產出清冊並發送郵件。")

    P_A2, P_A3, P_TRAF = 10.0, 5.0, 5.0

    # 1. 檔案上傳區
    st.subheader("📂 1. 當月原始資料上傳")
    c1, c2 = st.columns(2)
    file_template = c1.file_uploader("1. 上傳當月【獎勵金點數統計表】", type=['xlsx'])
    file_acc = c2.file_uploader("2. 上傳當月【處理交通事故案件統計表】", type=['xls', 'xlsx'])
    file_traf_list = st.file_uploader("3. 上傳當月【各單位_交通疏導統計】(可多選)", type=['xlsx'], accept_multiple_files=True)
    
    # 2. 印領清冊參數與名單設定
    st.subheader("📝 2. 印領清冊設定")
    point_value = st.number_input("💵 每點獎金金額", value=1.905, format="%.3f", step=0.001)

    st.markdown("##### 👥 共同作業及配合人員名單")
    st.caption("💡 提示：可直接在表格內修改金額、在最下方點擊「+」新增人員，或選取左側核取方塊按 Delete 刪除。也支援直接從 Excel 複製貼上！")
    
    # 預設底稿 (不包含序號，後端會自動產生)
    default_coworkers_data = [
        {"單位": "龍潭分局", "職別": "分局長", "姓名": "施宇峰", "金額": 301, "蓋章": ""},
        {"單位": "龍潭分局", "職別": "副分局長", "姓名": "何憶雯", "金額": 100, "蓋章": ""},
        {"單位": "龍潭分局", "職別": "副分局長", "姓名": "蔡志明", "金額": 100, "蓋章": ""},
        {"單位": "交通組", "職別": "業務單位主管", "姓名": "陳維明", "金額": 298, "蓋章": ""},
        {"單位": "交通組", "職別": "交通業務承辦人", "姓名": "盧冠仁", "金額": 298, "蓋章": ""},
        {"單位": "交通組", "職別": "交通業務承辦人", "姓名": "李峯甫", "金額": 298, "蓋章": ""},
        {"單位": "會計室", "職別": "主計", "姓名": "郭貞彣", "金額": 77, "蓋章": ""}
        # 你可以在此處繼續補充其他固定班底
    ]
    df_coworkers_default = pd.DataFrame(default_coworkers_data)

    # 顯示互動式資料表
    edited_df_coworkers = st.data_editor(
        df_coworkers_default,
        num_rows="dynamic", # 允許動態新增/刪除行
        use_container_width=True,
        hide_index=True,
        column_config={
            "金額": st.column_config.NumberColumn("金額", min_value=0, step=1, format="%d")
        }
    )

    if st.button("🚀 執行彙整、計算獎金與自動寄信", type="primary", use_container_width=True):
        if not (file_template and file_acc and file_traf_list):
            st.error("⚠️ 請確保上方 3 種點數統計資料皆已完成上傳！")
            return

        with st.spinner("正在計算數據、產出印領清冊並發送郵件..."):
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

                # --- C. 點數表裁剪與重新計算 ---
                final_sheets = {}
                summary_rows = []
                g_cite, g_acc, g_traf, g_all = 0, 0, 0, 0
                
                # 用於印領清冊的「直接執行人員」清單
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
                            
                            # 蒐集印領清冊資料 (直接執行人員)
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

                # --- D. 產生印領清冊 DataFrame ---
                df_direct_exec = pd.DataFrame(direct_exec_list)
                df_direct_exec.insert(0, '序號', range(1, len(df_direct_exec) + 1))
                direct_total_money = df_direct_exec['實領獎金'].sum()
                
                # 處理共同作業人員 (來自前端 Data Editor)
                df_coworkers_final = edited_df_coworkers.copy()
                # 濾除所有欄位都是空值的行，避免匯出多餘空白行
                df_coworkers_final.dropna(how='all', inplace=True)
                # 自動產生序號，插在最前面
                df_coworkers_final.insert(0, '序號', range(1, len(df_coworkers_final) + 1))
                
                coworkers_total_money = pd.to_numeric(df_coworkers_final['金額'], errors='coerce').fillna(0).sum()

                # 產生支領一覽表
                summary_data = [
                    {"項目": "直接執行人員", "金額": direct_total_money},
                    {"項目": "共同作業及配合人員", "金額": coworkers_total_money},
                    {"項目": "合計", "金額": direct_total_money + coworkers_total_money},
                    {"項目": "製表人", "金額": ""}
                ]
                df_payroll_summary = pd.DataFrame(summary_data)

                # --- E. 封裝成兩個 Excel 檔案 ---
                # 1. 點數統計表
                pts_output = io.BytesIO()
                df_pts_summary = pd.DataFrame([['單位名稱', '取締點數', '事故點數', '交整點數', '個人總點數']] + summary_rows + [['合計', g_cite, g_acc, g_traf, g_all]])
                with pd.ExcelWriter(pts_output, engine='xlsxwriter') as writer:
                    df_pts_summary.to_excel(writer, sheet_name='總表', header=False, index=False)
                    for sn, df_f in final_sheets.items():
                        df_f.to_excel(writer, sheet_name=sn, index=False)
                pts_excel_data = pts_output.getvalue()
                pts_filename = f"龍潭分局{ext_year}年{ext_month}月份_點數統計表.xlsx"

                # 2. 印領清冊
                payroll_output = io.BytesIO()
                with pd.ExcelWriter(payroll_output, engine='xlsxwriter') as writer:
                    df_direct_exec.to_excel(writer, sheet_name='直接執行人員', index=False)
                    if not df_coworkers_final.empty:
                        df_coworkers_final.to_excel(writer, sheet_name='共同作業及配合人員', index=False)
                    df_payroll_summary.to_excel(writer, sheet_name='獎勵金支領一覽表', index=False)
                payroll_excel_data = payroll_output.getvalue()
                payroll_filename = f"龍潭分局{ext_year}年{ext_month}月份_獎勵金印領清冊.xlsx"

                # --- F. 自動發送郵件 ---
                ok, err = send_report_email_auto(payroll_excel_data, payroll_filename, ext_year, ext_month)
                
                if ok:
                    st.success(f"✅ 雙報表產出成功！印領清冊已自動備份至您的信箱。")
                else:
                    st.warning(f"⚠️ 報表已產出，但郵件發送失敗: {err}")

                c5, c6 = st.columns(2)
                c5.download_button(label="📥 下載【點數統計表】", data=pts_excel_data, file_name=pts_filename, use_container_width=True)
                c6.download_button(label="📥 下載【印領清冊】", data=payroll_excel_data, file_name=payroll_filename, use_container_width=True, type="primary")

            except Exception as e:
                st.error(f"❌ 發生錯誤：{str(e)}")

if __name__ == "__main__":
    p18_page()
