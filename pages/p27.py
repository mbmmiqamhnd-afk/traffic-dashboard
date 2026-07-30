import streamlit as st
import pandas as pd
import re
import io
import os
import smtplib
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ==========================================
# 0. 輔助函式：發送多個檔案附件至自己的信箱
# ==========================================
def send_multiple_files_email(file_buffers_dict):
    try:
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender
        msg["Subject"] = "🚔 交通違規舉發績效結算表 (批次產出)"
        
        body_text = (
            "長官您好，\n\n"
            "系統已自動產出最新結算的【交通違規舉發績效結算表與個人明細】。\n\n"
            "本次共產出以下單位報表，詳見附件：\n"
        )
        for fname in file_buffers_dict.keys():
            body_text += f"- {fname}\n"
            
        body_text += "\n本信件由交通執法自動化分析引擎發送。"
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        for fname, buffer in file_buffers_dict.items():
            part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(buffer.getvalue())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(fname)}")
            msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
            
        return True, None
    except KeyError:
        return False, "系統找不到 Email 寄件設定。請確認 `secrets.toml` 中已設定 `[email]` 區塊，並包含 `user` 與 `password`。"
    except smtplib.SMTPAuthenticationError:
        return False, "Email 帳號或密碼驗證失敗！請確認使用的是「應用程式密碼」而非一般登入密碼。"
    except Exception as e:
        return False, str(e)

# ==========================================
# 1. 頁面基本設定與側邊欄
# ==========================================
st.set_page_config(page_title="舉發績效結算", page_icon="👮", layout="wide")

st.title("⚡ 員警交通違規舉發績效結算")
st.markdown("本頁面專門處理**員警個人績效配分**與**門檻結算**。系統會依據您上傳的檔案**逐一產出各單位的獨立報表**，並自動優化排版。")

st.sidebar.header("⚙️ 結算參數設定")
st.sidebar.info("💡 **智慧基準偵測**：\n系統將自動分析上傳的檔名。若檔名包含「交通分隊」，將自動套用 **800分** 基準；其餘單位一律自動套用 **400分** 基準。")

st.sidebar.markdown("---")
st.sidebar.subheader("📂 步驟 1：上傳配分表")
db_file = st.sidebar.file_uploader("外部配分表 (如：檔案 B)", type=["xlsx"], key="db_file")

st.sidebar.subheader("📂 步驟 2：上傳原始舉發資料")
data_files = st.sidebar.file_uploader("批次選擇多個單位的半年期 Excel 檔案 (支援原始無配分欄位格式)", type=["xlsx", "xls"], accept_multiple_files=True, key="data_files")

# ==========================================
# 2. 核心結算邏輯
# ==========================================
if db_file and data_files:
    with st.spinner("🔄 正在逐一讀取與結算各單位資料..."):
        try:
            df_db = pd.read_excel(db_file)
            if '違規條款' not in df_db.columns:
                st.error("❌ 配分表缺少『違規條款』欄位！請檢查檔案格式。")
                st.stop()
        except Exception as e:
            st.error(f"❌ 讀取配分表失敗：{e}")
            st.stop()

        db_map = {}
        for _, row in df_db.iterrows():
            rule = str(row.get('違規條款', '')).strip()
            if rule:
                s = pd.to_numeric(row.get('攔舉配分', 0), errors='coerce')
                d = pd.to_numeric(row.get('逕舉配分', 0), errors='coerce')
                db_map[rule] = {
                    'stop': 0 if pd.isna(s) else int(s),
                    'dir': 0 if pd.isna(d) else int(d)
                }
                
        def extract_officer_name(df_head):
            for r_idx, row in df_head.iterrows():
                for c_idx, val in enumerate(row.values):
                    val_str = str(val).strip()
                    if "舉發員警" in val_str:
                        clean = re.sub(r'舉發員警[:：]?', '', val_str).strip()
                        if clean: return clean
                        if c_idx + 1 < len(row.values):
                            next_val = str(row.values[c_idx + 1]).strip()
                            if next_val and next_val.lower() != 'nan':
                                return next_val
            return ""

        all_summaries = {}
        all_output_buffers = {}

        for f in data_files:
            f.seek(0)
            
            unit_name = re.sub(r'\.[a-zA-Z0-9]+$', '', f.name)
            is_800_quota = "交通分隊" in unit_name
            quota = 800 if is_800_quota else 400
            threshold_7x = quota * 7
            unit_type_label = f"{unit_name} (基準 {quota} 分)"
            
            processed_sheets = []
            
            try:
                xls = pd.ExcelFile(f)
                for sheet_name in xls.sheet_names:
                    raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    raw_df = raw_df.astype('object')
                    
                    header_idx = -1
                    officer_name = extract_officer_name(raw_df.head(20))
                    if not officer_name:
                        officer_name = sheet_name.strip()
                    
                    for idx, row in raw_df.iterrows():
                        if "違規條款" in row.astype(str).str.replace(" ", "").values:
                            header_idx = idx
                            break
                    
                    if header_idx == -1:
                        continue
                    
                    # 💡 提取列印日期 (在刪除空白欄位前先撈出來保護)
                    print_date_val = None
                    for r_search in range(min(5, len(raw_df))):
                        for c_search in range(len(raw_df.columns)):
                            val = raw_df.iat[r_search, c_search]
                            if pd.notna(val) and "列印日期" in str(val):
                                print_date_val = str(val).strip()
                                raw_df.iat[r_search, c_search] = None # 撈出後清空原位置，避免重複顯示
                                break
                        if print_date_val:
                            break
                    
                    # 清理無關欄位：找出標題為空或 nan 的欄位並刪除
                    header_row_temp = [str(x).strip().replace(" ", "") for x in raw_df.iloc[header_idx]]
                    cols_to_keep = [c for c, val in enumerate(header_row_temp) if val not in ["nan", "None", ""]]
                    
                    # 只保留有效欄位，並重置索引
                    raw_df = raw_df.iloc[:, cols_to_keep]
                    raw_df.columns = range(raw_df.shape[1])

                    # 重新取得清理後的標題列
                    header_row = [str(x).strip().replace(" ", "") for x in raw_df.iloc[header_idx]]
                    col_rule = header_row.index("違規條款") if "違規條款" in header_row else -1
                    col_s_cnt = header_row.index("攔停數") if "攔停數" in header_row else -1
                    col_d_cnt = header_row.index("逕舉數") if "逕舉數" in header_row else -1
                    
                    col_s_score = header_row.index("攔舉配分") if "攔舉配分" in header_row else -1
                    col_d_score = header_row.index("逕舉配分") if "逕舉配分" in header_row else -1
                    col_subtotal = header_row.index("小計") if "小計" in header_row else -1
                    
                    if col_rule == -1 or col_s_cnt == -1 or col_d_cnt == -1:
                        continue

                    # 動態插入欄位
                    if col_s_score == -1:
                        idx_s_score = col_s_cnt + 1
                        raw_df.insert(idx_s_score, f'new_{idx_s_score}', None)
                        raw_df.iat[header_idx, idx_s_score] = "攔舉配分"
                        col_s_score = idx_s_score
                        
                        if col_d_cnt >= idx_s_score: col_d_cnt += 1
                        if col_subtotal >= idx_s_score: col_subtotal += 1
                        
                        idx_d_score = col_d_cnt + 1
                        raw_df.insert(idx_d_score, f'new_{idx_d_score}', None)
                        raw_df.iat[header_idx, idx_d_score] = "逕舉配分"
                        col_d_score = idx_d_score
                        
                        if col_subtotal >= idx_d_score: col_subtotal += 1
                        
                        idx_subtotal = idx_d_score + 1
                        raw_df.insert(idx_subtotal, f'new_{idx_subtotal}', None)
                        raw_df.iat[header_idx, idx_subtotal] = "小計"
                        col_subtotal = idx_subtotal
                        raw_df.columns = range(raw_df.shape[1])

                    grand_total = 0
                    yellow_cells = []
                    
                    for r in range(header_idx + 1, len(raw_df)):
                        rule = str(raw_df.iloc[r, col_rule]).strip()
                        if not rule or "合計" in rule or "製表" in rule or "舉發單張數" in rule or rule.lower() == 'nan':
                            continue
                            
                        def safe_int(val):
                            try: return int(float(str(val).replace(",", "")))
                            except: return 0
                                
                        stop_cnt = safe_int(raw_df.iloc[r, col_s_cnt])
                        dir_cnt = safe_int(raw_df.iloc[r, col_d_cnt])
                        
                        if rule in db_map:
                            s_score = db_map[rule]['stop']
                            d_score = db_map[rule]['dir']
                        else:
                            s_score = 0
                            d_score = 0
                            
                        raw_df.iat[r, col_s_score] = s_score if s_score != 0 else 0
                        raw_df.iat[r, col_d_score] = d_score if d_score != 0 else 0
                        
                        row_subtotal = (s_score * stop_cnt) + (d_score * dir_cnt)
                        if col_subtotal != -1:
                            raw_df.iat[r, col_subtotal] = row_subtotal
                        
                        if s_score == 0 and d_score == 0:
                            yellow_cells.append((r + 1, col_s_score + 1))
                            yellow_cells.append((r + 1, col_d_score + 1))
                            
                        grand_total += row_subtotal
                        
                    processed_sheets.append({
                        "officer": officer_name.replace(" ", ""),
                        "df": raw_df,
                        "yellow_cells": yellow_cells,
                        "grand_total": grand_total,
                        "col_d_score": col_d_score,
                        "print_date": print_date_val  # 儲存提取出的列印日期
                    })
                    
            except Exception as e:
                st.error(f"❌ 處理單位檔案 {f.name} 時發生錯誤：{e}")
                continue

            # ==========================================
            # 3. 各單位獨立結算與報表產出
            # ==========================================
            if processed_sheets:
                df_raw_summary = pd.DataFrame([{
                    "員警姓名": s["officer"],
                    "本期總分": s["grand_total"]
                } for s in processed_sheets])
                
                df_summary = df_raw_summary.groupby('員警姓名', as_index=False)['本期總分'].sum()
                df_summary['上半年分數'] = 5800  
                
                def calc_rem(score):
                    if score >= threshold_7x: return score - threshold_7x
                    return score % quota
                    
                df_summary['上半年剩餘分數'] = df_summary['上半年分數'].apply(calc_rem)
                df_summary['最終總分'] = df_summary['本期總分'] + df_summary['上半年剩餘分數']

                all_summaries[unit_name] = df_summary

                # --- 產出 Excel 檔案 ---
                output = io.BytesIO()
                wb = Workbook()
                
                ws_summary = wb.active
                ws_summary.title = "績效結算總表"
                ws_summary['A1'] = "交通違規舉發績效結算表"
                ws_summary['A1'].font = Font(size=14, bold=True)
                ws_summary['A2'] = f"結算單位：{unit_type_label}"
                
                header = list(df_summary.columns)
                header_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
                for c_idx, title in enumerate(header, 1):
                    cell = ws_summary.cell(row=4, column=c_idx, value=title)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = header_fill
                    ws_summary.column_dimensions[get_column_letter(c_idx)].width = 16
                    
                for r_idx, row_data in enumerate(df_summary.values, 5):
                    for c_idx, val in enumerate(row_data, 1):
                        ws_summary.cell(row=r_idx, column=c_idx, value=val)

                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                sheet_name_counts = {}
                
                for s in processed_sheets:
                    base_name = s["officer"][:25]
                    if base_name not in sheet_name_counts:
                        sheet_name_counts[base_name] = 1
                        final_name = base_name
                    else:
                        sheet_name_counts[base_name] += 1
                        final_name = f"{base_name}({sheet_name_counts[base_name]})"
                        
                    ws_officer = wb.create_sheet(title=final_name)
                    
                    # 💡 建立置頂的全新第 1 列 (在原本資料的最上面)
                    num_cols = len(s["df"].columns)
                    title_row = ["員警開單績效統計表"] + [None] * (num_cols - 1)
                    if s["print_date"]:
                        # 放在倒數第2個欄位，讓排版平均
                        put_idx = max(1, num_cols - 2)
                        title_row[put_idx] = s["print_date"]
                        
                    ws_officer.append(title_row)
                    # 設定標題樣式
                    title_cell = ws_officer.cell(row=1, column=1)
                    title_cell.font = Font(size=14, bold=True)
                    if s["print_date"]:
                        date_cell = ws_officer.cell(row=1, column=put_idx + 1)
                        date_cell.font = Font(bold=True)
                        date_cell.alignment = Alignment(horizontal="right") # 靠右對齊更好看
                    
                    # 依序寫入剩下的資料列 (這些原本在第一列的開單日期，現在會自動變成第 2 列)
                    for r_idx, row in s["df"].iterrows():
                        cleaned_row = [val if pd.notna(val) else None for val in row.tolist()]
                        ws_officer.append(cleaned_row)
                        
                    # 標示無配分的黃色警示區塊 (因上方硬插了一列，故行數需 +1)
                    for r, c in s["yellow_cells"]:
                        ws_officer.cell(row=r + 1, column=c).fill = yellow_fill
                        
                    officer_summary = df_summary[df_summary['員警姓名'] == s['officer']].iloc[0]
                    
                    footer_start_row = ws_officer.max_row + 2
                    write_col = s["col_d_score"] + 1 if s["col_d_score"] != -1 else 5
                    
                    def write_footer(row_offset, title, val, color="000000"):
                        c_title = ws_officer.cell(row=footer_start_row + row_offset, column=write_col - 1, value=title)
                        c_title.font = Font(bold=True)
                        c_val = ws_officer.cell(row=footer_start_row + row_offset, column=write_col, value=val)
                        c_val.font = Font(bold=True, color=color)
                    
                    write_footer(0, "本期分數：", int(officer_summary["本期總分"]), "0000FF")
                    write_footer(1, "上半年分數：", int(officer_summary["上半年分數"]), "0000FF")
                    write_footer(2, "上半年剩餘分數：", int(officer_summary["上半年剩餘分數"]), "0000FF")
                    write_footer(3, "總分：", int(officer_summary["最終總分"]), "FF0000")
                    
                    for col_idx in range(1, len(s["df"].columns) + 1):
                        ws_officer.column_dimensions[get_column_letter(col_idx)].width = 14

                wb.save(output)
                all_output_buffers[f"績效結算_{unit_name}.xlsx"] = output

        # ==========================================
        # 4. 畫面展示與下載/寄件區塊
        # ==========================================
        if all_summaries:
            st.success("✅ 所有單位的原始報表已全數完成結算！")
            
            unit_names = list(all_summaries.keys())
            tabs = st.tabs([f"🏢 {name}" for name in unit_names])
            
            for i, name in enumerate(unit_names):
                with tabs[i]:
                    st.subheader(f"📊 {name} - 績效結算總表")
                    st.dataframe(all_summaries[name], use_container_width=True, hide_index=True)
                    
                    st.download_button(
                        label=f"📥 下載 {name} 完整報表",
                        data=all_output_buffers[f"績效結算_{name}.xlsx"].getvalue(),
                        file_name=f"績效結算_{name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"dl_{name}"
                    )

            st.divider()
            st.markdown("### 📧 批次寄送報表")
            st.info("點擊下方按鈕，系統會將上方產出的 **所有單位的 Excel 報表** 一次打包寄送至您的信箱！")
            
            if st.button("🚀 一鍵將所有報表寄至我的信箱", use_container_width=True):
                with st.spinner("信件發送中，請稍候…"):
                    ok, mail_err = send_multiple_files_email(all_output_buffers)
                    if ok:
                        st.success(f"✅ 信件發送成功！本次共寄出 {len(all_output_buffers)} 份報表。")
                        st.balloons()
                    else:
                        st.error(f"❌ 發信失敗: {mail_err}")
