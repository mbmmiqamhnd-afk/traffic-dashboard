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
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 💡 匯入系統原本的側邊欄設定
# ==========================================
try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        pass

# ==========================================
# 💡 系統後台設定
# ==========================================
TARGET_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1HaFu5PZkFDUg7WZGV9khyQ0itdGXhXUakP4_BClFTUg/edit"

def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
    elif "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        creds_dict = dict(st.secrets["connections"]["gsheets"])
    else:
        return None
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

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
        msg["Subject"] = "🚔 交通執法重點工作舉發績效結算表 (批次產出)"
        
        body_text = (
            "長官您好，\n\n"
            "系統已自動產出最新結算的【交通執法重點工作舉發績效結算表與個人明細】。\n\n"
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
# 主程式執行區塊
# ==========================================
def main():
    # 1. 頁面基本設定 (必須是第一個 Streamlit 指令)
    st.set_page_config(page_title="舉發績效結算", page_icon="👮", layout="wide")

    # 2. 呼叫側邊欄導航選單
    show_sidebar()

    # 3. 繪製主畫面 UI
    st.title("⚡ 員警交通執法重點工作舉發績效結算")
    st.markdown("本頁面專門處理**員警個人績效配分**與**門檻結算**。系統具備雙向雲端同步與**智慧修復引擎**，能自動補齊殘缺的資料。")

    # ==========================================
    # 💡 背景讀取配分資料與「自動修復引擎」
    # ==========================================
    db_map = {}
    df_db = None
    sheet_id = TARGET_GSHEET_URL.split("/d/")[1].split("/")[0]
    xlsx_export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

    # 抓取雲端資料
    try:
        df_db = pd.read_excel(xlsx_export_url, sheet_name="配分表")
        connection_status = True
    except Exception as e:
        connection_status = False
        error_msg = e

    # 解析與檢查缺漏
    incomplete_rules = [] 
    if connection_status:
        if '違規條款' not in df_db.columns:
            st.error("❌ 雲端配分表缺少『違規條款』欄位！")
            st.stop()
        if '違規事實' not in df_db.columns: df_db['違規事實'] = ""
        if '類別' not in df_db.columns: df_db['類別'] = ""
        if '取締項目' not in df_db.columns: df_db['取締項目'] = ""

        for row_idx, row in df_db.iterrows():
            rule = str(row.get('違規條款', '')).strip()
            if rule:
                s = pd.to_numeric(row.get('攔舉配分', 0), errors='coerce')
                d = pd.to_numeric(row.get('逕舉配分', 0), errors='coerce')
                fact = str(row.get('違規事實', '')).strip()
                cat = str(row.get('類別', '')).strip()
                item = str(row.get('取締項目', '')).strip()
                
                if fact.lower() == 'nan': fact = ""
                if cat.lower() == 'nan': cat = ""
                if item.lower() == 'nan': item = ""
                    
                db_map[rule] = {
                    'stop': 0 if pd.isna(s) else int(s),
                    'dir': 0 if pd.isna(d) else int(d),
                    'fact': fact,
                    'category': cat,
                    'item': item,
                    'gsheet_row': row_idx + 2 
                }
                
                if not cat and not item:
                    incomplete_rules.append(rule)

    def find_closest_reference(new_rule_str, existing_db):
        best_match = None
        max_prefix_len = 0
        for old_rule, data in existing_db.items():
            match_len = 0
            for c1, c2 in zip(new_rule_str, old_rule):
                if c1 == c2: match_len += 1
                else: break
            if match_len > max_prefix_len and (data['category'] or data['item']):
                max_prefix_len = match_len
                best_match = data
        return best_match

    # ==========================================
    # 💡 將狀態顯示改為「折疊面板」，保持主畫面乾淨
    # ==========================================
    with st.expander("⚙️ 系統運行狀態與結算基準 (點擊展開)", expanded=bool(incomplete_rules)):
        if connection_status:
            st.success("🟢 雲端配分資料庫連線正常")
        else:
            st.error(f"🔴 雲端資料庫連線失敗：{error_msg}")
            st.stop()
            
        st.info("💡 **自動結算基準**\n- **交通分隊**：800 分\n- **其他單位**：400 分\n*(系統將自動判讀報表單位)*")

        # 隱藏在這裡的修復按鈕
        if incomplete_rules:
            st.warning(f"⚠️ 偵測到雲端資料庫中有 **{len(incomplete_rules)}** 筆過往新增的條款缺乏「類別」與「取締項目」。")
            if st.button("🛠️ 一鍵智慧修復資料庫", use_container_width=True):
                with st.spinner("☁️ 正在透過智慧參照引擎修復雲端資料..."):
                    client = get_gspread_client()
                    if not client:
                        st.error("❌ 找不到 GCP 憑證，無法寫入雲端。")
                    else:
                        try:
                            worksheet = client.open_by_key(sheet_id).worksheet("配分表")
                            updates = []
                            for rule in incomplete_rules:
                                ref_data = find_closest_reference(rule, db_map)
                                if ref_data:
                                    target_row = db_map[rule]['gsheet_row']
                                    new_cat = ref_data['category']
                                    new_item = ref_data['item']
                                    
                                    updates.append({'range': f'E{target_row}', 'values': [[new_cat]]})
                                    updates.append({'range': f'F{target_row}', 'values': [[new_item]]})
                                    
                                    db_map[rule]['category'] = new_cat
                                    db_map[rule]['item'] = new_item
                                    
                            if updates:
                                worksheet.batch_update(updates)
                                st.success("✅ 修復完成！已成功將參照類別補上。")
                                st.rerun() 
                            else:
                                st.info("無可參照的相近條款。")
                        except Exception as e:
                            st.error(f"❌ 修復失敗：{e}")

    st.markdown("---")
    
    # 💡 檔案上傳區
    st.subheader("📂 步驟 1：上傳原始舉發資料")
    data_files = st.file_uploader("批次選擇多個單位的半年期 Excel 檔案", type=["xlsx", "xls"], accept_multiple_files=True, key="data_files")
    st.markdown("---")

    # ==========================================
    # 2. 核心結算邏輯
    # ==========================================
    if data_files:
        # --- 步驟 B：預掃描未知條款 ---
        missing_rules = {} 
        for f in data_files:
            f.seek(0)
            try:
                xls = pd.ExcelFile(f)
                for sheet_name in xls.sheet_names:
                    raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None).astype('object')
                    header_idx = -1
                    for idx, row in raw_df.iterrows():
                        if "違規條款" in row.astype(str).str.replace(" ", "").values:
                            header_idx = idx
                            break
                    if header_idx == -1: continue

                    header_row = [str(x).strip().replace(" ", "") for x in raw_df.iloc[header_idx]]
                    if "違規條款" not in header_row: continue
                    
                    col_rule = header_row.index("違規條款")
                    col_fact = next((i for i, c in enumerate(header_row) if "違規事實" in c), -1)
                    col_cat = next((i for i, c in enumerate(header_row) if "類別" in c), -1)
                    col_item = next((i for i, c in enumerate(header_row) if "取締項目" in c), -1)

                    for r in range(header_idx + 1, len(raw_df)):
                        rule = str(raw_df.iloc[r, col_rule]).strip()
                        if not rule or "合計" in rule or "製表" in rule or "舉發單張數" in rule or rule.lower() == 'nan':
                            continue
                            
                        if rule not in db_map:
                            if rule not in missing_rules:
                                missing_rules[rule] = {"fact": "", "category": "", "item": ""}
                                
                            if not missing_rules[rule]["fact"] and col_fact != -1:
                                val = str(raw_df.iloc[r, col_fact]).strip()
                                if val.lower() != 'nan': missing_rules[rule]["fact"] = val
                                
                            if not missing_rules[rule]["category"] and col_cat != -1:
                                val = str(raw_df.iloc[r, col_cat]).strip()
                                if val.lower() != 'nan': missing_rules[rule]["category"] = val
                                
                            if not missing_rules[rule]["item"] and col_item != -1:
                                val = str(raw_df.iloc[r, col_item]).strip()
                                if val.lower() != 'nan': missing_rules[rule]["item"] = val
            except Exception:
                pass

        for rule_key, data in missing_rules.items():
            if not data["category"] and not data["item"]:
                ref_data = find_closest_reference(rule_key, db_map)
                if ref_data:
                    data["category"] = ref_data["category"]
                    data["item"] = ref_data["item"]

        if missing_rules:
            st.warning(f"⚠️ 掃描完畢！發現 **{len(missing_rules)}** 筆新條款。系統已根據雲端資料庫智慧帶入相近類別，請確認配分：")
            
            missing_df = pd.DataFrame({
                "違規條款": list(missing_rules.keys()),
                "違規事實": [v["fact"] for v in missing_rules.values()],
                "類別": [v["category"] for v in missing_rules.values()],
                "取締項目": [v["item"] for v in missing_rules.values()],
                "攔舉配分": 0,
                "逕舉配分": 0
            })
            
            edited_missing = st.data_editor(
                missing_df,
                column_config={
                    "違規條款": st.column_config.TextColumn("違規條款 (未設定)", disabled=True),
                    "違規事實": st.column_config.TextColumn("違規事實", disabled=False),
                    "類別": st.column_config.TextColumn("類別 ✍️", disabled=False),
                    "取締項目": st.column_config.TextColumn("取締項目 ✍️", disabled=False),
                    "攔舉配分": st.column_config.NumberColumn("攔舉配分 ✍️", min_value=0, required=True),
                    "逕舉配分": st.column_config.NumberColumn("逕舉配分 ✍️", min_value=0, required=True),
                },
                hide_index=True,
                use_container_width=True,
                key="missing_rules_editor"
            )
            
            if not st.button("✅ 我已確認完畢，同步寫回雲端並開始結算", type="primary", use_container_width=True):
                st.stop() 
            else:
                with st.spinner("☁️ 正在將新條款同步寫回 Google 試算表..."):
                    client = get_gspread_client()
                    if not client:
                        st.error("❌ 找不到對應的 GCP 憑證設定。")
                        st.stop()
                    try:
                        worksheet = client.open_by_key(sheet_id).worksheet("配分表")
                        for _, row in edited_missing.iterrows():
                            rule_val = str(row["違規條款"])
                            s_val = int(row["攔舉配分"])
                            d_val = int(row["逕舉配分"])
                            fact_val = str(row["違規事實"]) if pd.notna(row["違規事實"]) else ""
                            cat_val = str(row["類別"]) if pd.notna(row["類別"]) else ""
                            item_val = str(row["取締項目"]) if pd.notna(row["取締項目"]) else ""
                            
                            worksheet.append_row([rule_val, s_val, d_val, fact_val, cat_val, item_val])
                            db_map[rule_val] = {'stop': s_val, 'dir': d_val}
                            
                        st.success("✅ 新條款與對應資訊已成功寫回 Google 試算表最下方！")
                    except Exception as write_err:
                        st.error(f"❌ 寫入雲端失敗，詳細錯誤：{write_err}")
                        st.stop()

        # --- 步驟 C：正式結算 ---
        with st.spinner("🔄 正在讀取並結算各單位資料，請稍候..."):
            for f in data_files:
                f.seek(0)
                
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

            def extract_unit_name(df_head):
                for r_idx, row in df_head.iterrows():
                    for c_idx, val in enumerate(row.values):
                        val_str = str(val).strip()
                        if "舉發單位" in val_str:
                            clean = re.sub(r'舉發單位[:：]?', '', val_str).strip()
                            if clean: return clean
                            if c_idx + 1 < len(row.values):
                                next_val = str(row.values[c_idx + 1]).strip()
                                if next_val and next_val.lower() != 'nan':
                                    return next_val
                return ""

            unit_collected_data = {}

            for f in data_files:
                try:
                    xls = pd.ExcelFile(f)
                    for sheet_name in xls.sheet_names:
                        raw_df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                        raw_df = raw_df.astype('object')
                        
                        header_idx = -1
                        officer_name = extract_officer_name(raw_df.head(20))
                        if not officer_name:
                            officer_name = sheet_name.strip()
                            
                        detected_unit = extract_unit_name(raw_df.head(20))
                        if not detected_unit:
                            detected_unit = re.sub(r'\.[a-zA-Z0-9]+$', '', f.name) 
                        
                        for idx, row in raw_df.iterrows():
                            if "違規條款" in row.astype(str).str.replace(" ", "").values:
                                header_idx = idx
                                break
                        
                        if header_idx == -1:
                            continue
                        
                        print_date_val = ""
                        issue_date_val = ""
                        unit_val = ""
                        officer_val = ""
                        
                        for r_search in range(header_idx):
                            for c_search in range(len(raw_df.columns)):
                                val = raw_df.iat[r_search, c_search]
                                if pd.notna(val) and str(val).strip() != "":
                                    val_str = str(val).strip()
                                    val_no_space = val_str.replace(" ", "")
                                    if "列印日期" in val_no_space:
                                        print_date_val = val_str
                                    elif "開單日期" in val_no_space:
                                        issue_date_val = val_str
                                    elif "舉發單位" in val_no_space:
                                        unit_val = val_str
                                    elif "舉發員警" in val_no_space:
                                        officer_val = val_str

                        raw_df = raw_df.iloc[header_idx:].reset_index(drop=True)
                        header_idx = 0 
                        
                        header_row_temp = [str(x).strip().replace(" ", "") for x in raw_df.iloc[header_idx]]
                        cols_to_keep = [c for c, val in enumerate(header_row_temp) if val not in ["nan", "None", ""]]
                        
                        raw_df = raw_df.iloc[:, cols_to_keep]
                        raw_df.columns = range(raw_df.shape[1])

                        header_row = [str(x).strip().replace(" ", "") for x in raw_df.iloc[header_idx]]
                        col_rule = header_row.index("違規條款") if "違規條款" in header_row else -1
                        col_s_cnt = header_row.index("攔停數") if "攔停數" in header_row else -1
                        col_d_cnt = header_row.index("逕舉數") if "逕舉數" in header_row else -1
                        
                        col_s_score = header_row.index("攔舉配分") if "攔舉配分" in header_row else -1
                        col_d_score = header_row.index("逕舉配分") if "逕舉配分" in header_row else -1
                        col_subtotal = header_row.index("小計") if "小計" in header_row else -1
                        
                        if col_rule == -1 or col_s_cnt == -1 or col_d_cnt == -1:
                            continue

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
                                yellow_cells.append((r, col_s_score))
                                yellow_cells.append((r, col_d_score))
                                
                            grand_total += row_subtotal
                            
                        sheet_data = {
                            "officer": officer_name.replace(" ", ""),
                            "df": raw_df,
                            "yellow_cells": yellow_cells,
                            "grand_total": grand_total,
                            "col_d_score": col_d_score,
                            "print_date": print_date_val,
                            "issue_date": issue_date_val,
                            "unit_val": unit_val,
                            "officer_val": officer_val
                        }
                        
                        if detected_unit not in unit_collected_data:
                            is_800 = "交通分隊" in detected_unit
                            quota_val = 800 if is_800 else 400
                            unit_collected_data[detected_unit] = {
                                "processed_sheets": [],
                                "quota": quota_val,
                                "threshold_7x": quota_val * 7,
                                "unit_type_label": f"{detected_unit} (基準 {quota_val} 分)"
                            }
                        unit_collected_data[detected_unit]["processed_sheets"].append(sheet_data)
                        
                except Exception as e:
                    st.error(f"❌ 處理檔案 {f.name} 時發生錯誤：{e}")
                    continue

            # ==========================================
            # 3. 各單位獨立結算與報表產出
            # ==========================================
            all_summaries = {}
            all_output_buffers = {}

            for unit_name, unit_info in unit_collected_data.items():
                processed_sheets = unit_info["processed_sheets"]
                quota = unit_info["quota"]
                threshold_7x = unit_info["threshold_7x"]
                unit_type_label = unit_info["unit_type_label"]

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

                    output = io.BytesIO()
                    wb = Workbook()
                    
                    ws_summary = wb.active
                    ws_summary.title = "績效結算總表"
                    ws_summary['A1'] = "交通執法重點工作舉發績效結算表"
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
                        num_cols = len(s["df"].columns)
                        
                        title_row = ["員警開單績效統計表"] + [None] * (num_cols - 1)
                        if s["print_date"]:
                            put_idx = min(3, num_cols - 1) 
                            title_row[put_idx] = s["print_date"]
                        ws_officer.append(title_row)
                        ws_officer['A1'].font = Font(size=14, bold=True)
                        if s["print_date"]:
                            date_cell = ws_officer.cell(row=1, column=put_idx + 1)
                            date_cell.font = Font(bold=True, color="0000FF") 
                            date_cell.alignment = Alignment(horizontal="left")
                            
                        if s["issue_date"]:
                            ws_officer.append([s["issue_date"]] + [None] * (num_cols - 1))
                            
                        if s["unit_val"]:
                            ws_officer.append([s["unit_val"]] + [None] * (num_cols - 1))
                        else:
                            ws_officer.append([f"舉發單位：{unit_name}"] + [None] * (num_cols - 1))
                            
                        if s["officer_val"]:
                            ws_officer.append([s["officer_val"]] + [None] * (num_cols - 1))
                        else:
                            ws_officer.append([f"舉發員警：{s['officer']}"] + [None] * (num_cols - 1))

                        current_row_offset = ws_officer.max_row 
                        
                        for r_idx, row in s["df"].iterrows():
                            cleaned_row = [val if pd.notna(val) and str(val).strip() != "" else None for val in row.tolist()]
                            ws_officer.append(cleaned_row)
                            
                        for r, c in s["yellow_cells"]:
                            excel_r = current_row_offset + r + 1
                            excel_c = c + 1
                            ws_officer.cell(row=excel_r, column=excel_c).fill = yellow_fill
                            
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
                    all_output_buffers[f"交通執法重點工作舉發績效結算表_{unit_name}.xlsx"] = output

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
                        
                        dl_filename = f"交通執法重點工作舉發績效結算表_{name}.xlsx"
                        st.download_button(
                            label=f"📥 下載 {name} 完整報表",
                            data=all_output_buffers[dl_filename].getvalue(),
                            file_name=dl_filename,
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

# 執行點
if __name__ == "__main__":
    main()
