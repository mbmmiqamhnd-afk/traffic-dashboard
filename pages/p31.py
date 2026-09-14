import io
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import urllib.parse as _ul

import streamlit as st
import pandas as pd
import numpy as np

# 載入自訂側邊欄
try:
    import menu
except ImportError:
    pass

# ==========================================
# 0. 輔助函式：發送單一檔案 Email
# ==========================================
def send_single_file_email(file_bytes, file_name, mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
    """使用 st.secrets 設定檔發送夾帶報表的電子郵件"""
    try:
        sender = st.secrets["email"]["user"]
        pwd = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = sender  # 寄給自己
        msg["Subject"] = f"🧪 毒駕專案敘獎統計 - {file_name}"
        
        body_text = (
            f"長官您好，\n\n"
            f"系統已自動完成「加強攔查取締施用毒品後駕車專案工作計畫」出力人員敘獎結算。\n"
            f"附件為最新產出之【{file_name}】。\n\n"
            f"依據桃園市政府警察局 115 年 9 月 10 日 桃警交字第 1150120088 號函，"
            f"本案辦理時限至民國 115 年 9 月 17 日（星期四）止，請查照。\n\n"
            f"本信件由交通執法自動化分析引擎發送。"
        )
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        # 解析 MIME 類型
        main_type, sub_type = mime_type.split('/') if '/' in mime_type else ("application", "octet-stream")
        part = MIMEBase(main_type, sub_type)
        part.set_payload(file_bytes.getvalue())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(file_name)}")
        msg.attach(part)

        # 透過 SMTP 發送 (SSL 465 埠)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# ==========================================
# 1. 資料處理與統計邏輯
# ==========================================
def classify_violation(fact_text, law_code=""):
    """
    結合法條代碼與違規事實文字，精準判定案件類別：
    - 涵蓋迷幻藥、麻醉藥品、藥駕、毒駕本體、毒駕累犯、毒駕拒測
    """
    text = f"{law_code} {fact_text}".strip()
    if not text or text == "nan":
        return "未填寫"

    is_drug = bool(re.search(r"毒|毒品|麻醉|迷幻|藥駕|第1項第2款|35102|35300175|35300178|35300184|35300199|35300211|35300253|35300256|35402002", text))
    is_alcohol = bool(re.search(r"酒|酒精|呼氣|吐氣|第1項第1款|第一項第一款|1項1款|35101|35700|35402001", text))
    is_refusal = bool(re.search(r"拒測|拒絕接受|拒絕|35402", text))
    is_recidivism = bool(re.search(r"累犯|二次以上|第2次|第3次|多次", text))
    is_slow_vehicle = bool(re.search(r"73|慢車|自行車|腳踏車|微型電動二輪車", text))
    is_impound = bool(re.search(r"第三十五條第一、三、四、五項之情形之一|第35條第1、3、4、5項之情形之一|移置|保管|35900", text))

    if is_refusal:
        return "🧪 毒駕拒測" if is_drug else ("🍺 酒駕拒測" if is_alcohol else "⚠️ 拒測(未指明類別)")
    if is_recidivism:
        return "🧪 毒駕累犯" if is_drug else ("🍺 酒駕累犯" if is_alcohol else "⚠️ 累犯")
    if is_slow_vehicle:
        return "🚲 慢車毒駕" if is_drug else ("🚲 慢車酒駕" if is_alcohol else "🚲 慢車其他違規")
    if is_impound:
        return "🚗 車輛移置保管（第35條第9項）"
    if is_drug:
        return "🧪 毒駕本體"
    elif is_alcohol:
        return "🍺 一般酒駕"

    return "📋 其他相關違規"

def classify_vehicle_type(v_name):
    """
    精準區分車種：
    - 大型車：大貨車、大客車、聯結車、曳引車 (唯有大型車可記功)
    - 小型車：汽車、自小客、小貨車
    - 機車：重機、輕機、微型電動二輪車 (微型電動二輪車比照機車標準)
    - 其餘慢車：電動輔助自行車、腳踏自行車 (不納入專案計點)
    """
    v = str(v_name).strip()
    if any(k in v for k in ["大貨", "大客", "聯結", "曳引", "大型車"]):
        return "大型車"
    elif any(k in v for k in ["汽車", "自小客", "自小貨", "小型車", "小貨", "小客"]):
        return "小型車"
    elif any(k in v for k in ["微型電動二輪車", "微型電動"]):
        return "機車"  # 慢車中微型電動二輪車含在機車的標準
    elif any(k in v for k in ["重機", "輕機", "機車"]):
        return "機車"
    elif any(k in v for k in ["電動輔助", "自行車", "腳踏", "慢車"]):
        return "其餘慢車(不納入)"
    return "其他車種"

def process_traffic_data(file):
    """讀取並清洗自選匯出 Excel 資料 (自動直取案件明細並對齊標題列)"""
    try:
        xls = pd.ExcelFile(file)
        sheet_options = xls.sheet_names
        target_sheet = next((s for s in sheet_options if "案件明細" in s), sheet_options[0])

        df_raw = pd.read_excel(file, sheet_name=target_sheet, header=None)

        header_row_index = None
        for idx in range(min(15, len(df_raw))):
            row_values = " ".join(df_raw.iloc[idx].dropna().astype(str).tolist())
            if '單號' in row_values and ('舉發員警1' in row_values or '舉發員警' in row_values):
                header_row_index = idx
                break

        if header_row_index is None:
            st.error("找不到資料標題列！請確認上傳檔案工作表【案件明細】中是否包含『單號』與『舉發員警1』。")
            return None

        new_cols = [str(val).strip() for val in df_raw.iloc[header_row_index]]
        df = df_raw.iloc[header_row_index + 1:].copy().reset_index(drop=True)
        df.columns = new_cols

        # 過濾底部公文頁尾列
        ticket_col = next((c for c in df.columns if "單號" in str(c)), None)
        if ticket_col:
            df = df[df[ticket_col].notna()]
            df = df[~df[ticket_col].astype(str).str.contains("列印人員|統計期間|製表人員|總計", na=False)]

        cols_to_keep = ['單號', '簡式車種名稱', '違規法條1', '違規事實1', '入案日', '舉發員警1']
        missing_cols = [col for col in cols_to_keep if col not in df.columns]
        if missing_cols:
            st.error(f"上傳的檔案缺少以下必要欄位：{', '.join(missing_cols)}，請確認自選匯出時是否有勾選。")
            return None

        df = df[cols_to_keep].dropna(how='all')
        return df

    except Exception as e:
        st.error(f"檔案解析失敗，錯誤訊息：{str(e)}")
        return None

def calculate_merits_for_officer(group):
    """計算單一員警的敘獎額度（拒測獨立統計為嘉獎一次，不計入車種計算）"""
    group = group.sort_values(by='入案日')

    heavy_cases = 0
    car_cases = 0
    moto_cases = 0
    other_slow_cases = 0
    refusal_cases = 0

    tickets = []

    for idx, row in group.iterrows():
        v_type = str(row['車種判定'])
        cat = str(row['案件分類'])
        tickets.append(str(row['單號']))

        # ✅ 拒測件數獨立統計，不計入車種計算中
        if '拒測' in cat:
            refusal_cases += 1
        else:
            if v_type == '大型車':
                heavy_cases += 1
            elif v_type == '小型車':
                car_cases += 1
            elif v_type == '機車':
                moto_cases += 1
            elif v_type == '其餘慢車(不納入)':
                other_slow_cases += 1

    # 1. 大型車專屬功次（唯有大型車記功）
    merit_cnt = heavy_cases * 1

    # 2. 其他車種與拒測累計點數（小型車2點、機車含微電車1點、毒駕拒測1點，其餘慢車0點）
    # 依最新規定：拒測核予嘉獎一次（1點）
    total_pts = (car_cases * 2) + (moto_cases * 1) + (refusal_cases * 1)
    
    # 3. 法定獎懲額度拆解（商數為嘉獎二次，餘數為嘉獎一次）
    num_commend_2 = total_pts // 2
    num_commend_1 = total_pts % 2

    # 4. 組合建議獎勵額度文字
    reward_parts = []
    if merit_cnt > 0:
        reward_parts.append(f"記功一次{merit_cnt}次" if merit_cnt > 1 else "記功一次")
    if num_commend_2 > 0:
        reward_parts.append(f"嘉獎二次{num_commend_2}次" if num_commend_2 > 1 else "嘉獎二次")
    if num_commend_1 > 0:
        reward_parts.append(f"嘉獎一次{num_commend_1}次" if num_commend_1 > 1 else "嘉獎一次")

    final_reward_text = "、".join(reward_parts) if reward_parts else "列入參考（未達標準）"

    # 具體出力事由：內容只到件數，後方標點符號與工作出力文字均已移除
    reasons = []
    if heavy_cases > 0:
        reasons.append(f"大型車毒駕{heavy_cases}件")
    if car_cases > 0:
        reasons.append(f"小型車毒駕{car_cases}件")
    if moto_cases > 0:
        reasons.append(f"機車(含微電車)毒駕{moto_cases}件")
    if refusal_cases > 0:
        reasons.append(f"毒駕拒測{refusal_cases}件")

    reason_str = "執行加強攔查取締毒駕專案工作計畫，查獲" + "、".join(reasons) if reasons else ""

    return pd.Series({
        '大型車(件)': heavy_cases,
        '小型車(件)': car_cases,
        '機車含微電車(件)': moto_cases,
        '其餘慢車(不納入)': other_slow_cases,
        '毒駕拒測(件)': refusal_cases,
        '計獎總件數': heavy_cases + car_cases + moto_cases + refusal_cases,
        '記功一次': merit_cnt,
        '嘉獎二次': num_commend_2,
        '嘉獎一次': num_commend_1,
        '建議獎勵額度': final_reward_text,
        '具體出力事由': reason_str,
        '舉發單號明細': ", ".join(tickets)
    })

# ==========================================
# 2. 主程式介面
# ==========================================
def main():
    st.set_page_config(page_title="毒駕專案 敘獎統計", page_icon="🧪", layout="wide")

    try:
        menu.show_sidebar()
    except Exception as e:
        st.sidebar.error("無法載入側邊欄，請確認根目錄下有 menu.py")

    st.title("🧪 加強取締施用毒品後駕車專案 - 自動敘獎統計系統")
    st.caption("依據內政部警政署加強攔查取締施用毒品後駕車專案工作計畫暨桃園市政府警察局 115年9月10日 桃警交字第1150120088號函辦理（時限：115年9月17日）")
    st.divider()

    uploaded_file = st.file_uploader("請上傳『自選匯出.xlsx』(資料來源需包含簡式車種名稱與違規事實)", type=["xlsx"])

    if uploaded_file is not None:
        with st.spinner("資料處理與專案案件過濾中，請稍候..."):
            df = process_traffic_data(uploaded_file)

        if df is not None:
            # 標註分類與車種判定
            df["案件分類"] = df.apply(lambda r: classify_violation(r["違規事實1"], r["違規法條1"]), axis=1)
            df["車種判定"] = df["簡式車種名稱"].apply(classify_vehicle_type)

            with st.expander("📄 檢視原始案件明細", expanded=False):
                st.dataframe(df, use_container_width=True)

            # 篩選毒駕專案案件
            df_drug = df[df["案件分類"].str.contains("🧪")].copy()

            # 指標看板
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("🧪 毒駕本體", len(df[df["案件分類"] == "🧪 毒駕本體"]))
            m2.metric("🧪 毒駕累犯", len(df[df["案件分類"] == "🧪 毒駕累犯"]))
            m3.metric("🧪 毒駕拒測", len(df[df["案件分類"] == "🧪 毒駕拒測"]))
            m4.metric("📌 專案計獎案件總數", len(df_drug))

            st.subheader("📊 員警專案敘獎統計表 (依警政署工作計畫標準)")

            # 按員警分組結算
            merit_stats = df_drug.groupby('舉發員警1').apply(calculate_merits_for_officer).reset_index()
            merit_stats = merit_stats.sort_values(
                by=['記功一次', '嘉獎二次', '嘉獎一次', '計獎總件數'], 
                ascending=[False, False, False, False]
            ).reset_index(drop=True)

            styled_df = (merit_stats.style
                         .background_gradient(subset=['嘉獎二次'], cmap='Reds')
                         .background_gradient(subset=['嘉獎一次'], cmap='Blues'))

            st.dataframe(styled_df, use_container_width=True)

            # 建立記憶體中的 Excel 檔案 (使用 openpyxl 引擎)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                merit_stats.to_excel(writer, index=False, sheet_name='出力人員敘獎名冊')
                df_sorted = df_drug.sort_values(by=['舉發員警1', '入案日']).reset_index(drop=True)
                df_sorted.to_excel(writer, index=False, sheet_name='專案案件明細')

            excel_data = output.getvalue()
            excel_filename = '加強取締施用毒品後駕車專案敘獎統計含明細.xlsx'

            st.divider()
            col_dl, col_mail = st.columns(2)

            with col_dl:
                st.download_button(
                    label="📥 下載敘獎名冊及明細 (Excel)",
                    data=excel_data,
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

            with col_mail:
                if st.button("📧 將此統計表一鍵寄至我的信箱", use_container_width=True):
                    with st.spinner("信件發送中，請稍候…"):
                        output.seek(0)
                        ok, mail_err = send_single_file_email(output, excel_filename)
                        if ok:
                            st.success("✅ 信件發送成功！統計報表 Excel 已夾帶至您的信箱。")
                        else:
                            st.error(f"❌ 發信失敗，請檢查系統信箱設定。錯誤訊息: {mail_err}")

            st.markdown("<br>", unsafe_allow_html=True)
            st.info("💡 **系統計算標準：**\n"
                    "1. **大型車專屬記功**：僅查獲大型車（大貨車、大客車、聯結車等）毒駕核予「記功一次」。\n"
                    "2. **小型車**：每件核給 2 點（嘉獎二次）。\n"
                    "3. **機車（含微電車）**：每件核給 1 點（慢車中之微型電動二輪車含在機車標準）。\n"
                    "4. **毒駕拒測**：每件核給 1 點（嘉獎一次），件數獨立計算不計入車種欄位。\n"
                    "5. **其餘慢車**：電動輔助自行車、腳踏自行車等依規定不納入專案計點。\n"
                    "6. **獎勵名目拆分**：非大型車點數嚴格拆解為「嘉獎二次」（點數 // 2）與「嘉獎一次」（點數 % 2），無跨級折算記功或嘉獎三次/六次情形。\n"
                    "7. **出力事由簡約化**：具體事由嚴格截止於查獲件數，無後綴標點符號與額外贅字。")

if __name__ == "__main__":
    main()
