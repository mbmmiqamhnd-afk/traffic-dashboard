import io
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header

import streamlit as st
import pandas as pd

# 引入共用側邊欄選單
try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        st.sidebar.title("🚓 交通執法系統")

st.set_page_config(
    page_title="毒駕專案取締敘獎統計",
    page_icon="🧪",
    layout="wide"
)

# 渲染側邊欄
show_sidebar()

# ==================== 頁面標題與公文資訊 ====================
st.title("🧪 加強取締施用毒品後駕車專案出力人員敘獎統計")
st.caption("依據內政部警政署加強攔查取締施用毒品後駕車專案工作計畫暨桃園市政府警察局 115年9月10日 桃警交字第1150120088號函辦理")

with st.expander("📌 專案公文重點提示與待辦時限", expanded=False):
    st.markdown("""
    * **主旨**：有關辦理本局執行內政部警政署加強攔查取締施用毒品後駕車專案工作計畫出力人員敘獎案，請依說明事項辦理，請查照。
    * **發文字號**：桃園市政府警察局 115 年 9 月 10 日 桃警交字第 1150120088 號函
    * **承辦窗口**：本局交通警察大隊
    * **辦理時限**：**民國 115 年 9 月 17 日（星期四）**
    * **車種敘獎標準規範**：
      1. **大型車**：每件核予「記功一次」。
      2. **小型車**：每件核給 2 點（嘉獎二次）。
      3. **機車（含微型電動二輪車）**：每件核給 1 點。
      4. **其餘慢車（電輔車、腳踏自行車）**：不納入專案計點。
      5. **人事獎懲拆解**：非大型車累積點數以商數（// 2）為「嘉獎二次」，餘數（% 2）為「嘉獎一次」。
    """)

# ==================== 違規事實與車種精準分類 ====================
def classify_violation(fact_text, law_code=""):
    text = f"{law_code} {fact_text}".strip()
    if not text or text == "nan":
        return "未填寫"

    is_drug = bool(re.search(r"毒|毒品|麻醉|迷幻|藥駕|第1項第2款|35102|35300175|35300178|35300184|35300199|35300211|35300253|35300256|35402002", text))
    is_alcohol = bool(re.search(r"酒|酒精|呼氣|吐氣|第1項第1款|第一項第一款|1項1款|35101|35700|35402001", text))
    is_refusal = bool(re.search(r"拒測|拒絕接受|拒絕|35402", text))
    is_recidivism = bool(re.search(r"累犯|二次以上|第2次|第3次|多次", text))
    is_slow_vehicle = bool(re.search(r"73|慢車|自行車|腳踏車|微型電動二輪車", text))
    is_impound = bool(re.search(r"第三十五條第一、三、四、五項之情形之一|第35條第1、3、4、5項之情形之一|移置|保管|35900", text))

    # 1. 拒測判別
    if is_refusal:
        return "🧪 毒駕拒測" if is_drug else ("🍺 酒駕拒測" if is_alcohol else "⚠️ 拒測(未指明類別)")

    # 2. 累犯判別
    if is_recidivism:
        return "🧪 毒駕累犯" if is_drug else ("🍺 酒駕累犯" if is_alcohol else "⚠️ 累犯")

    # 3. 慢車判別 (第 73 條)
    if is_slow_vehicle:
        return "🚲 慢車毒駕" if is_drug else ("🚲 慢車酒駕" if is_alcohol else "🚲 慢車其他違規")

    # 4. 車輛移置保管 (第 35 條第 9 項)
    if is_impound:
        return "🚗 車輛移置保管（第35條第9項）"

    # 5. 本體違規判別 (第 35 條)
    if is_drug:
        return "🧪 毒駕本體"
    elif is_alcohol:
        return "🍺 一般酒駕"

    return "📋 其他相關違規"

def classify_vehicle_type(v_name):
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

# ==================== 檔案上傳與自動直取【案件明細】 ====================
uploaded_file = st.file_uploader(
    "📂 請上傳「35條73條統計表」（自選匯出.xlsx）", 
    type=["xlsx", "xls"],
    help="系統將自動直取「案件明細」並依署頒計畫標準精算"
)

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_options = xls.sheet_names
    target_sheet = next((s for s in sheet_options if "案件明細" in s), sheet_options[0])
    
    df_raw_no_header = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=None)
    
    header_idx = None
    for idx in range(min(15, len(df_raw_no_header))):
        row_text = " ".join(df_raw_no_header.iloc[idx].dropna().astype(str).tolist())
        if any(k in row_text for k in ["單號", "違規事實", "違規法條", "舉發員警"]):
            header_idx = idx
            break

    if header_idx is not None:
        new_columns = [str(c).strip() for c in df_raw_no_header.iloc[header_idx].tolist()]
        df_clean = df_raw_no_header.iloc[header_idx + 1:].copy().reset_index(drop=True)
        df_clean.columns = new_columns
    else:
        df_clean = pd.read_excel(uploaded_file, sheet_name=target_sheet)

    # 過濾公文頁尾
    ticket_col = next((c for c in df_clean.columns if "單號" in str(c)), None)
    if ticket_col:
        df_clean = df_clean[df_clean[ticket_col].notna()]
        df_clean = df_clean[~df_clean[ticket_col].astype(str).str.contains("列印人員|統計期間|製表人員|總計", na=False)]

    # 定位關鍵欄位
    fact_col = next((c for c in df_clean.columns if any(k in str(c) for k in ["違規事實", "事實說明", "違規事項", "事實"])), None)
    law_col = next((c for c in df_clean.columns if any(k in str(c) for k in ["違規法條", "法條代碼", "法條"])), None)
    officer_col = next((c for c in df_clean.columns if any(k in str(c) for k in ["舉發員警", "員警", "警號", "姓名", "填單人"])), None)
    vehicle_col = next((c for c in df_clean.columns if any(k in str(c) for k in ["簡式車種名稱", "車種", "車種名稱"])), None)
    date_col = next((c for c in df_clean.columns if any(k in str(c) for k in ["入案日", "違規日", "日期"])), None)

    st.success(f"已自動鎖定【{target_sheet}】，共成功匯入 **{len(df_clean)}** 筆案件紀錄！")

    if fact_col and vehicle_col:
        # 標註分類與車種判定
        df_clean["案件分類"] = df_clean.apply(lambda r: classify_violation(r[fact_col], r[law_col] if law_col else ""), axis=1)
        df_clean["車種判定"] = df_clean[vehicle_col].apply(classify_vehicle_type)

        st.markdown("### 🎯 案件類別篩選")
        all_cats = sorted(df_clean["案件分類"].unique().tolist())
        quick_mode = st.radio(
            "快速切換模式：", 
            ["僅毒品專案（敘獎標準）", "全部案件（含酒駕與移置）", "自訂勾選"], 
            index=0, 
            horizontal=True
        )

        if quick_mode == "僅毒品專案（敘獎標準）":
            selected_cats = [c for c in all_cats if "🧪" in c or "🚲 慢車毒駕" in c]
        elif quick_mode == "全部案件（含酒駕與移置）":
            selected_cats = all_cats
        else:
            selected_cats = st.multiselect(
                "請勾選本次納入統計之案件分類：",
                options=all_cats,
                default=[c for c in all_cats if "🧪" in c or "🚲 慢車毒駕" in c]
            )

        df_filtered = df_clean[df_clean["案件分類"].isin(selected_cats)].copy()

        # 指標看板
        st.markdown("#### 📈 專案指標概覽")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("🧪 毒駕本體", len(df_clean[df_clean["案件分類"] == "🧪 毒駕本體"]))
        m2.metric("🧪 毒駕累犯", len(df_clean[df_clean["案件分類"] == "🧪 毒駕累犯"]))
        m3.metric("🧪 毒駕拒測", len(df_clean[df_clean["案件分類"] == "🧪 毒駕拒測"]))
        m4.metric("🚲 慢車毒駕", len(df_clean[df_clean["案件分類"] == "🚲 慢車毒駕"]))
        m5.metric("📌 本次統計納入", len(df_filtered))

        st.divider()

        # ==================== 署頒專案敘獎標準設定面板 ====================
        st.markdown("### ⚙️ 警政署專案工作計畫點數設定")
        with st.expander("🛠️ 點此確認或微調各車種折算點數（微電車已自動納入機車）", expanded=True):
            s1, s2, s3, s4 = st.columns(4)
            heavy_pts = s1.number_input("🚛 大型車毒駕每件記功次數：", min_value=1, max_value=2, value=1, step=1)
            car_pts = s2.number_input("🚗 小型車毒駕每件嘉獎點數：", min_value=1, max_value=3, value=2, step=1)
            moto_pts = s3.number_input("🛵 機車(含微電車)每件點數：", min_value=1, max_value=2, value=1, step=1)
            s4.info("🚲 其餘慢車（電輔車、自行車等）依規定不納入專案計點。")

        # 分頁展示
        tab_officer, tab_summary, tab_detail = st.tabs(["👮 出力員警敘獎建議名冊", "📊 車種與違規分佈統計", "📑 專案查獲案件明細表"])

        # TAB 1: 出力人員敘獎名冊
        with tab_officer:
            st.subheader("👮 查獲出力員警敘獎建議名冊（署頒計畫標準）")

            if officer_col:
                officer_list = df_filtered[officer_col].unique()
                summary_data = []

                for officer in officer_list:
                    sub = df_filtered[df_filtered[officer_col] == officer]
                    
                    heavy_cnt = len(sub[sub["車種判定"] == "大型車"])
                    car_cnt = len(sub[sub["車種判定"] == "小型車"])
                    moto_cnt = len(sub[sub["車種判定"] == "機車"])  # 包含微型電動二輪車
                    other_slow_cnt = len(sub[sub["車種判定"] == "其餘慢車(不納入)"])
                    refusal_cnt = len(sub[sub["案件分類"] == "🧪 毒駕拒測"])
                    reward_total_cnt = heavy_cnt + car_cnt + moto_cnt

                    # 1. 大型車專屬：記功一次
                    merit_total = heavy_cnt * heavy_pts

                    # 2. 其他車種：累計嘉獎點數 (小型車2點、機車含微電車1點)
                    total_pts = (car_cnt * car_pts) + (moto_cnt * moto_pts)

                    # 3. 依警察法規拆解為「嘉獎二次」與「嘉獎一次」
                    num_commend_2 = total_pts // 2
                    num_commend_1 = total_pts % 2

                    # 4. 組合建議獎勵額度文字
                    reward_parts = []
                    if merit_total > 0:
                        reward_parts.append(f"記功一次{merit_total}次" if merit_total > 1 else "記功一次")
                    if num_commend_2 > 0:
                        reward_parts.append(f"嘉獎二次{num_commend_2}次" if num_commend_2 > 1 else "嘉獎二次")
                    if num_commend_1 > 0:
                        reward_parts.append(f"嘉獎一次{num_commend_1}次" if num_commend_1 > 1 else "嘉獎一次")

                    final_reward_text = "、".join(reward_parts) if reward_parts else "列入參考（未達標準）"

                    reasons = []
                    if heavy_cnt > 0:
                        reasons.append(f"查獲大型車毒駕{heavy_cnt}件")
                    if car_cnt > 0:
                        reasons.append(f"查獲小型車毒駕{car_cnt}件")
                    if moto_cnt > 0:
                        reasons.append(f"查獲機車(含微電車)毒駕{moto_cnt}件")
                    if refusal_cnt > 0:
                        reasons.append(f"查獲毒駕拒測{refusal_cnt}件")

                    reason_str = "執行加強攔查取締毒駕專案工作計畫，" + "、".join(reasons) + "，工作出力。" if reasons else "執行毒駕專案工作出力。"

                    summary_data.append({
                        "員警姓名": officer,
                        "大型車(件)": heavy_cnt,
                        "小型車(件)": car_cnt,
                        "機車含微電車(件)": moto_cnt,
                        "其餘慢車(不計)": other_slow_cnt,
                        "毒駕拒測(件)": refusal_cnt,
                        "計獎總件數": reward_total_cnt,
                        "記功一次": merit_total,
                        "嘉獎二次": num_commend_2,
                        "嘉獎一次": num_commend_1,
                        "建議獎勵額度": final_reward_text,
                        "具體出力事由": reason_str
                    })

                officer_df = pd.DataFrame(summary_data).sort_values(
                    by=["記功一次", "嘉獎二次", "嘉獎一次", "計獎總件數"], 
                    ascending=[False, False, False, False]
                )
                
                st.dataframe(officer_df, use_container_width=True, hide_index=True)

                # 生成 Excel 二進位資料
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    officer_df.to_excel(writer, index=False, sheet_name="出力人員敘獎建議名冊")
                    df_filtered.to_excel(writer, index=False, sheet_name="專案查獲案件明細")
                excel_bytes = output.getvalue()
                excel_filename = "加強攔查取締施用毒品後駕車專案出力人員敘獎建議表.xlsx"

                # 檔案下載按鈕
                st.download_button(
                    label="📥 下載【加強取締施用毒品後駕車專案出力人員敘獎建議表】(Excel)",
                    data=excel_bytes,
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # ==================== 📧 一鍵寄送郵件給自己功能 ====================
                st.markdown("---")
                st.subheader("📧 將敘獎建議名冊發送至電子信箱")

                with st.expander("📬 點此發送郵件（附帶 Excel 附件與完整名冊 HTML）", expanded=True):
                    # 預設從 secrets 讀取，沒有則使用輸入框
                    default_sender = st.secrets.get("GMAIL_USER", "mbmmiqamhnd@gmail.com")
                    saved_password = st.secrets.get("GMAIL_PASSWORD", "")

                    c_mail1, c_mail2 = st.columns(2)
                    target_email = c_mail1.text_input("收件人信箱：", value="mbmmiqamhnd@gmail.com")
                    
                    if not saved_password:
                        app_pwd = c_mail2.text_input("Gmail 應用程式密碼 (16位英文字母)：", type="password", help="如已在 Streamlit Secrets 設定 GMAIL_PASSWORD 則會自動帶入")
                    else:
                        app_pwd = saved_password
                        c_mail2.success("✅ 已自動從 Streamlit Secrets 載入郵件密碼")

                    if st.button("🚀 立即發送郵件（含 Excel 附件）"):
                        if not app_pwd:
                            st.warning("⚠️ 請輸入 Gmail 應用程式密碼以進行發送。")
                        else:
                            try:
                                with st.spinner("正在發送郵件與上傳附件中..."):
                                    # 建立郵件訊息
                                    msg = MIMEMultipart()
                                    msg["From"] = default_sender
                                    msg["To"] = target_email
                                    mail_subject = "【專案報告】加強取締施用毒品後駕車專案出力人員敘獎建議名冊"
                                    msg["Subject"] = Header(mail_subject, "utf-8")

                                    # HTML 內容
                                    html_table = officer_df.to_html(index=False, classes="table table-bordered", border=1)
                                    html_content = f"""
                                    <div style="font-family: Arial, sans-serif; line-height: 1.6;">
                                        <h2 style="color: #1a73e8;">🚓 內政部警政署加強攔查取締施用毒品後駕車專案出力人員敘獎建議表</h2>
                                        <p><strong>承辦人：郭勝隆 巡官</strong></p>
                                        <p><strong>公文字號：</strong>桃園市政府警察局 115年9月10日 桃警交字第1150120088號函</p>
                                        <p><strong>辦理時限：</strong>民國 115 年 9 月 17 日（星期四）前函報交大</p>
                                        <hr>
                                        <h3>📋 出力人員敘獎建議名冊（共計 {len(officer_df)} 名同仁）</h3>
                                        {html_table}
                                        <hr>
                                        <p style="color: #555;">📎 完整案件明細與敘獎名冊 Excel 檔已隨信夾帶於附件，請查照。</p>
                                    </div>
                                    """
                                    msg.attach(MIMEText(html_content, "html", "utf-8"))

                                    # 附加 Excel 檔案
                                    part = MIMEApplication(excel_bytes)
                                    part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", excel_filename))
                                    msg.attach(part)

                                    # 發送郵件 (SMTP SSL/TLS)
                                    server = smtplib.SMTP("smtp.gmail.com", 587)
                                    server.starttls()
                                    server.login(default_sender, app_pwd)
                                    server.sendmail(default_sender, [target_email], msg.as_string())
                                    server.quit()

                                st.success(f"🎉 郵件已成功寄送至 {target_email}！內含 Excel 附件與完整敘獎名冊。")
                            except Exception as err:
                                st.error(f"❌ 郵件寄送失敗：{err}。請確認 Gmail 是否已啟用兩步驟驗證並產生「16位應用程式密碼」。")

            else:
                st.warning("此工作表未包含員警欄位。")

        # TAB 2: 分類分佈圖表
        with tab_summary:
            st.subheader("📊 車種與案件分類分佈統計")
            c1, c2 = st.columns(2)
            with c1:
                v_summary = df_filtered["車種判定"].value_counts().reset_index()
                v_summary.columns = ["車種大類", "案件數"]
                st.dataframe(v_summary, use_container_width=True, hide_index=True)
                st.bar_chart(v_summary.set_index("車種大類"))
            with c2:
                cat_summary = df_filtered["案件分類"].value_counts().reset_index()
                cat_summary.columns = ["違規分類", "案件數"]
                st.dataframe(cat_summary, use_container_width=True, hide_index=True)
                st.bar_chart(cat_summary.set_index("違規分類"))

        # TAB 3: 案件明細清單
        with tab_detail:
            st.subheader(f"📑 符合條件之案件清單（共 {len(df_filtered)} 筆）")
            display_cols = [c for c in [ticket_col, "車種判定", "案件分類", law_col, fact_col, officer_col, vehicle_col, date_col] if c and c in df_filtered.columns]
            other_cols = [c for c in df_filtered.columns if c not in display_cols]
            st.dataframe(df_filtered[display_cols + other_cols], use_container_width=True, hide_index=True)
    else:
        st.error("未能在此工作表中自動辨識到「違規事實」或「車種」欄位。")
        st.dataframe(df_clean.head(10), use_container_width=True)
else:
    st.info("💡 請上傳從 Gmail 下載之 `自選匯出.xlsx` 報表檔案。")
