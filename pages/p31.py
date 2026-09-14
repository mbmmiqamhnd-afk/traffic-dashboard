import io
import re
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
st.caption("依據桃園市政府警察局 115年9月10日 桃警交字第1150120088號函辦理")

with st.expander("📌 專案公文核心要旨與待辦時限", expanded=False):
    st.markdown("""
    * **主旨**：有關辦理本局執行內政部警政署加強攔查取締施用毒品後駕車專案工作計畫出力人員敘獎案，請依說明事項辦理，請查照。
    * **發文字號**：桃園市政府警察局 115 年 9 月 10 日 桃警交字第 1150120088 號函
    * **承辦窗口**：本局交通警察大隊
    * **辦理時限**：**民國 115 年 9 月 17 日（星期四）**
    * **具體待辦**：
      1. 彙整本分局執行毒駕專案查獲及線上攔檢出力員警名冊。
      2. 造具敘獎建議表陳核後，函報市警局交大辦理專案敘獎。
    """)

# ==================== 違規事實智慧分類邏輯 ====================
def classify_violation(text_input):
    """
    精準區分第 35 條與第 73 條之：
    - 毒駕本體 vs 酒駕本體
    - 毒駕拒測 vs 酒駕拒測
    - 毒駕累犯 vs 酒駕累犯
    - 慢車毒駕 vs 慢車酒駕
    """
    text = str(text_input).strip()
    if not text or text == "nan":
        return "未填寫"

    is_drug = bool(re.search(r"毒|毒品|麻醉|迷幻|第1項第2款|第一項第二款|1項2款|35102", text))
    is_alcohol = bool(re.search(r"酒|酒精|呼氣|吐氣|第1項第1款|第一項第一款|1項1款|35101", text))
    is_refusal = bool(re.search(r"拒測|拒絕接受|拒絕", text))
    is_recidivism = bool(re.search(r"累犯|二次以上|2次|3次|多次", text))
    is_slow_vehicle = bool(re.search(r"73|慢車|自行車|腳踏車|微型電動二輪車", text))

    # 1. 拒測判別
    if is_refusal:
        if is_drug:
            return "🧪 毒駕拒測"
        elif is_alcohol:
            return "🍺 酒駕拒測"
        elif "73" in text or "慢車" in text:
            return "🚲 慢車拒測"
        else:
            return "⚠️ 拒測(未指明類別)"

    # 2. 累犯判別
    if is_recidivism:
        if is_drug:
            return "🧪 毒駕累犯"
        elif is_alcohol:
            return "🍺 酒駕累犯"

    # 3. 慢車判別 (第 73 條)
    if is_slow_vehicle:
        if is_drug:
            return "🚲 慢車毒駕"
        elif is_alcohol:
            return "🚲 慢車酒駕"
        else:
            return "🚲 慢車其他違規"

    # 4. 本體違規判別 (第 35 條)
    if is_drug:
        return "🧪 毒駕本體"
    elif is_alcohol:
        return "🍺 一般酒駕"

    return "📋 其他相關違規"

# ==================== 檔案上傳與自動偵測 ====================
uploaded_file = st.file_uploader(
    "📂 請上傳「35條73條統計表」（自選匯出.xlsx）", 
    type=["xlsx", "xls"],
    help="支援交通舉發系統匯出之報表，可自動辨識工作表與案件明細"
)

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_options = xls.sheet_names
    
    # 預設選取可能包含「明細」或「法條」的工作表
    default_idx = 0
    for idx, s in enumerate(sheet_options):
        if any(k in s for k in ["明細", "案件", "法條", "統計"]):
            default_idx = idx
            break
            
    selected_sheet = st.selectbox("📑 選擇工作表：", sheet_options, index=default_idx)
    
    # 讀取選取的工作表
    df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
    
    # 自動搜尋正確的標題列（排除表頭列印時間等公文標頭）
    header_row_idx = 0
    for r in range(min(15, len(df_raw))):
        row_str = " ".join(df_raw.iloc[r].dropna().astype(str).tolist())
        if any(k in row_str for k in ["法條代碼", "違規事實", "單位", "單號", "員警"]):
            header_row_idx = r + 1
            break
            
    if header_row_idx > 0:
        df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, skiprows=header_row_idx)

    st.success(f"已讀取工作表【{selected_sheet}】，共 {len(df_raw)} 筆資料。")

    # 尋找關鍵欄位
    cols = df_raw.columns.tolist()
    fact_col = next((c for c in cols if any(k in str(c) for k in ["違規事實", "法條代碼", "事實", "違規事項", "法條"])), None)
    unit_col = next((c for c in cols if any(k in str(c) for k in ["單位", "派出所", "分隊", "所隊"])), None)
    officer_col = next((c for c in cols if any(k in str(c) for k in ["員警", "舉發人", "警號", "姓名", "填單人"])), None)

    if fact_col:
        # 標註分類
        df_raw["違規類別判定"] = df_raw[fact_col].apply(classify_violation)
        
        # 呈現分類篩選控制台
        st.markdown("### 🎯 案件類別篩選器（毒駕專案 / 酒駕區隔）")
        all_cats = sorted(df_raw["違規類別判定"].unique().tolist())
        
        # 預設勾選毒品專案項目
        default_selected = [c for c in all_cats if "🧪" in c or "🚲 慢車毒駕" in c]
        if not default_selected:
            default_selected = all_cats

        col_filter1, col_filter2 = st.columns([3, 1])
        with col_filter1:
            selected_cats = st.multiselect(
                "勾選要計入統計的案件類別（支援區分毒駕拒測/累犯）：",
                options=all_cats,
                default=default_selected
            )
        with col_filter2:
            quick_mode = st.radio(
                "快速模式：",
                ["自訂勾選", "全選毒品類", "全選全部"],
                index=0
            )
            if quick_mode == "全選毒品類":
                selected_cats = [c for c in all_cats if "🧪" in c or "🚲 慢車毒駕" in c]
            elif quick_mode == "全選全部":
                selected_cats = all_cats

        df_filtered = df_raw[df_raw["違規類別判定"].isin(selected_cats)].copy()

        # 顯示指標卡
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("毒駕本體", len(df_raw[df_raw["違規類別判定"] == "🧪 毒駕本體"]))
        m2.metric("毒駕累犯", len(df_raw[df_raw["違規類別判定"] == "🧪 毒駕累犯"]))
        m3.metric("毒駕拒測", len(df_raw[df_raw["違規類別判定"] == "🧪 毒駕拒測"]))
        m4.metric("慢車毒駕", len(df_raw[df_raw["違規類別判定"] == "🚲 慢車毒駕"]))
        m5.metric("本次納入合計", len(df_filtered))

        st.divider()

        # 分頁呈現各維度分析
        tab_unit, tab_officer, tab_detail = st.tabs(["🏢 所隊分局績效彙整", "👮 出力人員敘獎名冊", "📑 篩選後明細清冊"])

        # TAB 1: 單位統計
        with tab_unit:
            st.subheader("🏢 各所隊專案查獲取締件數統計")
            if unit_col:
                if "合計" in df_filtered.columns:
                    unit_summary = df_filtered.groupby(unit_col)["合計"].sum().reset_index()
                    unit_summary.columns = ["單位名稱", "查獲取締件數"]
                else:
                    unit_summary = df_filtered[unit_col].value_counts().reset_index()
                    unit_summary.columns = ["單位名稱", "查獲取締件數"]

                unit_summary = unit_summary.sort_values(by="查獲取締件數", ascending=False)
                
                c_u1, c_u2 = st.columns([1, 2])
                with c_u1:
                    st.dataframe(unit_summary, use_container_width=True, hide_index=True)
                with c_u2:
                    st.bar_chart(unit_summary.set_index("單位名稱"))
            else:
                st.info("此工作表未包含單位欄位，請切換至「各單位法條-件數統計報表」工作表。")

        # TAB 2: 出力人員敘獎建議表
        with tab_officer:
            st.subheader("👮 加強攔查取締毒駕專案出力人員敘獎名冊")
            
            # 敘獎標準參數調整
            with st.expander("⚙️ 專案敘獎建議額度規則設定", expanded=False):
                s1, s2 = st.columns(2)
                merit_2_threshold = s1.number_input("記功一次查獲件數門檻：", min_value=1, value=5)
                commend_2_threshold = s2.number_input("嘉獎二次查獲件數門檻：", min_value=1, value=2)

            if officer_col:
                # 統計員警件數
                agg_dict = {"查獲件數": (officer_col, "count")}
                if unit_col:
                    officer_stats = df_filtered.groupby([unit_col, officer_col]).size().reset_index(name="查獲件數")
                    officer_stats.columns = ["服務單位", "員警姓名", "查獲件數"]
                else:
                    officer_stats = df_filtered[officer_col].value_counts().reset_index()
                    officer_stats.columns = ["員警姓名", "查獲件數"]

                officer_stats = officer_stats.sort_values(by="查獲件數", ascending=False)

                # 試算建議額度
                def calculate_reward(count):
                    if count >= merit_2_threshold:
                        return "記功一次（專案主力）"
                    elif count >= commend_2_threshold:
                        return "嘉獎二次"
                    elif count >= 1:
                        return "嘉獎一次"
                    return "列入參考"

                officer_stats["建議獎勵"] = officer_stats["查獲件數"].apply(calculate_reward)
                officer_stats["事由"] = "執行加強攔查取締施用毒品後駕車專案工作計畫，查獲毒品後駕車相關違規案件出力。"

                st.dataframe(officer_stats, use_container_width=True, hide_index=True)

                # 匯出 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    officer_stats.to_excel(writer, index=False, sheet_name="出力人員敘獎建議名冊")
                st.download_button(
                    label="📥 下載【加強取締毒駕專案出力人員敘獎建議表】(Excel)",
                    data=output.getvalue(),
                    file_name="加強攔查取締施用毒品後駕車專案出力人員敘獎建議表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ 目前選取的工作表為「總量匯總表」，未包含【舉發員警】或【單號明細】欄位。")
                st.info("💡 建議：請於上方下拉選單切換至「案件明細」或包含員警清冊之工作表，即可自動彙整員警個人榜單與敘獎建請名冊。")

        # TAB 3: 篩選明細
        with tab_detail:
            st.subheader(f"📑 篩選後案件清單（共 {len(df_filtered)} 筆）")
            st.dataframe(df_filtered, use_container_width=True)
    else:
        st.error("無法在檔案中自動辨識到「違規事實」或「法條」欄位，請確認工作表內容。")
        st.dataframe(df_raw.head(10), use_container_width=True)
else:
    st.info("💡 請從左上方上傳從 Gmail 取得之 `自選匯出.xlsx` 報表，系統將自動區隔毒駕/酒駕/拒測/累犯進行分析。")
