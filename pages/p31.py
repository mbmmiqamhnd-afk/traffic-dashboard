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
st.caption("依據桃園市政府警察局 115年9月10日 桃警交字第1150120088號函辦理（辦理時限：115年9月17日）")

with st.expander("📌 專案公文重點提示與待辦時限", expanded=False):
    st.markdown("""
    * **主旨**：有關辦理本局執行內政部警政署加強攔查取締施用毒品後駕車專案工作計畫出力人員敘獎案，請依說明事項辦理，請查照。
    * **發文字號**：桃園市政府警察局 115 年 9 月 10 日 桃警交字第 1150120088 號函
    * **承辦窗口**：本局交通警察大隊
    * **辦理時限**：**民國 115 年 9 月 17 日（星期四）**
    * **具體待辦**：
      1. 彙整本分局執行毒駕專案查獲及線上攔檢出力員警名冊。
      2. 造具敘獎建議表陳核後，函報市警局交大辦理專案敘獎。
    """)

# ==================== 違規事實與法條代碼智慧分類 ====================
def classify_violation(fact_text, law_code=""):
    """
    結合法條代碼與違規事實文字，精準判定案件類別：
    - 35102...: 毒駕本體
    - 35300175/178/199/253: 毒駕累犯
    - 35402002: 拒絕接受毒品測試之檢定 (毒駕拒測)
    - 35101...: 一般酒駕
    - 35900...: 第35條第9項移置保管
    """
    text = f"{law_code} {fact_text}".strip()
    if not text or text == "nan":
        return "未填寫"

    is_drug = bool(re.search(r"毒|毒品|麻醉|迷幻|第1項第2款|第一項第二款|1項2款|35102|35300175|35300178|35300184|35300199|35300211|35300253|35300256|35402002", text))
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

# ==================== 檔案上傳與智慧標題識別 ====================
uploaded_file = st.file_uploader(
    "📂 請上傳「35條73條統計表」（自選匯出.xlsx）", 
    type=["xlsx", "xls"],
    help="支援自動跳過表頭公文資訊，精準定位「單號、違規事實1、違規法條1、舉發員警1」"
)

if uploaded_file:
    # 讀取全部工作表名稱
    xls = pd.ExcelFile(uploaded_file)
    sheet_options = xls.sheet_names
    
    # 預設選取「案件明細」
    default_idx = 0
    for idx, s in enumerate(sheet_options):
        if "案件明細" in s:
            default_idx = idx
            break
        elif any(k in s for k in ["明細", "案件", "法條"]):
            default_idx = idx

    selected_sheet = st.selectbox("📑 選擇工作表：", sheet_options, index=default_idx)
    
    # 1. 以無標題模式讀取前 20 列定位欄位
    df_raw_no_header = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)
    
    header_idx = None
    for idx in range(min(15, len(df_raw_no_header))):
        row_text = " ".join(df_raw_no_header.iloc[idx].dropna().astype(str).tolist())
        # 尋找包含「單號」或「違規事實」或「違規法條」之列
        if any(k in row_text for k in ["單號", "違規事實", "違規法條", "舉發員警"]):
            header_idx = idx
            break

    # 2. 定位標題列並下移切割資料
    if header_idx is not None:
        new_columns = [str(c).strip() for c in df_raw_no_header.iloc[header_idx].tolist()]
        df_clean = df_raw_no_header.iloc[header_idx + 1:].copy().reset_index(drop=True)
        df_clean.columns = new_columns
    else:
        df_clean = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

    # 3. 過濾底部的公文頁尾資訊（如：列印人員：郭勝隆、統計期間...）
    ticket_col = next((c for c in df_clean.columns if "單號" in str(c)), None)
    if ticket_col:
        df_clean = df_clean[df_clean[ticket_col].notna()]
        df_clean = df_clean[~df_clean[ticket_col].astype(str).str.contains("列印人員|統計期間|製表人員|總計", na=False)]

    # 4. 尋找核心欄位（支援「違規事實1」、「違規法條1」、「舉發員警1」等後綴）
    fact_col = next((c for c in df_clean.columns if any(k in str(c) for k in ["違規事實", "事實說明", "違規事項", "事實"])), None)
    law_col = next((c for c in df_clean.columns if any(k in str(c) for k in ["違規法條", "法條代碼", "法條"])), None)
    officer_col = next((c for c in df_clean.columns if any(k in str(c) for k in ["舉發員警", "員警", "警號", "姓名", "填單人"])), None)
    vehicle_col = next((c for c in df_clean.columns if any(k in str(c) for k in ["簡式車種名稱", "車種", "車種名稱"])), None)
    date_col = next((c for c in df_clean.columns if any(k in str(c) for k in ["入案日", "違規日", "日期"])), None)

    st.success(f"成功解析工作表【{selected_sheet}】，共計 **{len(df_clean)}** 筆有效案件紀錄。")

    if fact_col:
        # 標註分類標籤
        df_clean["案件分類"] = df_clean.apply(
            lambda r: classify_violation(
                r[fact_col], 
                r[law_col] if law_col else ""
            ), 
            axis=1
        )

        # 側邊/上方分類篩選控制台
        st.markdown("### 🎯 案件類別篩選（預設已勾選毒品專案）")
        all_cats = sorted(df_clean["案件分類"].unique().tolist())
        drug_default = [c for c in all_cats if "🧪" in c or "🚲 慢車毒駕" in c]

        # ✅ 這裡明確傳入 比例，保證不再報錯
        c_sel1, c_sel2 = st.columns()
        with c_sel1:
            selected_cats = st.multiselect(
                "請勾選本次納入統計之案件分類：",
                options=all_cats,
                default=drug_default if drug_default else all_cats
            )
        with c_sel2:
            quick_mode = st.radio("快速切換：", ["自訂", "僅毒品專案", "全部納入"], index=0, horizontal=True)
            if quick_mode == "僅毒品專案":
                selected_cats = [c for c in all_cats if "🧪" in c or "🚲 慢車毒駕" in c]
            elif quick_mode == "全部納入":
                selected_cats = all_cats

        # 篩選後的分析資料集
        df_filtered = df_clean[df_clean["案件分類"].isin(selected_cats)].copy()

        # 顯示指標卡
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("🧪 毒駕本體", len(df_clean[df_clean["案件分類"] == "🧪 毒駕本體"]))
        m2.metric("🧪 毒駕累犯", len(df_clean[df_clean["案件分類"] == "🧪 毒駕累犯"]))
        m3.metric("🧪 毒駕拒測", len(df_clean[df_clean["案件分類"] == "🧪 毒駕拒測"]))
        m4.metric("🚲 慢車毒駕", len(df_clean[df_clean["案件分類"] == "🚲 慢車毒駕"]))
        m5.metric("📌 本次統計納入", len(df_filtered))

        st.divider()

        # 分頁展示
        tab_officer, tab_summary, tab_detail = st.tabs(["👮 出力員警敘獎建議名冊", "📊 違規分類統計彙整", "📑 篩選後案件清冊"])

        # TAB 1: 出力人員敘獎名冊
        with tab_officer:
            st.subheader("👮 查獲及線上攔檢出力員警敘獎名冊")

            if officer_col:
                # 統計每位員警查獲件數
                officer_stats = df_filtered[officer_col].value_counts().reset_index()
                officer_stats.columns = ["員警姓名", "查獲件數"]

                # 依件數建議獎勵額度
                def get_reward_text(cnt):
                    if cnt >= 3:
                        return "嘉獎二次（專案主力）"
                    elif cnt >= 1:
                        return "嘉獎一次"
                    return "列入參考"

                officer_stats["建議獎勵"] = officer_stats["查獲件數"].apply(get_reward_text)
                officer_stats["具體事由"] = "執行加強攔查取締施用毒品後駕車專案工作計畫，查獲毒品後駕車違規案件出力。"

                st.dataframe(officer_stats, use_container_width=True, hide_index=True)

                # 匯出 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    officer_stats.to_excel(writer, index=False, sheet_name="出力人員敘獎名冊")
                    df_filtered.to_excel(writer, index=False, sheet_name="專案查獲案件明細")

                st.download_button(
                    label="📥 下載【加強取締施用毒品後駕車專案出力人員敘獎建議表】(Excel)",
                    data=output.getvalue(),
                    file_name="加強攔查取締施用毒品後駕車專案出力人員敘獎建議表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("此工作表未包含員警欄位。")

        # TAB 2: 分類分佈圖表
        with tab_summary:
            st.subheader("📊 案件分類分佈統計")
            summary_cat = df_clean["案件分類"].value_counts().reset_index()
            summary_cat.columns = ["案件類別", "件數"]
            c_g1, c_g2 = st.columns(2)
            with c_g1:
                st.dataframe(summary_cat, use_container_width=True, hide_index=True)
            with c_g2:
                st.bar_chart(summary_cat.set_index("案件類別"))

        # TAB 3: 案件明細清單
        with tab_detail:
            st.subheader(f"📑 符合條件之案件清單（共 {len(df_filtered)} 筆）")
            # 優先排列核心欄位
            display_cols = [c for c in [ticket_col, "案件分類", law_col, fact_col, officer_col, vehicle_col, date_col] if c and c in df_filtered.columns]
            other_cols = [c for c in df_filtered.columns if c not in display_cols]
            st.dataframe(df_filtered[display_cols + other_cols], use_container_width=True, hide_index=True)
    else:
        st.error("未能在此工作表中自動辨識到「違規事實」欄位。請切換工作表或檢查檔案格式。")
        st.dataframe(df_clean.head(10), use_container_width=True)
else:
    st.info("💡 請上傳從 Gmail 下載之 `自選匯出.xlsx` 報表檔案。")
