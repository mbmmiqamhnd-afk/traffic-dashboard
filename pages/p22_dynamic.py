import streamlit as st
import pandas as pd

st.set_page_config(page_title="綜合勤務規劃系統", layout="wide", page_icon="🚓")

# ==========================================
# 1. 建立「勤務版型設定檔」(Configuration Dictionary)
# ==========================================
# 這裡定義所有勤務的專屬特徵，未來新增勤務只要改這裡即可
DUTY_PROFILES = {
    "一般機動巡邏 (如防制危駕)": {
        "extra_cols": ["攜行裝備", "巡邏路段"],
        "default_focus": "採取全面機動巡邏，針對重點路段加強攔檢；發現異常立即通報，注意自身三安。",
        "col_config": {
            "巡邏路段": st.column_config.TextColumn("巡邏路段", width="large"),
            "攜行裝備": st.column_config.TextColumn("攜行裝備", width="medium")
        }
    },
    "定點場所臨檢 (如擴大臨檢)": {
        "extra_cols": ["臨檢目標場所"],
        "default_focus": "會合專案人員，準時進入目標場所執行威力掃蕩，全程開啟密錄器蒐證。",
        "col_config": {
            "臨檢目標場所": st.column_config.TextColumn("臨檢目標場所", width="large")
        }
    },
    "定點路檢 (如取締酒駕/砂石車)": {
        "extra_cols": ["攜行裝備", "定點路檢目標"],
        "default_focus": "於各聯外道路、重要路口設立檢測點，加強取締違規。",
        "col_config": {
            "定點路檢目標": st.column_config.TextColumn("定點路檢目標", width="large"),
            "攜行裝備": st.column_config.TextColumn("攜行裝備", width="medium")
        }
    }
}

# ==========================================
# 2. 側邊欄或主畫面：動態選擇器
# ==========================================
st.title("🚓 綜合勤務規劃系統 (動態介面測試)")

# 讓使用者選擇勤務類型
selected_duty = st.selectbox("📌 請選擇本次規劃的勤務類型", list(DUTY_PROFILES.keys()))

# 根據選擇，讀取對應的設定
profile = DUTY_PROFILES[selected_duty]

st.markdown("---")

# ==========================================
# 3. 動態渲染介面
# ==========================================
st.subheader(f"【{selected_duty}】勤務重點與編組")

# 動態帶入勤務重點預設值
focus_text = st.text_area("📢 勤務重點", value=profile["default_focus"], height=80)

# 基礎通用欄位 + 動態附加欄位
base_cols = ["組別", "無線電代號", "派遣單位", "姓名", "任務分工"]
dynamic_cols = base_cols + profile["extra_cols"]

# 建立預設假資料 (展示用，直接使用單位名稱)
default_units = ["石門", "中興", "聖亭", "龍潭", "高平", "三和", "交通分隊"]
default_data = []
for i, unit in enumerate(default_units[:3]): # 先塞三筆示範
    row = {
        "組別": f"第{i+1}組",
        "無線電代號": f"隆安{i+1}0",
        "派遣單位": unit,
        "姓名": "",
        "任務分工": "帶班"
    }
    # 把動態欄位也補上空值，確保 DataFrame 結構正確
    for col in profile["extra_cols"]:
        row[col] = ""
    default_data.append(row)

# 轉換為 DataFrame
df_template = pd.DataFrame(default_data, columns=dynamic_cols)

# 基礎的欄位寬度設定
base_config = {
    "組別": st.column_config.TextColumn("組別", width="small"),
    "無線電代號": st.column_config.TextColumn("無線電代號", width="small"),
    "派遣單位": st.column_config.TextColumn("派遣單位", width="small"),
    "姓名": st.column_config.TextColumn("姓名", width="small"),
    "任務分工": st.column_config.TextColumn("任務分工", width="medium"),
}

# 將基礎設定與動態設定合併
final_col_config = {**base_config, **profile["col_config"]}

# 渲染動態 Data Editor
st.caption("💡 下方表格已根據你選擇的勤務類型自動變換欄位：")
res_df = st.data_editor(
    df_template,
    num_rows="dynamic",
    use_container_width=True,
    column_config=final_col_config,
    key=f"editor_{selected_duty}" # 確保切換勤務時，快取不會互相干擾
)

# 預覽結果
with st.expander("🔍 檢視當前表格輸出的資料結構 (JSON)"):
    st.write(res_df.to_dict(orient="records"))
