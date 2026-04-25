import streamlit as st
import pandas as pd
import io
import re
import traceback
import smtplib
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

# ==========================================
# 0. 系統初始化
# ==========================================
st.set_page_config(page_title="交通業務與督導整合引擎", page_icon="🚓", layout="wide")

try:
    from menu import show_sidebar
    show_sidebar()
except:
    pass

st.markdown("""
    <style>
    @font-face { font-family: 'Kaiu'; src: url('kaiu.ttf'); }
    .stTextArea textarea {
        font-family: 'Kaiu', "標楷體", sans-serif !important;
        font-size: 19px !important;
        line-height: 1.7 !important;
        color: #1c1c1c !important;
    }
    .stTabs [data-baseweb="tab-list"] button {
        font-size: 18px;
        font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)


# ==========================================
# 1. 共用工具函數
# ==========================================
def d_safe_int(val):
    try:
        return int(float(str(val).split('.')[0].replace(',', '')))
    except:
        return 0


def d_normalize_code(c):
    c_str = str(c).strip().replace(".0", "").upper()
    return str(int(c_str)) if c_str.isdigit() else c_str


def d_parse_time(val):
    val_str = str(val).strip().replace("\n", "").replace(" ", "").replace("|", "-")
    if val_str in ["", "nan", "NaN"]:
        return None, None
    m_time = re.search(r'(\d{1,2})[~~\-－—–_]+(\d{1,2})', val_str)
    if m_time:
        return int(m_time.group(1)), int(m_time.group(2))
    return None, None


def d_detect_unit_type(unit_name: str, f_map: dict) -> str:
    """
    根據單位名稱優先判斷單位類型，
    若名稱無法判斷則從幹部職稱 f_map 反推。
    回傳 '所'、'分隊'、'隊' 其中之一。
    """
    if "分隊" in unit_name:
        return "分隊"
    if "隊" in unit_name and "分隊" not in unit_name:
        return "隊"
    if "所" in unit_name:
        return "所"
    # fallback：從職稱推斷
    titles = " ".join(f_map.values())
    if "分隊長" in titles or "小隊長" in titles:
        return "分隊"
    return "所"


# ==========================================
# 2. 裝備交接簿解析
# ==========================================
def d_extract_equip(e_file, hour):
    try:
        if e_file.name.endswith('csv'):
            df = pd.read_csv(e_file, header=None)
        else:
            df = pd.read_excel(e_file, header=None)

        df_s = df.astype(str)
        col_map = {"gun": None, "bullet": None, "radio": None, "vest": None}

        for r in range(min(10, len(df))):
            for c in range(len(df.columns)):
                v = str(df.iloc[r, c]).replace(" ", "").replace("　", "").replace("\n", "")
                if col_map["gun"] is None and ("手槍" in v or ("槍" in v and "手" in v)):
                    col_map["gun"] = c
                if col_map["bullet"] is None and ("子彈" in v or ("彈" in v and "子" in v)):
                    col_map["bullet"] = c
                if col_map["radio"] is None and "無線電" in v:
                    col_map["radio"] = c
                if col_map["vest"] is None and ("背心" in v or "防彈衣" in v):
                    col_map["vest"] = c

        defaults = [2, 3, 6, 11]
        col_map = {
            k: (v if v is not None else d)
            for k, v, d in zip(col_map.keys(), col_map.values(), defaults)
        }

        stop_r = len(df)
        for r in range(min(10, len(df)), len(df)):
            nums = re.findall(r'\d{1,2}', str(df.iloc[r, 0]))
            if nums and int(nums[0]) > hour and (int(nums[0]) - hour < 12):
                stop_r = r
                break

        sub = df.iloc[:stop_r]
        sub_s = df_s.iloc[:stop_r]

        def get_v(kw, k):
            rows = sub[sub_s.iloc[:, 1].str.contains(kw, na=False)]
            return d_safe_int(rows.iloc[-1, col_map[k]]) if not rows.empty else 0

        r_fix = sub[sub_s.iloc[:, 1].str.contains("送", na=False)]
        gf = d_safe_int(r_fix.iloc[-1, col_map["gun"]]) if not r_fix.empty else 0
        rf = d_safe_int(r_fix.iloc[-1, col_map["radio"]]) if not r_fix.empty else 0

        return {
            "gi": get_v("在", "gun"),  "go": get_v("出", "gun"),  "gf": gf,
            "bi": get_v("在", "bullet"), "bo": get_v("出", "bullet"),
            "ri": get_v("在", "radio"), "ro": get_v("出", "radio"), "rf": rf,
            "vi": get_v("在", "vest"),  "vo": get_v("出", "vest"),
        }
    except:
        return None


# ==========================================
# 3. 勤務表解析
# ==========================================
def d_extract_duty(d_file, hour):
    res = {
        'v_name': '未偵測',
        'cadre_status': '無幹部資料',
        'unit_name': '未偵測單位',
        'unit_type': '所',
    }
    try:
        df = pd.read_excel(d_file, header=None, dtype=str).fillna("")

        # --- 偵測單位名稱 ---
        for r in range(5):
            rt = "".join([str(x) for x in df.iloc[r].values])
            m = re.search(r'([\u4e00-\u9fa5]+(分局|派出所|分隊|局|隊))', rt)
            if m:
                res['unit_name'] = m.group(1)
                break

        # --- 建立人員代碼對照表 ---
        full = " ".join([str(x).strip() for x in df.values.flatten() if x])
        p = r'(?<![A-Za-z0-9])([A-Z]|[0-9]{1,2})\s*(所長|副所長|分隊長|小隊長|巡官|巡佐|警員|實習)\s*([\u4e00-\u9fa5]{2,4})'
        matches = re.findall(p, full)

        n_map, f_map = {}, {}
        for m in matches:
            code, title, name = d_normalize_code(m[0]), m[1].strip(), m[2]
            for t in ["所", "副", "巡", "警", "實", "員", "長", "隊"]:
                if name.endswith(t):
                    name = name[:-1]
            if len(name) >= 2:
                n_map[code] = name
                f_map[code] = f"{title}{name}"

        # --- 判斷單位類型（所 / 分隊 / 隊） ---
        res['unit_type'] = d_detect_unit_type(res['unit_name'], f_map)
        unit_type = res['unit_type']

        # --- 偵測時間欄位 ---
        tr_idx, t_cols, t_col_idx = 2, {}, -1
        for r in range(5):
            tmp = {}
            for c in range(len(df.columns)):
                sh, eh = d_parse_time(df.iloc[r, c])
                if sh is not None:
                    tmp[c] = (sh, eh)
            if len(tmp) > len(t_cols):
                tr_idx, t_cols = r, tmp

        for c, (sh, eh) in t_cols.items():
            ce = eh if eh > sh else eh + 24
            ch = hour if hour >= 6 or ce <= 24 else hour + 24
            if sh <= ch < ce:
                t_col_idx = c

        # --- 偵測值班人員 ---
        vr_idx = tr_idx + 1
        for r in range(tr_idx + 1, min(tr_idx + 4, len(df))):
            if "值" in str(df.iloc[r, 0]) + str(df.iloc[r, 1]):
                vr_idx = r
                break

        if t_col_idx != -1:
            raw = str(df.iloc[vr_idx, t_col_idx]).strip()
            mc = re.search(r'[A-Za-z0-9]{1,2}', raw)
            if mc:
                code = d_normalize_code(mc.group(0))
                res['v_name'] = f_map.get(code, f"未建檔:{code}")
            else:
                for c, n in n_map.items():
                    if n in raw:
                        res['v_name'] = f_map.get(c, n)
                        break

        # --- 幹部狀態分析 ---
        # 依單位類型決定幹部代碼預設職稱
        if unit_type == "分隊":
            default_titles = {"A": "分隊長", "B": "小隊長", "C": "幹部"}
        elif unit_type == "隊":
            default_titles = {"A": "隊長", "B": "副隊長", "C": "幹部"}
        else:
            default_titles = {"A": "所長", "B": "副所長", "C": "幹部"}

        # 依單位類型決定督勤場所用語
        location_word = unit_type  # 所 / 分隊 / 隊

        c_notes = []
        for code in ["A", "B", "C"]:
            full_name = f_map.get(code, default_titles[code])
            found, is_off, p_slots, d_names = False, False, [], set()

            for r in range(vr_idx, len(df)):
                dt = str(df.iloc[r, 0]) + str(df.iloc[r, 1])
                is_l = any(k in dt for k in ["休", "假", "輪", "輸", "補", "外"])
                for c, (sh, eh) in t_cols.items():
                    cell_codes = [d_normalize_code(x) for x in re.findall(r'[A-Za-z0-9]{1,2}', str(df.iloc[r, c]))]
                    if code in cell_codes:
                        found = True
                        if is_l:
                            is_off = True
                        else:
                            is_e = False
                            for k, kn in zip(
                                ["巡", "守", "臨", "交", "路"],
                                ["巡邏", "守望", "臨檢", "交整", "路檢"]
                            ):
                                if k in dt or k in str(df.iloc[r, c]):
                                    d_names.add(kn)
                                    is_e = True
                            if is_e:
                                p_slots.append((sh, eh))

            if not found or is_off:
                c_notes.append(f"{full_name}休假")
            else:
                if p_slots:
                    ms = min(s[0] for s in p_slots)
                    me = max(s[1] for s in p_slots)
                    e_str = "24" if me in (24, 0) else f"{me % 24:02d}"
                    duty_str = "、".join(sorted(d_names))
                    c_notes.append(
                        f"{full_name}在{location_word}督勤，編排{ms:02d}至{e_str}時段{duty_str}勤務"
                    )
                else:
                    c_notes.append(f"{full_name}在{location_word}督勤")

        res['cadre_status'] = "；".join(c_notes) + "。"

    except Exception:
        res['v_name'] = "解析失敗"
        res['cadre_status'] = f"解析錯誤：{traceback.format_exc()}"

    return res


# ==========================================
# 4. 報告文字產生（主詞自動對應）
# ==========================================
def build_report(dr, er, u_time, d_s5_s, d_end_s, d_s3_s):
    """
    依據 unit_type 決定所有主詞（該所 / 該分隊 / 該隊），
    組合並回傳完整報告文字。
    """
    unit_type = dr.get('unit_type', '所')
    subject = f"該{unit_type}"          # 該所 / 該分隊 / 該隊
    location = unit_type               # 所 / 分隊 / 隊（用於「在所督勤」等）

    eq = er if er else {
        "gi": 0, "go": 0, "gf": 0,
        "bi": 0, "bo": 0,
        "ri": 0, "ro": 0, "rf": 0,
        "vi": 0, "vo": 0,
    }
    fix = (
        f"（另有槍枝 {eq['gf']} 把、無線電 {eq['rf']} 臺送修中）"
        if (eq['gf'] + eq['rf']) > 0 else ""
    )

    lns = [
        f"{u_time.strftime('%H%M')}，{subject}值班{dr['v_name']}服裝整齊，佩件齊全，對槍、彈、無線電等裝備管制良好，領用情形均熟悉。",
        f"{subject}駐地監錄設備及天羅地網系統均運作正常，無故障，{d_s5_s}至{d_end_s}有逐日檢測2次以上紀錄。",
        f"{subject}{d_s3_s}至{d_end_s}勤前教育，幹部均有宣導「防制員警酒後駕車」、「員警駕車行駛交通優先權」及「追緝車輛執行原則」，參與同仁均有點閱。",
        f"{subject}環境內務擺設整齊清潔，符合規定。",
        (
            f"{subject}手槍出勤 {eq['go']} 把、在{location} {eq['gi']} 把，"
            f"子彈出勤 {eq['bo']} 顆、在{location} {eq['bi']} 顆，"
            f"無線電出勤 {eq['ro']} 臺、在{location} {eq['ri']} 臺；"
            f"防彈背心出勤 {eq['vo']} 件、在{location} {eq['vi']} 件，"
            f"幹部對械彈每日檢查管制良好，符合規定{fix}。"
        ),
        f"本日{dr['cadre_status']}",
        f"{subject}酒測聯單日期、編號均依規定填寫、黏貼，無跳號情形。",
    ]

    return "\n".join(f"{idx+1}、{line}" for idx, line in enumerate(lns))


# ==========================================
# 5. 郵件寄送功能
# ==========================================
def send_supervisor_email(insp_date: datetime, full_text: str) -> tuple[bool, str]:
    """
    將督導報告全文以純文字附件寄出。
    需在 st.secrets 設定：
      [email]
      user     = "yourmail@gmail.com"
      password = "your_app_password"
      to       = "recipient@example.com"   # 可選，預設寄給自己
    """
    try:
        cfg      = st.secrets["email"]
        sender   = cfg["user"]
        password = cfg["password"]
        recipient = cfg.get("to", sender)   # 若無 to 則寄給自己

        date_str  = insp_date.strftime("%Y%m%d")
        date_label = insp_date.strftime("%m月%d日")

        msg            = MIMEMultipart()
        msg["From"]    = sender
        msg["To"]      = recipient
        msg["Subject"] = f"勤務督導報告_{date_label}"

        # 郵件本文
        body = (
            f"龍潭分局勤務督導報告\n"
            f"督導日期：{date_label}\n"
            f"產製時間：{datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n"
            f"{'─'*40}\n\n"
            f"{full_text}"
        )
        msg.attach(MIMEText(body, "plain", "utf-8"))

        # txt 附件（utf-8-sig 讓 Windows 記事本正常開啟）
        attach_name = f"督導報告_{date_str}.txt"
        part = MIMEBase("text", "plain", charset="utf-8")
        part.set_payload(full_text.encode("utf-8-sig"))
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename*=UTF-8''{_ul.quote(attach_name)}"
        )
        msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, recipient, msg.as_string())

        return True, ""

    except KeyError:
        return False, "secrets.toml 缺少 [email] 設定（user / password）"
    except Exception as e:
        return False, str(e)


# ==========================================
# 6. 主介面
# ==========================================
st.header("📋 勤務督導報告自動生成")

c_s1, c_s2 = st.columns(2)
with c_s1:
    insp_date = st.date_input("選擇督導日期", datetime.now(), key="insp_d")
    num_units = st.number_input("單位數量", 1, 8, 3, key="num_u")
with c_s2:
    d_end   = insp_date - timedelta(days=1)
    d_s5    = insp_date - timedelta(days=5)
    d_s3    = insp_date - timedelta(days=3)
    d_end_s = d_end.strftime('%m月%d日')
    d_s5_s  = d_s5.strftime('%m月%d日')
    d_s3_s  = d_s3.strftime('%m月%d日')
    st.info(f"📅 區間：監錄({d_s5_s}–{d_end_s}) / 勤教({d_s3_s}–{d_end_s})")

u_tabs = st.tabs([f"🏢 單位 {i+1}" for i in range(num_units)] + ["📄 總匯整報告"])

# session_state 儲存各單位報告，key = "report_i"，跨 rerun 不遺失
if "unit_reports" not in st.session_state:
    st.session_state["unit_reports"] = {}

for i in range(num_units):
    with u_tabs[i]:
        u_time = st.time_input("抵達時間", datetime.now().time(), key=f"ut_{i}")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            u_duty = st.file_uploader("1. 上傳勤務表 (.xlsx)", type=['xlsx'], key=f"ud_{i}")
        with col_f2:
            u_eq = st.file_uploader("2. 上傳裝備交接簿 (.xlsx)", type=['xlsx'], key=f"ue_{i}")

        if u_duty and u_eq:
            # 用檔名組合當 cache key，檔案沒換就不重新解析
            file_key = f"{u_duty.name}|{u_eq.name}|{u_time.hour}"
            stored   = st.session_state["unit_reports"].get(i) or {}

            if stored.get("file_key") != file_key:
                with st.spinner("解析中…"):
                    dr = d_extract_duty(u_duty, u_time.hour)
                    er = d_extract_equip(u_eq, u_time.hour)
                report = build_report(dr, er, u_time, d_s5_s, d_end_s, d_s3_s)
                st.session_state["unit_reports"][i] = {
                    "file_key":  file_key,
                    "uname":     dr["unit_name"],
                    "unit_type": dr["unit_type"],
                    "v_name":    dr["v_name"],
                    "cadre":     dr["cadre_status"],
                    "report":    report,
                }

            rec = st.session_state["unit_reports"][i]
            st.success(f"✅ {rec['uname']}（{rec['unit_type']}）報告已完成")

            with st.expander("🔍 偵測結果核查"):
                st.write(f"**單位名稱：** {rec['uname']}")
                st.write(f"**單位類型：** {rec['unit_type']}")
                st.write(f"**值班人員：** {rec['v_name']}")
                st.write(f"**幹部狀態：** {rec['cadre']}")

            st.text_area("📋 預覽報告", rec["report"], height=380, key=f"txt_{i}")

        elif u_duty and not u_eq:
            st.warning("請補上裝備交接簿檔案。")
        elif u_eq and not u_duty:
            st.warning("請補上勤務表檔案。")
        # 注意：兩個都沒上傳時不清除 session_state，
        # 避免切換 tab 觸發 rerun 時把其他單位已存好的資料誤刪

# 總匯整：從 session_state 按單位編號順序組合
with u_tabs[-1]:
    saved = st.session_state.get("unit_reports", {})
    all_final_reports = [
        f"【{saved[i]['uname']} 督導報告】\n{saved[i]['report']}"
        for i in sorted(saved.keys())
        if i < num_units
    ]
    if all_final_reports:
        full_text = ("\n\n" + "─" * 40 + "\n\n").join(all_final_reports)
        st.text_area("📄 總匯整結果", full_text, height=700, key="full_report")

        col_dl, col_mail = st.columns(2)

        col_dl.download_button(
            label="⬇️ 下載報告 (.txt)",
            data=full_text.encode("utf-8-sig"),
            file_name=f"督導報告_{insp_date.strftime('%Y%m%d')}.txt",
            mime="text/plain",
            use_container_width=True,
        )

        if col_mail.button("📧 寄出督導報告郵件", use_container_width=True):
            with st.spinner("寄送中…"):
                ok, err_msg = send_supervisor_email(insp_date, full_text)
            if ok:
                st.success("✅ 郵件已成功寄出！")
            else:
                st.error(f"❌ 寄信失敗：{err_msg}")

        with st.expander("📮 郵件設定說明（點此展開）"):
            st.markdown("""
**需在 `.streamlit/secrets.toml` 加入以下設定：**

```toml
[email]
user     = "yourmail@gmail.com"      # 寄件 Gmail 帳號
password = "xxxx xxxx xxxx xxxx"     # Gmail 應用程式密碼（非登入密碼）
to       = "recipient@example.com"   # 收件人（可省略，省略則寄給自己）
```

> **Gmail 應用程式密碼取得方式：**  
> Google 帳戶 → 安全性 → 兩步驟驗證（需開啟）→ 應用程式密碼 → 產生
""")
    else:
        st.info("請先至各單位分頁上傳檔案並完成解析。")
