import streamlit as st
import openpyxl
import re
from datetime import datetime

st.set_page_config(page_title="勤務督導報告系統", layout="wide")

# ==========================================
# 函式定義
# ==========================================

def extract_duty_v2(file, current_hour: int) -> dict:
    try:
        wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))

        # 單位名稱（Row1）
        title_cell = str(all_rows[0][0]) if all_rows[0][0] else ''
        m_term = re.search(r'龍潭分局\s*([\u4e00-\u9fa5]+派出所|[\u4e00-\u9fa5]+隊)', title_cell)
        term = m_term.group(1) if m_term else '本所'

        # 番號→姓名對照表（Row44~48）
        code_map = {}
        for row in all_rows[43:48]:
            for grp in range(6):
                base = 1 + grp * 6
                if base + 1 >= len(row):
                    break
                code = row[base]
                name_cell = row[base + 1]
                if not code or code == '輪番番號':
                    continue
                if name_cell and isinstance(name_cell, str):
                    names = re.findall(r'[\u4e00-\u9fa5]{2,4}', name_cell)
                    if names:
                        code_map[str(code).strip()] = names[-1]

        # 找時段欄位
        time_headers = list(all_rows[2][13:])

        def find_col(hour):
            for i, h in enumerate(time_headers):
                if h is None:
                    continue
                h_str = str(h).split('\n')[0].strip()
                try:
                    if int(h_str) == hour:
                        return 13 + i
                except ValueError:
                    pass
            return None

        col = find_col(current_hour)

        # 值班番號（Row4）
        v_code = ''
        if col is not None:
            raw = all_rows[3][col]
            if raw is not None:
                v_code = str(raw).strip()

        v_name = code_map.get(v_code, f'番號{v_code}')
        roster = list(code_map.values())

        return {'term': term, 'v_name': v_name, 'roster': roster}

    except Exception as e:
        return {'term': '本所', 'v_name': '（解析失敗）', 'roster': [], '_error': str(e)}


def extract_equip_v2(file) -> dict:
    EQUIP_COLS = {
        2: '手槍', 3: '手槍子彈', 4: '65式步槍', 5: '65式步槍子彈',
        6: 'M16步槍', 7: 'M16步槍子彈', 8: '無線電', 9: '行動電腦',
        10: '防彈背心', 11: '酒測器', 12: '量知器', 13: '電擊器',
        14: '警用機車', 15: '巡邏車', 16: '防彈頭盔',
    }

    try:
        wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))

        groups = []
        seen_times = set()
        for i, row in enumerate(all_rows[3:], start=3):
            cell = row[0]
            if not cell or not isinstance(cell, str):
                continue
            if not re.search(r'\d{2}:\d{2}', cell):
                continue
            parts = cell.strip().split('\n')
            if len(parts) < 2:
                continue
            time_str = parts[0].strip()
            if time_str in seen_times:
                continue
            seen_times.add(time_str)
            from_officer = parts[1].strip() if len(parts) > 1 else ''
            to_officer = parts[2].strip() if len(parts) > 2 else ''
            if re.search(r'\d{1,3}-\d{2}-\d{2} \d{2}:\d{2}', time_str):
                groups.append({
                    'row_index': i, 'time_str': time_str,
                    'from_officer': from_officer, 'to_officer': to_officer,
                })

        if not groups:
            return {'ok': False, 'latest_time': '', 'from_officer': '',
                    'to_officer': '', 'summary': '無交接紀錄', 'anomalies': []}

        latest = groups[-1]
        base_i = latest['row_index']
        repair_row = all_rows[base_i + 2] if base_i + 2 < len(all_rows) else []
        inservice_row = all_rows[base_i]

        anomalies = []
        for col_i, name in EQUIP_COLS.items():
            if col_i < len(repair_row):
                val = repair_row[col_i]
                if val and isinstance(val, (int, float)) and val > 0:
                    anomalies.append(f'{name}送修{int(val)}件')

        gun_cnt = inservice_row[2] if 2 < len(inservice_row) else '?'
        bullet_cnt = inservice_row[3] if 3 < len(inservice_row) else '?'
        summary = (
            f"在所手槍{gun_cnt}支、子彈{bullet_cnt}發，"
            f"{'裝備全部正常' if not anomalies else '異常：' + '、'.join(anomalies)}"
        )

        return {
            'ok': len(anomalies) == 0,
            'latest_time': latest['time_str'],
            'from_officer': latest['from_officer'],
            'to_officer': latest['to_officer'],
            'summary': summary,
            'anomalies': anomalies,
        }

    except Exception as e:
        return {'ok': False, 'latest_time': '', 'from_officer': '',
                'to_officer': '', 'summary': f'解析失敗：{e}', 'anomalies': []}


# ==========================================
# Streamlit UI
# ==========================================

if 'unit_reports' not in st.session_state:
    st.session_state.unit_reports = {}

st.header("📋 勤務督導報告自動生成系統")

insp_date = st.date_input("選擇督導日期", datetime.now())
num_units = st.number_input("待督導單位數量", 1, 8, 3)

tab_labels = [f"🏢 單位 {i+1}" for i in range(num_units)] + ["📄 總匯整報告"]
u_tabs = st.tabs(tab_labels)

for i in range(num_units):
    with u_tabs[i]:
        u_time = st.time_input("抵達時間", datetime.now().time(), key=f"ut_{i}")
        col1, col2, col3 = st.columns(3)
        u_duty = col1.file_uploader("勤務表 (.xlsx)", type=['xlsx'], key=f"ud_{i}")
        u_eq   = col2.file_uploader("交接簿 (.xlsx)", type=['xlsx'], key=f"ue_{i}")
        u_pdf  = col3.file_uploader("刑案單 (.pdf)", type=['pdf'],
                                    accept_multiple_files=True, key=f"updf_{i}")

        if u_duty and u_eq:
            dr = extract_duty_v2(u_duty, u_time.hour)
            er = extract_equip_v2(u_eq)

            if '_error' in dr:
                st.warning(f"勤務表解析警告：{dr['_error']}")

            lns = [
                f"{u_time.strftime('%H%M')}，{dr['term']}值班{dr['v_name']}，{er['summary']}。"
            ]

            st.session_state.unit_reports[i] = "\n".join(
                [f"{idx+1}、{line}" for idx, line in enumerate(lns)]
            )
            st.text_area("報告預覽", st.session_state.unit_reports[i],
                         height=200, key=f"prev_{i}")

            if u_pdf:
                st.info("PDF 刑案單解析功能需要 OCR 環境，目前顯示上傳檔名供人工核對：")
                for pdf_file in u_pdf:
                    st.write(f"• {pdf_file.name}")

# 總匯整報告
with u_tabs[num_units]:
    st.subheader("📄 總匯整報告")
    all_reports = []
    for i in range(num_units):
        if i in st.session_state.unit_reports:
            all_reports.append(f"【單位 {i+1}】\n{st.session_state.unit_reports[i]}")

    if all_reports:
        full_report = "\n\n".join(all_reports)
        st.text_area("完整督導報告", full_report, height=400)
        st.download_button("⬇️ 下載報告", full_report,
                           file_name=f"督導報告_{insp_date}.txt",
                           mime="text/plain")
    else:
        st.info("請先在各單位頁籤上傳勤務表與交接簿，報告將自動匯整於此。")
