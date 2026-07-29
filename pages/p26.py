import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Font
import io
import os
import smtplib
import urllib.parse as _ul
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# 匯入系統原本的側邊欄設定
try:
    from menu import show_sidebar
except ImportError:
    def show_sidebar():
        pass

def send_csv_email(file_bytes, file_name):
    try:
        sender, pwd = st.secrets["email"]["user"], st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg["From"], msg["To"] = sender, sender
        
        msg["Subject"] = f"龍潭分局_行人及護老專案自動產出結果_{file_name}"
        
        body_text = (
            f"您好，\n\n"
            f"系統已自動產出「115年上半年行人及護老專案」相關檔案，附件為對應的 Excel 統計檔案。\n\n"
            f"本信件由交通執法自動化分析引擎發送。"
        )
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(file_bytes.getvalue())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{_ul.quote(file_name)}")
        msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, pwd)
            server.sendmail(sender, sender, msg.as_string())
            
        return True, None
    except Exception as e:
        return False, str(e)

def generate_excel_file(combined_date_str, total_hours):
    file_path = '376431843C_1150087037_ATTACH4.xlsx'
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到範本檔案 {file_path}")

    wb = openpyxl.load_workbook(file_path)
    ws = wb['警力統計']

    for row in range(3, 7):
        cell_c = ws.cell(row=row, column=3)
        cell_c.value = combined_date_str
        cell_c.alignment = Alignment(wrapText=True, vertical='center', horizontal='left')
        
        cell_d = ws.cell(row=row, column=4)
        cell_d.value = total_hours
        cell_d.alignment = Alignment(vertical='center', horizontal='center')
        cell_d.font = Font(name='微軟正黑體', size=12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    show_sidebar()

    # 注入 CSS 讓網頁預覽文字強制採用「標楷體」
    st.markdown("""
        <style>
        .kaiti-box {
            font-family: 'DFKai-SB', 'BiauKai', 'KaiTi', '標楷體', serif !important;
            font-size: 16px !important;
            line-height: 1.6 !important;
        }
        </style>
    """, unsafe_allow_html=True)

    st.title("👵 行人及護老專案交辦單與時數統計系統")
    st.info("本系統整合了「正式交辦單文字檔產生器（支援標楷體預覽）」與「護老專案時數統計與報表匯出」功能。")

    tab1, tab2 = st.tabs(["📄 1. 各單位交辦單產生器", "📊 2. 護老專案時數統計與匯出"])

    # ==========================================
    # TAB 1: 交辦單產生器 (標楷體預覽與純文字下載)
    # ==========================================
    with tab1:
        st.subheader("📋 桃市警龍潭分局交通組交(會)辦單")
        st.write("依據警察局來文指示[cite: 1]，選擇受文單位即可直接預覽標準標楷體格式之正式交辦單並下載檔案。")

        units = ["龍潭所", "聖亭所", "中興所", "石門所", "高平所", "勤務指揮中心", "龍潭交通分隊"]
        selected_unit = st.selectbox("選擇受文單位：", units, key="slip_unit")
        
        notice_content = f"""桃園市政府警察局龍潭分局交通組交(會)辦單
受文者：{selected_unit}[cite: 1]
交(會)辦日期：115年7月21日[cite: 1]
單位主管：組長楊孟竟[cite: 1]

交辦事項：
一、為辦理本分局115年1至6月各單位執行「行人及護老交通安全實施計畫」工作出力人員敘獎案，請統計所屬執勤時數並彙整敘獎人員名冊[cite: 1]。
二、獎勵規則：
    1、行人及護老交通安全實施計畫：
    (一)本案專責勤務人員執行成效良好，每半年執勤時數累計達40小時以上者核予嘉獎一次、80小時以上者核予嘉獎二次[cite: 1]。
    (二)略.......[cite: 1]
    (三)本案專責勤務及督導人員每人每半年獎勵額度以嘉獎二次為限，其中上半年未達獎勵額度之時數(勤務人員未達40小時或督導人員未達60小時)，得累計至當年度下半年計算[cite: 1]。
三、請至網路硬碟/交通組/巡官郭勝隆的資料夾/☆115年1至6月「行人及護老交通安全勤務實施計畫」工作出力人員敘獎區☆，下載115年1至6月「行人及護老交通安全勤務實施計畫」工作出力人員獎勵清冊，再依「115年辦公日曆表」，「1月至6月非假日之每日06-10時段及16-20時段，凡勤務分配表有顯示「護老專案」之服勤時數均得計算，再請於人事資訊整合管理系統登錄獎懲資料[cite: 1]。
四、115年1至6月「行人及護老交通安全勤務實施計畫」工作出力人員獎勵清冊由主管核章後附交辦單，逕送本組[cite: 1]。
五、登錄獎懲資料的流程：
    新增 > 主旨事由: 115年1至6月執行「行人及護老交通安全實施計畫」出力人員獎勵案 > 受理單位代碼: 交通組 > 受理人員代碼: 郭勝隆 > 再按新增 > 加入獎懲人員 > 獎懲事由: 115年1至6月執行「行人及護老交通安全」專責勤務達幾小時[cite: 1]。
六、請不用寫辛勞得力及備極辛勞，系統會自動帶入[cite: 1]。

辦理期限：115年7月28日前辦理完畢連同原件具報[cite: 1]。"""
        
        # 使用 HTML 容器強制套用標楷體樣式進行預覽
        st.markdown(f'<div class="kaiti-box"><pre style="font-family: inherit; white-space: pre-wrap;">{notice_content}</pre></div>', unsafe_allow_html=True)
        
        st.download_button(
            label=f"📥 下載【{selected_unit}】正式交辦單 (純文字檔)",
            data=notice_content.strip().encode('utf-8-sig'),
            file_name=f"115年上半年行人及護老專案交辦單_{selected_unit}.txt",
            mime="text/plain",
            use_container_width=True
        )

    # ==========================================
    # TAB 2: 時數統計與匯出
    # ==========================================
    with tab2:
        st.subheader("📝 勤務日期與時段設定")
        
        default_dates = []
        for m in range(1, 7):
            for d in [5, 12, 19, 26]:
                default_dates.append(f"{m}/{d}(06-10)")
                default_dates.append(f"{m}/{d}(16-20)")
                
        default_str = "、".join(default_dates)

        combined_date_str = st.text_area(
            "請確認或修改督勤日期清單 (各時段請以頓號「、」分隔)：", 
            value=default_str, 
            height=200
        )

        hours_per_shift = 4

        if combined_date_str.strip() == "":
            total_shifts = 0
        else:
            total_shifts = len(combined_date_str.split("、"))
            
        total_hours = total_shifts * hours_per_shift
        total_days = total_shifts // 2

        st.subheader("📋 運算結果預覽")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric(label="合計排定天數", value=f"{total_days} 天")
        with col2:
            st.metric(label="合計督勤班次", value=f"{total_shifts} 班")
        with col3:
            st.metric(label="總計時數 (D欄輸出值)", value=f"{total_hours} 小時")

        st.divider()

        st.subheader("📥 匯出與寄送督勤表")
        
        if total_shifts > 0:
            try:
                excel_data = generate_excel_file(combined_date_str, total_hours)
                file_name = "115年上半年督導行人及護老交通安全勤務表.xlsx"

                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        label="📥 下載督勤日期表 (Excel)",
                        data=excel_data,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col2:
                    if st.button("📧 將此報表一鍵寄至我的信箱", use_container_width=True):
                        with st.spinner("信件發送中，請稍候…"):
                            ok, mail_err = send_csv_email(excel_data, file_name)
                            if ok:
                                st.success("✅ 信件發送成功！報表已隨信夾帶至您的信箱。")
                            else:
                                st.error(f"❌ 發信失敗: {mail_err}")

            except FileNotFoundError as fnf_err:
                st.error(f"⚠️ 錯誤：{fnf_err}。請確認 `376431843C_1150087037_ATTACH4.xlsx` 檔案是否存在於主目錄中。")
            except Exception as e:
                st.error(f"⚠️ 發生未知的錯誤：{e}")
        else:
            st.warning("請先於上方輸入督勤日期，方可匯出表單。")

if __name__ == "__main__":
    main()
