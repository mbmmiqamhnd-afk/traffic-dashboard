import streamlit as st

# 設定網頁佈局為寬螢幕
st.set_page_config(layout="wide", page_title="勤務簡報系統")

# 自訂 CSS 來重現來源中的網格背景與投影片風格 [1-3]
def apply_custom_css():
    st.markdown("""
        <style>
        /* 網格背景設定 */
        .stApp {
            background-color: #f8f9fa;
            background-image: 
                linear-gradient(#bdc3c7 1px, transparent 1px),
                linear-gradient(90deg, #bdc3c7 1px, transparent 1px);
            background-size: 30px 30px;
        }
        
        /* 頂部標題列樣式 */
        .slide-header {
            background-color: #002060;
            color: white;
            padding: 15px 30px;
            font-size: 32px;
            font-weight: bold;
            border-left: 15px solid #aeb6bf;
            margin-bottom: 30px;
            box-shadow: 2px 2px 5px rgba(0,0,0,0.2);
        }
        
        /* 資訊卡片與表格樣式 */
        .info-box {
            background-color: white;
            border: 3px solid #002060;
            padding: 0;
            margin-bottom: 20px;
        }
        .info-title {
            background-color: #002060;
            color: white;
            padding: 8px;
            text-align: center;
            font-weight: bold;
            font-size: 20px;
        }
        .info-content {
            padding: 15px;
            font-size: 18px;
            text-align: center;
            font-weight: bold;
        }
        
        /* 任務分工表格 */
        table {
            width: 100%;
            border-collapse: collapse;
            background-color: white;
            border: 3px solid #002060;
        }
        th {
            background-color: #002060;
            color: white;
            font-size: 20px;
            padding: 10px;
            border: 2px solid #002060;
        }
        td {
            padding: 12px;
            font-size: 18px;
            font-weight: bold;
            text-align: center;
            border: 2px solid #002060;
        }
        .bg-light-blue {
            background-color: #dbe2ef;
        }

        /* 雨天備案警告標誌 */
        .warning-box {
            background-color: #ffc000;
            border: 4px solid #000000;
            border-radius: 8px;
            padding: 20px;
            text-align: center;
            height: 100%;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }
        .warning-title {
            font-size: 36px;
            font-weight: bold;
            color: black;
        }
        .warning-text {
            font-size: 24px;
            font-weight: bold;
            color: black;
            margin-top: 10px;
        }
        </style>
    """, unsafe_allow_html=True)

# 投影片 1：第一階段勤務規劃 (對應來源的巡邏組排版) [3, 4]
def slide_phase_1():
    st.markdown('<div class="slide-header">第一階段 | 第1巡邏組勤務規劃 (21:00 - 22:30)</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([1.2, 1])
    
    with col1:
        st.markdown("""
        <div class="info-box">
            <div class="info-title">負責單位 / 代號</div>
            <div class="info-content">聖亭所 / 隆安52</div>
        </div>
        
        <div class="info-box">
            <div class="info-title">巡邏地點</div>
            <div class="info-content">
                <div style="font-size: 24px; margin-bottom: 15px;">於易發生酒駕、危險駕車<br>路段加強攔檢</div>
                <span style="border: 2px solid #002060; border-radius: 20px; padding: 5px 15px; margin: 5px; display: inline-block;">中豐路</span>
                <span style="border: 2px solid #002060; border-radius: 20px; padding: 5px 15px; margin: 5px; display: inline-block;">中豐路中山段</span>
                <span style="border: 2px solid #002060; border-radius: 20px; padding: 5px 15px; margin: 5px; display: inline-block;">中豐路上林段</span>
                <br>
                <span style="border: 2px solid #002060; border-radius: 20px; padding: 5px 15px; margin: 5px; display: inline-block;">大昌路一段</span>
                <span style="border: 2px solid #002060; border-radius: 20px; padding: 5px 15px; margin: 5px; display: inline-block;">中正路</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    with col2:
        st.markdown("""
        <div class="info-box" style="border:none;">
            <div class="info-title">任務分工</div>
            <table>
                <tr>
                    <th>職稱</th>
                    <th>姓名</th>
                </tr>
                <tr>
                    <td class="bg-light-blue">副所長</td>
                    <td class="bg-light-blue">邱品淳</td>
                </tr>
                <tr>
                    <td>警員</td>
                    <td>傅維強</td>
                </tr>
                <tr>
                    <td>警員</td>
                    <td>劉兆敏</td>
                </tr>
            </table>
        </div>
        """, unsafe_allow_html=True)

# 投影片 2：第二階段路檢規劃 (對應來源中包含雨天備案的排版) [5, 6]
def slide_phase_2():
    st.markdown('<div class="slide-header">第二階段 | 第1路檢組勤務規劃 (22:30 - 24:00)</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1])
    
    with col1:
        st.markdown("""
        <div class="info-box">
            <div class="info-title">負責單位 / 代號</div>
            <div class="info-content">中興所及高平所 / 隆安72</div>
        </div>
        <div class="info-box">
            <div class="info-title">路檢地點</div>
            <div class="info-content">
                中正路三坑段與美國路口<br>
                <span style="font-size: 40px;">⬇</span><br>
                攔檢往龍潭市區方向車輛
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    with col2:
        st.markdown("""
        <div class="warning-box">
            <div style="font-size: 80px; line-height: 1;">⚠️</div>
            <div class="warning-title">雨天備案</div>
            <div class="warning-text">轄區治安要點巡邏</div>
        </div>
        """, unsafe_allow_html=True)
        
    with col3:
        st.markdown("""
        <div class="info-box" style="border:none;">
            <div class="info-title">任務分工</div>
            <table>
                <tr>
                    <th>職稱</th>
                    <th>姓名</th>
                </tr>
                <tr>
                    <td class="bg-light-blue">副所長</td>
                    <td class="bg-light-blue">薛德祥</td>
                </tr>
                <tr>
                    <td>警員</td>
                    <td>蔡震東</td>
                </tr>
                <tr>
                    <td>警員</td>
                    <td>洪祥皓</td>
                </tr>
                <tr>
                    <td>警員</td>
                    <td>張維忻</td>
                </tr>
            </table>
        </div>
        """, unsafe_allow_html=True)

# 主程式邏輯
def main():
    apply_custom_css()
    
    # 側邊欄：導覽列
    st.sidebar.title("簡報導覽")
    slide_selection = st.sidebar.radio(
        "選擇投影片頁面：",
        ["第一階段 - 巡邏組規劃", "第二階段 - 路檢組規劃 (含雨天備案)"]
    )
    
    # 根據選擇渲染對應的投影片
    if slide_selection == "第一階段 - 巡邏組規劃":
        slide_phase_1()
    elif slide_selection == "第二階段 - 路檢組規劃 (含雨天備案)":
        slide_phase_2()

if __name__ == "__main__":
    main()
