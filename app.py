import streamlit as st
import pandas as pd
import numpy as np
import json
import os
from utils.data_processor import DataProcessor
from utils.map_visualization import MapVisualization
from utils.funnel_analysis import FunnelAnalysis
from utils.retention_analysis import RetentionAnalysis

# ============================================================
# 頁面設定
# ============================================================
st.set_page_config(
    page_title="中華醫事科技大學 招生數據分析系統",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# 自訂樣式
# ============================================================
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1B3A5C;
        text-align: center;
        padding: 1rem 0;
        border-bottom: 3px solid #E8792F;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.3rem;
        font-weight: 600;
        color: #2C5F8A;
        margin-top: 1.5rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .metric-value {
        font-size: 2.5rem;
        font-weight: 700;
    }
    .metric-label {
        font-size: 0.9rem;
        opacity: 0.9;
    }
    .info-box {
        background-color: #f0f7ff;
        border-left: 5px solid #2C5F8A;
        padding: 1rem 1.5rem;
        border-radius: 0 8px 8px 0;
        margin: 1rem 0;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #f0f2f6;
        border-radius: 8px 8px 0 0;
        padding: 10px 20px;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# 台灣縣市座標（內建）
# ============================================================
TAIWAN_COUNTY_COORDS = {
    "台北市": (25.0330, 121.5654), "新北市": (25.0120, 121.4650),
    "桃園市": (24.9936, 121.3010), "台中市": (24.1477, 120.6736),
    "台南市": (22.9998, 120.2269), "高雄市": (22.6273, 120.3014),
    "基隆市": (25.1276, 121.7392), "新竹市": (24.8138, 120.9675),
    "新竹縣": (24.8387, 121.0178), "苗栗縣": (24.5602, 120.8214),
    "彰化縣": (24.0518, 120.5161), "南投縣": (23.9610, 120.9718),
    "雲林縣": (23.7092, 120.4313), "嘉義市": (23.4801, 120.4491),
    "嘉義縣": (23.4518, 120.2551), "屏東縣": (22.5519, 120.5487),
    "宜蘭縣": (24.7570, 121.7533), "花蓮縣": (23.9872, 121.6016),
    "台東縣": (22.7583, 121.1444), "澎湖縣": (23.5711, 119.5793),
    "金門縣": (24.4493, 118.3767), "連江縣": (26.1505, 119.9499),
    "臺北市": (25.0330, 121.5654), "新北市": (25.0120, 121.4650),
    "桃園市": (24.9936, 121.3010), "臺中市": (24.1477, 120.6736),
    "臺南市": (22.9998, 120.2269), "臺東縣": (22.7583, 121.1444),
}


# ============================================================
# 生成範例資料
# ============================================================
def generate_sample_applicants():
    """生成申請入學範例資料"""
    np.random.seed(42)
    n = 500

    departments = ["護理系", "食品科技系", "醫學檢驗生物技術系", "職業安全衛生系",
                   "環境與安全衛生工程系", "生物醫學工程系", "藥學系", "長期照護系",
                   "幼兒保育系", "資訊管理系", "化妝品應用與管理系", "餐旅管理系"]

    schools_by_city = {
        "台南市": ["台南高工", "台南女中", "長榮高中", "南光高中", "港明高中",
                   "新營高工", "北門高中", "曾文高中", "南科實中", "家齊高中",
                   "後壁高中", "白河商工", "新化高中", "善化高中", "歸仁國中"],
        "高雄市": ["高雄中學", "高雄女中", "鳳山高中", "前鎮高中", "三信家商",
                   "中山工商", "高雄高工", "鳳新高中", "路竹高中", "岡山高中"],
        "嘉義縣": ["嘉義高中", "嘉義女中", "民雄農工", "東石高中", "協志工商"],
        "嘉義市": ["嘉義高商", "輔仁中學", "嘉華中學", "興華高中", "宏仁女中"],
        "屏東縣": ["屏東高中", "屏東女中", "屏東高工", "潮州高中", "恆春工商",
                   "東港高中", "佳冬農校", "內埔農工"],
        "雲林縣": ["斗六高中", "虎尾高中", "北港農工", "土庫商工", "西螺農工"],
        "彰化縣": ["彰化高中", "員林高中", "鹿港高中", "秀水高工", "大慶商工"],
        "台中市": ["台中一中", "台中女中", "豐原高中", "大甲高中", "明道中學"],
        "新北市": ["板橋高中", "中和高中", "三重商工", "淡水商工", "樹林高中"],
        "台北市": ["建國中學", "北一女中", "大安高工", "松山工農", "內湖高中"],
        "桃園市": ["武陵高中", "中壢高中", "桃園高中", "楊梅高中", "永豐高中"],
        "新竹市": ["新竹高中", "新竹女中", "光復中學", "建功高中"],
        "新竹縣": ["竹北高中", "關西高中", "湖口高中"],
        "苗栗縣": ["苗栗高中", "大湖農工", "苑裡高中"],
        "南投縣": ["南投高中", "中興高中", "草屯商工"],
        "宜蘭縣": ["宜蘭高中", "羅東高中", "蘇澳海事"],
        "花蓮縣": ["花蓮高中", "花蓮女中", "花蓮高工"],
        "台東縣": ["台東高中", "台東女中", "台東專科"],
        "澎湖縣": ["馬公高中", "澎湖海事"],
    }

    # 設定各縣市權重（台南和鄰近縣市較高）
    city_weights = {
        "台南市": 0.35, "高雄市": 0.20, "嘉義縣": 0.08, "嘉義市": 0.06,
        "屏東縣": 0.08, "雲林縣": 0.05, "彰化縣": 0.04, "台中市": 0.04,
        "新北市": 0.02, "台北市": 0.01, "桃園市": 0.01, "新竹市": 0.01,
        "新竹縣": 0.005, "苗栗縣": 0.005, "南投縣": 0.01, "宜蘭縣": 0.005,
        "花蓮縣": 0.005, "台東縣": 0.005, "澎湖縣": 0.005,
    }

    cities = list(city_weights.keys())
    weights = list(city_weights.values())
    weights = [w / sum(weights) for w in weights]

    stages = ["第一階段報名", "通過第一階段", "完成二階面試", "錄取", "已報到"]
    stage_probs = {
        "第一階段報名": 1.0,
        "通過第一階段": 0.75,
        "完成二階面試": 0.60,
        "錄取": 0.45,
        "已報到": 0.35,
    }

    academic_years = [111, 112, 113]
    year_weights = [0.25, 0.35, 0.40]

    records = []
    for i in range(n):
        year = np.random.choice(academic_years, p=year_weights)
        city = np.random.choice(cities, p=weights)
        school = np.random.choice(schools_by_city[city])
        dept = np.random.choice(departments)

        # 決定該學生走到哪個階段
        final_stage = "第一階段報名"
        for s in stages:
            if np.random.random() < stage_probs[s]:
                final_stage = s
            else:
                break

        is_enrolled = "是" if final_stage == "已報到" else "否"

        records.append({
            "學年度": year,
            "學生編號": f"A{year}{i:04d}",
            "報考科系": dept,
            "畢業學校": school,
            "畢業學校縣市": city,
            "住家縣市": city if np.random.random() < 0.85 else np.random.choice(cities, p=weights),
            "入學管道": "申請入學",
            "階段狀態": final_stage,
            "最終入學": is_enrolled,
        })

    return pd.DataFrame(records)


def generate_sample_retention():
    """生成在學穩定度範例資料"""
    np.random.seed(123)
    n = 800

    channels = ["申請入學", "統測分發", "繁星推薦", "技優甄審", "單獨招生", "運動績優"]
    channel_weights = [0.30, 0.25, 0.15, 0.10, 0.15, 0.05]

    departments = ["護理系", "食品科技系", "醫學檢驗生物技術系", "職業安全衛生系",
                   "環境與安全衛生工程系", "長期照護系", "幼兒保育系", "資訊管理系",
                   "化妝品應用與管理系", "餐旅管理系"]

    # 不同管道的休退學率設定（模擬真實差異）
    dropout_rates = {
        "申請入學": {"休學": 0.08, "退學": 0.05},
        "統測分發": {"休學": 0.12, "退學": 0.10},
        "繁星推薦": {"休學": 0.06, "退學": 0.03},
        "技優甄審": {"休學": 0.10, "退學": 0.07},
        "單獨招生": {"休學": 0.15, "退學": 0.12},
        "運動績優": {"休學": 0.13, "退學": 0.09},
    }

    academic_years = [108, 109, 110, 111, 112]
    semesters = [1, 2]

    records = []
    for i in range(n):
        year = np.random.choice(academic_years)
        channel = np.random.choice(channels, p=channel_weights)
        dept = np.random.choice(departments)

        rates = dropout_rates[channel]
        rand = np.random.random()
        if rand < rates["退學"]:
            status = "退學"
            dropout_sem = f"{year + np.random.randint(0, 3)}-{np.random.choice(semesters)}"
            suspend_sem = ""
        elif rand < rates["退學"] + rates["休學"]:
            status = "休學"
            dropout_sem = ""
            suspend_sem = f"{year + np.random.randint(0, 3)}-{np.random.choice(semesters)}"
        else:
            status = "在學" if year >= 111 else "畢業"
            dropout_sem = ""
            suspend_sem = ""

        records.append({
            "學年度": year,
            "學生編號": f"R{year}{i:04d}",
            "入學管道": channel,
            "入學科系": dept,
            "入學學期": f"{year}-1",
            "目前狀態": status,
            "休學學期": suspend_sem,
            "退學學期": dropout_sem,
            "休退學原因": np.random.choice(
                ["學業因素", "經濟因素", "志趣不合", "家庭因素", "健康因素", ""]
            ) if status in ["休學", "退學"] else "",
        })

    return pd.DataFrame(records)


# ============================================================
# Session State 初始化
# ============================================================
if 'applicant_data' not in st.session_state:
    st.session_state.applicant_data = None
if 'retention_data' not in st.session_state:
    st.session_state.retention_data = None


# ============================================================
# 側邊欄
# ============================================================
with st.sidebar:
    st.markdown("## 🎓 招生分析系統")
    st.markdown("**中華醫事科技大學**")
    st.markdown("入學服務處")
    st.markdown("---")

    st.markdown("### 📁 資料上傳")

    # 申請入學資料上傳
    st.markdown("#### 一、申請入學資料")
    uploaded_applicant = st.file_uploader(
        "上傳申請入學資料 (CSV/Excel)",
        type=["csv", "xlsx"],
        key="applicant_upload",
        help="欄位需包含：學年度、報考科系、畢業學校、畢業學校縣市、住家縣市、階段狀態、最終入學"
    )

    if uploaded_applicant is not None:
        try:
            if uploaded_applicant.name.endswith('.csv'):
                st.session_state.applicant_data = pd.read_csv(uploaded_applicant)
            else:
                st.session_state.applicant_data = pd.read_excel(uploaded_applicant)
            st.success(f"✅ 已載入 {len(st.session_state.applicant_data)} 筆申請資料")
        except Exception as e:
            st.error(f"❌ 檔案讀取錯誤：{e}")

    # 休退學資料上傳
    st.markdown("#### 二、在學穩定度資料")
    uploaded_retention = st.file_uploader(
        "上傳休退學資料 (CSV/Excel)",
        type=["csv", "xlsx"],
        key="retention_upload",
        help="欄位需包含：學年度、入學管道、入學科系、目前狀態"
    )

    if uploaded_retention is not None:
        try:
            if uploaded_retention.name.endswith('.csv'):
                st.session_state.retention_data = pd.read_csv(uploaded_retention)
            else:
                st.session_state.retention_data = pd.read_excel(uploaded_retention)
            st.success(f"✅ 已載入 {len(st.session_state.retention_data)} 筆在學資料")
        except Exception as e:
            st.error(f"❌ 檔案讀取錯誤：{e}")

    st.markdown("---")

    # 使用範例資料
    st.markdown("### 🧪 快速體驗")
    if st.button("📋 載入範例資料", use_container_width=True, type="primary"):
        st.session_state.applicant_data = generate_sample_applicants()
        st.session_state.retention_data = generate_sample_retention()
        st.success("✅ 已載入範例資料！")
        st.rerun()

    st.markdown("---")
    st.markdown("### ℹ️ 系統資訊")
    st.markdown("""
    - **版本**：v2.0
    - **開發**：入學服務處
    - **更新**：2024年12月
    """)


# ============================================================
# 主頁面
# ============================================================
st.markdown('<div class="main-header">🎓 中華醫事科技大學 招生數據分析系統</div>', unsafe_allow_html=True)

# 檢查是否有資料
if st.session_state.applicant_data is None and st.session_state.retention_data is None:
    st.markdown("""
    <div class="info-box">
        <h3>👋 歡迎使用招生數據分析系統</h3>
        <p>本系統提供以下分析功能：</p>
        <ol>
            <li><strong>地圖分布分析</strong>：以台灣地圖呈現學生來源分布</li>
            <li><strong>招生漏斗分析</strong>：追蹤「報名→二階→錄取→報到」各階段轉換</li>
            <li><strong>來源學校分析</strong>：評估各校學生最終入學轉換率</li>
            <li><strong>在學穩定度分析</strong>：比較不同入學管道的休退學情況</li>
        </ol>
        <p>👈 請從左側上傳資料，或點擊 <strong>「載入範例資料」</strong> 快速體驗系統功能。</p>
    </div>
    """, unsafe_allow_html=True)

    # 顯示所需資料格式說明
    st.markdown("---")
    st.markdown("### 📋 資料格式說明")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 申請入學資料欄位")
        st.dataframe(pd.DataFrame({
            "欄位名稱": ["學年度", "學生編號", "報考科系", "畢業學校", "畢業學校縣市",
                        "住家縣市", "入學管道", "階段狀態", "最終入學"],
            "說明": ["如：113", "匿名編號", "本校科系名稱", "高中職校名", "縣市名稱",
                    "縣市名稱", "申請入學/統測分發等", "第一階段報名/通過/錄取/已報到", "是/否"],
            "範例": ["113", "A1130001", "護理系", "台南高工", "台南市",
                    "台南市", "申請入學", "已報到", "是"]
        }), use_container_width=True)

    with col2:
        st.markdown("#### 在學穩定度資料欄位")
        st.dataframe(pd.DataFrame({
            "欄位名稱": ["學年度", "學生編號", "入學管道", "入學科系", "目前狀態",
                        "休學學期", "退學學期", "休退學原因"],
            "說明": ["入學學年", "匿名編號", "入學管道", "就讀科系", "在學/休學/退學/畢業",
                    "如：111-2", "如：112-1", "原因說明"],
            "範例": ["110", "R1100001", "申請入學", "護理系", "在學",
                    "", "", ""]
        }), use_container_width=True)

    st.stop()


# ============================================================
# 有資料時顯示分析頁面
# ============================================================
tabs = st.tabs([
    "📊 總覽儀表板",
    "🗺️ 地圖分布分析",
    "🔽 招生漏斗分析",
    "🏫 來源學校分析",
    "📉 在學穩定度分析",
    "📥 資料檢視與匯出"
])


# ============================================================
# TAB 0: 總覽儀表板
# ============================================================
with tabs[0]:
    st.markdown("## 📊 招生總覽儀表板")

    if st.session_state.applicant_data is not None:
        df = st.session_state.applicant_data

        # 學年度篩選
        years = sorted(df['學年度'].unique())
        selected_year_overview = st.selectbox("選擇學年度", years, index=len(years)-1, key="overview_year")
        df_year = df[df['學年度'] == selected_year_overview]

        # KPI 指標
        total_applicants = len(df_year)
        enrolled = len(df_year[df_year['最終入學'] == '是'])
        enrollment_rate = (enrolled / total_applicants * 100) if total_applicants > 0 else 0
        school_count = df_year['畢業學校'].nunique()
        city_count = df_year['畢業學校縣市'].nunique()
        dept_count = df_year['報考科系'].nunique()

        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("📝 總申請人數", f"{total_applicants:,}")
        with col2:
            st.metric("✅ 最終入學人數", f"{enrolled:,}")
        with col3:
            st.metric("📈 入學轉換率", f"{enrollment_rate:.1f}%")
        with col4:
            st.metric("🏫 來源學校數", f"{school_count}")
        with col5:
            st.metric("🗺️ 來源縣市數", f"{city_count}")

        st.markdown("---")

        col_left, col_right = st.columns(2)

        with col_left:
            st.markdown("### 📊 各科系申請人數")
            import plotly.express as px

            dept_stats = df_year.groupby('報考科系').agg(
                申請人數=('學生編號', 'count'),
                入學人數=('最終入學', lambda x: (x == '是').sum())
            ).reset_index()
            dept_stats['轉換率'] = (dept_stats['入學人數'] / dept_stats['申請人數'] * 100).round(1)
            dept_stats = dept_stats.sort_values('申請人數', ascending=True)

            fig = px.bar(
                dept_stats, y='報考科系', x=['申請人數', '入學人數'],
                orientation='h', barmode='group',
                color_discrete_sequence=['#667eea', '#28a745'],
                title=f"{selected_year_overview} 學年度各科系申請與入學人數"
            )
            fig.update_layout(height=500, legend_title="", yaxis_title="", xaxis_title="人數")
            st.plotly_chart(fig, use_container_width=True)

        with col_right:
            st.markdown("### 🗺️ 來源縣市分布")
            city_stats = df_year['畢業學校縣市'].value_counts().reset_index()
            city_stats.columns = ['縣市', '人數']
            city_stats['佔比'] = (city_stats['人數'] / city_stats['人數'].sum() * 100).round(1)

            fig = px.pie(
                city_stats.head(10), values='人數', names='縣市',
                title=f"{selected_year_overview} 學年度學生來源縣市分布 (前10)",
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            fig.update_traces(textposition='inside', textinfo='percent+label')
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)

        # 歷年趨勢
        if len(years) > 1:
            st.markdown("### 📈 歷年申請與入學趨勢")
            yearly_stats = df.groupby('學年度').agg(
                申請人數=('學生編號', 'count'),
                入學人數=('最終入學', lambda x: (x == '是').sum())
            ).reset_index()
            yearly_stats['轉換率(%)'] = (yearly_stats['入學人數'] / yearly_stats['申請人數'] * 100).round(1)

            fig = px.line(
                yearly_stats, x='學年度', y=['申請人數', '入學人數'],
                markers=True, title="歷年申請與入學人數趨勢",
                color_discrete_sequence=['#667eea', '#28a745']
            )

            # 加上轉換率在次要 Y 軸
            import plotly.graph_objects as go
            fig.add_trace(go.Scatter(
                x=yearly_stats['學年度'], y=yearly_stats['轉換率(%)'],
                mode='lines+markers+text', name='轉換率(%)',
                yaxis='y2', line=dict(color='#E8792F', dash='dash'),
                text=yearly_stats['轉換率(%)'].apply(lambda x: f'{x}%'),
                textposition='top center'
            ))
            fig.update_layout(
                yaxis2=dict(title='轉換率(%)', overlaying='y', side='right', range=[0, 100]),
                height=400
            )
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("請上傳申請入學資料以檢視總覽儀表板")


# ============================================================
# TAB 1: 地圖分布分析
# ============================================================
with tabs[1]:
    st.markdown("## 🗺️ 台灣地圖 - 學生來源分布")

    if st.session_state.applicant_data is not None:
        df = st.session_state.applicant_data

        col_filter1, col_filter2, col_filter3 = st.columns(3)
        with col_filter1:
            years = sorted(df['學年度'].unique())
            sel_year_map = st.selectbox("學年度", years, index=len(years)-1, key="map_year")
        with col_filter2:
            dept_options = ["全部科系"] + sorted(df['報考科系'].unique().tolist())
            sel_dept_map = st.selectbox("科系篩選", dept_options, key="map_dept")
        with col_filter3:
            stage_options = ["全部階段", "已報到（入學）"]
            sel_stage = st.selectbox("階段篩選", stage_options, key="map_stage")

        # 篩選資料
        df_filtered = df[df['學年度'] == sel_year_map].copy()
        if sel_dept_map != "全部科系":
            df_filtered = df_filtered[df_filtered['報考科系'] == sel_dept_map]
        if sel_stage == "已報到（入學）":
            df_filtered = df_filtered[df_filtered['最終入學'] == '是']

        if len(df_filtered) > 0:
            # 使用 Folium 繪製地圖
            import folium
            from streamlit_folium import st_folium

            # 中華醫事科大位置
            hwu_lat, hwu_lon = 23.0048, 120.2210

            m = folium.Map(
                location=[23.5, 120.9],
                zoom_start=8,
                tiles='CartoDB positron'
            )

            # 標記中華醫事科大
            folium.Marker(
                [hwu_lat, hwu_lon],
                popup="中華醫事科技大學",
                tooltip="中華醫事科技大學",
                icon=folium.Icon(color='red', icon='university', prefix='fa')
            ).add_to(m)

            # 依照住家縣市統計
            city_counts = df_filtered['住家縣市'].value_counts().to_dict()

            # 已入學的統計
            enrolled_by_city = df_filtered[df_filtered['最終入學'] == '是']['住家縣市'].value_counts().to_dict()

            max_count = max(city_counts.values()) if city_counts else 1

            for city, count in city_counts.items():
                if city in TAIWAN_COUNTY_COORDS:
                    lat, lon = TAIWAN_COUNTY_COORDS[city]
                    enrolled = enrolled_by_city.get(city, 0)
                    rate = (enrolled / count * 100) if count > 0 else 0

                    # 圓圈大小根據人數
                    radius = max(8, min(40, count / max_count * 40))

                    # 顏色根據轉換率
                    if rate >= 40:
                        color = '#28a745'
                    elif rate >= 25:
                        color = '#ffc107'
                    else:
                        color = '#dc3545'

                    folium.CircleMarker(
                        location=[lat, lon],
                        radius=radius,
                        popup=folium.Popup(
                            f"""<div style='font-family: Microsoft JhengHei; width:200px;'>
                                <h4>{city}</h4>
                                <p>📝 申請人數：<b>{count}</b></p>
                                <p>✅ 入學人數：<b>{enrolled}</b></p>
                                <p>📈 轉換率：<b>{rate:.1f}%</b></p>
                            </div>""",
                            max_width=250
                        ),
                        tooltip=f"{city}：{count}人（入學{enrolled}人）",
                        color=color,
                        fill=True,
                        fillColor=color,
                        fillOpacity=0.6,
                        weight=2
                    ).add_to(m)

                    # 加上人數標籤
                    folium.Marker(
                        [lat, lon],
                        icon=folium.DivIcon(
                            html=f'<div style="font-size:11px;font-weight:bold;text-align:center;color:#333;">{count}</div>',
                            icon_size=(30, 15),
                            icon_anchor=(15, 7)
                        )
                    ).add_to(m)

            # 圖例
            legend_html = """
            <div style="position: fixed; bottom: 50px; left: 50px; z-index: 1000;
                        background: white; padding: 15px; border-radius: 8px;
                        border: 2px solid #ccc; font-size: 13px; font-family: Microsoft JhengHei;">
                <b>🎯 轉換率圖例</b><br>
                <i style="background:#28a745;width:12px;height:12px;display:inline-block;border-radius:50%;"></i> ≥40% 高轉換<br>
                <i style="background:#ffc107;width:12px;height:12px;display:inline-block;border-radius:50%;"></i> 25-40% 中轉換<br>
                <i style="background:#dc3545;width:12px;height:12px;display:inline-block;border-radius:50%;"></i> <25% 低轉換<br>
                <br><b>⭕ 圓圈大小 = 申請人數</b>
            </div>
            """
            m.get_root().html.add_child(folium.Element(legend_html))

            st_folium(m, width=1200, height=650)

            # 地圖下方的統計表
            st.markdown("### 📋 各縣市詳細統計")
            city_detail = df_filtered.groupby('住家縣市').agg(
                申請人數=('學生編號', 'count'),
                入學人數=('最終入學', lambda x: (x == '是').sum()),
                來源學校數=('畢業學校', 'nunique'),
                報考科系數=('報考科系', 'nunique')
            ).reset_index()
            city_detail['轉換率(%)'] = (city_detail['入學人數'] / city_detail['申請人數'] * 100).round(1)
            city_detail = city_detail.sort_values('申請人數', ascending=False)

            st.dataframe(
                city_detail.style.background_gradient(subset=['申請人數'], cmap='Blues')
                    .background_gradient(subset=['轉換率(%)'], cmap='RdYlGn'),
                use_container_width=True,
                height=400
            )
        else:
            st.warning("篩選條件下無資料")
    else:
        st.info("請上傳申請入學資料以進行地圖分析")


# ============================================================
# TAB 2: 招生漏斗分析
# ============================================================
with tabs[2]:
    st.markdown("## 🔽 招生漏斗分析")
    st.markdown("追蹤學生從「第一階段報名」到「最終報到入學」各階段的轉換與流失情況")

    if st.session_state.applicant_data is not None:
        df = st.session_state.applicant_data

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            years = sorted(df['學年度'].unique())
            sel_year_funnel = st.selectbox("學年度", years, index=len(years)-1, key="funnel_year")
        with col_f2:
            dept_options = ["全部科系"] + sorted(df['報考科系'].unique().tolist())
            sel_dept_funnel = st.selectbox("科系篩選", dept_options, key="funnel_dept")

        df_f = df[df['學年度'] == sel_year_funnel].copy()
        if sel_dept_funnel != "全部科系":
            df_f = df_f[df_f['報考科系'] == sel_dept_funnel]

        if len(df_f) > 0:
            # 定義階段順序
            stage_order = ["第一階段報名", "通過第一階段", "完成二階面試", "錄取", "已報到"]
            stage_hierarchy = {s: i for i, s in enumerate(stage_order)}

            df_f['階段序號'] = df_f['階段狀態'].map(stage_hierarchy)

            # 計算各階段人數（累積方式：到達該階段及以上的人數）
            funnel_data = []
            for i, stage in enumerate(stage_order):
                count = len(df_f[df_f['階段序號'] >= i])
                funnel_data.append({
                    '階段': stage,
                    '人數': count,
                    '佔比': (count / len(df_f) * 100) if len(df_f) > 0 else 0
                })

            funnel_df = pd.DataFrame(funnel_data)

            # 漏斗圖
            import plotly.graph_objects as go

            fig_funnel = go.Figure(go.Funnel(
                y=funnel_df['階段'],
                x=funnel_df['人數'],
                textinfo="value+percent initial",
                textposition="inside",
                marker=dict(
                    color=['#667eea', '#7c8cf5', '#a3b1ff', '#48bb78', '#28a745']
                ),
                connector=dict(line=dict(color="#ccc", width=2))
            ))
            fig_funnel.update_layout(
                title=f"{sel_year_funnel} 學年度 {sel_dept_funnel} 招生漏斗",
                height=500,
                font=dict(size=14)
            )
            st.plotly_chart(fig_funnel, use_container_width=True)

            # 各階段轉換率
            st.markdown("### 📊 各階段轉換率分析")
            conversion_data = []
            for i in range(1, len(funnel_df)):
                prev = funnel_df.iloc[i-1]['人數']
                curr = funnel_df.iloc[i]['人數']
                rate = (curr / prev * 100) if prev > 0 else 0
                loss = prev - curr
                conversion_data.append({
                    '從': funnel_df.iloc[i-1]['階段'],
                    '到': funnel_df.iloc[i]['階段'],
                    '前一階段人數': prev,
                    '到達人數': curr,
                    '流失人數': loss,
                    '轉換率(%)': round(rate, 1),
                    '流失率(%)': round(100 - rate, 1)
                })

            conv_df = pd.DataFrame(conversion_data)
            st.dataframe(
                conv_df.style.background_gradient(subset=['轉換率(%)'], cmap='RdYlGn')
                    .background_gradient(subset=['流失率(%)'], cmap='RdYlGn_r'),
                use_container_width=True
            )

            # 各科系漏斗比較
            if sel_dept_funnel == "全部科系":
                st.markdown("### 🏫 各科系轉換率比較")
                dept_funnel = []
                for dept in df_f['報考科系'].unique():
                    df_dept = df_f[df_f['報考科系'] == dept]
                    total = len(df_dept)
                    enrolled = len(df_dept[df_dept['階段序號'] >= 4])
                    rate = (enrolled / total * 100) if total > 0 else 0
                    dept_funnel.append({
                        '科系': dept,
                        '申請人數': total,
                        '最終報到': enrolled,
                        '總轉換率(%)': round(rate, 1)
                    })

                dept_funnel_df = pd.DataFrame(dept_funnel).sort_values('總轉換率(%)', ascending=True)

                fig_dept = px.bar(
                    dept_funnel_df, y='科系', x='總轉換率(%)',
                    orientation='h', text='總轉換率(%)',
                    color='總轉換率(%)',
                    color_continuous_scale='RdYlGn',
                    title="各科系最終報到轉換率"
                )
                fig_dept.update_traces(textposition='outside', texttemplate='%{text}%')
                fig_dept.update_layout(height=500, yaxis_title="")
                st.plotly_chart(fig_dept, use_container_width=True)
        else:
            st.warning("篩選條件下無資料")
    else:
        st.info("請上傳申請入學資料")


# ============================================================
# TAB 3: 來源學校分析
# ============================================================
with tabs[3]:
    st.markdown("## 🏫 來源學校轉換率分析")
    st.markdown("評估哪些學校的學生最終來本校就讀的機會最高")

    if st.session_state.applicant_data is not None:
        df = st.session_state.applicant_data

        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            years = sorted(df['學年度'].unique())
            sel_year_school = st.selectbox("學年度", ["全部學年度"] + [str(y) for y in years], key="school_year")
        with col_s2:
            dept_options = ["全部科系"] + sorted(df['報考科系'].unique().tolist())
            sel_dept_school = st.selectbox("科系篩選", dept_options, key="school_dept")
        with col_s3:
            min_applicants = st.slider("最低申請人數門檻", 1, 20, 3, key="min_app")

        df_s = df.copy()
        if sel_year_school != "全部學年度":
            df_s = df_s[df_s['學年度'] == int(sel_year_school)]
        if sel_dept_school != "全部科系":
            df_s = df_s[df_s['報考科系'] == sel_dept_school]

        if len(df_s) > 0:
            # 學校層級統計
            school_stats = df_s.groupby(['畢業學校', '畢業學校縣市']).agg(
                申請人數=('學生編號', 'count'),
                入學人數=('最終入學', lambda x: (x == '是').sum()),
                報考科系數=('報考科系', 'nunique'),
                主要報考科系=('報考科系', lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else ''),
            ).reset_index()

            school_stats['轉換率(%)'] = (school_stats['入學人數'] / school_stats['申請人數'] * 100).round(1)
            school_stats = school_stats[school_stats['申請人數'] >= min_applicants]
            school_stats = school_stats.sort_values('入學人數', ascending=False)

            # 重點學校指標
            col_k1, col_k2, col_k3, col_k4 = st.columns(4)
            with col_k1:
                st.metric("🏫 來源學校總數", len(school_stats))
            with col_k2:
                top_school = school_stats.iloc[0] if len(school_stats) > 0 else None
                st.metric("🥇 入學最多學校",
                         top_school['畢業學校'] if top_school is not None else "N/A")
            with col_k3:
                high_rate = school_stats[school_stats['轉換率(%)'] >= 50]
                st.metric("⭐ 高轉換率學校數", f"{len(high_rate)} 校 (≥50%)")
            with col_k4:
                avg_rate = school_stats['轉換率(%)'].mean()
                st.metric("📊 平均轉換率", f"{avg_rate:.1f}%")

            st.markdown("---")

            col_chart1, col_chart2 = st.columns(2)

            with col_chart1:
                st.markdown("### 📊 入學人數 Top 15 學校")
                top15 = school_stats.head(15).sort_values('入學人數', ascending=True)

                fig = go.Figure()
                fig.add_trace(go.Bar(
                    y=top15['畢業學校'], x=top15['申請人數'],
                    name='申請人數', orientation='h',
                    marker_color='#667eea', opacity=0.7
                ))
                fig.add_trace(go.Bar(
                    y=top15['畢業學校'], x=top15['入學人數'],
                    name='入學人數', orientation='h',
                    marker_color='#28a745'
                ))
                fig.update_layout(barmode='overlay', height=550, yaxis_title="",
                                 xaxis_title="人數", legend=dict(x=0.7, y=0.1))
                st.plotly_chart(fig, use_container_width=True)

            with col_chart2:
                st.markdown("### 📈 轉換率 Top 15 學校")
                top_rate = school_stats.sort_values('轉換率(%)', ascending=True).tail(15)

                fig = px.bar(
                    top_rate, y='畢業學校', x='轉換率(%)',
                    orientation='h', text='轉換率(%)',
                    color='轉換率(%)', color_continuous_scale='RdYlGn'
                )
                fig.update_traces(textposition='outside', texttemplate='%{text}%')
                fig.update_layout(height=550, yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)

            # 詳細表格
            st.markdown("### 📋 來源學校完整統計")

            # 加入評級
            def rating(row):
                if row['轉換率(%)'] >= 50 and row['入學人數'] >= 3:
                    return "⭐⭐⭐ 重點經營"
                elif row['轉換率(%)'] >= 30 and row['入學人數'] >= 2:
                    return "⭐⭐ 持續關注"
                elif row['入學人數'] >= 1:
                    return "⭐ 一般往來"
                else:
                    return "🔸 待開發"

            school_stats['經營建議'] = school_stats.apply(rating, axis=1)

            st.dataframe(
                school_stats.style
                    .background_gradient(subset=['轉換率(%)'], cmap='RdYlGn')
                    .background_gradient(subset=['入學人數'], cmap='Blues'),
                use_container_width=True,
                height=500
            )

            # 歷年比較（如果選了全部學年度）
            if sel_year_school == "全部學年度" and len(years) > 1:
                st.markdown("### 📈 重點學校歷年入學趨勢")
                top_schools = school_stats.head(10)['畢業學校'].tolist()

                yearly_school = df_s[df_s['畢業學校'].isin(top_schools)].groupby(
                    ['學年度', '畢業學校']
                ).agg(
                    入學人數=('最終入學', lambda x: (x == '是').sum())
                ).reset_index()

                fig = px.line(
                    yearly_school, x='學年度', y='入學人數',
                    color='畢業學校', markers=True,
                    title="重點來源學校歷年入學人數趨勢"
                )
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("篩選條件下無資料")
    else:
        st.info("請上傳申請入學資料")


# ============================================================
# TAB 4: 在學穩定度分析
# ============================================================
with tabs[4]:
    st.markdown("## 📉 在學穩定度分析")
    st.markdown("比較不同入學管道學生的休學、退學情況，評估各管道學生穩定度")

    if st.session_state.retention_data is not None:
        df_r = st.session_state.retention_data.copy()

        col_r1, col_r2 = st.columns(2)
        with col_r1:
            r_years = sorted(df_r['學年度'].unique())
            sel_year_ret = st.selectbox("入學學年度", ["全部學年度"] + [str(y) for y in r_years], key="ret_year")
        with col_r2:
            r_depts = ["全部科系"] + sorted(df_r['入學科系'].unique().tolist())
            sel_dept_ret = st.selectbox("科系篩選", r_depts, key="ret_dept")

        if sel_year_ret != "全部學年度":
            df_r = df_r[df_r['學年度'] == int(sel_year_ret)]
        if sel_dept_ret != "全部科系":
            df_r = df_r[df_r['入學科系'] == sel_dept_ret]

        if len(df_r) > 0:
            # 總覽指標
            total_students = len(df_r)
            active = len(df_r[df_r['目前狀態'].isin(['在學', '畢業'])])
            suspended = len(df_r[df_r['目前狀態'] == '休學'])
            dropped = len(df_r[df_r['目前狀態'] == '退學'])
            graduated = len(df_r[df_r['目前狀態'] == '畢業'])

            col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
            with col_m1:
                st.metric("👥 總人數", total_students)
            with col_m2:
                st.metric("📚 在學/畢業", active, f"{active/total_students*100:.1f}%")
            with col_m3:
                st.metric("⏸️ 休學", suspended, f"{suspended/total_students*100:.1f}%")
            with col_m4:
                st.metric("❌ 退學", dropped, f"{dropped/total_students*100:.1f}%")
            with col_m5:
                st.metric("🎓 畢業", graduated, f"{graduated/total_students*100:.1f}%")

            st.markdown("---")

            col_left, col_right = st.columns(2)

            with col_left:
                st.markdown("### 📊 各入學管道穩定度比較")
                channel_stats = df_r.groupby('入學管道').agg(
                    總人數=('學生編號', 'count'),
                    在學畢業=('目前狀態', lambda x: x.isin(['在學', '畢業']).sum()),
                    休學=('目前狀態', lambda x: (x == '休學').sum()),
                    退學=('目前狀態', lambda x: (x == '退學').sum()),
                ).reset_index()

                channel_stats['穩定率(%)'] = (channel_stats['在學畢業'] / channel_stats['總人數'] * 100).round(1)
                channel_stats['休學率(%)'] = (channel_stats['休學'] / channel_stats['總人數'] * 100).round(1)
                channel_stats['退學率(%)'] = (channel_stats['退學'] / channel_stats['總人數'] * 100).round(1)
                channel_stats = channel_stats.sort_values('穩定率(%)', ascending=True)

                fig = go.Figure()
                fig.add_trace(go.Bar(
                    y=channel_stats['入學管道'],
                    x=channel_stats['穩定率(%)'],
                    name='穩定率', orientation='h',
                    marker_color='#28a745',
                    text=channel_stats['穩定率(%)'].apply(lambda x: f'{x}%'),
                    textposition='inside'
                ))
                fig.add_trace(go.Bar(
                    y=channel_stats['入學管道'],
                    x=channel_stats['休學率(%)'],
                    name='休學率', orientation='h',
                    marker_color='#ffc107',
                    text=channel_stats['休學率(%)'].apply(lambda x: f'{x}%'),
                    textposition='inside'
                ))
                fig.add_trace(go.Bar(
                    y=channel_stats['入學管道'],
                    x=channel_stats['退學率(%)'],
                    name='退學率', orientation='h',
                    marker_color='#dc3545',
                    text=channel_stats['退學率(%)'].apply(lambda x: f'{x}%'),
                    textposition='inside'
                ))
                fig.update_layout(
                    barmode='stack', height=400,
                    xaxis_title='百分比(%)', yaxis_title='',
                    title='各入學管道學生狀態分布'
                )
                st.plotly_chart(fig, use_container_width=True)

            with col_right:
                st.markdown("### 📊 各科系穩定度比較")
                dept_stats = df_r.groupby('入學科系').agg(
                    總人數=('學生編號', 'count'),
                    休退學人數=('目前狀態', lambda x: x.isin(['休學', '退學']).sum()),
                ).reset_index()
                dept_stats['休退學率(%)'] = (dept_stats['休退學人數'] / dept_stats['總人數'] * 100).round(1)
                dept_stats = dept_stats.sort_values('休退學率(%)', ascending=True)

                fig = px.bar(
                    dept_stats, y='入學科系', x='休退學率(%)',
                    orientation='h', text='休退學率(%)',
                    color='休退學率(%)',
                    color_continuous_scale='RdYlGn_r',
                    title='各科系休退學率'
                )
                fig.update_traces(textposition='outside', texttemplate='%{text}%')
                fig.update_layout(height=400, yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)

            # 休退學原因分析
            st.markdown("### 📋 休退學原因分析")
            col_reason1, col_reason2 = st.columns(2)

            dropout_df = df_r[df_r['目前狀態'].isin(['休學', '退學'])]

            with col_reason1:
                if '休退學原因' in dropout_df.columns:
                    reason_stats = dropout_df[dropout_df['休退學原因'] != '']['休退學原因'].value_counts()
                    if len(reason_stats) > 0:
                        fig = px.pie(
                            values=reason_stats.values,
                            names=reason_stats.index,
                            title='休退學原因分布',
                            color_discrete_sequence=px.colors.qualitative.Set2
                        )
                        fig.update_traces(textposition='inside', textinfo='percent+label')
                        fig.update_layout(height=400)
                        st.plotly_chart(fig, use_container_width=True)

            with col_reason2:
                if '休退學原因' in dropout_df.columns:
                    # 各管道的休退學原因交叉分析
                    cross = pd.crosstab(
                        dropout_df['入學管道'],
                        dropout_df['休退學原因']
                    )
                    if len(cross) > 0:
                        fig = px.imshow(
                            cross, text_auto=True,
                            color_continuous_scale='YlOrRd',
                            title='入學管道 × 休退學原因 交叉分析',
                            aspect='auto'
                        )
                        fig.update_layout(height=400)
                        st.plotly_chart(fig, use_container_width=True)

            # 歷年穩定度趨勢
            if sel_year_ret == "全部學年度" and len(r_years) > 1:
                st.markdown("### 📈 歷年各管道休退學率趨勢")
                df_r_full = st.session_state.retention_data.copy()
                if sel_dept_ret != "全部科系":
                    df_r_full = df_r_full[df_r_full['入學科系'] == sel_dept_ret]

                yearly_channel = df_r_full.groupby(['學年度', '入學管道']).agg(
                    總人數=('學生編號', 'count'),
                    休退學=('目前狀態', lambda x: x.isin(['休學', '退學']).sum())
                ).reset_index()
                yearly_channel['休退學率(%)'] = (yearly_channel['休退學'] / yearly_channel['總人數'] * 100).round(1)

                fig = px.line(
                    yearly_channel, x='學年度', y='休退學率(%)',
                    color='入學管道', markers=True,
                    title='歷年各入學管道休退學率趨勢'
                )
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)

            # 詳細統計表
            st.markdown("### 📋 各入學管道完整統計")
            st.dataframe(
                channel_stats.style
                    .background_gradient(subset=['穩定率(%)'], cmap='RdYlGn')
                    .background_gradient(subset=['退學率(%)'], cmap='RdYlGn_r'),
                use_container_width=True
            )
        else:
            st.warning("篩選條件下無資料")
    else:
        st.info("請上傳在學穩定度資料")


# ============================================================
# TAB 5: 資料檢視與匯出
# ============================================================
with tabs[5]:
    st.markdown("## 📥 資料檢視與匯出")

    tab_data1, tab_data2 = st.tabs(["申請入學資料", "在學穩定度資料"])

    with tab_data1:
        if st.session_state.applicant_data is not None:
            df_view = st.session_state.applicant_data
            st.markdown(f"**共 {len(df_view)} 筆資料**")

            # 搜尋與篩選
            col_search1, col_search2, col_search3 = st.columns(3)
            with col_search1:
                search_school = st.text_input("搜尋學校名稱", key="search_school")
            with col_search2:
                filter_year = st.multiselect("篩選學年度", df_view['學年度'].unique(), key="filter_year_data")
            with col_search3:
                filter_dept = st.multiselect("篩選科系", df_view['報考科系'].unique(), key="filter_dept_data")

            df_display = df_view.copy()
            if search_school:
                df_display = df_display[df_display['畢業學校'].str.contains(search_school, na=False)]
            if filter_year:
                df_display = df_display[df_display['學年度'].isin(filter_year)]
            if filter_dept:
                df_display = df_display[df_display['報考科系'].isin(filter_dept)]

            st.dataframe(df_display, use_container_width=True, height=400)

            # 匯出功能
            csv = df_display.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 下載篩選後資料 (CSV)",
                data=csv,
                file_name="applicant_data_filtered.csv",
                mime="text/csv"
            )
        else:
            st.info("尚未載入申請入學資料")

    with tab_data2:
        if st.session_state.retention_data is not None:
            df_view2 = st.session_state.retention_data
            st.markdown(f"**共 {len(df_view2)} 筆資料**")
            st.dataframe(df_view2, use_container_width=True, height=400)

            csv2 = df_view2.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 下載在學穩定度資料 (CSV)",
                data=csv2,
                file_name="retention_data.csv",
                mime="text/csv"
            )
        else:
            st.info("尚未載入在學穩定度資料")


# ============================================================
# 頁尾
# ============================================================
st.markdown("---")
st.markdown(
    """<div style='text-align:center; color:#999; font-size:0.85rem;'>
        中華醫事科技大學 入學服務處 招生數據分析系統 v2.0 | 
        © 2024 Chung Hwa University of Medical Technology
    </div>""",
    unsafe_allow_html=True
)
