import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

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
    .info-box {
        background-color: #f0f7ff;
        border-left: 5px solid #2C5F8A;
        padding: 1rem 1.5rem;
        border-radius: 0 8px 8px 0;
        margin: 1rem 0;
    }
    .file-badge {
        display: inline-block;
        background: #e8f5e9;
        border: 1px solid #81c784;
        border-radius: 12px;
        padding: 2px 10px;
        margin: 2px 4px;
        font-size: 0.82rem;
    }
    .merge-info {
        background: #fff3e0;
        border-left: 4px solid #ff9800;
        padding: 0.8rem 1.2rem;
        border-radius: 0 8px 8px 0;
        margin: 0.5rem 0;
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 台灣縣市座標
# ============================================================
TAIWAN_COUNTY_COORDS = {
    "台北市": (25.0330, 121.5654), "臺北市": (25.0330, 121.5654),
    "新北市": (25.0120, 121.4650), "桃園市": (24.9936, 121.3010),
    "台中市": (24.1477, 120.6736), "臺中市": (24.1477, 120.6736),
    "台南市": (22.9998, 120.2269), "臺南市": (22.9998, 120.2269),
    "高雄市": (22.6273, 120.3014), "基隆市": (25.1276, 121.7392),
    "新竹市": (24.8138, 120.9675), "新竹縣": (24.8387, 121.0178),
    "苗栗縣": (24.5602, 120.8214), "彰化縣": (24.0518, 120.5161),
    "南投縣": (23.9610, 120.9718), "雲林縣": (23.7092, 120.4313),
    "嘉義市": (23.4801, 120.4491), "嘉義縣": (23.4518, 120.2551),
    "屏東縣": (22.5519, 120.5487), "宜蘭縣": (24.7570, 121.7533),
    "花蓮縣": (23.9872, 121.6016),
    "台東縣": (22.7583, 121.1444), "臺東縣": (22.7583, 121.1444),
    "澎湖縣": (23.5711, 119.5793),
    "金門縣": (24.4493, 118.3767), "連江縣": (26.1505, 119.9499),
}

# ============================================================
# 必要欄位定義（用來驗證上傳檔案）
# ============================================================
APPLICANT_REQUIRED_COLS = ["學年度", "學生編號", "報考科系", "畢業學校", "畢業學校縣市",
                           "住家縣市", "入學管道", "階段狀態", "最終入學"]
RETENTION_REQUIRED_COLS = ["學年度", "學生編號", "入學管道", "入學科系", "目前狀態"]


# ============================================================
# 檔案讀取 + 驗證 工具函式
# ============================================================
def read_uploaded_file(uploaded_file):
    """讀取 CSV 或 Excel，回傳 DataFrame"""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        return df
    except Exception as e:
        st.error(f"❌ 無法讀取 {uploaded_file.name}：{e}")
        return None


def validate_columns(df, required_cols, file_name):
    """檢查必要欄位是否存在"""
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"❌ {file_name} 缺少欄位：{', '.join(missing)}")
        return False
    return True


def merge_dataframes(existing_df, new_df, key_col="學生編號"):
    """
    合併資料：
    - 新增不重複的列
    - 重複的學生編號以新檔案為準（更新）
    """
    if existing_df is None or len(existing_df) == 0:
        return new_df.copy()

    # 移除舊資料中與新檔案重複的學生編號
    mask = ~existing_df[key_col].isin(new_df[key_col])
    kept = existing_df[mask]
    merged = pd.concat([kept, new_df], ignore_index=True)
    return merged


# ============================================================
# 生成範例資料
# ============================================================
@st.cache_data
def generate_sample_applicants():
    np.random.seed(42)
    n_per_year = {111: 150, 112: 180, 113: 200}

    departments = [
        "護理系", "食品科技系", "醫學檢驗生物技術系", "職業安全衛生系",
        "環境與安全衛生工程系", "生物醫學工程系", "藥學系", "長期照護系",
        "幼兒保育系", "資訊管理系", "化妝品應用與管理系", "餐旅管理系"
    ]

    schools_by_city = {
        "台南市": ["台南高工", "台南女中", "長榮高中", "南光高中", "港明高中",
                   "新營高工", "北門高中", "曾文高中", "南科實中", "家齊高中",
                   "後壁高中", "白河商工", "新化高中", "善化高中"],
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
        "新竹市": ["新竹高中", "新竹女中", "光復中學"],
        "新竹縣": ["竹北高中", "關西高中", "湖口高中"],
        "苗栗縣": ["苗栗高中", "大湖農工", "苑裡高中"],
        "南投縣": ["南投高中", "中興高中", "草屯商工"],
        "宜蘭縣": ["宜蘭高中", "羅東高中", "蘇澳海事"],
        "花蓮縣": ["花蓮高中", "花蓮女中", "花蓮高工"],
        "台東縣": ["台東高中", "台東女中"],
        "澎湖縣": ["馬公高中", "澎湖海事"],
    }

    city_weights = {
        "台南市": 0.35, "高雄市": 0.20, "嘉義縣": 0.08, "嘉義市": 0.06,
        "屏東縣": 0.08, "雲林縣": 0.05, "彰化縣": 0.04, "台中市": 0.04,
        "新北市": 0.02, "台北市": 0.01, "桃園市": 0.01, "新竹市": 0.01,
        "新竹縣": 0.005, "苗栗縣": 0.005, "南投縣": 0.01, "宜蘭縣": 0.005,
        "花蓮縣": 0.005, "台東縣": 0.005, "澎湖縣": 0.005,
    }

    cities = list(city_weights.keys())
    weights = list(city_weights.values())
    total_w = sum(weights)
    weights = [w / total_w for w in weights]

    stages = ["第一階段報名", "通過第一階段", "完成二階面試", "錄取", "已報到"]
    stage_cumulative_prob = [1.0, 0.75, 0.60, 0.45, 0.35]

    records = []
    idx = 0
    for year, n in n_per_year.items():
        for i in range(n):
            city = np.random.choice(cities, p=weights)
            school = np.random.choice(schools_by_city[city])
            dept = np.random.choice(departments)

            final_stage_idx = 0
            for si in range(len(stages)):
                if np.random.random() < stage_cumulative_prob[si]:
                    final_stage_idx = si
                else:
                    break

            final_stage = stages[final_stage_idx]
            is_enrolled = "是" if final_stage == "已報到" else "否"
            home_city = city if np.random.random() < 0.85 else np.random.choice(cities, p=weights)

            records.append({
                "學年度": year,
                "學生編號": f"A{year}{idx:04d}",
                "報考科系": dept,
                "畢業學校": school,
                "畢業學校縣市": city,
                "住家縣市": home_city,
                "入學管道": "申請入學",
                "階段狀態": final_stage,
                "最終入學": is_enrolled,
            })
            idx += 1

    return pd.DataFrame(records)


@st.cache_data
def generate_sample_retention():
    np.random.seed(123)

    channels = ["申請入學", "統測分發", "繁星推薦", "技優甄審", "單獨招生", "運動績優"]
    channel_weights = [0.30, 0.25, 0.15, 0.10, 0.15, 0.05]
    departments = [
        "護理系", "食品科技系", "醫學檢驗生物技術系", "職業安全衛生系",
        "環境與安全衛生工程系", "長期照護系", "幼兒保育系", "資訊管理系",
        "化妝品應用與管理系", "餐旅管理系"
    ]
    dropout_rates = {
        "申請入學": {"休學": 0.08, "退學": 0.05},
        "統測分發": {"休學": 0.12, "退學": 0.10},
        "繁星推薦": {"休學": 0.06, "退學": 0.03},
        "技優甄審": {"休學": 0.10, "退學": 0.07},
        "單獨招生": {"休學": 0.15, "退學": 0.12},
        "運動績優": {"休學": 0.13, "退學": 0.09},
    }
    reasons = ["學業因素", "經濟因素", "志趣不合", "家庭因素", "健康因素"]
    n_per_year = {108: 150, 109: 160, 110: 170, 111: 180, 112: 160}

    records = []
    idx = 0
    for year, n in n_per_year.items():
        for i in range(n):
            channel = np.random.choice(channels, p=channel_weights)
            dept = np.random.choice(departments)
            rates = dropout_rates[channel]
            rand_val = np.random.random()

            if rand_val < rates["退學"]:
                status = "退學"
                reason = np.random.choice(reasons)
            elif rand_val < rates["退學"] + rates["休學"]:
                status = "休學"
                reason = np.random.choice(reasons)
            else:
                status = "在學" if year >= 111 else "畢業"
                reason = ""

            records.append({
                "學年度": year,
                "學生編號": f"R{year}{idx:04d}",
                "入學管道": channel,
                "入學科系": dept,
                "入學學期": f"{year}-1",
                "目前狀態": status,
                "休學學期": f"{year + np.random.randint(0, 3)}-{np.random.choice([1, 2])}" if status == "休學" else "",
                "退學學期": f"{year + np.random.randint(0, 3)}-{np.random.choice([1, 2])}" if status == "退學" else "",
                "休退學原因": reason,
            })
            idx += 1

    return pd.DataFrame(records)


# ============================================================
# Session State 初始化
# ============================================================
if 'applicant_data' not in st.session_state:
    st.session_state.applicant_data = None
if 'retention_data' not in st.session_state:
    st.session_state.retention_data = None
# ★ 新增：追蹤已匯入的檔案清單
if 'applicant_files' not in st.session_state:
    st.session_state.applicant_files = []  # [{"name": ..., "rows": ..., "years": ...}]
if 'retention_files' not in st.session_state:
    st.session_state.retention_files = []
# ★ 新增：追蹤上傳計數器，用於重置 file_uploader
if 'upload_counter' not in st.session_state:
    st.session_state.upload_counter = 0


# ============================================================
# 側邊欄
# ============================================================
with st.sidebar:
    st.markdown("## 🎓 招生分析系統")
    st.markdown("**中華醫事科技大學**")
    st.markdown("入學服務處")
    st.markdown("---")

    # ============================================================
    # ★★★ 核心修改：多檔案上傳 + 累積合併 ★★★
    # ============================================================
    st.markdown("### 📁 資料上傳")
    st.caption("💡 可一次選擇多個檔案，也可多次上傳，資料會自動累積合併")

    # ── 申請入學資料 ──
    st.markdown("#### 一、申請入學資料")
    uploaded_applicants = st.file_uploader(
        "上傳申請入學資料 (CSV/Excel)",
        type=["csv", "xlsx"],
        accept_multiple_files=True,       # ★ 允許多檔
        key=f"applicant_upload_{st.session_state.upload_counter}"
    )

    if uploaded_applicants:
        new_count = 0
        for f in uploaded_applicants:
            # 避免重複匯入同一檔名
            already = [rec['name'] for rec in st.session_state.applicant_files]
            if f.name in already:
                continue

            df_new = read_uploaded_file(f)
            if df_new is not None and validate_columns(df_new, APPLICANT_REQUIRED_COLS, f.name):
                st.session_state.applicant_data = merge_dataframes(
                    st.session_state.applicant_data, df_new, key_col="學生編號"
                )
                years_in_file = sorted(df_new['學年度'].unique().tolist())
                st.session_state.applicant_files.append({
                    "name": f.name,
                    "rows": len(df_new),
                    "years": years_in_file
                })
                new_count += 1

        if new_count > 0:
            st.success(f"✅ 新增 {new_count} 個檔案，目前共 {len(st.session_state.applicant_data)} 筆申請資料")

    # 顯示已匯入檔案
    if st.session_state.applicant_files:
        st.markdown("**已匯入檔案：**")
        for rec in st.session_state.applicant_files:
            yrs = ", ".join(str(y) for y in rec['years'])
            st.markdown(
                f'<span class="file-badge">📄 {rec["name"]}（{rec["rows"]}筆，{yrs}學年）</span>',
                unsafe_allow_html=True
            )
        total_rows = len(st.session_state.applicant_data) if st.session_state.applicant_data is not None else 0
        total_years = sorted(st.session_state.applicant_data['學年度'].unique()) if st.session_state.applicant_data is not None else []
        st.markdown(
            f'<div class="merge-info">📊 合併後：<b>{total_rows}</b> 筆 ｜ 學年度：<b>{", ".join(str(y) for y in total_years)}</b></div>',
            unsafe_allow_html=True
        )

    # ── 在學穩定度資料 ──
    st.markdown("#### 二、在學穩定度資料")
    uploaded_retentions = st.file_uploader(
        "上傳休退學資料 (CSV/Excel)",
        type=["csv", "xlsx"],
        accept_multiple_files=True,       # ★ 允許多檔
        key=f"retention_upload_{st.session_state.upload_counter}"
    )

    if uploaded_retentions:
        new_count = 0
        for f in uploaded_retentions:
            already = [rec['name'] for rec in st.session_state.retention_files]
            if f.name in already:
                continue

            df_new = read_uploaded_file(f)
            if df_new is not None and validate_columns(df_new, RETENTION_REQUIRED_COLS, f.name):
                st.session_state.retention_data = merge_dataframes(
                    st.session_state.retention_data, df_new, key_col="學生編號"
                )
                years_in_file = sorted(df_new['學年度'].unique().tolist())
                st.session_state.retention_files.append({
                    "name": f.name,
                    "rows": len(df_new),
                    "years": years_in_file
                })
                new_count += 1

        if new_count > 0:
            st.success(f"✅ 新增 {new_count} 個檔案，目前共 {len(st.session_state.retention_data)} 筆在學資料")

    if st.session_state.retention_files:
        st.markdown("**已匯入檔案：**")
        for rec in st.session_state.retention_files:
            yrs = ", ".join(str(y) for y in rec['years'])
            st.markdown(
                f'<span class="file-badge">📄 {rec["name"]}（{rec["rows"]}筆，{yrs}學年）</span>',
                unsafe_allow_html=True
            )
        total_rows = len(st.session_state.retention_data) if st.session_state.retention_data is not None else 0
        total_years = sorted(st.session_state.retention_data['學年度'].unique()) if st.session_state.retention_data is not None else []
        st.markdown(
            f'<div class="merge-info">📊 合併後：<b>{total_rows}</b> 筆 ｜ 學年度：<b>{", ".join(str(y) for y in total_years)}</b></div>',
            unsafe_allow_html=True
        )

    st.markdown("---")

    # ── 清除資料按鈕 ──
    st.markdown("### 🔧 資料管理")

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("🗑️ 清除申請資料", use_container_width=True):
            st.session_state.applicant_data = None
            st.session_state.applicant_files = []
            st.session_state.upload_counter += 1
            st.rerun()
    with col_btn2:
        if st.button("🗑️ 清除在學資料", use_container_width=True):
            st.session_state.retention_data = None
            st.session_state.retention_files = []
            st.session_state.upload_counter += 1
            st.rerun()

    if st.button("🗑️ 清除全部資料", use_container_width=True, type="secondary"):
        st.session_state.applicant_data = None
        st.session_state.retention_data = None
        st.session_state.applicant_files = []
        st.session_state.retention_files = []
        st.session_state.upload_counter += 1
        st.rerun()

    st.markdown("---")

    # ── 範例資料 ──
    st.markdown("### 🧪 快速體驗")
    if st.button("📋 載入範例資料", use_container_width=True, type="primary"):
        st.session_state.applicant_data = generate_sample_applicants()
        st.session_state.retention_data = generate_sample_retention()
        st.session_state.applicant_files = [
            {"name": "範例_申請入學.csv", "rows": len(st.session_state.applicant_data),
             "years": sorted(st.session_state.applicant_data['學年度'].unique().tolist())}
        ]
        st.session_state.retention_files = [
            {"name": "範例_在學穩定度.csv", "rows": len(st.session_state.retention_data),
             "years": sorted(st.session_state.retention_data['學年度'].unique().tolist())}
        ]
        st.rerun()

    st.markdown("---")
    st.markdown("### ℹ️ 系統資訊")
    st.markdown("版本：v2.1 | 多檔合併版")


# ============================================================
# 主頁面標題
# ============================================================
st.markdown('<div class="main-header">🎓 中華醫事科技大學 招生數據分析系統</div>', unsafe_allow_html=True)

# ============================================================
# 無資料：歡迎頁面
# ============================================================
if st.session_state.applicant_data is None and st.session_state.retention_data is None:
    st.markdown("""
    <div class="info-box">
        <h3>👋 歡迎使用招生數據分析系統</h3>
        <p>本系統支援 <strong>多檔案上傳與跨學年度合併</strong>，您可以：</p>
        <ul>
            <li>📂 一次選取多個學年度檔案同時匯入</li>
            <li>📂 分批上傳不同學年度檔案，系統自動累積合併</li>
            <li>🔄 重複的學生編號會自動以新資料覆蓋</li>
        </ul>
        <h4>分析功能：</h4>
        <ol>
            <li><strong>地圖分布分析</strong>：以台灣地圖呈現學生來源分布</li>
            <li><strong>招生漏斗分析</strong>：追蹤「報名→二階→錄取→報到」各階段轉換</li>
            <li><strong>來源學校分析</strong>：評估各校學生最終入學轉換率</li>
            <li><strong>在學穩定度分析</strong>：比較不同入學管道的休退學情況</li>
        </ol>
        <p>👈 請從左側上傳資料，或點擊 <strong>「載入範例資料」</strong> 快速體驗。</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📋 資料格式說明")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 申請入學資料欄位")
        st.dataframe(pd.DataFrame({
            "欄位名稱": APPLICANT_REQUIRED_COLS,
            "範例": ["113", "A1130001", "護理系", "台南高工", "台南市",
                    "台南市", "申請入學", "已報到", "是"]
        }), use_container_width=True)
        st.info("💡 每個學年度一個檔案，例如 `111申請入學.csv`、`112申請入學.csv`")

    with col2:
        st.markdown("#### 在學穩定度資料欄位")
        st.dataframe(pd.DataFrame({
            "欄位名稱": RETENTION_REQUIRED_COLS + ["休學學期", "退學學期", "休退學原因"],
            "範例": ["110", "R1100001", "申請入學", "護理系", "在學", "", "", ""]
        }), use_container_width=True)
        st.info("💡 同樣可分學年度上傳，系統會自動合併")

    st.stop()


# ============================================================
# 有資料：分頁分析
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
        df = st.session_state.applicant_data.copy()
        years = sorted(df['學年度'].unique())

        # ★ 多學年度篩選器
        year_options = ["全部學年度合計"] + [str(y) for y in years]
        selected_year_str = st.selectbox("選擇學年度", year_options, key="ov_year")

        if selected_year_str == "全部學年度合計":
            df_year = df.copy()
            display_year = "全部學年度"
        else:
            sel_y = int(selected_year_str)
            df_year = df[df['學年度'] == sel_y]
            display_year = f"{sel_y} 學年度"

        total_app = len(df_year)
        enrolled = len(df_year[df_year['最終入學'] == '是'])
        rate = (enrolled / total_app * 100) if total_app > 0 else 0
        school_cnt = df_year['畢業學校'].nunique()
        city_cnt = df_year['畢業學校縣市'].nunique()

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("📝 總申請人數", f"{total_app:,}")
        c2.metric("✅ 最終入學", f"{enrolled:,}")
        c3.metric("📈 轉換率", f"{rate:.1f}%")
        c4.metric("🏫 來源學校數", school_cnt)
        c5.metric("🗺️ 來源縣市數", city_cnt)

        # 顯示涵蓋學年度
        st.caption(f"📅 資料涵蓋學年度：{', '.join(str(y) for y in years)}（共 {len(years)} 個學年度）")

        st.markdown("---")
        left_col, right_col = st.columns(2)

        with left_col:
            dept_stats = df_year.groupby('報考科系').agg(
                申請人數=('學生編號', 'count'),
                入學人數=('最終入學', lambda x: (x == '是').sum())
            ).reset_index().sort_values('申請人數', ascending=True)

            fig = px.bar(
                dept_stats, y='報考科系', x=['申請人數', '入學人數'],
                orientation='h', barmode='group',
                color_discrete_sequence=['#667eea', '#28a745'],
                title=f"{display_year} 各科系申請與入學人數"
            )
            fig.update_layout(height=500, yaxis_title="", xaxis_title="人數")
            st.plotly_chart(fig, use_container_width=True)

        with right_col:
            city_stats = df_year['畢業學校縣市'].value_counts().reset_index()
            city_stats.columns = ['縣市', '人數']
            fig = px.pie(
                city_stats.head(10), values='人數', names='縣市',
                title=f"{display_year} 學生來源縣市分布 (前10)",
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            fig.update_traces(textposition='inside', textinfo='percent+label')
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)

        # 歷年趨勢（僅在有多學年資料時顯示）
        if len(years) > 1:
            st.markdown("### 📈 歷年申請與入學趨勢")
            yearly = df.groupby('學年度').agg(
                申請人數=('學生編號', 'count'),
                入學人數=('最終入學', lambda x: (x == '是').sum())
            ).reset_index()
            yearly['轉換率'] = (yearly['入學人數'] / yearly['申請人數'] * 100).round(1)

            fig = go.Figure()
            fig.add_trace(go.Bar(x=yearly['學年度'].astype(str), y=yearly['申請人數'],
                                name='申請人數', marker_color='#667eea'))
            fig.add_trace(go.Bar(x=yearly['學年度'].astype(str), y=yearly['入學人數'],
                                name='入學人數', marker_color='#28a745'))
            fig.add_trace(go.Scatter(
                x=yearly['學年度'].astype(str), y=yearly['轉換率'], name='轉換率(%)',
                yaxis='y2', mode='lines+markers+text',
                line=dict(color='#E8792F', width=3),
                text=yearly['轉換率'].apply(lambda x: f'{x}%'),
                textposition='top center'
            ))
            fig.update_layout(
                barmode='group', height=400,
                yaxis=dict(title='人數'),
                yaxis2=dict(title='轉換率(%)', overlaying='y', side='right', range=[0, 100])
            )
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("請上傳申請入學資料或載入範例資料")


# ============================================================
# TAB 1: 地圖分布分析
# ============================================================
with tabs[1]:
    st.markdown("## 🗺️ 台灣地圖 - 學生來源分布")

    if st.session_state.applicant_data is not None:
        df = st.session_state.applicant_data.copy()

        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            years = sorted(df['學年度'].unique())
            map_yr_opts = ["全部學年度"] + [str(y) for y in years]
            sel_y = st.selectbox("學年度", map_yr_opts, key="map_y")
        with fc2:
            dept_opts = ["全部科系"] + sorted(df['報考科系'].unique().tolist())
            sel_d = st.selectbox("科系篩選", dept_opts, key="map_d")
        with fc3:
            stage_opts = ["全部階段（所有申請者）", "僅已報到（最終入學）"]
            sel_s = st.selectbox("階段篩選", stage_opts, key="map_s")

        df_f = df.copy()
        if sel_y != "全部學年度":
            df_f = df_f[df_f['學年度'] == int(sel_y)]
        if sel_d != "全部科系":
            df_f = df_f[df_f['報考科系'] == sel_d]
        if "僅已報到" in sel_s:
            df_f = df_f[df_f['最終入學'] == '是']

        if len(df_f) > 0:
            try:
                import folium
                from streamlit_folium import st_folium

                m = folium.Map(location=[23.5, 120.9], zoom_start=8, tiles='CartoDB positron')
                folium.Marker(
                    [23.0048, 120.2210],
                    popup="中華醫事科技大學",
                    tooltip="📍 中華醫事科技大學",
                    icon=folium.Icon(color='red', icon='star', prefix='fa')
                ).add_to(m)

                city_counts = df_f['住家縣市'].value_counts().to_dict()
                enrolled_counts = df_f[df_f['最終入學'] == '是']['住家縣市'].value_counts().to_dict()
                max_count = max(city_counts.values()) if city_counts else 1

                for city, count in city_counts.items():
                    if city in TAIWAN_COUNTY_COORDS:
                        lat, lon = TAIWAN_COUNTY_COORDS[city]
                        enr = enrolled_counts.get(city, 0)
                        r = (enr / count * 100) if count > 0 else 0
                        radius = max(8, min(40, count / max_count * 40))

                        color = '#28a745' if r >= 40 else '#ffc107' if r >= 25 else '#dc3545'

                        popup_html = f"""
                        <div style='font-family:sans-serif;width:180px;'>
                            <h4 style='margin:0;'>{city}</h4>
                            <p style='margin:4px 0;'>📝 申請人數：<b>{count}</b></p>
                            <p style='margin:4px 0;'>✅ 入學人數：<b>{enr}</b></p>
                            <p style='margin:4px 0;'>📈 轉換率：<b>{r:.1f}%</b></p>
                        </div>"""
                        folium.CircleMarker(
                            location=[lat, lon], radius=radius,
                            popup=folium.Popup(popup_html, max_width=200),
                            tooltip=f"{city}：{count}人（入學{enr}人）",
                            color=color, fill=True, fillColor=color,
                            fillOpacity=0.6, weight=2
                        ).add_to(m)

                st_folium(m, width=1100, height=600)

            except ImportError:
                city_df = df_f['住家縣市'].value_counts().reset_index()
                city_df.columns = ['縣市', '人數']
                fig = px.bar(city_df, x='縣市', y='人數', title="各縣市申請人數分布",
                            color='人數', color_continuous_scale='Blues')
                fig.update_layout(height=500)
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("### 📋 各縣市詳細統計")
            city_detail = df_f.groupby('住家縣市').agg(
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
                use_container_width=True, height=400
            )
        else:
            st.warning("篩選條件下無資料")
    else:
        st.info("請上傳申請入學資料")


# ============================================================
# TAB 2: 招生漏斗分析
# ============================================================
with tabs[2]:
    st.markdown("## 🔽 招生漏斗分析")

    if st.session_state.applicant_data is not None:
        df = st.session_state.applicant_data.copy()

        f1, f2 = st.columns(2)
        with f1:
            years = sorted(df['學年度'].unique())
            fn_yr_opts = ["全部學年度"] + [str(y) for y in years]
            sy = st.selectbox("學年度", fn_yr_opts, key="fn_y")
        with f2:
            d_opts = ["全部科系"] + sorted(df['報考科系'].unique().tolist())
            sd = st.selectbox("科系篩選", d_opts, key="fn_d")

        df_fn = df.copy()
        if sy != "全部學年度":
            df_fn = df_fn[df_fn['學年度'] == int(sy)]
        if sd != "全部科系":
            df_fn = df_fn[df_fn['報考科系'] == sd]

        if len(df_fn) > 0:
            stage_order = ["第一階段報名", "通過第一階段", "完成二階面試", "錄取", "已報到"]
            stage_map = {s: i for i, s in enumerate(stage_order)}
            df_fn['stage_idx'] = df_fn['階段狀態'].map(stage_map).fillna(0).astype(int)

            funnel_data = []
            total = len(df_fn)
            for i, stg in enumerate(stage_order):
                cnt = len(df_fn[df_fn['stage_idx'] >= i])
                funnel_data.append({'階段': stg, '人數': cnt, '佔比': round(cnt / total * 100, 1)})

            funnel_df = pd.DataFrame(funnel_data)

            fig = go.Figure(go.Funnel(
                y=funnel_df['階段'], x=funnel_df['人數'],
                textinfo="value+percent initial",
                textposition="inside",
                marker=dict(color=['#667eea', '#7c8cf5', '#a3b1ff', '#48bb78', '#28a745']),
                connector=dict(line=dict(color="#ccc", width=2))
            ))
            fig.update_layout(title=f"招生漏斗 ({sy} {sd})", height=500, font=dict(size=14))
            st.plotly_chart(fig, use_container_width=True)

            st.markdown("### 📊 各階段轉換率")
            conv_rows = []
            for i in range(1, len(funnel_df)):
                prev = funnel_df.iloc[i - 1]['人數']
                curr = funnel_df.iloc[i]['人數']
                cr = (curr / prev * 100) if prev > 0 else 0
                conv_rows.append({
                    '從': funnel_df.iloc[i - 1]['階段'],
                    '到': funnel_df.iloc[i]['階段'],
                    '前階段人數': prev, '到達人數': curr,
                    '流失人數': prev - curr,
                    '轉換率(%)': round(cr, 1),
                    '流失率(%)': round(100 - cr, 1)
                })
            conv_df = pd.DataFrame(conv_rows)
            st.dataframe(
                conv_df.style.background_gradient(subset=['轉換率(%)'], cmap='RdYlGn')
                    .background_gradient(subset=['流失率(%)'], cmap='RdYlGn_r'),
                use_container_width=True
            )

            if sd == "全部科系":
                st.markdown("### 🏫 各科系最終報到轉換率")
                dept_fn = []
                for dept in df_fn['報考科系'].unique():
                    dd = df_fn[df_fn['報考科系'] == dept]
                    t = len(dd)
                    e = len(dd[dd['stage_idx'] >= 4])
                    dept_fn.append({'科系': dept, '申請人數': t, '報到人數': e,
                                   '轉換率(%)': round(e / t * 100, 1) if t > 0 else 0})
                dept_fn_df = pd.DataFrame(dept_fn).sort_values('轉換率(%)', ascending=True)
                fig = px.bar(dept_fn_df, y='科系', x='轉換率(%)', orientation='h',
                            text='轉換率(%)', color='轉換率(%)', color_continuous_scale='RdYlGn',
                            title="各科系最終報到轉換率")
                fig.update_traces(textposition='outside', texttemplate='%{text}%')
                fig.update_layout(height=500, yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)

            # ★ 新增：多學年度漏斗比較
            if sy == "全部學年度" and len(years) > 1:
                st.markdown("### 📈 歷年各階段人數比較")
                yearly_funnel = []
                for yr in years:
                    dy = df_fn[df_fn['學年度'] == yr].copy()
                    dy['stage_idx'] = dy['階段狀態'].map(stage_map).fillna(0).astype(int)
                    for i, stg in enumerate(stage_order):
                        cnt = len(dy[dy['stage_idx'] >= i])
                        yearly_funnel.append({'學年度': str(yr), '階段': stg, '人數': cnt})
                yf_df = pd.DataFrame(yearly_funnel)
                fig = px.bar(yf_df, x='階段', y='人數', color='學年度', barmode='group',
                            title='歷年各階段人數比較', color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("篩選條件下無資料")
    else:
        st.info("請上傳申請入學資料")


# ============================================================
# TAB 3: 來源學校分析
# ============================================================
with tabs[3]:
    st.markdown("## 🏫 來源學校轉換率分析")

    if st.session_state.applicant_data is not None:
        df = st.session_state.applicant_data.copy()

        s1, s2, s3 = st.columns(3)
        with s1:
            years = sorted(df['學年度'].unique())
            yr_opts = ["全部學年度"] + [str(y) for y in years]
            sy_sch = st.selectbox("學年度", yr_opts, key="sch_y")
        with s2:
            d_opts = ["全部科系"] + sorted(df['報考科系'].unique().tolist())
            sd_sch = st.selectbox("科系篩選", d_opts, key="sch_d")
        with s3:
            min_app = st.slider("最低申請人數門檻", 1, 20, 3, key="min_a")

        df_s = df.copy()
        if sy_sch != "全部學年度":
            df_s = df_s[df_s['學年度'] == int(sy_sch)]
        if sd_sch != "全部科系":
            df_s = df_s[df_s['報考科系'] == sd_sch]

        if len(df_s) > 0:
            sch_stats = df_s.groupby(['畢業學校', '畢業學校縣市']).agg(
                申請人數=('學生編號', 'count'),
                入學人數=('最終入學', lambda x: (x == '是').sum()),
                報考科系數=('報考科系', 'nunique'),
            ).reset_index()
            sch_stats['轉換率(%)'] = (sch_stats['入學人數'] / sch_stats['申請人數'] * 100).round(1)
            sch_stats = sch_stats[sch_stats['申請人數'] >= min_app].sort_values('入學人數', ascending=False)

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("🏫 來源學校總數", len(sch_stats))
            if len(sch_stats) > 0:
                k2.metric("🥇 入學最多", sch_stats.iloc[0]['畢業學校'])
            hi = len(sch_stats[sch_stats['轉換率(%)'] >= 50])
            k3.metric("⭐ 高轉換學校", f"{hi} 校")
            k4.metric("📊 平均轉換率", f"{sch_stats['轉換率(%)'].mean():.1f}%")

            st.markdown("---")
            lc, rc = st.columns(2)

            with lc:
                st.markdown("### 📊 入學人數 Top 15")
                top15 = sch_stats.head(15).sort_values('入學人數', ascending=True)
                fig = go.Figure()
                fig.add_trace(go.Bar(y=top15['畢業學校'], x=top15['申請人數'],
                                    name='申請人數', orientation='h', marker_color='#667eea', opacity=0.7))
                fig.add_trace(go.Bar(y=top15['畢業學校'], x=top15['入學人數'],
                                    name='入學人數', orientation='h', marker_color='#28a745'))
                fig.update_layout(barmode='overlay', height=550, yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)

            with rc:
                st.markdown("### 📈 轉換率 Top 15")
                top_r = sch_stats.sort_values('轉換率(%)', ascending=True).tail(15)
                fig = px.bar(top_r, y='畢業學校', x='轉換率(%)', orientation='h',
                            text='轉換率(%)', color='轉換率(%)', color_continuous_scale='RdYlGn')
                fig.update_traces(textposition='outside', texttemplate='%{text}%')
                fig.update_layout(height=550, yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)

            def get_rating(row):
                if row['轉換率(%)'] >= 50 and row['入學人數'] >= 3:
                    return "⭐⭐⭐ 重點經營"
                elif row['轉換率(%)'] >= 30 and row['入學人數'] >= 2:
                    return "⭐⭐ 持續關注"
                elif row['入學人數'] >= 1:
                    return "⭐ 一般往來"
                else:
                    return "🔸 待開發"

            sch_stats['經營建議'] = sch_stats.apply(get_rating, axis=1)

            st.markdown("### 📋 完整統計表（含經營建議）")
            st.dataframe(
                sch_stats.style
                    .background_gradient(subset=['轉換率(%)'], cmap='RdYlGn')
                    .background_gradient(subset=['入學人數'], cmap='Blues'),
                use_container_width=True, height=500
            )

            if sy_sch == "全部學年度" and len(years) > 1:
                st.markdown("### 📈 重點學校歷年入學趨勢")
                top_schools = sch_stats.head(8)['畢業學校'].tolist()
                ys = df_s[df_s['畢業學校'].isin(top_schools)].groupby(
                    ['學年度', '畢業學校']
                ).agg(入學人數=('最終入學', lambda x: (x == '是').sum())).reset_index()
                fig = px.line(ys, x='學年度', y='入學人數', color='畢業學校',
                             markers=True, title="重點來源學校歷年入學趨勢")
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

    if st.session_state.retention_data is not None:
        df_r = st.session_state.retention_data.copy()

        r1, r2 = st.columns(2)
        with r1:
            r_years = sorted(df_r['學年度'].unique())
            ry_opts = ["全部學年度"] + [str(y) for y in r_years]
            sel_ry = st.selectbox("入學學年度", ry_opts, key="ret_y")
        with r2:
            rd_opts = ["全部科系"] + sorted(df_r['入學科系'].unique().tolist())
            sel_rd = st.selectbox("科系篩選", rd_opts, key="ret_d")

        df_rf = df_r.copy()
        if sel_ry != "全部學年度":
            df_rf = df_rf[df_rf['學年度'] == int(sel_ry)]
        if sel_rd != "全部科系":
            df_rf = df_rf[df_rf['入學科系'] == sel_rd]

        if len(df_rf) > 0:
            total = len(df_rf)
            active = len(df_rf[df_rf['目前狀態'].isin(['在學', '畢業'])])
            susp = len(df_rf[df_rf['目前狀態'] == '休學'])
            drop = len(df_rf[df_rf['目前狀態'] == '退學'])
            grad = len(df_rf[df_rf['目前狀態'] == '畢業'])

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("👥 總人數", total)
            m2.metric("📚 在學/畢業", active, f"{active / total * 100:.1f}%")
            m3.metric("⏸️ 休學", susp, f"{susp / total * 100:.1f}%")
            m4.metric("❌ 退學", drop, f"{drop / total * 100:.1f}%")
            m5.metric("🎓 畢業", grad, f"{grad / total * 100:.1f}%")

            st.caption(f"📅 資料涵蓋學年度：{', '.join(str(y) for y in r_years)}")
            st.markdown("---")

            cl, cr = st.columns(2)

            with cl:
                st.markdown("### 📊 各入學管道穩定度")
                ch_stats = df_rf.groupby('入學管道').agg(
                    總人數=('學生編號', 'count'),
                    在學畢業=('目前狀態', lambda x: x.isin(['在學', '畢業']).sum()),
                    休學=('目前狀態', lambda x: (x == '休學').sum()),
                    退學=('目前狀態', lambda x: (x == '退學').sum()),
                ).reset_index()
                ch_stats['穩定率'] = (ch_stats['在學畢業'] / ch_stats['總人數'] * 100).round(1)
                ch_stats['休學率'] = (ch_stats['休學'] / ch_stats['總人數'] * 100).round(1)
                ch_stats['退學率'] = (ch_stats['退學'] / ch_stats['總人數'] * 100).round(1)
                ch_stats = ch_stats.sort_values('穩定率', ascending=True)

                fig = go.Figure()
                fig.add_trace(go.Bar(y=ch_stats['入學管道'], x=ch_stats['穩定率'],
                                    name='穩定率%', orientation='h', marker_color='#28a745',
                                    text=ch_stats['穩定率'].apply(lambda x: f'{x}%'), textposition='inside'))
                fig.add_trace(go.Bar(y=ch_stats['入學管道'], x=ch_stats['休學率'],
                                    name='休學率%', orientation='h', marker_color='#ffc107',
                                    text=ch_stats['休學率'].apply(lambda x: f'{x}%'), textposition='inside'))
                fig.add_trace(go.Bar(y=ch_stats['入學管道'], x=ch_stats['退學率'],
                                    name='退學率%', orientation='h', marker_color='#dc3545',
                                    text=ch_stats['退學率'].apply(lambda x: f'{x}%'), textposition='inside'))
                fig.update_layout(barmode='stack', height=400, xaxis_title='百分比(%)', yaxis_title='')
                st.plotly_chart(fig, use_container_width=True)

            with cr:
                st.markdown("### 📊 各科系休退學率")
                dp_stats = df_rf.groupby('入學科系').agg(
                    總人數=('學生編號', 'count'),
                    休退學=('目前狀態', lambda x: x.isin(['休學', '退學']).sum()),
                ).reset_index()
                dp_stats['休退學率(%)'] = (dp_stats['休退學'] / dp_stats['總人數'] * 100).round(1)
                dp_stats = dp_stats.sort_values('休退學率(%)', ascending=True)

                fig = px.bar(dp_stats, y='入學科系', x='休退學率(%)', orientation='h',
                            text='休退學率(%)', color='休退學率(%)',
                            color_continuous_scale='RdYlGn_r')
                fig.update_traces(textposition='outside', texttemplate='%{text}%')
                fig.update_layout(height=400, yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("### 📋 休退學原因分析")
            do_df = df_rf[df_rf['目前狀態'].isin(['休學', '退學'])]

            rc1, rc2 = st.columns(2)
            with rc1:
                if '休退學原因' in do_df.columns and len(do_df) > 0:
                    reason_s = do_df[do_df['休退學原因'] != '']['休退學原因'].value_counts()
                    if len(reason_s) > 0:
                        fig = px.pie(values=reason_s.values, names=reason_s.index,
                                    title='休退學原因分布',
                                    color_discrete_sequence=px.colors.qualitative.Set2)
                        fig.update_traces(textposition='inside', textinfo='percent+label')
                        fig.update_layout(height=400)
                        st.plotly_chart(fig, use_container_width=True)

            with rc2:
                if '休退學原因' in do_df.columns and len(do_df) > 0:
                    valid_do = do_df[do_df['休退學原因'] != '']
                    if len(valid_do) > 0:
                        cross = pd.crosstab(valid_do['入學管道'], valid_do['休退學原因'])
                        if len(cross) > 0:
                            fig = px.imshow(cross, text_auto=True, color_continuous_scale='YlOrRd',
                                          title='入學管道 × 休退學原因', aspect='auto')
                            fig.update_layout(height=400)
                            st.plotly_chart(fig, use_container_width=True)

            if sel_ry == "全部學年度" and len(r_years) > 1:
                st.markdown("### 📈 歷年各管道休退學率趨勢")
                yc = df_r.copy()
                if sel_rd != "全部科系":
                    yc = yc[yc['入學科系'] == sel_rd]
                yc = yc.groupby(['學年度', '入學管道']).agg(
                    總人數=('學生編號', 'count'),
                    休退學=('目前狀態', lambda x: x.isin(['休學', '退學']).sum())
                ).reset_index()
                yc['休退學率(%)'] = (yc['休退學'] / yc['總人數'] * 100).round(1)
                fig = px.line(yc, x='學年度', y='休退學率(%)', color='入學管道',
                             markers=True, title='歷年各入學管道休退學率趨勢')
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("### 📋 各入學管道完整統計")
            st.dataframe(
                ch_stats.style
                    .background_gradient(subset=['穩定率'], cmap='RdYlGn')
                    .background_gradient(subset=['退學率'], cmap='RdYlGn_r'),
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

    dt1, dt2, dt3 = st.tabs(["申請入學資料", "在學穩定度資料", "匯入紀錄"])

    with dt1:
        if st.session_state.applicant_data is not None:
            dv = st.session_state.applicant_data
            st.markdown(f"**共 {len(dv)} 筆資料，涵蓋 {dv['學年度'].nunique()} 個學年度**")

            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                search_txt = st.text_input("搜尋學校名稱", key="srch_sch")
            with sc2:
                filter_yr = st.multiselect("篩選學年度", sorted(dv['學年度'].unique()), key="flt_yr")
            with sc3:
                filter_dept = st.multiselect("篩選科系", sorted(dv['報考科系'].unique()), key="flt_dept")

            dv_show = dv.copy()
            if search_txt:
                dv_show = dv_show[dv_show['畢業學校'].str.contains(search_txt, na=False)]
            if filter_yr:
                dv_show = dv_show[dv_show['學年度'].isin(filter_yr)]
            if filter_dept:
                dv_show = dv_show[dv_show['報考科系'].isin(filter_dept)]

            st.markdown(f"篩選後：**{len(dv_show)}** 筆")
            st.dataframe(dv_show, use_container_width=True, height=400)

            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                csv_data = dv_show.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 下載篩選後資料 (CSV)", csv_data,
                                 "applicant_filtered.csv", "text/csv")
            with col_dl2:
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    dv_show.to_excel(writer, index=False, sheet_name='申請入學')
                st.download_button("📥 下載篩選後資料 (Excel)", buffer.getvalue(),
                                 "applicant_filtered.xlsx",
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("尚未載入申請入學資料")

    with dt2:
        if st.session_state.retention_data is not None:
            dv2 = st.session_state.retention_data
            st.markdown(f"**共 {len(dv2)} 筆資料，涵蓋 {dv2['學年度'].nunique()} 個學年度**")

            rc1, rc2 = st.columns(2)
            with rc1:
                r_filter_yr = st.multiselect("篩選學年度", sorted(dv2['學年度'].unique()), key="r_flt_yr")
            with rc2:
                r_filter_ch = st.multiselect("篩選入學管道", sorted(dv2['入學管道'].unique()), key="r_flt_ch")

            dv2_show = dv2.copy()
            if r_filter_yr:
                dv2_show = dv2_show[dv2_show['學年度'].isin(r_filter_yr)]
            if r_filter_ch:
                dv2_show = dv2_show[dv2_show['入學管道'].isin(r_filter_ch)]

            st.markdown(f"篩選後：**{len(dv2_show)}** 筆")
            st.dataframe(dv2_show, use_container_width=True, height=400)

            col_dl3, col_dl4 = st.columns(2)
            with col_dl3:
                csv2 = dv2_show.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 下載篩選後資料 (CSV)", csv2,
                                 "retention_filtered.csv", "text/csv")
            with col_dl4:
                buffer2 = BytesIO()
                with pd.ExcelWriter(buffer2, engine='openpyxl') as writer:
                    dv2_show.to_excel(writer, index=False, sheet_name='在學穩定度')
                st.download_button("📥 下載篩選後資料 (Excel)", buffer2.getvalue(),
                                 "retention_filtered.xlsx",
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("尚未載入在學穩定度資料")

    # ★ 新增：匯入紀錄頁面
    with dt3:
        st.markdown("### 📋 已匯入檔案紀錄")

        st.markdown("#### 申請入學資料")
        if st.session_state.applicant_files:
            app_log = pd.DataFrame(st.session_state.applicant_files)
            app_log['學年度'] = app_log['years'].apply(lambda x: ', '.join(str(y) for y in x))
            app_log = app_log.rename(columns={'name': '檔案名稱', 'rows': '資料筆數'})
            st.dataframe(app_log[['檔案名稱', '資料筆數', '學年度']], use_container_width=True)
            total = len(st.session_state.applicant_data) if st.session_state.applicant_data is not None else 0
            st.info(f"合併後總計：{total} 筆（已自動去除重複學生編號）")
        else:
            st.caption("尚無匯入紀錄")

        st.markdown("#### 在學穩定度資料")
        if st.session_state.retention_files:
            ret_log = pd.DataFrame(st.session_state.retention_files)
            ret_log['學年度'] = ret_log['years'].apply(lambda x: ', '.join(str(y) for y in x))
            ret_log = ret_log.rename(columns={'name': '檔案名稱', 'rows': '資料筆數'})
            st.dataframe(ret_log[['檔案名稱', '資料筆數', '學年度']], use_container_width=True)
            total = len(st.session_state.retention_data) if st.session_state.retention_data is not None else 0
            st.info(f"合併後總計：{total} 筆（已自動去除重複學生編號）")
        else:
            st.caption("尚無匯入紀錄")


# ============================================================
# 頁尾
# ============================================================
st.markdown("---")
st.markdown(
    "<div style='text-align:center;color:#999;font-size:0.85rem;'>"
    "中華醫事科技大學 入學服務處 招生數據分析系統 v2.1 | © 2024"
    "</div>",
    unsafe_allow_html=True
)
