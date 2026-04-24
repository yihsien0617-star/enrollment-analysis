import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import hashlib

# ============================================================
# 頁面設定
# ============================================================
st.set_page_config(
    page_title="中華醫事科技大學 招生數據分析系統",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem; font-weight: 700; color: #1B3A5C;
        text-align: center; padding: 1rem 0;
        border-bottom: 3px solid #E8792F; margin-bottom: 2rem;
    }
    .info-box {
        background-color: #f0f7ff; border-left: 5px solid #2C5F8A;
        padding: 1rem 1.5rem; border-radius: 0 8px 8px 0; margin: 1rem 0;
    }
    .file-tag {
        display: inline-block; background: #e8f5e9; border: 1px solid #81c784;
        border-radius: 16px; padding: 3px 12px; margin: 3px 2px; font-size: 0.82rem;
    }
    .file-tag-warn {
        display: inline-block; background: #fff3e0; border: 1px solid #ffb74d;
        border-radius: 16px; padding: 3px 12px; margin: 3px 2px; font-size: 0.82rem;
    }
    .merge-box {
        background: linear-gradient(135deg, #e8f5e9 0%, #f1f8e9 100%);
        border: 2px solid #66bb6a; border-radius: 10px;
        padding: 1rem 1.5rem; margin: 0.8rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 座標 & 欄位定義
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

APPLICANT_REQUIRED = ["學年度", "學生編號", "報考科系", "畢業學校",
                      "畢業學校縣市", "住家縣市", "入學管道", "階段狀態", "最終入學"]
RETENTION_REQUIRED = ["學年度", "學生編號", "入學管道", "入學科系", "目前狀態"]


# ============================================================
# 工具函式
# ============================================================
def file_hash(uploaded_file):
    """根據檔名 + 大小產生唯一識別碼"""
    return hashlib.md5(f"{uploaded_file.name}_{uploaded_file.size}".encode()).hexdigest()


def read_file(uploaded_file):
    try:
        uploaded_file.seek(0)
        if uploaded_file.name.endswith('.csv'):
            return pd.read_csv(uploaded_file)
        else:
            return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ 無法讀取 {uploaded_file.name}：{e}")
        return None


def validate_cols(df, required, fname):
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"❌ {fname} 缺少欄位：{', '.join(missing)}")
        return False
    return True


def merge_into(existing_df, new_df):
    """合併：以學生編號+學年度為 key，新資料覆蓋舊資料"""
    if existing_df is None or len(existing_df) == 0:
        return new_df.copy()
    
    key_cols = ['學生編號']
    if '學年度' in new_df.columns and '學年度' in existing_df.columns:
        key_cols = ['學生編號', '學年度']
    
    # 標記新資料
    new_keys = new_df[key_cols].drop_duplicates()
    merged_mark = existing_df.merge(new_keys, on=key_cols, how='left', indicator=True)
    kept = existing_df[merged_mark['_merge'] == 'left_only']
    
    result = pd.concat([kept, new_df], ignore_index=True)
    return result


# ============================================================
# Session State 初始化（資料永久保存在 session 中）
# ============================================================
def init_state():
    defaults = {
        'applicant_data': None,          # 合併後的 DataFrame
        'retention_data': None,           # 合併後的 DataFrame
        'applicant_file_log': [],         # [{name, rows, years, hash}]
        'retention_file_log': [],
        'processed_hashes_app': set(),    # 已處理過的檔案 hash
        'processed_hashes_ret': set(),
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()


# ============================================================
# 範例資料產生器
# ============================================================
@st.cache_data
def generate_sample_applicants():
    np.random.seed(42)
    departments = ["護理系", "食品科技系", "醫學檢驗生物技術系", "職業安全衛生系",
                   "環境與安全衛生工程系", "生物醫學工程系", "藥學系", "長期照護系",
                   "幼兒保育系", "資訊管理系", "化妝品應用與管理系", "餐旅管理系"]
    schools_by_city = {
        "台南市": ["台南高工", "台南女中", "長榮高中", "南光高中", "港明高中",
                   "新營高工", "北門高中", "曾文高中", "南科實中", "家齊高中"],
        "高雄市": ["高雄中學", "高雄女中", "鳳山高中", "前鎮高中", "三信家商",
                   "中山工商", "高雄高工", "鳳新高中", "路竹高中"],
        "嘉義縣": ["嘉義高中", "民雄農工", "東石高中", "協志工商"],
        "嘉義市": ["嘉義高商", "輔仁中學", "嘉華中學"],
        "屏東縣": ["屏東高中", "屏東女中", "屏東高工", "潮州高中", "東港高中"],
        "雲林縣": ["斗六高中", "虎尾高中", "北港農工", "土庫商工"],
        "彰化縣": ["彰化高中", "員林高中", "鹿港高中", "大慶商工"],
        "台中市": ["台中一中", "豐原高中", "大甲高中", "明道中學"],
        "新北市": ["板橋高中", "中和高中", "三重商工"],
        "台北市": ["建國中學", "大安高工", "松山工農"],
        "桃園市": ["武陵高中", "中壢高中"],
        "南投縣": ["南投高中", "草屯商工"],
        "花蓮縣": ["花蓮高中", "花蓮高工"],
        "台東縣": ["台東高中"],
        "澎湖縣": ["馬公高中"],
    }
    city_w = {"台南市": .35, "高雄市": .20, "嘉義縣": .08, "嘉義市": .06,
              "屏東縣": .08, "雲林縣": .05, "彰化縣": .04, "台中市": .04,
              "新北市": .02, "台北市": .01, "桃園市": .01,
              "南投縣": .01, "花蓮縣": .01, "台東縣": .01, "澎湖縣": .01}
    cities = list(city_w.keys())
    weights = np.array(list(city_w.values()))
    weights /= weights.sum()

    stages = ["第一階段報名", "通過第一階段", "完成二階面試", "錄取", "已報到"]
    cum_p = [1.0, 0.75, 0.60, 0.45, 0.35]

    records = []
    idx = 0
    for year, n in {111: 150, 112: 180, 113: 200}.items():
        for _ in range(n):
            city = np.random.choice(cities, p=weights)
            school = np.random.choice(schools_by_city[city])
            fi = 0
            for si in range(len(stages)):
                if np.random.random() < cum_p[si]:
                    fi = si
                else:
                    break
            records.append({
                "學年度": year, "學生編號": f"A{year}{idx:04d}",
                "報考科系": np.random.choice(departments),
                "畢業學校": school, "畢業學校縣市": city,
                "住家縣市": city if np.random.random() < .85 else np.random.choice(cities, p=weights),
                "入學管道": "申請入學", "階段狀態": stages[fi],
                "最終入學": "是" if fi == 4 else "否",
            })
            idx += 1
    return pd.DataFrame(records)


@st.cache_data
def generate_sample_retention():
    np.random.seed(123)
    channels = ["申請入學", "統測分發", "繁星推薦", "技優甄審", "單獨招生", "運動績優"]
    ch_w = [.30, .25, .15, .10, .15, .05]
    depts = ["護理系", "食品科技系", "醫學檢驗生物技術系", "職業安全衛生系",
             "長期照護系", "幼兒保育系", "資訊管理系", "化妝品應用與管理系", "餐旅管理系"]
    d_rates = {"申請入學": (.08, .05), "統測分發": (.12, .10), "繁星推薦": (.06, .03),
               "技優甄審": (.10, .07), "單獨招生": (.15, .12), "運動績優": (.13, .09)}
    reasons = ["學業因素", "經濟因素", "志趣不合", "家庭因素", "健康因素"]
    records = []
    idx = 0
    for year, n in {108: 150, 109: 160, 110: 170, 111: 180, 112: 160}.items():
        for _ in range(n):
            ch = np.random.choice(channels, p=ch_w)
            sr, dr = d_rates[ch]
            rv = np.random.random()
            if rv < dr:
                status, reason = "退學", np.random.choice(reasons)
            elif rv < dr + sr:
                status, reason = "休學", np.random.choice(reasons)
            else:
                status, reason = ("在學" if year >= 111 else "畢業"), ""
            records.append({
                "學年度": year, "學生編號": f"R{year}{idx:04d}",
                "入學管道": ch, "入學科系": np.random.choice(depts),
                "入學學期": f"{year}-1", "目前狀態": status,
                "休退學原因": reason,
            })
            idx += 1
    return pd.DataFrame(records)


# ============================================================
# ★★★ 核心：匯入資料處理函式 ★★★
# ============================================================
def process_uploaded_files(uploaded_files, data_key, log_key, hash_key, required_cols, label):
    """
    處理上傳的多個檔案，合併到 session_state 中。
    用 hash 追蹤已處理過的檔案，避免重複匯入。
    """
    if not uploaded_files:
        return

    new_count = 0
    for f in uploaded_files:
        fh = file_hash(f)
        # 已經處理過 → 跳過
        if fh in st.session_state[hash_key]:
            continue

        df_new = read_file(f)
        if df_new is None:
            continue
        if not validate_cols(df_new, required_cols, f.name):
            continue

        # 合併
        st.session_state[data_key] = merge_into(st.session_state[data_key], df_new)

        # 記錄
        years_list = sorted(df_new['學年度'].unique().tolist())
        st.session_state[log_key].append({
            'name': f.name,
            'rows': len(df_new),
            'years': years_list,
            'hash': fh
        })
        st.session_state[hash_key].add(fh)
        new_count += 1

    if new_count > 0:
        total = len(st.session_state[data_key])
        st.success(f"✅ 成功匯入 {new_count} 個{label}檔案！目前共 {total:,} 筆資料")


def display_file_status(data_key, log_key, label):
    """顯示已匯入檔案的狀態"""
    if st.session_state[log_key]:
        df_all = st.session_state[data_key]
        total_rows = len(df_all) if df_all is not None else 0
        all_years = sorted(df_all['學年度'].unique()) if df_all is not None else []

        st.markdown(
            f'<div class="merge-box">'
            f'📊 <b>{label}</b>：合計 <b>{total_rows:,}</b> 筆 ｜ '
            f'學年度 <b>{", ".join(str(y) for y in all_years)}</b> ｜ '
            f'已匯入 <b>{len(st.session_state[log_key])}</b> 個檔案'
            f'</div>',
            unsafe_allow_html=True
        )
        for rec in st.session_state[log_key]:
            yrs = ", ".join(str(y) for y in rec['years'])
            st.markdown(
                f'<span class="file-tag">📄 {rec["name"]}（{rec["rows"]}筆，{yrs}學年）</span>',
                unsafe_allow_html=True
            )
    else:
        st.caption(f"尚未匯入{label}")


# ============================================================
# 側邊欄
# ============================================================
with st.sidebar:
    st.markdown("## 🎓 招生分析系統")
    st.markdown("**中華醫事科技大學** 入學服務處")
    st.markdown("---")

    # ── 一、申請入學資料 ──
    st.markdown("### 📂 一、申請入學資料")
    st.caption("支援多檔上傳，不同學年度會自動合併")
    
    app_files = st.file_uploader(
        "選擇檔案（可多選）",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        key="uploader_app"
    )
    
    # ★ 用 hash 機制處理，不怕重跑
    process_uploaded_files(
        app_files, 'applicant_data', 'applicant_file_log',
        'processed_hashes_app', APPLICANT_REQUIRED, "申請入學"
    )
    display_file_status('applicant_data', 'applicant_file_log', '申請入學')

    st.markdown("---")

    # ── 二、在學穩定度資料 ──
    st.markdown("### 📂 二、在學穩定度資料")
    st.caption("支援多檔上傳，不同學年度會自動合併")
    
    ret_files = st.file_uploader(
        "選擇檔案（可多選）",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        key="uploader_ret"
    )
    
    process_uploaded_files(
        ret_files, 'retention_data', 'retention_file_log',
        'processed_hashes_ret', RETENTION_REQUIRED, "在學穩定度"
    )
    display_file_status('retention_data', 'retention_file_log', '在學穩定度')

    st.markdown("---")

    # ── 清除資料 ──
    st.markdown("### 🔧 資料管理")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("🗑️ 清除申請", use_container_width=True):
            st.session_state.applicant_data = None
            st.session_state.applicant_file_log = []
            st.session_state.processed_hashes_app = set()
            st.rerun()
    with c2:
        if st.button("🗑️ 清除在學", use_container_width=True):
            st.session_state.retention_data = None
            st.session_state.retention_file_log = []
            st.session_state.processed_hashes_ret = set()
            st.rerun()

    if st.button("🗑️ 清除全部資料", use_container_width=True):
        for k in ['applicant_data', 'retention_data']:
            st.session_state[k] = None
        for k in ['applicant_file_log', 'retention_file_log']:
            st.session_state[k] = []
        for k in ['processed_hashes_app', 'processed_hashes_ret']:
            st.session_state[k] = set()
        st.rerun()

    st.markdown("---")
    st.markdown("### 🧪 快速體驗")
    if st.button("📋 載入範例資料", use_container_width=True, type="primary"):
        st.session_state.applicant_data = generate_sample_applicants()
        st.session_state.retention_data = generate_sample_retention()
        st.session_state.applicant_file_log = [
            {"name": "範例_申請入學.csv",
             "rows": len(st.session_state.applicant_data),
             "years": sorted(st.session_state.applicant_data['學年度'].unique().tolist()),
             "hash": "sample_app"}
        ]
        st.session_state.retention_file_log = [
            {"name": "範例_在學穩定度.csv",
             "rows": len(st.session_state.retention_data),
             "years": sorted(st.session_state.retention_data['學年度'].unique().tolist()),
             "hash": "sample_ret"}
        ]
        st.session_state.processed_hashes_app = {"sample_app"}
        st.session_state.processed_hashes_ret = {"sample_ret"}
        st.rerun()

    st.markdown("---")
    st.caption("v2.1 多學年度合併版 | © 2024")


# ============================================================
# 主頁面標題
# ============================================================
st.markdown('<div class="main-header">🎓 中華醫事科技大學 招生數據分析系統</div>', unsafe_allow_html=True)

# ============================================================
# 無資料 → 歡迎頁
# ============================================================
if st.session_state.applicant_data is None and st.session_state.retention_data is None:
    st.markdown("""
    <div class="info-box">
        <h3>👋 歡迎使用招生數據分析系統</h3>
        <p>本系統支援 <b>多檔案上傳、多學年度自動合併</b>：</p>
        <ul>
            <li>📂 <b>一次選多個檔案</b>：按住 Ctrl / Cmd 點選多個檔案一起上傳</li>
            <li>📂 <b>分批上傳</b>：先傳 111 學年，再傳 112 學年，資料自動累加</li>
            <li>🔄 相同學生編號+學年度的資料會自動以新檔為準</li>
        </ul>
        <h4>📊 分析功能：</h4>
        <ol>
            <li><b>總覽儀表板</b>：歷年申請與入學趨勢</li>
            <li><b>地圖分布分析</b>：台灣地圖呈現學生來源</li>
            <li><b>招生漏斗分析</b>：追蹤各階段轉換</li>
            <li><b>來源學校分析</b>：評估各校轉換率</li>
            <li><b>在學穩定度分析</b>：比較不同管道休退學情況</li>
        </ol>
        <p>👈 請從左側上傳資料，或點擊 <b>「載入範例資料」</b> 快速體驗。</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📋 資料格式說明")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 申請入學資料")
        st.dataframe(pd.DataFrame({
            "欄位": APPLICANT_REQUIRED,
            "範例": ["113", "A1130001", "護理系", "台南高工", "台南市",
                    "台南市", "申請入學", "已報到", "是"]
        }), use_container_width=True)
        st.info("💡 建議每個學年度存一個檔案，例如 `111申請入學.csv`、`112申請入學.csv`")
    with c2:
        st.markdown("#### 在學穩定度資料")
        st.dataframe(pd.DataFrame({
            "欄位": RETENTION_REQUIRED + ["休退學原因"],
            "範例": ["110", "R1100001", "申請入學", "護理系", "在學", ""]
        }), use_container_width=True)
        st.info("💡 同樣每學年度一個檔案，系統自動合併")
    st.stop()


# ============================================================
# 有資料 → 分頁分析
# ============================================================
tabs = st.tabs([
    "📊 總覽儀表板",
    "🗺️ 地圖分布",
    "🔽 招生漏斗",
    "🏫 來源學校",
    "📉 在學穩定度",
    "📥 資料檢視與匯出"
])

# ── 共用：學年度篩選器 ──
def year_selector(df, key_prefix):
    years = sorted(df['學年度'].unique())
    opts = ["全部學年度"] + [str(y) for y in years]
    sel = st.selectbox("選擇學年度", opts, key=f"{key_prefix}_yr")
    if sel == "全部學年度":
        return df.copy(), "全部學年度", years
    else:
        y = int(sel)
        return df[df['學年度'] == y].copy(), f"{y} 學年度", years


# ============================================================
# TAB 0: 總覽
# ============================================================
with tabs[0]:
    st.markdown("## 📊 招生總覽儀表板")
    if st.session_state.applicant_data is not None:
        df = st.session_state.applicant_data.copy()
        df_yr, disp, all_years = year_selector(df, "ov")

        total = len(df_yr)
        enrolled = (df_yr['最終入學'] == '是').sum()
        rate = enrolled / total * 100 if total else 0

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("📝 申請人數", f"{total:,}")
        c2.metric("✅ 入學人數", f"{enrolled:,}")
        c3.metric("📈 轉換率", f"{rate:.1f}%")
        c4.metric("🏫 來源學校", df_yr['畢業學校'].nunique())
        c5.metric("🗺️ 來源縣市", df_yr['畢業學校縣市'].nunique())
        st.caption(f"📅 資料涵蓋：{', '.join(str(y) for y in all_years)} 學年度（共 {len(all_years)} 個學年度）")

        st.markdown("---")
        lc, rc = st.columns(2)

        with lc:
            dept = df_yr.groupby('報考科系').agg(
                申請=('學生編號', 'count'),
                入學=('最終入學', lambda x: (x == '是').sum())
            ).reset_index().sort_values('申請', ascending=True)
            fig = px.bar(dept, y='報考科系', x=['申請', '入學'], orientation='h',
                        barmode='group', color_discrete_sequence=['#667eea', '#28a745'],
                        title=f"{disp} 各科系申請與入學")
            fig.update_layout(height=500, yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)

        with rc:
            city = df_yr['畢業學校縣市'].value_counts().head(10).reset_index()
            city.columns = ['縣市', '人數']
            fig = px.pie(city, values='人數', names='縣市',
                        title=f"{disp} 來源縣市 Top 10",
                        color_discrete_sequence=px.colors.qualitative.Set3)
            fig.update_traces(textposition='inside', textinfo='percent+label')
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)

        if len(all_years) > 1:
            st.markdown("### 📈 歷年趨勢")
            yearly = df.groupby('學年度').agg(
                申請=('學生編號', 'count'),
                入學=('最終入學', lambda x: (x == '是').sum())
            ).reset_index()
            yearly['轉換率'] = (yearly['入學'] / yearly['申請'] * 100).round(1)

            fig = go.Figure()
            fig.add_trace(go.Bar(x=yearly['學年度'].astype(str), y=yearly['申請'],
                                name='申請', marker_color='#667eea'))
            fig.add_trace(go.Bar(x=yearly['學年度'].astype(str), y=yearly['入學'],
                                name='入學', marker_color='#28a745'))
            fig.add_trace(go.Scatter(
                x=yearly['學年度'].astype(str), y=yearly['轉換率'], name='轉換率(%)',
                yaxis='y2', mode='lines+markers+text',
                line=dict(color='#E8792F', width=3),
                text=yearly['轉換率'].apply(lambda v: f'{v}%'), textposition='top center'))
            fig.update_layout(
                barmode='group', height=420,
                yaxis=dict(title='人數'),
                yaxis2=dict(title='轉換率(%)', overlaying='y', side='right', range=[0, 100]))
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("請先上傳申請入學資料")


# ============================================================
# TAB 1: 地圖
# ============================================================
with tabs[1]:
    st.markdown("## 🗺️ 台灣地圖 - 學生來源分布")
    if st.session_state.applicant_data is not None:
        df = st.session_state.applicant_data.copy()

        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            df_map, disp_m, _ = year_selector(df, "map")
        with fc2:
            d_opts = ["全部科系"] + sorted(df['報考科系'].unique().tolist())
            sel_d = st.selectbox("科系", d_opts, key="map_d")
        with fc3:
            sel_s = st.selectbox("階段", ["全部申請者", "僅已報到"], key="map_s")

        if sel_d != "全部科系":
            df_map = df_map[df_map['報考科系'] == sel_d]
        if sel_s == "僅已報到":
            df_map = df_map[df_map['最終入學'] == '是']

        if len(df_map) > 0:
            try:
                import folium
                from streamlit_folium import st_folium

                m = folium.Map(location=[23.5, 120.9], zoom_start=8, tiles='CartoDB positron')
                folium.Marker([23.0048, 120.2210], popup="中華醫事科技大學",
                             tooltip="📍 中華醫事科技大學",
                             icon=folium.Icon(color='red', icon='star', prefix='fa')).add_to(m)

                city_cnt = df_map['住家縣市'].value_counts().to_dict()
                enr_cnt = df_map[df_map['最終入學'] == '是']['住家縣市'].value_counts().to_dict()
                mx = max(city_cnt.values()) if city_cnt else 1

                for city, cnt in city_cnt.items():
                    if city in TAIWAN_COUNTY_COORDS:
                        lat, lon = TAIWAN_COUNTY_COORDS[city]
                        enr = enr_cnt.get(city, 0)
                        r = enr / cnt * 100 if cnt else 0
                        radius = max(8, min(40, cnt / mx * 40))
                        color = '#28a745' if r >= 40 else '#ffc107' if r >= 25 else '#dc3545'
                        popup = f"<b>{city}</b><br>申請 {cnt}人<br>入學 {enr}人<br>轉換率 {r:.1f}%"
                        folium.CircleMarker(
                            [lat, lon], radius=radius,
                            popup=folium.Popup(popup, max_width=200),
                            tooltip=f"{city}：{cnt}人",
                            color=color, fill=True, fillColor=color,
                            fillOpacity=0.6, weight=2).add_to(m)

                st_folium(m, width=1100, height=600)
            except ImportError:
                st.warning("地圖套件未安裝，改用長條圖")
                city_df = df_map['住家縣市'].value_counts().reset_index()
                city_df.columns = ['縣市', '人數']
                fig = px.bar(city_df, x='縣市', y='人數', color='人數', color_continuous_scale='Blues')
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("### 📋 各縣市統計")
            cd = df_map.groupby('住家縣市').agg(
                申請=('學生編號', 'count'),
                入學=('最終入學', lambda x: (x == '是').sum()),
                學校數=('畢業學校', 'nunique')
            ).reset_index()
            cd['轉換率(%)'] = (cd['入學'] / cd['申請'] * 100).round(1)
            cd = cd.sort_values('申請', ascending=False)
            st.dataframe(cd.style.background_gradient(subset=['申請'], cmap='Blues')
                        .background_gradient(subset=['轉換率(%)'], cmap='RdYlGn'),
                        use_container_width=True, height=400)
        else:
            st.warning("篩選條件下無資料")
    else:
        st.info("請先上傳申請入學資料")


# ============================================================
# TAB 2: 漏斗
# ============================================================
with tabs[2]:
    st.markdown("## 🔽 招生漏斗分析")
    if st.session_state.applicant_data is not None:
        df = st.session_state.applicant_data.copy()

        fc1, fc2 = st.columns(2)
        with fc1:
            df_fn, disp_fn, all_yrs = year_selector(df, "fn")
        with fc2:
            d_opts = ["全部科系"] + sorted(df['報考科系'].unique().tolist())
            sel_fd = st.selectbox("科系", d_opts, key="fn_d")
        if sel_fd != "全部科系":
            df_fn = df_fn[df_fn['報考科系'] == sel_fd]

        if len(df_fn) > 0:
            stages = ["第一階段報名", "通過第一階段", "完成二階面試", "錄取", "已報到"]
            s_map = {s: i for i, s in enumerate(stages)}
            df_fn['si'] = df_fn['階段狀態'].map(s_map).fillna(0).astype(int)

            total = len(df_fn)
            funnel = [{'階段': s, '人數': (df_fn['si'] >= i).sum(),
                       '佔比': round((df_fn['si'] >= i).sum() / total * 100, 1)}
                      for i, s in enumerate(stages)]
            fdf = pd.DataFrame(funnel)

            fig = go.Figure(go.Funnel(
                y=fdf['階段'], x=fdf['人數'], textinfo="value+percent initial",
                marker=dict(color=['#667eea', '#7c8cf5', '#a3b1ff', '#48bb78', '#28a745']),
                connector=dict(line=dict(color="#ccc", width=2))))
            fig.update_layout(title=f"招生漏斗（{disp_fn} {sel_fd}）", height=500, font=dict(size=14))
            st.plotly_chart(fig, use_container_width=True)

            st.markdown("### 📊 各階段轉換率")
            conv = []
            for i in range(1, len(fdf)):
                p, c = fdf.iloc[i - 1]['人數'], fdf.iloc[i]['人數']
                cr = c / p * 100 if p else 0
                conv.append({'從': fdf.iloc[i - 1]['階段'], '到': fdf.iloc[i]['階段'],
                            '前階段': p, '到達': c, '流失': p - c,
                            '轉換率(%)': round(cr, 1), '流失率(%)': round(100 - cr, 1)})
            st.dataframe(pd.DataFrame(conv).style
                        .background_gradient(subset=['轉換率(%)'], cmap='RdYlGn')
                        .background_gradient(subset=['流失率(%)'], cmap='RdYlGn_r'),
                        use_container_width=True)

            if sel_fd == "全部科系":
                st.markdown("### 🏫 各科系報到轉換率")
                dept_fn = []
                for d in df_fn['報考科系'].unique():
                    dd = df_fn[df_fn['報考科系'] == d]
                    t, e = len(dd), (dd['si'] >= 4).sum()
                    dept_fn.append({'科系': d, '申請': t, '報到': e,
                                   '轉換率(%)': round(e / t * 100, 1) if t else 0})
                dfd = pd.DataFrame(dept_fn).sort_values('轉換率(%)', ascending=True)
                fig = px.bar(dfd, y='科系', x='轉換率(%)', orientation='h',
                            text='轉換率(%)', color='轉換率(%)', color_continuous_scale='RdYlGn')
                fig.update_traces(textposition='outside', texttemplate='%{text}%')
                fig.update_layout(height=500, yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)

            if disp_fn == "全部學年度" and len(all_yrs) > 1:
                st.markdown("### 📈 歷年各階段比較")
                rows = []
                for yr in all_yrs:
                    dy = df_fn[df_fn['學年度'] == yr]
                    for i, s in enumerate(stages):
                        rows.append({'學年度': str(yr), '階段': s, '人數': (dy['si'] >= i).sum()})
                fig = px.bar(pd.DataFrame(rows), x='階段', y='人數', color='學年度',
                            barmode='group', color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("篩選條件下無資料")
    else:
        st.info("請先上傳申請入學資料")


# ============================================================
# TAB 3: 來源學校
# ============================================================
with tabs[3]:
    st.markdown("## 🏫 來源學校轉換率分析")
    if st.session_state.applicant_data is not None:
        df = st.session_state.applicant_data.copy()

        sc1, sc2, sc3 = st.columns(3)
        with sc1:
            df_sch, disp_s, all_yrs = year_selector(df, "sch")
        with sc2:
            d_opts = ["全部科系"] + sorted(df['報考科系'].unique().tolist())
            sd = st.selectbox("科系", d_opts, key="sch_d")
        with sc3:
            min_a = st.slider("最低申請人數", 1, 20, 3, key="sch_min")
        if sd != "全部科系":
            df_sch = df_sch[df_sch['報考科系'] == sd]

        if len(df_sch) > 0:
            ss = df_sch.groupby(['畢業學校', '畢業學校縣市']).agg(
                申請=('學生編號', 'count'),
                入學=('最終入學', lambda x: (x == '是').sum()),
                科系數=('報考科系', 'nunique')
            ).reset_index()
            ss['轉換率(%)'] = (ss['入學'] / ss['申請'] * 100).round(1)
            ss = ss[ss['申請'] >= min_a].sort_values('入學', ascending=False)

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("🏫 學校數", len(ss))
            if len(ss) > 0:
                k2.metric("🥇 入學最多", ss.iloc[0]['畢業學校'])
            k3.metric("⭐ 高轉換(≥50%)", f"{(ss['轉換率(%)'] >= 50).sum()} 校")
            k4.metric("📊 平均轉換率", f"{ss['轉換率(%)'].mean():.1f}%")

            lc, rc = st.columns(2)
            with lc:
                top = ss.head(15).sort_values('入學', ascending=True)
                fig = go.Figure()
                fig.add_trace(go.Bar(y=top['畢業學校'], x=top['申請'], name='申請',
                                    orientation='h', marker_color='#667eea', opacity=.7))
                fig.add_trace(go.Bar(y=top['畢業學校'], x=top['入學'], name='入學',
                                    orientation='h', marker_color='#28a745'))
                fig.update_layout(barmode='overlay', height=550, yaxis_title="",
                                 title="入學人數 Top 15")
                st.plotly_chart(fig, use_container_width=True)
            with rc:
                top_r = ss.sort_values('轉換率(%)', ascending=True).tail(15)
                fig = px.bar(top_r, y='畢業學校', x='轉換率(%)', orientation='h',
                            text='轉換率(%)', color='轉換率(%)', color_continuous_scale='RdYlGn',
                            title="轉換率 Top 15")
                fig.update_traces(textposition='outside', texttemplate='%{text}%')
                fig.update_layout(height=550, yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)

            def rating(r):
                if r['轉換率(%)'] >= 50 and r['入學'] >= 3:
                    return "⭐⭐⭐ 重點經營"
                elif r['轉換率(%)'] >= 30 and r['入學'] >= 2:
                    return "⭐⭐ 持續關注"
                elif r['入學'] >= 1:
                    return "⭐ 一般往來"
                return "🔸 待開發"

            ss['經營建議'] = ss.apply(rating, axis=1)
            st.markdown("### 📋 完整統計（含經營建議）")
            st.dataframe(ss.style.background_gradient(subset=['轉換率(%)'], cmap='RdYlGn')
                        .background_gradient(subset=['入學'], cmap='Blues'),
                        use_container_width=True, height=500)

            if disp_s == "全部學年度" and len(all_yrs) > 1:
                st.markdown("### 📈 重點學校歷年趨勢")
                tops = ss.head(8)['畢業學校'].tolist()
                ys = df_sch[df_sch['畢業學校'].isin(tops)].groupby(
                    ['學年度', '畢業學校']).agg(
                    入學=('最終入學', lambda x: (x == '是').sum())).reset_index()
                fig = px.line(ys, x='學年度', y='入學', color='畢業學校', markers=True)
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("篩選條件下無資料")
    else:
        st.info("請先上傳申請入學資料")


# ============================================================
# TAB 4: 在學穩定度
# ============================================================
with tabs[4]:
    st.markdown("## 📉 在學穩定度分析")
    if st.session_state.retention_data is not None:
        df_r = st.session_state.retention_data.copy()

        r1, r2 = st.columns(2)
        with r1:
            df_rf, disp_r, r_years = year_selector(df_r, "ret")
        with r2:
            rd = ["全部科系"] + sorted(df_r['入學科系'].unique().tolist())
            sel_rd = st.selectbox("科系", rd, key="ret_d")
        if sel_rd != "全部科系":
            df_rf = df_rf[df_rf['入學科系'] == sel_rd]

        if len(df_rf) > 0:
            tot = len(df_rf)
            act = df_rf['目前狀態'].isin(['在學', '畢業']).sum()
            sus = (df_rf['目前狀態'] == '休學').sum()
            drp = (df_rf['目前狀態'] == '退學').sum()
            grd = (df_rf['目前狀態'] == '畢業').sum()

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("👥 總人數", tot)
            m2.metric("📚 在學/畢業", act, f"{act / tot * 100:.1f}%")
            m3.metric("⏸️ 休學", sus, f"{sus / tot * 100:.1f}%")
            m4.metric("❌ 退學", drp, f"{drp / tot * 100:.1f}%")
            m5.metric("🎓 畢業", grd, f"{grd / tot * 100:.1f}%")

            st.markdown("---")
            cl, cr = st.columns(2)

            with cl:
                st.markdown("### 各入學管道穩定度")
                ch = df_rf.groupby('入學管道').agg(
                    總=('學生編號', 'count'),
                    穩定=('目前狀態', lambda x: x.isin(['在學', '畢業']).sum()),
                    休學=('目前狀態', lambda x: (x == '休學').sum()),
                    退學=('目前狀態', lambda x: (x == '退學').sum()),
                ).reset_index()
                ch['穩定率'] = (ch['穩定'] / ch['總'] * 100).round(1)
                ch['休學率'] = (ch['休學'] / ch['總'] * 100).round(1)
                ch['退學率'] = (ch['退學'] / ch['總'] * 100).round(1)
                ch = ch.sort_values('穩定率', ascending=True)

                fig = go.Figure()
                fig.add_trace(go.Bar(y=ch['入學管道'], x=ch['穩定率'], name='穩定率%',
                                    orientation='h', marker_color='#28a745',
                                    text=ch['穩定率'].apply(lambda v: f'{v}%'), textposition='inside'))
                fig.add_trace(go.Bar(y=ch['入學管道'], x=ch['休學率'], name='休學率%',
                                    orientation='h', marker_color='#ffc107',
                                    text=ch['休學率'].apply(lambda v: f'{v}%'), textposition='inside'))
                fig.add_trace(go.Bar(y=ch['入學管道'], x=ch['退學率'], name='退學率%',
                                    orientation='h', marker_color='#dc3545',
                                    text=ch['退學率'].apply(lambda v: f'{v}%'), textposition='inside'))
                fig.update_layout(barmode='stack', height=400, xaxis_title='%', yaxis_title='')
                st.plotly_chart(fig, use_container_width=True)

            with cr:
                st.markdown("### 各科系休退學率")
                dp = df_rf.groupby('入學科系').agg(
                    總=('學生編號', 'count'),
                    休退=('目前狀態', lambda x: x.isin(['休學', '退學']).sum())
                ).reset_index()
                dp['休退學率(%)'] = (dp['休退'] / dp['總'] * 100).round(1)
                dp = dp.sort_values('休退學率(%)', ascending=True)
                fig = px.bar(dp, y='入學科系', x='休退學率(%)', orientation='h',
                            text='休退學率(%)', color='休退學率(%)', color_continuous_scale='RdYlGn_r')
                fig.update_traces(textposition='outside', texttemplate='%{text}%')
                fig.update_layout(height=400, yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("### 📋 休退學原因")
            do = df_rf[df_rf['目前狀態'].isin(['休學', '退學'])]
            if '休退學原因' in do.columns:
                valid = do[do['休退學原因'].astype(str).str.strip() != '']
                if len(valid) > 0:
                    r1c, r2c = st.columns(2)
                    with r1c:
                        rs = valid['休退學原因'].value_counts()
                        fig = px.pie(values=rs.values, names=rs.index, title='原因分布',
                                    color_discrete_sequence=px.colors.qualitative.Set2)
                        fig.update_traces(textposition='inside', textinfo='percent+label')
                        fig.update_layout(height=400)
                        st.plotly_chart(fig, use_container_width=True)
                    with r2c:
                        cross = pd.crosstab(valid['入學管道'], valid['休退學原因'])
                        if len(cross) > 0:
                            fig = px.imshow(cross, text_auto=True, color_continuous_scale='YlOrRd',
                                          title='入學管道 × 原因', aspect='auto')
                            fig.update_layout(height=400)
                            st.plotly_chart(fig, use_container_width=True)

            if disp_r == "全部學年度" and len(r_years) > 1:
                st.markdown("### 📈 歷年休退學率趨勢")
                yc = df_r.copy()
                if sel_rd != "全部科系":
                    yc = yc[yc['入學科系'] == sel_rd]
                yc = yc.groupby(['學年度', '入學管道']).agg(
                    總=('學生編號', 'count'),
                    休退=('目前狀態', lambda x: x.isin(['休學', '退學']).sum())
                ).reset_index()
                yc['休退學率(%)'] = (yc['休退'] / yc['總'] * 100).round(1)
                fig = px.line(yc, x='學年度', y='休退學率(%)', color='入學管道', markers=True)
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("### 📋 各管道完整統計")
            st.dataframe(ch.style
                        .background_gradient(subset=['穩定率'], cmap='RdYlGn')
                        .background_gradient(subset=['退學率'], cmap='RdYlGn_r'),
                        use_container_width=True)
        else:
            st.warning("篩選條件下無資料")
    else:
        st.info("請先上傳在學穩定度資料")


# ============================================================
# TAB 5: 資料檢視與匯出
# ============================================================
with tabs[5]:
    st.markdown("## 📥 資料檢視與匯出")
    dt1, dt2, dt3 = st.tabs(["申請入學資料", "在學穩定度資料", "📋 匯入紀錄"])

    with dt1:
        if st.session_state.applicant_data is not None:
            dv = st.session_state.applicant_data
            st.markdown(f"**{len(dv):,} 筆，{dv['學年度'].nunique()} 個學年度**")

            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                stxt = st.text_input("搜尋學校", key="s_sch")
            with fc2:
                fy = st.multiselect("學年度", sorted(dv['學年度'].unique()), key="s_yr")
            with fc3:
                fd = st.multiselect("科系", sorted(dv['報考科系'].unique()), key="s_dept")

            show = dv.copy()
            if stxt:
                show = show[show['畢業學校'].str.contains(stxt, na=False)]
            if fy:
                show = show[show['學年度'].isin(fy)]
            if fd:
                show = show[show['報考科系'].isin(fd)]

            st.markdown(f"篩選後：**{len(show):,}** 筆")
            st.dataframe(show, use_container_width=True, height=400)

            c1, c2 = st.columns(2)
            with c1:
                st.download_button("📥 CSV", show.to_csv(index=False).encode('utf-8-sig'),
                                 "applicant.csv", "text/csv")
            with c2:
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as w:
                    show.to_excel(w, index=False, sheet_name='申請入學')
                st.download_button("📥 Excel", buf.getvalue(), "applicant.xlsx",
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("尚未載入")

    with dt2:
        if st.session_state.retention_data is not None:
            dv2 = st.session_state.retention_data
            st.markdown(f"**{len(dv2):,} 筆，{dv2['學年度'].nunique()} 個學年度**")

            rc1, rc2 = st.columns(2)
            with rc1:
                ry = st.multiselect("學年度", sorted(dv2['學年度'].unique()), key="r2_yr")
            with rc2:
                rch = st.multiselect("入學管道", sorted(dv2['入學管道'].unique()), key="r2_ch")

            show2 = dv2.copy()
            if ry:
                show2 = show2[show2['學年度'].isin(ry)]
            if rch:
                show2 = show2[show2['入學管道'].isin(rch)]

            st.markdown(f"篩選後：**{len(show2):,}** 筆")
            st.dataframe(show2, use_container_width=True, height=400)

            c3, c4 = st.columns(2)
            with c3:
                st.download_button("📥 CSV", show2.to_csv(index=False).encode('utf-8-sig'),
                                 "retention.csv", "text/csv")
            with c4:
                buf2 = BytesIO()
                with pd.ExcelWriter(buf2, engine='openpyxl') as w:
                    show2.to_excel(w, index=False, sheet_name='在學穩定度')
                st.download_button("📥 Excel", buf2.getvalue(), "retention.xlsx",
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("尚未載入")

    with dt3:
        st.markdown("### 📋 已匯入檔案紀錄")

        st.markdown("#### 申請入學資料")
        if st.session_state.applicant_file_log:
            for i, rec in enumerate(st.session_state.applicant_file_log, 1):
                yrs = ", ".join(str(y) for y in rec['years'])
                st.markdown(f"**{i}.** 📄 `{rec['name']}`　→　{rec['rows']:,} 筆　｜　學年度：{yrs}")
            total = len(st.session_state.applicant_data) if st.session_state.applicant_data is not None else 0
            all_y = sorted(st.session_state.applicant_data['學年度'].unique()) if st.session_state.applicant_data is not None else []
            st.success(f"✅ 合併後：{total:,} 筆（學年度 {', '.join(str(y) for y in all_y)}，重複學生已去除）")
        else:
            st.caption("尚無")

        st.markdown("---")
        st.markdown("#### 在學穩定度資料")
        if st.session_state.retention_file_log:
            for i, rec in enumerate(st.session_state.retention_file_log, 1):
                yrs = ", ".join(str(y) for y in rec['years'])
                st.markdown(f"**{i}.** 📄 `{rec['name']}`　→　{rec['rows']:,} 筆　｜　學年度：{yrs}")
            total = len(st.session_state.retention_data) if st.session_state.retention_data is not None else 0
            all_y = sorted(st.session_state.retention_data['學年度'].unique()) if st.session_state.retention_data is not None else []
            st.success(f"✅ 合併後：{total:,} 筆（學年度 {', '.join(str(y) for y in all_y)}，重複學生已去除）")
        else:
            st.caption("尚無")


# ============================================================
# 頁尾
# ============================================================
st.markdown("---")
st.markdown(
    "<div style='text-align:center;color:#999;font-size:0.85rem;'>"
    "中華醫事科技大學 入學服務處 招生數據分析系統 v2.1 | © 2024</div>",
    unsafe_allow_html=True
)
