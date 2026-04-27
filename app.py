import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import folium
from streamlit_folium import st_folium
import hashlib
import io

# ============================================================
# 頁面設定
# ============================================================
st.set_page_config(
    page_title="中華醫事科技大學 - 招生數據分析系統",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# Session State 初始化
# ============================================================
DEFAULT_STATES = {
    "phase1_data": pd.DataFrame(),   # 一階報名
    "phase2_data": pd.DataFrame(),   # 二階甄試
    "final_data": pd.DataFrame(),    # 最終入學
    "merged_data": pd.DataFrame(),   # 合併後主資料
    "uploaded_hashes": {},
    "upload_log": [],
}
for key, default in DEFAULT_STATES.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ============================================================
# 工具函式
# ============================================================
def safe_int(val):
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return 0

def safe_sort_years(year_series):
    unique_vals = year_series.dropna().unique()
    str_vals = list(set(
        str(int(float(v))) if not isinstance(v, str) else v.strip()
        for v in unique_vals
    ))
    str_vals.sort(key=lambda x: safe_int(x))
    return str_vals

def compute_file_hash(file_bytes, filename):
    return hashlib.md5(filename.encode() + file_bytes).hexdigest()

def clean_column_names(df, phase="phase1"):
    """根據不同階段彈性清理欄位名稱"""
    col_mapping = {}
    for col in df.columns:
        c = str(col).strip().replace(" ", "")
        if "學年" in c:
            col_mapping[col] = "學年度"
        elif any(k in c for k in ["科系", "報考", "系所", "系別", "錄取科系"]):
            col_mapping[col] = "報考科系"
        elif "姓名" in c:
            col_mapping[col] = "姓名"
        elif any(k in c for k in ["畢業", "來源", "學校", "高中"]):
            col_mapping[col] = "畢業學校"
        elif any(k in c for k in ["經緯", "座標", "坐標"]):
            col_mapping[col] = "經緯度"
        elif any(k in c for k in ["身分證", "身份證"]) or "ID" in c.upper():
            col_mapping[col] = "身分證字號"
        elif "緯度" in c or "lat" in c.lower():
            col_mapping[col] = "緯度"
        elif "經度" in c or "lon" in c.lower() or "lng" in c.lower():
            col_mapping[col] = "經度"
        elif any(k in c for k in ["面試", "甄試", "筆試", "成績", "分數", "總分"]):
            col_mapping[col] = col  # 保留原始欄位名
        elif any(k in c for k in ["錄取", "報到", "入學", "狀態"]):
            col_mapping[col] = "入學狀態"
    df = df.rename(columns=col_mapping)
    return df

def standardize_data(df):
    if "學年度" in df.columns:
        df["學年度"] = df["學年度"].apply(
            lambda x: str(int(float(x))) if pd.notna(x) else None
        )
    for col in ["報考科系", "畢業學校", "姓名", "身分證字號"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": None, "None": None, "": None})
    return df

def parse_coordinates(df):
    if "緯度" in df.columns and "經度" in df.columns:
        df["緯度"] = pd.to_numeric(df["緯度"], errors="coerce")
        df["經度"] = pd.to_numeric(df["經度"], errors="coerce")
        return df
    if "經緯度" not in df.columns:
        return df

    lats, lons = [], []
    for val in df["經緯度"]:
        try:
            val_str = str(val).strip()
            if val_str in ("", "nan", "None", "NaN"):
                lats.append(None); lons.append(None); continue
            val_str = val_str.replace("(", "").replace(")", "").replace("（", "").replace("）", "")
            parts = val_str.replace("，", ",").split(",")
            if len(parts) == 2:
                a, b = float(parts[0].strip()), float(parts[1].strip())
                if 21 < a < 26 and 119 < b < 123:
                    lats.append(a); lons.append(b)
                elif 21 < b < 26 and 119 < a < 123:
                    lats.append(b); lons.append(a)
                else:
                    lats.append(None); lons.append(None)
            else:
                lats.append(None); lons.append(None)
        except Exception:
            lats.append(None); lons.append(None)
    df["緯度"] = lats
    df["經度"] = lons
    return df

def merge_all_phases():
    """
    以姓名（+學年度）為主鍵，將三階段資料合併
    一階為主表，二階及最終入學左合併進來
    """
    p1 = st.session_state.phase1_data.copy()
    p2 = st.session_state.phase2_data.copy()
    pf = st.session_state.final_data.copy()

    if p1.empty:
        st.session_state.merged_data = pd.DataFrame()
        return

    # 建立比對鍵
    merge_keys = ["姓名"]
    if "學年度" in p1.columns:
        merge_keys = ["姓名", "學年度"]

    # 標記階段
    p1["一階報名"] = "✅"

    # 合併二階
    if not p2.empty:
        p2_cols = [c for c in p2.columns if c in merge_keys or c not in p1.columns]
        if not all(k in p2.columns for k in merge_keys):
            # 只有姓名可用
            p2_cols = [c for c in p2.columns if c == "姓名" or c not in p1.columns]
            merge_keys_p2 = ["姓名"]
        else:
            merge_keys_p2 = merge_keys

        p2_merge = p2[p2_cols].copy()
        p2_merge["二階甄試"] = "✅"
        # 避免重複欄位衝突，加後綴
        p1 = p1.merge(p2_merge, on=merge_keys_p2, how="left", suffixes=("", "_二階"))
    else:
        p1["二階甄試"] = None

    # 合併最終入學
    if not pf.empty:
        pf_cols = [c for c in pf.columns if c in merge_keys or c not in p1.columns]
        if not all(k in pf.columns for k in merge_keys):
            merge_keys_pf = ["姓名"]
            pf_cols = [c for c in pf.columns if c == "姓名" or c not in p1.columns]
        else:
            merge_keys_pf = merge_keys

        pf_merge = pf[pf_cols].copy()
        pf_merge["最終入學"] = "✅"
        p1 = p1.merge(pf_merge, on=merge_keys_pf, how="left", suffixes=("", "_入學"))
    else:
        p1["最終入學"] = None

    # 填充階段標記
    for stage_col in ["二階甄試", "最終入學"]:
        if stage_col in p1.columns:
            p1[stage_col] = p1[stage_col].fillna("❌")

    # 建立招生漏斗狀態
    def get_stage(row):
        if row.get("最終入學") == "✅":
            return "已入學"
        elif row.get("二階甄試") == "✅":
            return "二階未入學"
        else:
            return "僅一階"
    p1["目前狀態"] = p1.apply(get_stage, axis=1)

    # 去重
    if "身分證字號" in p1.columns and "學年度" in p1.columns:
        p1 = p1.drop_duplicates(subset=["身分證字號", "學年度"], keep="last")
    elif "姓名" in p1.columns and "學年度" in p1.columns:
        p1 = p1.drop_duplicates(subset=["姓名", "學年度"], keep="last")

    st.session_state.merged_data = p1.reset_index(drop=True)

def generate_sample_data():
    """產生三階段示範資料"""
    import random
    random.seed(42)

    schools = [
        ("台南市私立長榮高中", 22.983, 120.215),
        ("台南市立台南一中", 22.993, 120.205),
        ("高雄市立高雄女中", 22.627, 120.305),
        ("高雄市私立義大高中", 22.652, 120.350),
        ("台南市立南寧高中", 22.955, 120.190),
        ("嘉義市立嘉義高中", 23.480, 120.440),
        ("屏東縣立屏東高中", 22.669, 120.488),
        ("台中市立台中一中", 24.148, 120.680),
        ("台南市私立光華高中", 22.978, 120.212),
        ("台南市立永仁高中", 23.025, 120.230),
    ]

    departments = [
        "護理系", "醫學檢驗暨生物技術系", "職業安全衛生系",
        "食品營養系", "幼兒保育系", "長期照護系",
        "視光系", "藥學系", "運動健康與休閒系"
    ]

    surnames = "陳林黃張李王吳劉蔡楊許鄭謝洪曾邱"
    given_names = "雅婷志明家豪淑芬建宏美玲俊傑怡君宗翰佩珊書豪"

    all_phase1 = []
    all_phase2 = []
    all_final = []

    for year in [112, 113, 114]:
        # 一階：大量報名
        n1 = random.randint(200, 300)
        year_names = []
        for _ in range(n1):
            school_name, base_lat, base_lon = random.choice(schools)
            dept = random.choice(departments)
            name = random.choice(surnames) + random.choice(given_names) + random.choice(given_names)
            lat = base_lat + random.uniform(-0.02, 0.02)
            lon = base_lon + random.uniform(-0.02, 0.02)
            fake_id = f"{'ABCDEFGHIJ'[random.randint(0,9)]}{random.randint(100000000, 299999999)}"
            year_names.append(name)
            all_phase1.append({
                "學年度": str(year), "報考科系": dept, "姓名": name,
                "畢業學校": school_name,
                "經緯度": f"{lat:.4f}, {lon:.4f}",
                "身分證字號": fake_id
            })

        # 二階：約60%進入
        n2 = int(n1 * 0.6)
        phase2_names = random.sample(year_names, n2)
        for name in phase2_names:
            all_phase2.append({
                "學年度": str(year), "姓名": name,
                "面試成績": random.randint(50, 100),
                "筆試成績": random.randint(40, 100),
            })

        # 最終入學：約40%
        nf = int(n1 * 0.4)
        final_names = random.sample(phase2_names, min(nf, len(phase2_names)))
        for name in final_names:
            all_final.append({
                "學年度": str(year), "姓名": name,
            })

    return pd.DataFrame(all_phase1), pd.DataFrame(all_phase2), pd.DataFrame(all_final)

# ============================================================
# 側邊欄 — 三階段資料匯入
# ============================================================
with st.sidebar:
    st.markdown("## 🎓 HWU 招生分析系統")
    st.markdown("""
    <div style='background:#FFF3E8; padding:10px; border-radius:8px; margin-bottom:10px;'>
    <small>📌 <b>三階段匯入說明</b><br>
    ① 一階報名：完整欄位（學年度、科系、姓名、學校、座標、身分證）<br>
    ② 二階甄試：姓名 + 成績欄位即可<br>
    ③ 最終入學：僅需姓名（+學年度）即可<br>
    系統會以<b>姓名</b>自動串接比對</small>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # === 一階報名 ===
    st.markdown("### 📋 ① 一階報名資料")
    st.caption("需含：學年度、報考科系、姓名、畢業學校、經緯度、身分證字號")
    phase1_files = st.file_uploader(
        "上傳一階 Excel", type=["xlsx", "xls"],
        accept_multiple_files=True, key="p1_upload"
    )

    # === 二階甄試 ===
    st.markdown("### 📝 ② 二階甄試資料")
    st.caption("需含：姓名（+學年度），可含成績欄位")
    phase2_files = st.file_uploader(
        "上傳二階 Excel", type=["xlsx", "xls"],
        accept_multiple_files=True, key="p2_upload"
    )

    # === 最終入學 ===
    st.markdown("### 🎓 ③ 最終入學資料")
    st.caption("最少僅需：姓名（+學年度）")
    final_files = st.file_uploader(
        "上傳入學 Excel", type=["xlsx", "xls"],
        accept_multiple_files=True, key="pf_upload"
    )

    # 處理上傳
    def process_upload(files, phase_key, phase_label, required_cols=None):
        if not files:
            return
        for uf in files:
            file_bytes = uf.read()
            uf.seek(0)
            fhash = compute_file_hash(file_bytes, uf.name + phase_key)
            if fhash in st.session_state.uploaded_hashes:
                continue
            try:
                new_df = pd.read_excel(uf)
                new_df = clean_column_names(new_df, phase=phase_key)
                new_df = standardize_data(new_df)
                if phase_key == "phase1":
                    new_df = parse_coordinates(new_df)

                # 檢查必要欄位
                if "姓名" not in new_df.columns:
                    st.session_state.upload_log.append(
                        f"❌ {uf.name}（{phase_label}）：缺少「姓名」欄位"
                    )
                    continue

                existing = st.session_state[f"{phase_key}_data"]
                if existing.empty:
                    st.session_state[f"{phase_key}_data"] = new_df
                else:
                    combined = pd.concat([existing, new_df], ignore_index=True)
                    combined = combined.drop_duplicates(keep="last")
                    st.session_state[f"{phase_key}_data"] = combined

                row_count = len(new_df)
                year_info = ""
                if "學年度" in new_df.columns:
                    yrs = safe_sort_years(new_df["學年度"])
                    year_info = f"（{', '.join(yrs)}）"

                st.session_state.uploaded_hashes[fhash] = uf.name
                st.session_state.upload_log.append(
                    f"✅ {uf.name}【{phase_label}】：{row_count} 筆 {year_info}"
                )
            except Exception as e:
                st.session_state.upload_log.append(
                    f"❌ {uf.name}（{phase_label}）：{str(e)}"
                )

    process_upload(phase1_files, "phase1", "一階")
    process_upload(phase2_files, "phase2", "二階")
    process_upload(final_files, "final", "入學")

    # 自動合併
    merge_all_phases()

    # 示範資料
    st.markdown("---")
    if st.button("📊 載入三階段示範資料", use_container_width=True):
        s1, s2, s3 = generate_sample_data()
        s1 = parse_coordinates(s1)
        st.session_state.phase1_data = s1
        st.session_state.phase2_data = s2
        st.session_state.final_data = s3
        merge_all_phases()
        st.session_state.upload_log.append("✅ 三階段示範資料已載入")
        st.rerun()

    # 匯入紀錄
    if st.session_state.upload_log:
        st.markdown("### 📋 匯入紀錄")
        for log in st.session_state.upload_log[-12:]:
            st.caption(log)

    # 資料狀態
    st.markdown("---")
    st.markdown("### 📊 資料狀態")
    p1_n = len(st.session_state.phase1_data)
    p2_n = len(st.session_state.phase2_data)
    pf_n = len(st.session_state.final_data)
    mg_n = len(st.session_state.merged_data)

    st.markdown(f"""
    | 階段 | 筆數 |
    |------|------|
    | ① 一階報名 | **{p1_n:,}** |
    | ② 二階甄試 | **{p2_n:,}** |
    | ③ 最終入學 | **{pf_n:,}** |
    | 🔗 合併後 | **{mg_n:,}** |
    """)

    # 清除
    st.markdown("---")
    if st.button("🗑️ 清除所有資料", use_container_width=True, type="secondary"):
        for key in DEFAULT_STATES:
            st.session_state[key] = DEFAULT_STATES[key].__class__()
        st.rerun()

# ============================================================
# 主畫面
# ============================================================
st.title("🎓 中華醫事科技大學 招生數據分析系統")

data = st.session_state.merged_data

if data.empty:
    st.markdown("""
    <div style='text-align:center; padding:50px 20px;'>
        <h2>👈 請先從左側匯入各階段資料</h2>
        <div style='display:flex; justify-content:center; gap:30px; margin-top:30px;'>
            <div style='background:#FFF3E8; padding:20px; border-radius:12px; width:220px;'>
                <h3>📋 一階報名</h3>
                <p style='font-size:13px; color:gray;'>學年度、科系、姓名<br>學校、座標、身分證</p>
            </div>
            <div style='background:#FFF3E8; padding:20px; border-radius:12px; width:220px;'>
                <h3>📝 二階甄試</h3>
                <p style='font-size:13px; color:gray;'>姓名、學年度<br>+ 成績欄位（選填）</p>
            </div>
            <div style='background:#FFF3E8; padding:20px; border-radius:12px; width:220px;'>
                <h3>🎓 最終入學</h3>
                <p style='font-size:13px; color:gray;'>僅需：姓名<br>（+學年度更精準）</p>
            </div>
        </div>
        <p style='color:gray; margin-top:20px;'>或點擊「載入三階段示範資料」快速體驗</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ============================================================
# 分頁
# ============================================================
tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🔄 招生漏斗",
    "🗺️ 地圖視覺化",
    "📊 科系招生分析",
    "🏫 來源學校分析",
    "📈 跨年度比較",
    "🔍 資料檢視與匯出"
])

# ============================================================
# TAB 0: 招生漏斗分析
# ============================================================
with tab0:
    st.header("🔄 招生漏斗分析")

    if "學年度" in data.columns:
        yr_options = ["全部"] + safe_sort_years(data["學年度"])
        sel_yr_funnel = st.selectbox("選擇學年度", yr_options, key="funnel_yr")
    else:
        sel_yr_funnel = "全部"

    funnel_data = data.copy()
    if sel_yr_funnel != "全部" and "學年度" in funnel_data.columns:
        funnel_data = funnel_data[funnel_data["學年度"] == sel_yr_funnel]

    # 統計各階段
    n_total = len(funnel_data)
    n_phase1 = len(funnel_data[funnel_data.get("一階報名", pd.Series()) == "✅"]) if "一階報名" in funnel_data.columns else n_total
    n_phase2 = len(funnel_data[funnel_data.get("二階甄試", pd.Series()) == "✅"]) if "二階甄試" in funnel_data.columns else 0
    n_final = len(funnel_data[funnel_data.get("最終入學", pd.Series()) == "✅"]) if "最終入學" in funnel_data.columns else 0

    # KPI
    kpi_cols = st.columns(4)
    with kpi_cols[0]:
        st.metric("📋 一階報名", f"{n_phase1:,} 人")
    with kpi_cols[1]:
        rate_21 = n_phase2 / n_phase1 * 100 if n_phase1 > 0 else 0
        st.metric("📝 二階甄試", f"{n_phase2:,} 人", f"轉換率 {rate_21:.1f}%")
    with kpi_cols[2]:
        rate_f1 = n_final / n_phase1 * 100 if n_phase1 > 0 else 0
        st.metric("🎓 最終入學", f"{n_final:,} 人", f"總錄取率 {rate_f1:.1f}%")
    with kpi_cols[3]:
        rate_f2 = n_final / n_phase2 * 100 if n_phase2 > 0 else 0
        st.metric("🎯 二階→入學", f"{rate_f2:.1f}%")

    # 漏斗圖
    col_fun1, col_fun2 = st.columns([2, 1])

    with col_fun1:
        fig_funnel = go.Figure(go.Funnel(
            y=["一階報名", "二階甄試", "最終入學"],
            x=[n_phase1, n_phase2, n_final],
            textinfo="value+percent initial",
            marker=dict(color=["#FDD7B4", "#E8792F", "#B84500"]),
            connector=dict(line=dict(color="#E8792F", width=2)),
        ))
        fig_funnel.update_layout(title="招生漏斗", height=400)
        st.plotly_chart(fig_funnel, use_container_width=True)

    with col_fun2:
        # 各階段流失分析
        st.markdown("**📉 流失分析**")
        lost_12 = n_phase1 - n_phase2
        lost_2f = n_phase2 - n_final
        st.markdown(f"""
        | 階段 | 流失人數 | 流失率 |
        |------|---------|--------|
        | 一階→二階 | {lost_12:,} | {lost_12/n_phase1*100:.1f}% |
        | 二階→入學 | {lost_2f:,} | {lost_2f/n_phase2*100:.1f}% |
        """) if n_phase1 > 0 and n_phase2 > 0 else st.info("需要多階段資料")

    # 分科系漏斗
    if "報考科系" in funnel_data.columns:
        st.subheader("📊 各科系招生漏斗")
        dept_funnel = []
        for dept in sorted(funnel_data["報考科系"].dropna().unique()):
            d = funnel_data[funnel_data["報考科系"] == dept]
            d1 = len(d)
            d2 = len(d[d["二階甄試"] == "✅"]) if "二階甄試" in d.columns else 0
            df_ = len(d[d["最終入學"] == "✅"]) if "最終入學" in d.columns else 0
            dept_funnel.append({
                "科系": dept,
                "一階報名": d1,
                "二階甄試": d2,
                "最終入學": df_,
                "一→二轉換率": f"{d2/d1*100:.1f}%" if d1 > 0 else "0%",
                "總錄取率": f"{df_/d1*100:.1f}%" if d1 > 0 else "0%",
            })
        dept_funnel_df = pd.DataFrame(dept_funnel)
        st.dataframe(dept_funnel_df, use_container_width=True, hide_index=True)

        # 分科系柱狀比較
        fig_dept_funnel = go.Figure()
        fig_dept_funnel.add_trace(go.Bar(
            name="一階報名", x=dept_funnel_df["科系"], y=dept_funnel_df["一階報名"],
            marker_color="#FDD7B4"
        ))
        fig_dept_funnel.add_trace(go.Bar(
            name="二階甄試", x=dept_funnel_df["科系"], y=dept_funnel_df["二階甄試"],
            marker_color="#E8792F"
        ))
        fig_dept_funnel.add_trace(go.Bar(
            name="最終入學", x=dept_funnel_df["科系"], y=dept_funnel_df["最終入學"],
            marker_color="#B84500"
        ))
        fig_dept_funnel.update_layout(barmode="group", title="各科系三階段比較", height=450)
        st.plotly_chart(fig_dept_funnel, use_container_width=True)

    # 跨年度漏斗趨勢
    if "學年度" in data.columns and data["學年度"].nunique() > 1 and sel_yr_funnel == "全部":
        st.subheader("📈 歷年漏斗趨勢")
        yearly_funnel = []
        for yr in safe_sort_years(data["學年度"]):
            yd = data[data["學年度"] == yr]
            y1 = len(yd)
            y2 = len(yd[yd["二階甄試"] == "✅"]) if "二階甄試" in yd.columns else 0
            yf = len(yd[yd["最終入學"] == "✅"]) if "最終入學" in yd.columns else 0
            yearly_funnel.append({"學年度": yr, "一階報名": y1, "二階甄試": y2, "最終入學": yf})

        yf_df = pd.DataFrame(yearly_funnel)
        fig_yf = go.Figure()
        fig_yf.add_trace(go.Scatter(x=yf_df["學年度"], y=yf_df["一階報名"],
                                     mode="lines+markers", name="一階報名", line=dict(color="#FDD7B4", width=3)))
        fig_yf.add_trace(go.Scatter(x=yf_df["學年度"], y=yf_df["二階甄試"],
                                     mode="lines+markers", name="二階甄試", line=dict(color="#E8792F", width=3)))
        fig_yf.add_trace(go.Scatter(x=yf_df["學年度"], y=yf_df["最終入學"],
                                     mode="lines+markers", name="最終入學", line=dict(color="#B84500", width=3)))
        fig_yf.update_layout(title="歷年三階段人數趨勢", height=400)
        st.plotly_chart(fig_yf, use_container_width=True)

# ============================================================
# TAB 1: 地圖視覺化
# ============================================================
with tab1:
    st.header("🗺️ 報考生分布地圖")

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        if "學年度" in data.columns:
            year_options = ["全部"] + safe_sort_years(data["學年度"])
            sel_year_map = st.selectbox("選擇學年度", year_options, key="map_year")
        else:
            sel_year_map = "全部"
    with col_f2:
        if "報考科系" in data.columns:
            dept_options = ["全部"] + sorted(data["報考科系"].dropna().unique().tolist())
            sel_dept_map = st.selectbox("選擇科系", dept_options, key="map_dept")
        else:
            sel_dept_map = "全部"
    with col_f3:
        if "目前狀態" in data.columns:
            stage_options = ["全部"] + sorted(data["目前狀態"].dropna().unique().tolist())
            sel_stage_map = st.selectbox("招生階段", stage_options, key="map_stage")
        else:
            sel_stage_map = "全部"

    map_data = data.copy()
    if sel_year_map != "全部" and "學年度" in map_data.columns:
        map_data = map_data[map_data["學年度"] == sel_year_map]
    if sel_dept_map != "全部" and "報考科系" in map_data.columns:
        map_data = map_data[map_data["報考科系"] == sel_dept_map]
    if sel_stage_map != "全部" and "目前狀態" in map_data.columns:
        map_data = map_data[map_data["目前狀態"] == sel_stage_map]

    has_coords = "緯度" in map_data.columns and "經度" in map_data.columns
    if has_coords:
        valid_coords = map_data.dropna(subset=["緯度", "經度"]).copy()
        valid_coords["緯度"] = pd.to_numeric(valid_coords["緯度"], errors="coerce")
        valid_coords["經度"] = pd.to_numeric(valid_coords["經度"], errors="coerce")
        valid_coords = valid_coords.dropna(subset=["緯度", "經度"])

        if len(valid_coords) > 0:
            col_m1, col_m2 = st.columns([3, 1])
            with col_m1:
                m = folium.Map(location=[23.5, 120.5], zoom_start=8, tiles="CartoDB positron")
                folium.Marker(
                    location=[22.9908, 120.2133],
                    popup="中華醫事科技大學",
                    icon=folium.Icon(color="red", icon="star", prefix="fa"),
                ).add_to(m)

                color_map = {"已入學": "#B84500", "二階未入學": "#E8792F", "僅一階": "#FDD7B4"}
                display_coords = valid_coords if len(valid_coords) <= 1000 else valid_coords.sample(1000, random_state=42)
                if len(valid_coords) > 1000:
                    st.caption(f"⚡ 地圖顯示抽樣 1,000 筆（共 {len(valid_coords):,} 筆）")

                for _, row in display_coords.iterrows():
                    status = row.get("目前狀態", "僅一階")
                    color = color_map.get(status, "#E8792F")
                    popup_parts = []
                    for field in ["畢業學校", "報考科系", "學年度", "目前狀態"]:
                        if field in row.index and pd.notna(row.get(field)):
                            popup_parts.append(f"{field}：{row[field]}")
                    folium.CircleMarker(
                        location=[float(row["緯度"]), float(row["經度"])],
                        radius=5, color=color, fill=True, fill_color=color, fill_opacity=0.7,
                        popup=folium.Popup("<br>".join(popup_parts), max_width=200) if popup_parts else None,
                    ).add_to(m)

                st_folium(m, width=None, height=550, use_container_width=True)

            with col_m2:
                st.metric("篩選結果", f"{len(map_data):,} 人")
                st.metric("有座標", f"{len(valid_coords):,} 人")
                coverage = len(valid_coords) / len(map_data) * 100 if len(map_data) > 0 else 0
                st.metric("座標涵蓋率", f"{coverage:.1f}%")

                if "目前狀態" in valid_coords.columns:
                    st.markdown("**📊 階段分布**")
                    for status, cnt in valid_coords["目前狀態"].value_counts().items():
                        st.caption(f"{'🟤' if status=='已入學' else '🟠' if status=='二階未入學' else '🟡'} {status}：{cnt}")

                if "畢業學校" in valid_coords.columns:
                    st.markdown("**📍 主要來源學校**")
                    for school, cnt in valid_coords["畢業學校"].value_counts().head(8).items():
                        st.caption(f"• {school}：{cnt}")
        else:
            st.warning("⚠️ 篩選後無有效座標資料")
    else:
        st.warning("⚠️ 無座標欄位，請在一階資料中包含經緯度")

# ============================================================
# TAB 2: 科系招生分析
# ============================================================
with tab2:
    st.header("📊 科系招生分析")

    if "報考科系" not in data.columns:
        st.warning("資料中未包含「報考科系」欄位")
    else:
        if "學年度" in data.columns:
            yr_opts_t2 = ["全部"] + safe_sort_years(data["學年度"])
            sel_year_t2 = st.selectbox("選擇學年度", yr_opts_t2, key="t2_year")
        else:
            sel_year_t2 = "全部"

        t2_data = data.copy()
        if sel_year_t2 != "全部" and "學年度" in t2_data.columns:
            t2_data = t2_data[t2_data["學年度"] == sel_year_t2]

        dept_counts = t2_data["報考科系"].value_counts().reset_index()
        dept_counts.columns = ["科系", "報考人數"]

        col_c1, col_c2 = st.columns(2)
        with col_c1:
            fig_bar = px.bar(
                dept_counts, x="報考人數", y="科系", orientation="h",
                title="各科系報考人數", color="報考人數",
                color_continuous_scale=["#FDD7B4", "#E8792F", "#B84500"],
            )
            fig_bar.update_layout(yaxis={"categoryorder": "total ascending"},
                                  height=max(400, len(dept_counts) * 35), showlegend=False)
            st.plotly_chart(fig_bar, use_container_width=True)

        with col_c2:
            # 各科系入學率
            if "最終入學" in t2_data.columns:
                dept_enroll = []
                for dept in dept_counts["科系"]:
                    dd = t2_data[t2_data["報考科系"] == dept]
                    total = len(dd)
                    enrolled = len(dd[dd["最終入學"] == "✅"])
                    dept_enroll.append({"科系": dept, "報名": total, "入學": enrolled,
                                        "入學率": enrolled / total * 100 if total > 0 else 0})
                de_df = pd.DataFrame(dept_enroll)

                fig_enroll = go.Figure()
                fig_enroll.add_trace(go.Bar(
                    name="報名", x=de_df["科系"], y=de_df["報名"], marker_color="#FDD7B4"))
                fig_enroll.add_trace(go.Bar(
                    name="入學", x=de_df["科系"], y=de_df["入學"], marker_color="#B84500"))
                fig_enroll.update_layout(barmode="overlay", title="各科系報名 vs 入學", height=max(400, len(dept_counts) * 35))
                st.plotly_chart(fig_enroll, use_container_width=True)
            else:
                fig_pie = px.pie(dept_counts, values="報考人數", names="科系", title="科系佔比分布",
                                 color_discrete_sequence=px.colors.sequential.Oranges_r)
                fig_pie.update_layout(height=max(400, len(dept_counts) * 35))
                st.plotly_chart(fig_pie, use_container_width=True)

        # 歷年趨勢
        if "學年度" in data.columns and data["學年度"].nunique() > 1 and sel_year_t2 == "全部":
            st.subheader("📈 各科系歷年趨勢")
            yearly_dept = data.groupby(["學年度", "報考科系"]).size().reset_index(name="人數")
            yearly_dept["排序鍵"] = yearly_dept["學年度"].apply(safe_int)
            yearly_dept = yearly_dept.sort_values("排序鍵")
            fig_line = px.line(yearly_dept, x="學年度", y="人數", color="報考科系",
                               markers=True, title="各科系歷年報考人數趨勢")
            fig_line.update_layout(height=450)
            st.plotly_chart(fig_line, use_container_width=True)

        # KPI
        st.subheader("📋 科系統計摘要")
        kpi_cols = st.columns(4)
        with kpi_cols[0]:
            st.metric("總報考人數", f"{len(t2_data):,}")
        with kpi_cols[1]:
            st.metric("科系數", f"{t2_data['報考科系'].nunique()}")
        with kpi_cols[2]:
            top_dept = dept_counts.iloc[0]["科系"] if len(dept_counts) > 0 else "-"
            st.metric("最熱門科系", top_dept)
        with kpi_cols[3]:
            if "最終入學" in t2_data.columns:
                total_enrolled = len(t2_data[t2_data["最終入學"] == "✅"])
                overall_rate = total_enrolled / len(t2_data) * 100 if len(t2_data) > 0 else 0
                st.metric("總入學率", f"{overall_rate:.1f}%")
            else:
                avg_count = dept_counts["報考人數"].mean() if len(dept_counts) > 0 else 0
                st.metric("平均每科系", f"{avg_count:.0f} 人")

# ============================================================
# TAB 3: 來源學校分析
# ============================================================
with tab3:
    st.header("🏫 來源學校分析")

    if "畢業學校" not in data.columns:
        st.warning("資料中未包含「畢業學校」欄位")
    else:
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            if "學年度" in data.columns:
                yr_opts_t3 = ["全部"] + safe_sort_years(data["學年度"])
                sel_year_t3 = st.selectbox("選擇學年度", yr_opts_t3, key="t3_year")
            else:
                sel_year_t3 = "全部"
        with col_f2:
            top_n = st.slider("顯示前 N 所學校", 5, 30, 15, key="t3_topn")
        with col_f3:
            if "目前狀態" in data.columns:
                stage_opts_t3 = ["全部"] + sorted(data["目前狀態"].dropna().unique().tolist())
                sel_stage_t3 = st.selectbox("階段篩選", stage_opts_t3, key="t3_stage")
            else:
                sel_stage_t3 = "全部"

        t3_data = data.copy()
        if sel_year_t3 != "全部" and "學年度" in t3_data.columns:
            t3_data = t3_data[t3_data["學年度"] == sel_year_t3]
        if sel_stage_t3 != "全部" and "目前狀態" in t3_data.columns:
            t3_data = t3_data[t3_data["目前狀態"] == sel_stage_t3]

        school_counts = t3_data["畢業學校"].value_counts().head(top_n).reset_index()
        school_counts.columns = ["畢業學校", "人數"]

        fig_school = px.bar(
            school_counts, x="人數", y="畢業學校", orientation="h",
            title=f"Top {top_n} 來源學校" + (f"（{sel_stage_t3}）" if sel_stage_t3 != "全部" else ""),
            color="人數", color_continuous_scale=["#FDD7B4", "#E8792F", "#B84500"],
        )
        fig_school.update_layout(yaxis={"categoryorder": "total ascending"},
                                  height=max(400, top_n * 30), showlegend=False)
        st.plotly_chart(fig_school, use_container_width=True)

        # 各學校入學轉換率
        if "最終入學" in t3_data.columns:
            st.subheader("🎯 各學校入學轉換率")
            school_conv = []
            for school in school_counts["畢業學校"]:
                sd = t3_data[t3_data["畢業學校"] == school]
                total = len(sd)
                enrolled = len(sd[sd["最終入學"] == "✅"])
                school_conv.append({
                    "學校": school, "報名": total, "入學": enrolled,
                    "轉換率": f"{enrolled/total*100:.1f}%" if total > 0 else "0%"
                })
            st.dataframe(pd.DataFrame(school_conv), use_container_width=True, hide_index=True)

        # 交叉分析
        if "報考科系" in t3_data.columns:
            st.subheader("🔀 來源學校 × 科系 交叉分析")
            top_schools_list = school_counts["畢業學校"].tolist()
            cross_data = t3_data[t3_data["畢業學校"].isin(top_schools_list)]
            if len(cross_data) > 0:
                cross_table = pd.crosstab(cross_data["畢業學校"], cross_data["報考科系"])
                fig_heat = px.imshow(cross_table, title="來源學校 vs 報考科系（人數）",
                                     color_continuous_scale="Oranges", aspect="auto")
                fig_heat.update_layout(height=max(400, top_n * 28))
                st.plotly_chart(fig_heat, use_container_width=True)

        # 管理建議
        st.subheader("⭐ 來源學校管理建議")
        all_school_counts = t3_data["畢業學校"].value_counts()
        total = len(t3_data)
        recs = []
        cumulative = 0
        for school, cnt in all_school_counts.items():
            cumulative += cnt
            ratio = cnt / total * 100 if total > 0 else 0
            cum_ratio = cumulative / total * 100 if total > 0 else 0
            level = "⭐⭐⭐ 重點經營" if cum_ratio <= 50 else ("⭐⭐ 持續關注" if cum_ratio <= 80 else "⭐ 一般維護")
            recs.append({"學校": school, "人數": cnt, "佔比(%)": round(ratio, 1),
                         "累積佔比(%)": round(cum_ratio, 1), "建議等級": level})
        st.dataframe(pd.DataFrame(recs).head(top_n), use_container_width=True, hide_index=True)

# ============================================================
# TAB 4: 跨年度比較
# ============================================================
with tab4:
    st.header("📈 跨年度比較分析")

    if "學年度" not in data.columns or data["學年度"].nunique() < 2:
        st.info("💡 需要至少兩個學年度的資料才能進行跨年比較。")
    else:
        years_sorted = safe_sort_years(data["學年度"])

        # 總人數趨勢
        st.subheader("📊 歷年招生趨勢")
        yearly_stats = []
        for yr in years_sorted:
            yd = data[data["學年度"] == yr]
            row = {"學年度": yr, "一階報名": len(yd)}
            if "二階甄試" in yd.columns:
                row["二階甄試"] = len(yd[yd["二階甄試"] == "✅"])
            if "最終入學" in yd.columns:
                row["最終入學"] = len(yd[yd["最終入學"] == "✅"])
            yearly_stats.append(row)
        ys_df = pd.DataFrame(yearly_stats)

        fig_trend = go.Figure()
        for col_name, color in [("一階報名", "#FDD7B4"), ("二階甄試", "#E8792F"), ("最終入學", "#B84500")]:
            if col_name in ys_df.columns:
                fig_trend.add_trace(go.Bar(name=col_name, x=ys_df["學年度"], y=ys_df[col_name],
                                           marker_color=color, text=ys_df[col_name], textposition="outside"))
        fig_trend.update_layout(barmode="group", title="歷年各階段人數", height=400)
        st.plotly_chart(fig_trend, use_container_width=True)

        # 年度增減
        col_inc1, col_inc2 = st.columns([1, 2])
        with col_inc1:
            st.markdown("**📊 年度增減**")
            for i in range(1, len(ys_df)):
                prev_n = ys_df.iloc[i-1]["一階報名"]
                curr_n = ys_df.iloc[i]["一階報名"]
                diff = curr_n - prev_n
                pct = diff / prev_n * 100 if prev_n > 0 else 0
                st.metric(f"{ys_df.iloc[i]['學年度']} 學年",
                          f"{curr_n:,} 人", f"{diff:+d}（{pct:+.1f}%）")

        with col_inc2:
            if "報考科系" in data.columns and len(years_sorted) >= 2:
                st.markdown("**📊 科系增減比較**")
                year_a, year_b = years_sorted[-2], years_sorted[-1]
                da = data[data["學年度"] == year_a]["報考科系"].value_counts()
                db = data[data["學年度"] == year_b]["報考科系"].value_counts()
                all_depts = sorted(set(da.index) | set(db.index))
                comparison = []
                for dept in all_depts:
                    a_val, b_val = int(da.get(dept, 0)), int(db.get(dept, 0))
                    diff = b_val - a_val
                    comparison.append({"科系": dept, f"{year_a}": a_val, f"{year_b}": b_val,
                                       "增減": diff, "趨勢": "📈" if diff > 0 else ("📉" if diff < 0 else "➡️")})
                st.dataframe(pd.DataFrame(comparison).sort_values("增減", ascending=False),
                             use_container_width=True, hide_index=True)

        # 來源學校趨勢
        if "畢業學校" in data.columns:
            st.subheader("🏫 來源學校年度變化")
            top_schools_overall = data["畢業學校"].value_counts().head(10).index.tolist()
            school_year = data.groupby(["學年度", "畢業學校"]).size().reset_index(name="人數")
            school_year_top = school_year[school_year["畢業學校"].isin(top_schools_overall)].copy()
            school_year_top["排序鍵"] = school_year_top["學年度"].apply(safe_int)
            school_year_top = school_year_top.sort_values("排序鍵")
            fig_st = px.line(school_year_top, x="學年度", y="人數", color="畢業學校",
                             markers=True, title="Top 10 來源學校歷年趨勢")
            fig_st.update_layout(height=450)
            st.plotly_chart(fig_st, use_container_width=True)

# ============================================================
# TAB 5: 資料檢視與匯出
# ============================================================
with tab5:
    st.header("🔍 資料檢視與匯出")

    # 顯示各階段原始資料
    with st.expander("📋 各階段原始資料", expanded=False):
        sub_tabs = st.tabs(["一階報名", "二階甄試", "最終入學"])
        with sub_tabs[0]:
            if not st.session_state.phase1_data.empty:
                st.caption(f"共 {len(st.session_state.phase1_data):,} 筆")
                st.dataframe(st.session_state.phase1_data.head(100), use_container_width=True, hide_index=True)
            else:
                st.info("尚無一階資料")
        with sub_tabs[1]:
            if not st.session_state.phase2_data.empty:
                st.caption(f"共 {len(st.session_state.phase2_data):,} 筆")
                st.dataframe(st.session_state.phase2_data.head(100), use_container_width=True, hide_index=True)
            else:
                st.info("尚無二階資料")
        with sub_tabs[2]:
            if not st.session_state.final_data.empty:
                st.caption(f"共 {len(st.session_state.final_data):,} 筆")
                st.dataframe(st.session_state.final_data.head(100), use_container_width=True, hide_index=True)
            else:
                st.info("尚無入學資料")

    # 合併後資料篩選
    st.subheader("🔗 合併後資料")

    col_e1, col_e2, col_e3, col_e4 = st.columns(4)
    export_data = data.copy()

    with col_e1:
        if "學年度" in data.columns:
            yr_options = ["全部"] + safe_sort_years(data["學年度"])
            sel_yr_exp = st.selectbox("學年度", yr_options, key="exp_yr")
            if sel_yr_exp != "全部":
                export_data = export_data[export_data["學年度"] == sel_yr_exp]
    with col_e2:
        if "報考科系" in data.columns:
            dept_opts = ["全部"] + sorted(data["報考科系"].dropna().unique().tolist())
            sel_dept_exp = st.selectbox("科系", dept_opts, key="exp_dept")
            if sel_dept_exp != "全部":
                export_data = export_data[export_data["報考科系"] == sel_dept_exp]
    with col_e3:
        if "目前狀態" in data.columns:
            stage_opts = ["全部"] + sorted(data["目前狀態"].dropna().unique().tolist())
            sel_stage_exp = st.selectbox("階段", stage_opts, key="exp_stage")
            if sel_stage_exp != "全部":
                export_data = export_data[export_data["目前狀態"] == sel_stage_exp]
    with col_e4:
        if "畢業學校" in data.columns:
            search_school = st.text_input("搜尋學校", key="exp_school")
            if search_school:
                export_data = export_data[export_data["畢業學校"].str.contains(search_school, na=False)]

    st.caption(f"篩選結果：{len(export_data):,} 筆")

    # 隱藏敏感欄位
    hide_cols = ["身分證字號"]
    display_cols = [c for c in export_data.columns if c not in hide_cols]
    st.dataframe(export_data[display_cols], use_container_width=True, hide_index=True, height=400)

    # 匯出
    st.subheader("📥 匯出資料")
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        csv_buf = io.BytesIO()
        export_data.to_csv(csv_buf, index=False, encoding="utf-8-sig")
        st.download_button("⬇️ 下載 CSV", csv_buf.getvalue(),
                           "招生分析資料.csv", "text/csv", use_container_width=True)
    with col_dl2:
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            export_data.to_excel(writer, index=False, sheet_name="招生資料")
        st.download_button("⬇️ 下載 Excel", excel_buf.getvalue(),
                           "招生分析資料.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

    # 資料品質報告
    st.subheader("📋 資料品質報告")
    quality = []
    for col in data.columns:
        non_null = data[col].notna().sum()
        total_rows = len(data)
        quality.append({
            "欄位": col, "非空筆數": non_null, "總筆數": total_rows,
            "完整率": f"{non_null/total_rows*100:.1f}%" if total_rows > 0 else "0%",
            "空值數": total_rows - non_null,
        })
    st.dataframe(pd.DataFrame(quality), use_container_width=True, hide_index=True)
