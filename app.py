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
    "phase1_data": pd.DataFrame(),
    "phase2_data": pd.DataFrame(),
    "final_data": pd.DataFrame(),
    "merged_data": pd.DataFrame(),
    "uploaded_hashes": {},
    "upload_log": [],
}
for key, default in DEFAULT_STATES.items():
    if key not in st.session_state:
        if isinstance(default, pd.DataFrame):
            st.session_state[key] = default.copy()
        elif isinstance(default, dict):
            st.session_state[key] = {}
        elif isinstance(default, list):
            st.session_state[key] = []
        else:
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

def deduplicate_columns(df):
    """處理重複欄位名稱：重複的欄位加上 _2, _3 等後綴"""
    cols = list(df.columns)
    seen = {}
    new_cols = []
    for c in cols:
        c_str = str(c).strip()
        if c_str in seen:
            seen[c_str] += 1
            new_cols.append("{}_{}".format(c_str, seen[c_str]))
        else:
            seen[c_str] = 1
            new_cols.append(c_str)
    df.columns = new_cols
    return df

def clean_column_names(df, phase="phase1"):
    # 先處理重複欄位
    df = deduplicate_columns(df)

    col_mapping = {}
    for col in df.columns:
        c = str(col).strip().replace(" ", "")
        if "學年" in c and col not in col_mapping.values():
            col_mapping[col] = "學年度"
        elif any(k in c for k in ["科系", "報考", "系所", "系別", "錄取科系"]) and "學年" not in c:
            if "報考科系" not in col_mapping.values():
                col_mapping[col] = "報考科系"
        elif "姓名" in c and "姓名" not in col_mapping.values():
            col_mapping[col] = "姓名"
        elif any(k in c for k in ["畢業", "來源", "高中"]) and "學校" in c:
            if "畢業學校" not in col_mapping.values():
                col_mapping[col] = "畢業學校"
        elif "學校" in c and "畢業學校" not in col_mapping.values() and "科系" not in c:
            col_mapping[col] = "畢業學校"
        elif any(k in c for k in ["經緯", "座標", "坐標"]):
            if "經緯度" not in col_mapping.values():
                col_mapping[col] = "經緯度"
        elif any(k in c for k in ["身分證", "身份證"]) or "ID" in c.upper():
            if "身分證字號" not in col_mapping.values():
                col_mapping[col] = "身分證字號"
        elif ("緯度" in c or c.lower() == "lat") and "經" not in c:
            if "緯度" not in col_mapping.values():
                col_mapping[col] = "緯度"
        elif ("經度" in c or c.lower() in ("lon", "lng")) and "緯" not in c:
            if "經度" not in col_mapping.values():
                col_mapping[col] = "經度"
        elif any(k in c for k in ["面試", "甄試", "筆試", "成績", "分數", "總分"]):
            pass  # 保留原始欄位名
        elif any(k in c for k in ["錄取", "報到", "入學", "狀態"]):
            if "入學狀態" not in col_mapping.values():
                col_mapping[col] = "入學狀態"
    df = df.rename(columns=col_mapping)

    # 再次移除重命名後可能出現的重複
    df = deduplicate_columns(df)
    return df

def standardize_data(df):
    if "學年度" in df.columns:
        # 確保是 Series
        col_data = df["學年度"]
        if isinstance(col_data, pd.DataFrame):
            col_data = col_data.iloc[:, 0]
        df = df.copy()
        df["學年度"] = col_data.apply(
            lambda x: str(int(float(x))) if pd.notna(x) else None
        )

    for col_name in ["報考科系", "畢業學校", "姓名", "身分證字號"]:
        if col_name in df.columns:
            col_data = df[col_name]
            # 如果同名欄位多於一個，取第一個
            if isinstance(col_data, pd.DataFrame):
                col_data = col_data.iloc[:, 0]
            df = df.copy()
            df[col_name] = col_data.astype(str).str.strip()
            df[col_name] = df[col_name].replace({"nan": None, "None": None, "": None})
    return df

def parse_coordinates(df):
    if "緯度" in df.columns and "經度" in df.columns:
        lat_col = df["緯度"]
        lon_col = df["經度"]
        if isinstance(lat_col, pd.DataFrame):
            lat_col = lat_col.iloc[:, 0]
        if isinstance(lon_col, pd.DataFrame):
            lon_col = lon_col.iloc[:, 0]
        df = df.copy()
        df["緯度"] = pd.to_numeric(lat_col, errors="coerce")
        df["經度"] = pd.to_numeric(lon_col, errors="coerce")
        return df

    if "經緯度" not in df.columns:
        return df

    coord_col = df["經緯度"]
    if isinstance(coord_col, pd.DataFrame):
        coord_col = coord_col.iloc[:, 0]

    lats, lons = [], []
    for val in coord_col:
        try:
            val_str = str(val).strip()
            if val_str in ("", "nan", "None", "NaN"):
                lats.append(None)
                lons.append(None)
                continue
            for ch in ["(", ")", "（", "）", "「", "」"]:
                val_str = val_str.replace(ch, "")
            parts = val_str.replace("，", ",").split(",")
            if len(parts) == 2:
                a = float(parts[0].strip())
                b = float(parts[1].strip())
                if 21 < a < 26 and 119 < b < 123:
                    lats.append(a)
                    lons.append(b)
                elif 21 < b < 26 and 119 < a < 123:
                    lats.append(b)
                    lons.append(a)
                else:
                    lats.append(None)
                    lons.append(None)
            else:
                lats.append(None)
                lons.append(None)
        except Exception:
            lats.append(None)
            lons.append(None)

    df = df.copy()
    df["緯度"] = lats
    df["經度"] = lons
    return df

def merge_all_phases():
    p1 = st.session_state.phase1_data.copy()
    p2 = st.session_state.phase2_data.copy()
    pf = st.session_state.final_data.copy()

    if p1.empty:
        st.session_state.merged_data = pd.DataFrame()
        return

    merge_keys = ["姓名"]
    if "學年度" in p1.columns:
        merge_keys = ["姓名", "學年度"]

    p1["一階報名"] = "✅"

    if not p2.empty:
        available_keys_p2 = [k for k in merge_keys if k in p2.columns]
        if not available_keys_p2:
            available_keys_p2 = ["姓名"] if "姓名" in p2.columns else []
        if available_keys_p2:
            p2_extra_cols = [c for c in p2.columns if c not in p1.columns]
            p2_cols = available_keys_p2 + p2_extra_cols
            p2_merge = p2[p2_cols].drop_duplicates(subset=available_keys_p2, keep="last").copy()
            p2_merge["二階甄試"] = "✅"
            p1 = p1.merge(p2_merge, on=available_keys_p2, how="left", suffixes=("", "_二階"))
        else:
            p1["二階甄試"] = None
    else:
        p1["二階甄試"] = None

    if not pf.empty:
        available_keys_pf = [k for k in merge_keys if k in pf.columns]
        if not available_keys_pf:
            available_keys_pf = ["姓名"] if "姓名" in pf.columns else []
        if available_keys_pf:
            pf_extra_cols = [c for c in pf.columns if c not in p1.columns]
            pf_cols = available_keys_pf + pf_extra_cols
            pf_merge = pf[pf_cols].drop_duplicates(subset=available_keys_pf, keep="last").copy()
            pf_merge["最終入學"] = "✅"
            p1 = p1.merge(pf_merge, on=available_keys_pf, how="left", suffixes=("", "_入學"))
        else:
            p1["最終入學"] = None
    else:
        p1["最終入學"] = None

    for stage_col in ["二階甄試", "最終入學"]:
        if stage_col in p1.columns:
            p1[stage_col] = p1[stage_col].fillna("❌")

    def get_stage(row):
        if row.get("最終入學") == "✅":
            return "已入學"
        elif row.get("二階甄試") == "✅":
            return "二階未入學"
        else:
            return "僅一階"

    p1["目前狀態"] = p1.apply(get_stage, axis=1)

    if "身分證字號" in p1.columns and "學年度" in p1.columns:
        mask = p1["身分證字號"].notna() & (p1["身分證字號"] != "")
        p1_with_id = p1[mask].drop_duplicates(subset=["身分證字號", "學年度"], keep="last")
        p1_no_id = p1[~mask].drop_duplicates(subset=["姓名", "學年度"], keep="last")
        p1 = pd.concat([p1_with_id, p1_no_id], ignore_index=True)
    elif "姓名" in p1.columns and "學年度" in p1.columns:
        p1 = p1.drop_duplicates(subset=["姓名", "學年度"], keep="last")

    st.session_state.merged_data = p1.reset_index(drop=True)

def generate_sample_data():
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
        n1 = random.randint(200, 300)
        year_names = []
        for _ in range(n1):
            school_name, base_lat, base_lon = random.choice(schools)
            dept = random.choice(departments)
            name = random.choice(surnames) + random.choice(given_names) + random.choice(given_names)
            lat = base_lat + random.uniform(-0.02, 0.02)
            lon = base_lon + random.uniform(-0.02, 0.02)
            fake_id = "{}{}".format(
                "ABCDEFGHIJ"[random.randint(0, 9)],
                random.randint(100000000, 299999999)
            )
            year_names.append(name)
            all_phase1.append({
                "學年度": str(year), "報考科系": dept, "姓名": name,
                "畢業學校": school_name,
                "經緯度": "{:.4f}, {:.4f}".format(lat, lon),
                "身分證字號": fake_id
            })
        n2 = int(n1 * 0.6)
        phase2_names = random.sample(year_names, n2)
        for name in phase2_names:
            all_phase2.append({
                "學年度": str(year), "姓名": name,
                "面試成績": random.randint(50, 100),
                "筆試成績": random.randint(40, 100),
            })
        nf = int(n1 * 0.4)
        final_names = random.sample(phase2_names, min(nf, len(phase2_names)))
        for name in final_names:
            all_final.append({
                "學年度": str(year), "姓名": name,
            })
    return pd.DataFrame(all_phase1), pd.DataFrame(all_phase2), pd.DataFrame(all_final)

# ============================================================
# 側邊欄
# ============================================================
with st.sidebar:
    st.markdown("## 🎓 HWU 招生分析系統")
    st.markdown(
        "<div style='background:#FFF3E8; padding:10px; border-radius:8px; margin-bottom:10px;'>"
        "<small>📌 <b>三階段匯入說明</b><br>"
        "① 一階報名：完整欄位（學年度、科系、姓名、學校、座標、身分證）<br>"
        "② 二階甄試：姓名 + 學年度 + 成績欄位<br>"
        "③ 最終入學：僅需姓名（+學年度）<br>"
        "系統以<b>姓名+學年度</b>自動串接比對</small>"
        "</div>",
        unsafe_allow_html=True
    )
    st.markdown("---")

    st.markdown("### 📋 ① 一階報名資料")
    st.caption("需含：學年度、報考科系、姓名、畢業學校、經緯度、身分證字號")
    phase1_files = st.file_uploader(
        "上傳一階 Excel", type=["xlsx", "xls"],
        accept_multiple_files=True, key="p1_upload"
    )

    st.markdown("### 📝 ② 二階甄試資料")
    st.caption("需含：姓名（+學年度），可含成績欄位")
    phase2_files = st.file_uploader(
        "上傳二階 Excel", type=["xlsx", "xls"],
        accept_multiple_files=True, key="p2_upload"
    )

    st.markdown("### 🎓 ③ 最終入學資料")
    st.caption("最少僅需：姓名（+學年度）")
    final_files = st.file_uploader(
        "上傳入學 Excel", type=["xlsx", "xls"],
        accept_multiple_files=True, key="pf_upload"
    )

    def process_upload(files, phase_key, phase_label):
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

                # 偵錯：記錄原始欄位名稱
                original_cols = list(new_df.columns)

                new_df = clean_column_names(new_df, phase=phase_key)
                new_df = standardize_data(new_df)
                if phase_key == "phase1":
                    new_df = parse_coordinates(new_df)

                if "姓名" not in new_df.columns:
                    msg = "❌ {}（{}）：缺少「姓名」欄位。原始欄位：{}".format(
                        uf.name, phase_label,
                        ", ".join(str(c) for c in original_cols[:10])
                    )
                    st.session_state.upload_log.append(msg)
                    continue

                existing = st.session_state["{}_data".format(phase_key)]
                if existing.empty:
                    st.session_state["{}_data".format(phase_key)] = new_df
                else:
                    # 對齊欄位後合併
                    combined = pd.concat([existing, new_df], ignore_index=True, sort=False)
                    combined = combined.drop_duplicates(keep="last")
                    st.session_state["{}_data".format(phase_key)] = combined

                row_count = len(new_df)
                year_info = ""
                if "學年度" in new_df.columns:
                    yrs = safe_sort_years(new_df["學年度"])
                    year_info = "（{}）".format(", ".join(yrs))
                col_info = ", ".join(list(new_df.columns)[:8])
                st.session_state.uploaded_hashes[fhash] = uf.name
                msg = "✅ {}【{}】：{} 筆 {} [{}]".format(
                    uf.name, phase_label, row_count, year_info, col_info
                )
                st.session_state.upload_log.append(msg)
            except Exception as e:
                msg = "❌ {}（{}）：{}".format(uf.name, phase_label, str(e))
                st.session_state.upload_log.append(msg)

    process_upload(phase1_files, "phase1", "一階")
    process_upload(phase2_files, "phase2", "二階")
    process_upload(final_files, "final", "入學")

    merge_all_phases()

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

    if st.session_state.upload_log:
        st.markdown("### 📋 匯入紀錄")
        for log in st.session_state.upload_log[-12:]:
            st.caption(log)

    st.markdown("---")
    st.markdown("### 📊 資料狀態")
    p1_n = len(st.session_state.phase1_data)
    p2_n = len(st.session_state.phase2_data)
    pf_n = len(st.session_state.final_data)
    mg_n = len(st.session_state.merged_data)

    status_table = (
        "| 階段 | 筆數 |\n"
        "|------|------|\n"
        "| ① 一階報名 | **{:,}** |\n"
        "| ② 二階甄試 | **{:,}** |\n"
        "| ③ 最終入學 | **{:,}** |\n"
        "| 🔗 合併後 | **{:,}** |"
    ).format(p1_n, p2_n, pf_n, mg_n)
    st.markdown(status_table)

    st.markdown("---")
    if st.button("🗑️ 清除所有資料", use_container_width=True, type="secondary"):
        for key in DEFAULT_STATES:
            if isinstance(DEFAULT_STATES[key], pd.DataFrame):
                st.session_state[key] = pd.DataFrame()
            elif isinstance(DEFAULT_STATES[key], dict):
                st.session_state[key] = {}
            elif isinstance(DEFAULT_STATES[key], list):
                st.session_state[key] = []
        st.rerun()

# ============================================================
# 主畫面
# ============================================================
st.title("🎓 中華醫事科技大學 招生數據分析系統")

data = st.session_state.merged_data

if data.empty:
    st.markdown(
        "<div style='text-align:center; padding:50px 20px;'>"
        "<h2>👈 請先從左側匯入各階段資料</h2>"
        "<div style='display:flex; justify-content:center; gap:30px; margin-top:30px;'>"
        "<div style='background:#FFF3E8; padding:20px; border-radius:12px; width:220px;'>"
        "<h3>📋 一階報名</h3>"
        "<p style='font-size:13px; color:gray;'>學年度、科系、姓名<br>學校、座標、身分證</p>"
        "</div>"
        "<div style='background:#FFF3E8; padding:20px; border-radius:12px; width:220px;'>"
        "<h3>📝 二階甄試</h3>"
        "<p style='font-size:13px; color:gray;'>姓名、學年度<br>+ 成績欄位（選填）</p>"
        "</div>"
        "<div style='background:#FFF3E8; padding:20px; border-radius:12px; width:220px;'>"
        "<h3>🎓 最終入學</h3>"
        "<p style='font-size:13px; color:gray;'>僅需：姓名<br>（+學年度更精準）</p>"
        "</div>"
        "</div>"
        "<p style='color:gray; margin-top:20px;'>或點擊「載入三階段示範資料」快速體驗</p>"
        "</div>",
        unsafe_allow_html=True
    )
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

    n_total = len(funnel_data)
    n_phase1 = len(funnel_data[funnel_data["一階報名"] == "✅"]) if "一階報名" in funnel_data.columns else n_total
    n_phase2 = len(funnel_data[funnel_data["二階甄試"] == "✅"]) if "二階甄試" in funnel_data.columns else 0
    n_final = len(funnel_data[funnel_data["最終入學"] == "✅"]) if "最終入學" in funnel_data.columns else 0

    kpi_cols = st.columns(4)
    with kpi_cols[0]:
        st.metric("📋 一階報名", "{:,} 人".format(n_phase1))
    with kpi_cols[1]:
        rate_21 = n_phase2 / n_phase1 * 100 if n_phase1 > 0 else 0
        st.metric("📝 二階甄試", "{:,} 人".format(n_phase2), "轉換率 {:.1f}%".format(rate_21))
    with kpi_cols[2]:
        rate_f1 = n_final / n_phase1 * 100 if n_phase1 > 0 else 0
        st.metric("🎓 最終入學", "{:,} 人".format(n_final), "總錄取率 {:.1f}%".format(rate_f1))
    with kpi_cols[3]:
        rate_f2 = n_final / n_phase2 * 100 if n_phase2 > 0 else 0
        st.metric("🎯 二階→入學", "{:.1f}%".format(rate_f2))

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
        st.markdown("**📉 流失分析**")
        if n_phase1 > 0 and n_phase2 > 0:
            lost_12 = n_phase1 - n_phase2
            lost_2f = n_phase2 - n_final
            loss_table = (
                "| 階段 | 流失人數 | 流失率 |\n"
                "|------|---------|--------|\n"
                "| 一階→二階 | {:,} | {:.1f}% |\n"
                "| 二階→入學 | {:,} | {:.1f}% |"
            ).format(
                lost_12, lost_12 / n_phase1 * 100,
                lost_2f, lost_2f / n_phase2 * 100 if n_phase2 > 0 else 0
            )
            st.markdown(loss_table)
        else:
            st.info("需要多階段資料才能計算流失率")

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
                "一→二轉換率": "{:.1f}%".format(d2 / d1 * 100) if d1 > 0 else "0%",
                "總錄取率": "{:.1f}%".format(df_ / d1 * 100) if d1 > 0 else "0%",
            })
        dept_funnel_df = pd.DataFrame(dept_funnel)
        st.dataframe(dept_funnel_df, use_container_width=True, hide_index=True)

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
        fig_yf.add_trace(go.Scatter(
            x=yf_df["學年度"], y=yf_df["一階報名"],
            mode="lines+markers", name="一階報名", line=dict(color="#FDD7B4", width=3)
        ))
        fig_yf.add_trace(go.Scatter(
            x=yf_df["學年度"], y=yf_df["二階甄試"],
            mode="lines+markers", name="二階甄試", line=dict(color="#E8792F", width=3)
        ))
        fig_yf.add_trace(go.Scatter(
            x=yf_df["學年度"], y=yf_df["最終入學"],
            mode="lines+markers", name="最終入學", line=dict(color="#B84500", width=3)
        ))
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
            year_options_map = ["全部"] + safe_sort_years(data["學年度"])
            sel_year_map = st.selectbox("選擇學年度", year_options_map, key="map_year")
        else:
            sel_year_map = "全部"
    with col_f2:
        if "報考科系" in data.columns:
            dept_options_map = ["全部"] + sorted(data["報考科系"].dropna().unique().tolist())
            sel_dept_map = st.selectbox("選擇科系", dept_options_map, key="map_dept")
        else:
            sel_dept_map = "全部"
    with col_f3:
        if "目前狀態" in data.columns:
            stage_options_map = ["全部"] + sorted(data["目前狀態"].dropna().unique().tolist())
            sel_stage_map = st.selectbox("招生階段", stage_options_map, key="map_stage")
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
                if len(valid_coords) > 1000:
                    display_coords = valid_coords.sample(1000, random_state=42)
                    st.caption("⚡ 地圖顯示抽樣 1,000 筆（共 {:,} 筆）".format(len(valid_coords)))
                else:
                    display_coords = valid_coords

                for _, row in display_coords.iterrows():
                    status = row.get("目前狀態", "僅一階")
                    color = color_map.get(status, "#E8792F")
                    popup_parts = []
                    for field in ["畢業學校", "報考科系", "學年度", "目前狀態"]:
                        if field in row.index and pd.notna(row.get(field)):
                            popup_parts.append("{}：{}".format(field, row[field]))
                    popup_obj = None
                    if popup_parts:
                        popup_obj = folium.Popup("<br>".join(popup_parts), max_width=200)
                    folium.CircleMarker(
                        location=[float(row["緯度"]), float(row["經度"])],
                        radius=5, color=color, fill=True, fill_color=color, fill_opacity=0.7,
                        popup=popup_obj,
                    ).add_to(m)

                st_folium(m, width=None, height=550, use_container_width=True)

            with col_m2:
                st.metric("篩選結果", "{:,} 人".format(len(map_data)))
                st.metric("有座標", "{:,} 人".format(len(valid_coords)))
                coverage = len(valid_coords) / len(map_data) * 100 if len(map_data) > 0 else 0
                st.metric("座標涵蓋率", "{:.1f}%".format(coverage))

                if "目前狀態" in valid_coords.columns:
                    st.markdown("**📊 階段分布**")
                    for status_val, cnt in valid_coords["目前狀態"].value_counts().items():
                        if status_val == "已入學":
                            icon = "🟤"
                        elif status_val == "二階未入學":
                            icon = "🟠"
                        else:
                            icon = "🟡"
                        st.caption("{} {}：{}".format(icon, status_val, cnt))

                if "畢業學校" in valid_coords.columns:
                    st.markdown("**📍 主要來源學校**")
                    for school_val, cnt in valid_coords["畢業學校"].value_counts().head(8).items():
                        st.caption("• {}：{}".format(school_val, cnt))
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
            fig_bar.update_layout(
                yaxis={"categoryorder": "total ascending"},
                height=max(400, len(dept_counts) * 35),
                showlegend=False
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        with col_c2:
            if "最終入學" in t2_data.columns:
                dept_enroll = []
                for dept in dept_counts["科系"]:
                    dd = t2_data[t2_data["報考科系"] == dept]
                    total_d = len(dd)
                    enrolled_d = len(dd[dd["最終入學"] == "✅"])
                    dept_enroll.append({
                        "科系": dept, "報名": total_d, "入學": enrolled_d,
                        "入學率": enrolled_d / total_d * 100 if total_d > 0 else 0
                    })
                de_df = pd.DataFrame(dept_enroll)
                fig_enroll = go.Figure()
                fig_enroll.add_trace(go.Bar(
                    name="報名", x=de_df["科系"], y=de_df["報名"], marker_color="#FDD7B4"
                ))
                fig_enroll.add_trace(go.Bar(
                    name="入學", x=de_df["科系"], y=de_df["入學"], marker_color="#B84500"
                ))
                fig_enroll.update_layout(
                    barmode="overlay", title="各科系報名 vs 入學",
                    height=max(400, len(dept_counts) * 35)
                )
                st.plotly_chart(fig_enroll, use_container_width=True)
            else:
                fig_pie = px.pie(
                    dept_counts, values="報考人數", names="科系", title="科系佔比分布",
                    color_discrete_sequence=px.colors.sequential.Oranges_r
                )
                fig_pie.update_layout(height=max(400, len(dept_counts) * 35))
                st.plotly_chart(fig_pie, use_container_width=True)

        if "學年度" in data.columns and data["學年度"].nunique() > 1 and sel_year_t2 == "全部":
            st.subheader("📈 各科系歷年趨勢")
            yearly_dept = data.groupby(["學年度", "報考科系"]).size().reset_index(name="人數")
            yearly_dept["排序鍵"] = yearly_dept["學年度"].apply(safe_int)
            yearly_dept = yearly_dept.sort_values("排序鍵")
            fig_line = px.line(
                yearly_dept, x="學年度", y="人數", color="報考科系",
                markers=True, title="各科系歷年報考人數趨勢"
            )
            fig_line.update_layout(height=450)
            st.plotly_chart(fig_line, use_container_width=True)

        st.subheader("📋 科系統計摘要")
        kpi_dept_cols = st.columns(4)
        with kpi_dept_cols[0]:
            st.metric("總報考人數", "{:,}".format(len(t2_data)))
        with kpi_dept_cols[1]:
            st.metric("科系數", "{}".format(t2_data["報考科系"].nunique()))
        with kpi_dept_cols[2]:
            top_dept = dept_counts.iloc[0]["科系"] if len(dept_counts) > 0 else "-"
            st.metric("最熱門科系", top_dept)
        with kpi_dept_cols[3]:
            if "最終入學" in t2_data.columns:
                total_enrolled = len(t2_data[t2_data["最終入學"] == "✅"])
                overall_rate = total_enrolled / len(t2_data) * 100 if len(t2_data) > 0 else 0
                st.metric("總入學率", "{:.1f}%".format(overall_rate))
            else:
                avg_count = dept_counts["報考人數"].mean() if len(dept_counts) > 0 else 0
                st.metric("平均每科系", "{:.0f} 人".format(avg_count))

# ============================================================
# TAB 3: 來源學校分析
# ============================================================
with tab3:
    st.header("🏫 來源學校分析")

    if "畢業學校" not in data.columns:
        st.warning("資料中未包含「畢業學校」欄位")
    else:
        col_t3_f1, col_t3_f2, col_t3_f3 = st.columns(3)
        with col_t3_f1:
            if "學年度" in data.columns:
                yr_opts_t3 = ["全部"] + safe_sort_years(data["學年度"])
                sel_year_t3 = st.selectbox("選擇學年度", yr_opts_t3, key="t3_year")
            else:
                sel_year_t3 = "全部"
        with col_t3_f2:
            top_n = st.slider("顯示前 N 所學校", 5, 30, 15, key="t3_topn")
        with col_t3_f3:
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

        bar_title = "Top {} 來源學校".format(top_n)
        if sel_stage_t3 != "全部":
            bar_title += "（{}）".format(sel_stage_t3)
        fig_school = px.bar(
            school_counts, x="人數", y="畢業學校", orientation="h",
            title=bar_title, color="人數",
            color_continuous_scale=["#FDD7B4", "#E8792F", "#B84500"],
        )
        fig_school.update_layout(
            yaxis={"categoryorder": "total ascending"},
            height=max(400, top_n * 30),
            showlegend=False
        )
        st.plotly_chart(fig_school, use_container_width=True)

        if "最終入學" in t3_data.columns:
            st.subheader("🎯 各學校入學轉換率")
            school_conv = []
            for school_name in school_counts["畢業學校"]:
                sd = t3_data[t3_data["畢業學校"] == school_name]
                total_s = len(sd)
                enrolled_s = len(sd[sd["最終入學"] == "✅"])
                rate_s = "{:.1f}%".format(enrolled_s / total_s * 100) if total_s > 0 else "0%"
                school_conv.append({
                    "學校": school_name, "報名": total_s, "入學": enrolled_s,
                    "轉換率": rate_s
                })
            st.dataframe(pd.DataFrame(school_conv), use_container_width=True, hide_index=True)

        if "報考科系" in t3_data.columns:
            st.subheader("🔀 來源學校 × 科系 交叉分析")
            top_schools_list = school_counts["畢業學校"].tolist()
            cross_data = t3_data[t3_data["畢業學校"].isin(top_schools_list)]
            if len(cross_data) > 0:
                cross_table = pd.crosstab(cross_data["畢業學校"], cross_data["報考科系"])
                fig_heat = px.imshow(
                    cross_table, title="來源學校 vs 報考科系（人數）",
                    color_continuous_scale="Oranges", aspect="auto"
                )
                fig_heat.update_layout(height=max(400, top_n * 28))
                st.plotly_chart(fig_heat, use_container_width=True)

        st.subheader("⭐ 來源學校管理建議")
        all_school_counts = t3_data["畢業學校"].value_counts()
        total_t3 = len(t3_data)
        recs = []
        cumulative = 0
        for school_name_rec, cnt in all_school_counts.items():
            cumulative += cnt
            ratio = cnt / total_t3 * 100 if total_t3 > 0 else 0
            cum_ratio = cumulative / total_t3 * 100 if total_t3 > 0 else 0
            if cum_ratio <= 50:
                level = "⭐⭐⭐ 重點經營"
            elif cum_ratio <= 80:
                level = "⭐⭐ 持續關注"
            else:
                level = "⭐ 一般維護"
            recs.append({
                "學校": school_name_rec,
                "人數": cnt,
                "佔比(%)": round(ratio, 1),
                "累積佔比(%)": round(cum_ratio, 1),
                "建議等級": level
            })
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

        st.subheader("📊 歷年招生趨勢")
        yearly_stats = []
        for yr in years_sorted:
            yd = data[data["學年度"] == yr]
            row_ys = {"學年度": yr, "一階報名": len(yd)}
            if "二階甄試" in yd.columns:
                row_ys["二階甄試"] = len(yd[yd["二階甄試"] == "✅"])
            if "最終入學" in yd.columns:
                row_ys["最終入學"] = len(yd[yd["最終入學"] == "✅"])
            yearly_stats.append(row_ys)
        ys_df = pd.DataFrame(yearly_stats)

        fig_trend = go.Figure()
        for col_name, color in [("一階報名", "#FDD7B4"), ("二階甄試", "#E8792F"), ("最終入學", "#B84500")]:
            if col_name in ys_df.columns:
                fig_trend.add_trace(go.Bar(
                    name=col_name, x=ys_df["學年度"], y=ys_df[col_name],
                    marker_color=color, text=ys_df[col_name], textposition="outside"
                ))
        fig_trend.update_layout(barmode="group", title="歷年各階段人數", height=400)
        st.plotly_chart(fig_trend, use_container_width=True)

        col_inc1, col_inc2 = st.columns([1, 2])
        with col_inc1:
            st.markdown("**📊 年度增減**")
            for i in range(1, len(ys_df)):
                prev_n = ys_df.iloc[i - 1]["一階報名"]
                curr_n = ys_df.iloc[i]["一階報名"]
                diff = curr_n - prev_n
                pct = diff / prev_n * 100 if prev_n > 0 else 0
                st.metric(
                    "{} 學年".format(ys_df.iloc[i]["學年度"]),
                    "{:,} 人".format(curr_n),
                    "{:+d}（{:+.1f}%）".format(diff, pct)
                )

        with col_inc2:
            if "報考科系" in data.columns and len(years_sorted) >= 2:
                st.markdown("**📊 科系增減比較**")
                year_a = years_sorted[-2]
                year_b = years_sorted[-1]
                da = data[data["學年度"] == year_a]["報考科系"].value_counts()
                db = data[data["學年度"] == year_b]["報考科系"].value_counts()
                all_depts = sorted(set(da.index) | set(db.index))
                comparison = []
                for dept in all_depts:
                    a_val = int(da.get(dept, 0))
                    b_val = int(db.get(dept, 0))
                    diff_val = b_val - a_val
                    if diff_val > 0:
                        trend_icon = "📈"
                    elif diff_val < 0:
                        trend_icon = "📉"
                    else:
                        trend_icon = "➡️"
                    comparison.append({
                        "科系": dept,
                        year_a: a_val,
                        year_b: b_val,
                        "增減": diff_val,
                        "趨勢": trend_icon
                    })
                comp_df = pd.DataFrame(comparison).sort_values("增減", ascending=False)
                st.dataframe(comp_df, use_container_width=True, hide_index=True)

        if "畢業學校" in data.columns:
            st.subheader("🏫 來源學校年度變化")
            top_schools_overall = data["畢業學校"].value_counts().head(10).index.tolist()
            school_year = data.groupby(["學年度", "畢業學校"]).size().reset_index(name="人數")
            school_year_top = school_year[school_year["畢業學校"].isin(top_schools_overall)].copy()
            school_year_top["排序鍵"] = school_year_top["學年度"].apply(safe_int)
            school_year_top = school_year_top.sort_values("排序鍵")
            fig_st = px.line(
                school_year_top, x="學年度", y="人數", color="畢業學校",
                markers=True, title="Top 10 來源學校歷年趨勢"
            )
            fig_st.update_layout(height=450)
            st.plotly_chart(fig_st, use_container_width=True)

# ============================================================
# TAB 5: 資料檢視與匯出
# ============================================================
with tab5:
    st.header("🔍 資料檢視與匯出")

    with st.expander("📋 各階段原始資料", expanded=False):
        sub_tabs = st.tabs(["一階報名", "二階甄試", "最終入學"])
        with sub_tabs[0]:
            if not st.session_state.phase1_data.empty:
                st.caption("共 {:,} 筆".format(len(st.session_state.phase1_data)))
                st.dataframe(st.session_state.phase1_data.head(100), use_container_width=True, hide_index=True)
            else:
                st.info("尚無一階資料")
        with sub_tabs[1]:
            if not st.session_state.phase2_data.empty:
                st.caption("共 {:,} 筆".format(len(st.session_state.phase2_data)))
                st.dataframe(st.session_state.phase2_data.head(100), use_container_width=True, hide_index=True)
            else:
                st.info("尚無二階資料")
        with sub_tabs[2]:
            if not st.session_state.final_data.empty:
                st.caption("共 {:,} 筆".format(len(st.session_state.final_data)))
                st.dataframe(st.session_state.final_data.head(100), use_container_width=True, hide_index=True)
            else:
                st.info("尚無入學資料")

    st.subheader("🔗 合併後資料")

    col_e1, col_e2, col_e3, col_e4 = st.columns(4)
    export_data = data.copy()

    with col_e1:
        if "學年度" in data.columns:
            yr_options_exp = ["全部"] + safe_sort_years(data["學年度"])
            sel_yr_exp = st.selectbox("學年度", yr_options_exp, key="exp_yr")
            if sel_yr_exp != "全部":
                export_data = export_data[export_data["學年度"] == sel_yr_exp]
    with col_e2:
        if "報考科系" in data.columns:
            dept_opts_exp = ["全部"] + sorted(data["報考科系"].dropna().unique().tolist())
            sel_dept_exp = st.selectbox("科系", dept_opts_exp, key="exp_dept")
            if sel_dept_exp != "全部":
                export_data = export_data[export_data["報考科系"] == sel_dept_exp]
    with col_e3:
        if "目前狀態" in data.columns:
            stage_opts_exp = ["全部"] + sorted(data["目前狀態"].dropna().unique().tolist())
            sel_stage_exp = st.selectbox("階段", stage_opts_exp, key="exp_stage")
            if sel_stage_exp != "全部":
                export_data = export_data[export_data["目前狀態"] == sel_stage_exp]
    with col_e4:
        if "畢業學校" in data.columns:
            search_school = st.text_input("搜尋學校", key="exp_school")
            if search_school:
                export_data = export_data[
                    export_data["畢業學校"].astype(str).str.contains(search_school, na=False)
                ]

    st.caption("篩選結果：{:,} 筆".format(len(export_data)))

    hide_cols = ["身分證字號"]
    display_cols = [c for c in export_data.columns if c not in hide_cols]
    st.dataframe(export_data[display_cols], use_container_width=True, hide_index=True, height=400)

    st.subheader("📥 匯出資料")
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        csv_buf = io.BytesIO()
        export_data.to_csv(csv_buf, index=False, encoding="utf-8-sig")
        st.download_button(
            "⬇️ 下載 CSV", csv_buf.getvalue(),
            "招生分析資料.csv", "text/csv", use_container_width=True
        )
    with col_dl2:
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            export_data.to_excel(writer, index=False, sheet_name="招生資料")
        st.download_button(
            "⬇️ 下載 Excel", excel_buf.getvalue(),
            "招生分析資料.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.subheader("📋 資料品質報告")
    quality = []
    for col_q in data.columns:
        non_null = data[col_q].notna().sum()
        total_rows = len(data)
        # 處理可能的 DataFrame 類型（重複欄名）
        if isinstance(non_null, pd.Series):
            non_null = int(non_null.iloc[0])
        rate_q = "{:.1f}%".format(non_null / total_rows * 100) if total_rows > 0 else "0%"
        quality.append({
            "欄位": col_q,
            "非空筆數": non_null,
            "總筆數": total_rows,
            "完整率": rate_q,
            "空值數": total_rows - non_null,
        })
    st.dataframe(pd.DataFrame(quality), use_container_width=True, hide_index=True)
