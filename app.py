# -*- coding: utf-8 -*-
"""
中華醫事科技大學 招生數據分析系統 v6.2
- 新增：班級名稱→科系自動映射（Phase 3）
- 修正：最終入學科系欄位偵測
- 跨階段名稱正規化
- 多年度標籤式介面
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import re

# ============================================================
# 頁面設定
# ============================================================
st.set_page_config(
    page_title="HWU 招生數據分析系統",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CSS
# ============================================================
st.markdown("""
<style>
    .main-header {
        font-size:2.2rem;font-weight:800;
        background:linear-gradient(135deg,#1e3a5f 0%,#2d6a4f 100%);
        -webkit-background-clip:text;-webkit-text-fill-color:transparent;
        text-align:center;padding:10px 0;margin-bottom:5px;
    }
    .sub-header{font-size:1.0rem;color:#6c757d;text-align:center;margin-bottom:20px;}
    .metric-card{
        background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);
        padding:20px;border-radius:15px;text-align:center;color:white;margin:5px 0;
        box-shadow:0 4px 15px rgba(0,0,0,.1);
    }
    .metric-card h3{margin:0;font-size:.85rem;opacity:.9;}
    .metric-card h1{margin:5px 0 0 0;font-size:1.8rem;}
    .metric-green{background:linear-gradient(135deg,#43e97b 0%,#38f9d7 100%);color:#1a1a2e;}
    .metric-orange{background:linear-gradient(135deg,#f093fb 0%,#f5576c 100%);}
    .metric-blue{background:linear-gradient(135deg,#4facfe 0%,#00f2fe 100%);color:#1a1a2e;}
    .metric-gold{background:linear-gradient(135deg,#f6d365 0%,#fda085 100%);color:#1a1a2e;}
    .section-divider{
        height:3px;background:linear-gradient(90deg,#667eea,#764ba2,#f093fb);
        border:none;border-radius:2px;margin:25px 0;
    }
    .info-box{background:#f0f4ff;border-left:4px solid #667eea;padding:15px;border-radius:0 10px 10px 0;margin:10px 0;font-size:.9rem;}
    .warning-box{background:#fff8e1;border-left:4px solid #ffa726;padding:15px;border-radius:0 10px 10px 0;margin:10px 0;font-size:.9rem;}
    .success-box{background:#e8f5e9;border-left:4px solid #4caf50;padding:15px;border-radius:0 10px 10px 0;margin:10px 0;font-size:.9rem;}
    .mapping-box{background:#f3e5f5;border-left:4px solid #9c27b0;padding:15px;border-radius:0 10px 10px 0;margin:10px 0;font-size:.9rem;}
    .year-tag{
        display:inline-block;background:#e3f2fd;color:#1565c0;
        padding:3px 12px;border-radius:12px;font-size:.85rem;margin:2px 3px;font-weight:600;
    }
    .channel-tag{
        display:inline-block;background:#e8f5e9;color:#2e7d32;
        padding:2px 10px;border-radius:12px;font-size:.8rem;margin:2px 3px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🎓 中華醫事科技大學 招生數據分析系統</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Enrollment Analytics v6.2 ｜ 班級→科系自動映射 ｜ 跨階段名稱正規化</div>', unsafe_allow_html=True)
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

# ============================================================
# 常數
# ============================================================
FINAL_CH_CANDIDATES = [
    "入學方式", "入學管道", "錄取管道", "招生管道", "管道",
    "入學途徑", "錄取方式", "報名管道"
]
HWU = {"lat": 22.9340, "lon": 120.2756}

# ── 中華醫事科技大學 科系關鍵字庫 ──
DEPT_KEYWORDS = [
    ("長期照護", "長期照護學位學程"),
    ("職業安全衛生", "職業安全衛生系"),
    ("食品營養", "食品營養系"),
    ("環境與安全衛生工程", "環境與安全衛生工程系"),
    ("醫務暨健康事務管理", "醫務暨健康事務管理系"),
    ("醫務管理", "醫務暨健康事務管理系"),
    ("健康事務管理", "醫務暨健康事務管理系"),
    ("醫學檢驗生物技術", "醫學檢驗生物技術系"),
    ("醫學檢驗", "醫學檢驗生物技術系"),
    ("醫檢", "醫學檢驗生物技術系"),
    ("護理", "護理系"),
    ("藥學", "藥學系"),
    ("視光", "視光系"),
    ("製藥", "製藥工程系"),
    ("生物科技", "生物科技系"),
    ("生科", "生物科技系"),
    ("寵物美容", "寵物美容學位學程"),
    ("寵物", "寵物美容學位學程"),
    ("幼兒保育", "幼兒保育系"),
    ("幼保", "幼兒保育系"),
    ("運動健康與休閒", "運動健康與休閒系"),
    ("運動休閒", "運動健康與休閒系"),
    ("運休", "運動健康與休閒系"),
    ("資訊管理", "資訊管理系"),
    ("資管", "資訊管理系"),
    ("多媒體設計", "多媒體設計系"),
    ("多媒體", "多媒體設計系"),
    ("餐旅管理", "餐旅管理系"),
    ("餐旅", "餐旅管理系"),
    ("觀光休閒", "觀光休閒事業管理系"),
    ("觀光", "觀光休閒事業管理系"),
    ("化妝品應用", "化妝品應用與管理系"),
    ("化妝品", "化妝品應用與管理系"),
    ("妝管", "化妝品應用與管理系"),
    ("美妝", "化妝品應用與管理系"),
    ("調理保健", "調理保健技術系"),
    ("調理", "調理保健技術系"),
    ("語言治療", "語言治療系"),
    ("語治", "語言治療系"),
    ("牙體技術", "牙體技術系"),
    ("牙技", "牙體技術系"),
    ("職安", "職業安全衛生系"),
    ("環安", "環境與安全衛生工程系"),
    ("食營", "食品營養系"),
]


# ============================================================
# 工具函式
# ============================================================
def detect_school_col(df):
    for kw in ["畢業學校", "來源學校", "原就讀學校", "高中職校名",
                "學校名稱", "校名", "畢業高中職", "畢業學校名稱",
                "就讀學校", "原學校"]:
        for c in df.columns:
            if kw in str(c):
                return c
    return None


def detect_dept_col(df):
    for kw in ["科系", "系所", "報名科系", "錄取科系", "志願科系",
                "就讀科系", "科系名稱", "系科", "學系", "錄取系所",
                "就讀系所", "註冊科系", "系別", "科別", "報名系所"]:
        for c in df.columns:
            if kw in str(c):
                return c
    return None


def detect_class_col(df):
    """偵測班級欄位"""
    for kw in ["班級", "班級名稱", "就讀班級", "編班", "班別",
                "入學班級", "註冊班級", "分班", "class"]:
        for c in df.columns:
            if kw in str(c).lower():
                return c
    return None


def detect_final_ch_col(df):
    for kw in FINAL_CH_CANDIDATES:
        for c in df.columns:
            if kw in str(c):
                return c
    return None


def norm_school(name):
    if not isinstance(name, str):
        return str(name).strip()
    name = name.strip()
    name = re.sub(r"\s+", "", name)
    name = name.replace("臺", "台").replace("（", "(").replace("）", ")")
    for sfx in ["附設進修學校", "進修學校", "進修部"]:
        name = name.replace(sfx, "")
    return name


def norm_dept(name):
    if not isinstance(name, str):
        return str(name).strip()
    name = name.strip()
    name = re.sub(r"\s+", "", name)
    name = name.replace("臺", "台").replace("（", "(").replace("）", ")")
    name = name.replace("　", "")
    return name


def class_to_dept(class_name):
    """從班級名稱萃取科系（使用關鍵字庫）"""
    if not isinstance(class_name, str):
        return None
    clean = class_name.strip()
    clean = re.sub(r"\s+", "", clean)
    for kw, dept in DEPT_KEYWORDS:
        if kw in clean:
            return dept
    return None


def auto_map_class_to_dept(df, class_col, p1_depts=None):
    """
    為 DataFrame 新增「_mapped_dept」欄位
    1. 先用關鍵字庫匹配
    2. 若有 p1 科系清單，嘗試模糊比對
    """
    df = df.copy()
    classes = df[class_col].fillna("").astype(str).str.strip().unique()

    mapping = {}
    for cls in classes:
        dept = class_to_dept(cls)
        if dept:
            mapping[cls] = dept
            continue
        if p1_depts:
            for d in p1_depts:
                d_clean = norm_dept(d)
                d_core = d_clean.rstrip("系科學程學位")
                if len(d_core) >= 2 and d_core in cls:
                    mapping[cls] = d
                    break

    df["_mapped_dept"] = df[class_col].map(mapping)
    return df, mapping


def detect_lat_lon_cols(df):
    lat_col = lon_col = None
    for c in df.columns:
        s = str(c).strip().lower()
        if lat_col is None and any(k in s for k in ["緯度", "lat", "latitude"]):
            lat_col = c
        if lon_col is None and any(k in s for k in ["經度", "lon", "lng", "longitude"]):
            lon_col = c
    return lat_col, lon_col


def build_geo_from_p1(p1):
    sc = detect_school_col(p1)
    lat_c, lon_c = detect_lat_lon_cols(p1)
    if sc is None or lat_c is None or lon_c is None:
        return None
    g = p1[[sc, lat_c, lon_c]].copy()
    g.columns = ["學校_raw", "lat", "lon"]
    g["lat"] = pd.to_numeric(g["lat"], errors="coerce")
    g["lon"] = pd.to_numeric(g["lon"], errors="coerce")
    g = g.dropna(subset=["lat", "lon"])
    g["_std"] = g["學校_raw"].apply(norm_school)
    geo = g.groupby("_std").agg(lat=("lat", "mean"), lon=("lon", "mean")).reset_index()
    return geo


def enrich_geo(df, geo):
    sc = detect_school_col(df)
    if sc is None or geo is None or geo.empty:
        return df
    df = df.copy()
    df["_std"] = df[sc].apply(norm_school)
    for c in ["緯度", "經度", "lat", "lon", "latitude", "longitude"]:
        if c in df.columns and c != sc:
            df.drop(columns=[c], inplace=True, errors="ignore")
    df = df.merge(geo, on="_std", how="left").drop(columns=["_std"], errors="ignore")
    return df


def eff_stars(r):
    if r >= 70:
        return "⭐⭐⭐"
    elif r >= 40:
        return "⭐⭐"
    else:
        return "⭐"


# ============================================================
# 統計構建 - 支援班級→科系映射
# ============================================================
def get_dept_series(df, p1_depts=None):
    """
    從 df 取得科系 Series（正規化後）
    優先使用科系欄位，若無則嘗試班級→科系映射
    """
    dc = detect_dept_col(df)
    if dc:
        s = df[dc].dropna().apply(norm_dept)
        return s, {"method": "direct", "col": dc, "mapped": 0, "unmapped": 0}

    cc = detect_class_col(df)
    if cc:
        df_mapped, mapping = auto_map_class_to_dept(df, cc, p1_depts)
        s = df_mapped["_mapped_dept"].dropna().apply(norm_dept)
        n_mapped = df_mapped["_mapped_dept"].notna().sum()
        n_unmapped = df_mapped["_mapped_dept"].isna().sum()
        return s, {
            "method": "class_mapping",
            "col": cc,
            "mapped": n_mapped,
            "unmapped": n_unmapped,
            "mapping": mapping
        }

    return pd.Series(dtype=str), {"method": "none", "col": None}


def build_dept_stats(p1, p2=None, p3=None):
    """建立科系統計 - 支援班級映射"""
    dc1 = detect_dept_col(p1)
    if dc1 is None:
        return None

    p1_depts = p1[dc1].dropna().unique().tolist()

    tmp1 = p1[dc1].dropna().apply(norm_dept)
    s = tmp1.value_counts().reset_index()
    s.columns = ["_dept_std", "一階人數"]

    name_map = {}
    for raw in p1[dc1].dropna().unique():
        name_map[norm_dept(raw)] = raw

    if p2 is not None:
        s2, info2 = get_dept_series(p2, p1_depts)
        if not s2.empty:
            t2 = s2.value_counts().reset_index()
            t2.columns = ["_dept_std", "二階人數"]
            s = s.merge(t2, on="_dept_std", how="left")
    if "二階人數" not in s.columns:
        s["二階人數"] = np.nan

    p3_info = None
    if p3 is not None:
        s3, p3_info = get_dept_series(p3, p1_depts)
        if not s3.empty:
            t3 = s3.value_counts().reset_index()
            t3.columns = ["_dept_std", "最終入學"]
            s = s.merge(t3, on="_dept_std", how="left")
    if "最終入學" not in s.columns:
        s["最終入學"] = np.nan

    s["二階人數"] = s["二階人數"].fillna(0).astype(int)
    s["最終入學"] = s["最終入學"].fillna(0).astype(int)
    s["一→二階(%)"] = (s["二階人數"] / s["一階人數"] * 100).round(1)
    s["一→最終(%)"] = (s["最終入學"] / s["一階人數"] * 100).round(1)
    s["流失人數"] = s["一階人數"] - s["最終入學"]
    s["效率評等"] = s["一→最終(%)"].apply(eff_stars)

    s["科系"] = s["_dept_std"].map(name_map).fillna(s["_dept_std"])
    s = s.drop(columns=["_dept_std"])

    col_order = ["科系", "一階人數", "二階人數", "最終入學",
                 "一→二階(%)", "一→最終(%)", "流失人數", "效率評等"]
    s = s[[c for c in col_order if c in s.columns]]
    return s, p3_info


def build_school_stats(p1, p2=None, p3=None):
    sc1 = detect_school_col(p1)
    if sc1 is None:
        return None

    tmp1 = p1[[sc1]].copy()
    tmp1["_sch_std"] = tmp1[sc1].apply(norm_school)
    s = tmp1["_sch_std"].value_counts().reset_index()
    s.columns = ["_sch_std", "一階人數"]

    if p2 is not None:
        sc2 = detect_school_col(p2)
        if sc2:
            tmp2 = p2[[sc2]].copy()
            tmp2["_sch_std"] = tmp2[sc2].apply(norm_school)
            t2 = tmp2["_sch_std"].value_counts().reset_index()
            t2.columns = ["_sch_std", "二階人數"]
            s = s.merge(t2, on="_sch_std", how="left")
    if "二階人數" not in s.columns:
        s["二階人數"] = np.nan

    if p3 is not None:
        sc3 = detect_school_col(p3)
        if sc3:
            tmp3 = p3[[sc3]].copy()
            tmp3["_sch_std"] = tmp3[sc3].apply(norm_school)
            t3 = tmp3["_sch_std"].value_counts().reset_index()
            t3.columns = ["_sch_std", "最終入學"]
            s = s.merge(t3, on="_sch_std", how="left")
    if "最終入學" not in s.columns:
        s["最終入學"] = np.nan

    s["二階人數"] = s["二階人數"].fillna(0).astype(int)
    s["最終入學"] = s["最終入學"].fillna(0).astype(int)
    s["一→二階(%)"] = (s["二階人數"] / s["一階人數"] * 100).round(1)
    s["一→最終(%)"] = (s["最終入學"] / s["一階人數"] * 100).round(1)
    s["流失人數"] = s["一階人數"] - s["最終入學"]
    s["效率評等"] = s["一→最終(%)"].apply(eff_stars)

    name_map = tmp1.drop_duplicates("_sch_std").set_index("_sch_std")[sc1]
    s["學校"] = s["_sch_std"].map(name_map).fillna(s["_sch_std"])
    s = s.drop(columns=["_sch_std"])

    col_order = ["學校", "一階人數", "二階人數", "最終入學",
                 "一→二階(%)", "一→最終(%)", "流失人數", "效率評等"]
    s = s[[c for c in col_order if c in s.columns]]
    return s


# ============================================================
# 視覺化
# ============================================================
def fig_funnel(labels, values, title="招生漏斗"):
    colors = ["#2196F3", "#FF9800", "#4CAF50", "#E91E63"]
    fig = go.Figure(go.Funnel(
        y=labels, x=values, textinfo="value+percent initial",
        marker=dict(color=colors[:len(labels)]),
        connector=dict(line=dict(color="royalblue", width=2))
    ))
    fig.update_layout(title=title, height=420, font=dict(size=14))
    return fig


def fig_bar_h(df, y, x, title, color="#667eea"):
    fig = px.bar(df.sort_values(x, ascending=True),
                 x=x, y=y, orientation="h", text=x, title=title)
    fig.update_traces(marker_color=color, texttemplate="%{text:.1f}%",
                      textposition="outside")
    fig.update_layout(height=max(380, len(df) * 28),
                      xaxis_title="轉換率 (%)", yaxis_title="")
    return fig


def fig_grouped_bar(df, y, vals, title):
    fig = go.Figure()
    colors = ["#2196F3", "#FF9800", "#4CAF50"]
    for i, v in enumerate(vals):
        if v in df.columns:
            fig.add_trace(go.Bar(
                name=v, y=df[y], x=df[v], orientation="h",
                marker_color=colors[i % 3], text=df[v], textposition="outside"
            ))
    fig.update_layout(barmode="group", title=title,
                      height=max(400, len(df) * 35),
                      yaxis=dict(autorange="reversed"))
    return fig


def fig_map(df, size_col, title, color_col=None):
    if "lat" not in df.columns or "lon" not in df.columns:
        return None
    m = df.dropna(subset=["lat", "lon"]).copy()
    if m.empty:
        return None
    m[size_col] = pd.to_numeric(m[size_col], errors="coerce").fillna(1)
    sc = detect_school_col(m)
    fig = px.scatter_mapbox(
        m, lat="lat", lon="lon", size=size_col,
        color=color_col if color_col and color_col in m.columns else None,
        hover_name=sc if sc and sc in m.columns else None,
        hover_data={size_col: True, "lat": ":.4f", "lon": ":.4f"},
        title=title, size_max=30, zoom=7,
        center={"lat": HWU["lat"], "lon": HWU["lon"]},
        mapbox_style="carto-positron"
    )
    fig.add_trace(go.Scattermapbox(
        lat=[HWU["lat"]], lon=[HWU["lon"]],
        mode="markers+text",
        marker=dict(size=18, color="red", symbol="star"),
        text=["中華醫事科技大學"], textposition="top center",
        name="本校", showlegend=True
    ))
    fig.update_layout(height=600, margin=dict(l=0, r=0, t=40, b=0))
    return fig


def fig_heatmap(df, x, y, v, title):
    pv = df.pivot_table(index=y, columns=x, values=v, aggfunc="sum").fillna(0)
    fig = px.imshow(pv, text_auto=True, aspect="auto",
                    color_continuous_scale="YlOrRd", title=title)
    fig.update_layout(height=max(400, len(pv) * 25))
    return fig


# ============================================================
# Session State
# ============================================================
if "years" not in st.session_state:
    st.session_state["years"] = {}
if "all_files" not in st.session_state:
    st.session_state["all_files"] = {}
if "analysis_ready" not in st.session_state:
    st.session_state["analysis_ready"] = False
if "analysis_version" not in st.session_state:
    st.session_state["analysis_version"] = 0
if "custom_mappings" not in st.session_state:
    st.session_state["custom_mappings"] = {}

# ============================================================
# Sidebar
# ============================================================
with st.sidebar:
    st.header("📂 資料管理")

    uploaded = st.file_uploader(
        "上傳所有招生資料 (Excel/CSV)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="可一次上傳多個年度的檔案"
    )
    if uploaded:
        for uf in uploaded:
            if uf.name not in st.session_state["all_files"]:
                try:
                    if uf.name.endswith(".csv"):
                        df = pd.read_csv(uf)
                    else:
                        df = pd.read_excel(uf)
                    st.session_state["all_files"][uf.name] = df
                    st.success(f"✅ {uf.name}（{len(df)} 筆）")
                except Exception as e:
                    st.error(f"❌ {uf.name}: {e}")

    if st.session_state["all_files"]:
        st.caption(f"📁 已載入 {len(st.session_state['all_files'])} 個檔案")
        if st.button("🗑️ 清除全部檔案"):
            st.session_state["all_files"] = {}
            st.session_state["years"] = {}
            st.session_state["analysis_ready"] = False
            st.rerun()

    st.markdown("---")
    st.header("📅 年度管理")
    st.markdown(
        '<div style="font-size:.8rem;color:#888;margin-bottom:8px;">'
        '每個年度指定一階/二階/最終入學<br>'
        '📌 最終入學若只有班級欄位，系統會自動映射科系</div>',
        unsafe_allow_html=True
    )

    new_year = st.text_input("新增年度標籤：", placeholder="例如：113學年",
                             key="new_year_input")
    if st.button("➕ 新增年度") and new_year:
        new_year = new_year.strip()
        if new_year and new_year not in st.session_state["years"]:
            st.session_state["years"][new_year] = {
                "p1": None, "p2": None, "p3": None,
                "channel_col": None, "selected_channels": None,
                "class_col_override": None,
            }
            st.success(f"✅ 已新增「{new_year}」")
            st.rerun()
        elif new_year in st.session_state["years"]:
            st.warning("⚠️ 此年度已存在")

    file_opts = ["-- 未選擇 --"] + list(st.session_state["all_files"].keys())

    for yr in list(st.session_state["years"].keys()):
        ydata = st.session_state["years"][yr]
        with st.expander(f"📅 {yr}", expanded=True):
            ydata["p1"] = st.selectbox(
                "🔵 一階（含經緯度）", file_opts, key=f"p1_{yr}",
                index=file_opts.index(ydata["p1"]) if ydata["p1"] in file_opts else 0
            )
            ydata["p2"] = st.selectbox(
                "🟠 二階", file_opts, key=f"p2_{yr}",
                index=file_opts.index(ydata["p2"]) if ydata["p2"] in file_opts else 0
            )
            ydata["p3"] = st.selectbox(
                "🟢 最終入學", file_opts, key=f"p3_{yr}",
                index=file_opts.index(ydata["p3"]) if ydata["p3"] in file_opts else 0
            )

            if ydata["p1"] and ydata["p1"] != "-- 未選擇 --":
                p1df = st.session_state["all_files"][ydata["p1"]]
                lat_c, lon_c = detect_lat_lon_cols(p1df)
                if lat_c and lon_c:
                    n_valid = p1df[[lat_c, lon_c]].dropna().shape[0]
                    st.caption(f"📍 經緯度：{lat_c}/{lon_c}（{n_valid}筆有效）")
                else:
                    st.caption("⚠️ 一階未偵測到經緯度欄位")

            if ydata["p3"] and ydata["p3"] != "-- 未選擇 --":
                p3df = st.session_state["all_files"][ydata["p3"]]

                dc3 = detect_dept_col(p3df)
                cc3 = detect_class_col(p3df)
                if dc3:
                    st.caption(f"✅ 最終入學有科系欄位：「{dc3}」")
                elif cc3:
                    st.markdown(
                        f'<div class="mapping-box">'
                        f'📋 偵測到班級欄位：「{cc3}」<br>'
                        f'系統將自動從班級名稱映射科系</div>',
                        unsafe_allow_html=True
                    )
                    sample = p3df[cc3].value_counts().head(8)
                    for cn, cnt in sample.items():
                        dept = class_to_dept(str(cn))
                        tag = f" → {dept}" if dept else " → ❓未映射"
                        st.caption(f"　{cn}（{cnt}人）{tag}")
                else:
                    st.caption("⚠️ 未偵測到科系或班級欄位")
                    manual_cc = st.selectbox(
                        "手動指定班級欄位：",
                        ["-- 無 --"] + list(p3df.columns),
                        key=f"mcc_{yr}"
                    )
                    if manual_cc != "-- 無 --":
                        ydata["class_col_override"] = manual_cc

                ch_col = detect_final_ch_col(p3df)
                if ch_col:
                    st.caption(f"📌 入學方式欄位：「{ch_col}」")
                    vals = p3df[ch_col].fillna("(空白)").astype(str).str.strip()
                    vals = vals.replace("", "(空白)")
                    ch_dist = vals.value_counts()
                    all_chs = ch_dist.index.tolist()
                    for cn, cnt in ch_dist.head(8).items():
                        st.markdown(
                            f'<span class="channel-tag">{cn}</span> {cnt}人',
                            unsafe_allow_html=True
                        )
                    if len(ch_dist) > 8:
                        st.caption(f"... 共 {len(ch_dist)} 種管道")
                    sel_chs = st.multiselect(
                        "納入分析的管道：", all_chs, default=all_chs,
                        key=f"chs_{yr}"
                    )
                    ydata["channel_col"] = ch_col
                    ydata["selected_channels"] = sel_chs
                else:
                    st.caption("⚠️ 未偵測到入學方式欄位")
                    manual = st.selectbox(
                        "手動選擇入學方式欄位：",
                        ["-- 無 --"] + list(p3df.columns),
                        key=f"mch_{yr}"
                    )
                    if manual != "-- 無 --":
                        ydata["channel_col"] = manual
                        vals = p3df[manual].fillna("(空白)").value_counts()
                        all_chs = vals.index.tolist()
                        sel_chs = st.multiselect(
                            "納入管道：", all_chs, default=all_chs,
                            key=f"mchs_{yr}"
                        )
                        ydata["selected_channels"] = sel_chs

            if st.button(f"🗑️ 刪除 {yr}", key=f"del_{yr}"):
                del st.session_state["years"][yr]
                st.rerun()

    st.markdown("---")
    if st.button("🔄 更新分析", type="primary", use_container_width=True):
        st.session_state["analysis_ready"] = True
        st.session_state["analysis_version"] += 1
        st.success(f"✅ 分析已更新！版本 #{st.session_state['analysis_version']}")

    if st.session_state["analysis_ready"]:
        n_years = len([
            y for y in st.session_state["years"]
            if st.session_state["years"][y].get("p1") not in [None, "-- 未選擇 --"]
        ])
        ver = st.session_state["analysis_version"]
        st.markdown(
            f'<div class="success-box">✅ 分析就緒　{n_years} 個年度<br>'
            f'版本 #{ver}</div>',
            unsafe_allow_html=True
        )


# ============================================================
# 取得某年度三階段資料（含班級映射）
# ============================================================
def get_year_dfs(yr):
    ydata = st.session_state["years"].get(yr, {})
    p1 = p2 = p3 = None
    geo = None
    s1 = ydata.get("p1")
    s2 = ydata.get("p2")
    s3 = ydata.get("p3")
    if s1 and s1 != "-- 未選擇 --" and s1 in st.session_state["all_files"]:
        p1 = st.session_state["all_files"][s1].copy()
        geo = build_geo_from_p1(p1)
    if s2 and s2 != "-- 未選擇 --" and s2 in st.session_state["all_files"]:
        p2 = st.session_state["all_files"][s2].copy()
        if geo is not None:
            p2 = enrich_geo(p2, geo)
    if s3 and s3 != "-- 未選擇 --" and s3 in st.session_state["all_files"]:
        p3 = st.session_state["all_files"][s3].copy()
        ch_col = ydata.get("channel_col")
        sel_chs = ydata.get("selected_channels")
        if ch_col and ch_col in p3.columns and sel_chs:
            p3[ch_col] = p3[ch_col].fillna("(空白)").astype(str).str.strip()
            p3.loc[p3[ch_col] == "", ch_col] = "(空白)"
            p3 = p3[p3[ch_col].isin(sel_chs)]
        if geo is not None:
            p3 = enrich_geo(p3, geo)

        # ── 班級→科系映射 ──
        dc3 = detect_dept_col(p3)
        if dc3 is None:
            cc3 = detect_class_col(p3)
            if cc3 is None:
                cc3 = ydata.get("class_col_override")
            if cc3 and cc3 in p3.columns:
                p1_depts = None
                if p1 is not None:
                    dc1 = detect_dept_col(p1)
                    if dc1:
                        p1_depts = p1[dc1].dropna().unique().tolist()
                p3, mapping = auto_map_class_to_dept(p3, cc3, p1_depts)
                p3["科系"] = p3["_mapped_dept"]
                p3 = p3.drop(columns=["_mapped_dept"], errors="ignore")

    return p1, p2, p3, geo, ydata.get("channel_col")


# ============================================================
# 檢查就緒
# ============================================================
if not st.session_state["analysis_ready"]:
    st.markdown("""
    <div class="warning-box">
    <h4>⏳ 尚未執行分析</h4>
    <ol>
    <li>上傳招生資料檔案（所有年度）</li>
    <li>新增年度標籤（如 113學年）</li>
    <li>為每個年度指定一階/二階/最終入學</li>
    <li>點擊 <b>🔄 更新分析</b></li>
    </ol>
    <p>📌 最終入學若只有班級名稱，系統會自動映射科系<br>
    📍 經緯度從一階資料的「緯度/經度」欄位讀取</p>
    </div>""", unsafe_allow_html=True)
    st.stop()

valid_years = [
    yr for yr in st.session_state["years"]
    if st.session_state["years"][yr].get("p1") not in [None, "-- 未選擇 --"]
]
if not valid_years:
    st.warning("⚠️ 沒有任何年度設定了一階資料。")
    st.stop()


# ============================================================
# 欄位診斷
# ============================================================
def show_field_diagnosis(p1, p2, p3, yr_label):
    with st.expander(f"🔍 {yr_label} 欄位偵測診斷", expanded=False):
        diag = []
        for phase, df, label in [("一階", p1, "P1"), ("二階", p2, "P2"), ("最終", p3, "P3")]:
            if df is None:
                diag.append({"階段": phase, "科系欄位": "—", "班級欄位": "—",
                             "學校欄位": "—", "筆數": 0})
                continue
            dc = detect_dept_col(df)
            cc = detect_class_col(df)
            sc = detect_school_col(df)
            diag.append({
                "階段": phase,
                "科系欄位": dc if dc else "❌",
                "班級欄位": cc if cc else "—",
                "學校欄位": sc if sc else "❌",
                "筆數": len(df)
            })
        st.dataframe(pd.DataFrame(diag), use_container_width=True, hide_index=True)

        if p3 is not None:
            dc3 = detect_dept_col(p3)
            if dc3:
                vals = p3[dc3].dropna().apply(norm_dept).unique()
                st.caption(f"最終入學科系（直接）：{', '.join(vals[:10])}")
            elif "科系" in p3.columns:
                mapped = p3["科系"].dropna().unique()
                unmapped = p3["科系"].isna().sum()
                st.markdown(
                    f'<div class="mapping-box">'
                    f'📋 班級→科系映射結果：<br>'
                    f'✅ 已映射：{len(p3) - unmapped} 筆<br>'
                    f'❌ 未映射：{unmapped} 筆<br>'
                    f'📌 映射到的科系：{", ".join(str(d) for d in mapped[:10])}</div>',
                    unsafe_allow_html=True
                )
                if unmapped > 0:
                    cc3 = detect_class_col(p3)
                    if cc3 is None:
                        for c in p3.columns:
                            if "班" in str(c):
                                cc3 = c
                                break
                    if cc3 and cc3 in p3.columns:
                        unmapped_classes = p3[p3["科系"].isna()][cc3].value_counts()
                        st.caption("未映射的班級名稱：")
                        for cn, cnt in unmapped_classes.head(10).items():
                            st.caption(f"　❓ {cn}（{cnt}人）")

        dc1 = detect_dept_col(p1) if p1 is not None else None
        if dc1 and p3 is not None and "科系" in p3.columns:
            set1 = set(p1[dc1].dropna().apply(norm_dept).unique())
            set3 = set(p3["科系"].dropna().apply(norm_dept).unique())
            matched = set1 & set3
            only1 = set1 - set3
            only3 = set3 - set1
            st.markdown(
                f"**科系匹配：** 共同 {len(matched)} ｜ "
                f"僅一階 {len(only1)} ｜ 僅最終 {len(only3)}"
            )
            if only1:
                st.caption(f"　僅一階：{', '.join(sorted(only1)[:10])}")
            if only3:
                st.caption(f"　僅最終：{', '.join(sorted(only3)[:10])}")


# ============================================================
# 年度單獨分析模組
# ============================================================
def render_year_analysis(yr):
    p1, p2, p3, geo, ch_col = get_year_dfs(yr)
    if p1 is None:
        st.warning(f"⚠️ {yr}：一階資料未指定或無法讀取。")
        return

    mod_opts = [
        "📊 總覽儀表板", "🔄 招生漏斗", "📈 入學管道",
        "🗺️ 地理分布", "🏫 科系熱力圖", "🎯 來源學校", "⚠️ 流失預警"
    ]
    mod = st.radio("選擇分析模組：", mod_opts, horizontal=True, key=f"mod_{yr}")
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    n1 = len(p1)
    n2 = len(p2) if p2 is not None else None
    n3 = len(p3) if p3 is not None else None

    # ── 總覽 ──
    if "總覽" in mod:
        st.subheader(f"📊 {yr} — 總覽儀表板")
        cols = st.columns(5)
        with cols[0]:
            st.markdown(
                f'<div class="metric-card"><h3>一階報名</h3><h1>{n1:,}</h1></div>',
                unsafe_allow_html=True
            )
        with cols[1]:
            v = f"{n2:,}" if n2 else "—"
            st.markdown(
                f'<div class="metric-card metric-orange"><h3>二階報到</h3><h1>{v}</h1></div>',
                unsafe_allow_html=True
            )
        with cols[2]:
            v = f"{n3:,}" if n3 else "—"
            st.markdown(
                f'<div class="metric-card metric-green"><h3>最終入學</h3><h1>{v}</h1></div>',
                unsafe_allow_html=True
            )
        with cols[3]:
            r = f"{n2 / n1 * 100:.1f}%" if n2 and n1 else "—"
            st.markdown(
                f'<div class="metric-card metric-blue"><h3>一→二階</h3><h1>{r}</h1></div>',
                unsafe_allow_html=True
            )
        with cols[4]:
            r = f"{n3 / n1 * 100:.1f}%" if n3 and n1 else "—"
            st.markdown(
                f'<div class="metric-card metric-gold"><h3>一→最終</h3><h1>{r}</h1></div>',
                unsafe_allow_html=True
            )

        if geo is not None:
            st.caption(f"📍 經緯度資料庫（一階）：{len(geo)} 所學校")

        if p3 is not None and "科系" in p3.columns:
            s3_name = st.session_state["years"][yr].get("p3")
            dc3_orig = None
            if s3_name and s3_name in st.session_state["all_files"]:
                dc3_orig = detect_dept_col(st.session_state["all_files"][s3_name])
            if dc3_orig is None:
                n_mapped = p3["科系"].notna().sum()
                n_unmapped = p3["科系"].isna().sum()
                status = "✅ 全部映射成功" if n_unmapped == 0 else f"❌ 未映射 {n_unmapped} 筆"
                st.markdown(
                    f'<div class="mapping-box">'
                    f'📋 <b>班級→科系映射</b>：已映射 {n_mapped} 筆 | {status}'
                    f'</div>',
                    unsafe_allow_html=True
                )

        show_field_diagnosis(p1, p2, p3, yr)

        if p3 is not None and ch_col and ch_col in p3.columns:
            st.markdown("---")
            st.subheader("🟢 最終入學管道分布")
            cd = p3[ch_col].value_counts().reset_index()
            cd.columns = ["入學管道", "人數"]
            cd["佔比(%)"] = (cd["人數"] / cd["人數"].sum() * 100).round(1)
            c1, c2 = st.columns(2)
            with c1:
                fig = px.pie(cd, names="入學管道", values="人數",
                             title="管道佔比", hole=.35)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = px.bar(
                    cd.sort_values("人數", ascending=True),
                    x="人數", y="入學管道", orientation="h", text="人數",
                    title="各管道人數", color="佔比(%)",
                    color_continuous_scale="Viridis"
                )
                fig.update_layout(
                    height=max(400, len(cd) * 28),
                    yaxis=dict(autorange="reversed")
                )
                st.plotly_chart(fig, use_container_width=True)

        fl, fv = ["一階報名"], [n1]
        if n2:
            fl.append("二階報到")
            fv.append(n2)
        if n3:
            fl.append("最終入學")
            fv.append(n3)
        if len(fv) > 1:
            st.plotly_chart(
                fig_funnel(fl, fv, f"{yr} 招生漏斗"),
                use_container_width=True
            )

        c1, c2 = st.columns(2)
        dc = detect_dept_col(p1)
        sc = detect_school_col(p1)
        with c1:
            if dc:
                dd = p1[dc].value_counts().reset_index()
                dd.columns = ["科系", "人數"]
                fig = px.pie(dd, names="科系", values="人數",
                             title="一階科系分布", hole=.4)
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            if sc:
                sd = p1[sc].value_counts().head(10).reset_index()
                sd.columns = ["學校", "人數"]
                fig = px.bar(sd, x="人數", y="學校", orientation="h",
                             title="來源學校 TOP 10", text="人數")
                fig.update_traces(marker_color="#667eea")
                fig.update_layout(yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig, use_container_width=True)

        result = build_dept_stats(p1, p2, p3)
        if result is not None:
            ds, p3_info = result
            st.markdown("---")
            st.subheader("各科系三階段概覽")

            if p3_info and p3_info.get("method") == "class_mapping":
                mp = p3_info.get("mapped", 0)
                ump = p3_info.get("unmapped", 0)
                mapping = p3_info.get("mapping", {})
                col_name = p3_info.get("col", "?")
                st.markdown(
                    f'<div class="mapping-box">'
                    f'📋 最終入學透過<b>班級→科系映射</b>（欄位：{col_name}）<br>'
                    f'✅ 已映射 {mp} 筆 ｜ ❌ 未映射 {ump} 筆</div>',
                    unsafe_allow_html=True
                )
                if mapping:
                    with st.expander("🔍 查看班級→科系映射表", expanded=False):
                        mdf = pd.DataFrame([
                            {"班級名稱": k, "映射科系": v}
                            for k, v in sorted(mapping.items())
                        ])
                        st.dataframe(mdf, use_container_width=True, hide_index=True)

            if p3 is not None and ds["最終入學"].sum() == 0:
                st.markdown(
                    '<div class="warning-box">'
                    '⚠️ 最終入學全部為 0！可能原因：<br>'
                    '• 班級名稱無法映射到科系（請檢查映射表）<br>'
                    '• 請展開「🔍 欄位偵測診斷」查看未映射班級<br>'
                    '• 可在 DEPT_KEYWORDS 中新增對應關鍵字</div>',
                    unsafe_allow_html=True
                )

            st.dataframe(
                ds.sort_values("一階人數", ascending=False),
                use_container_width=True, hide_index=True
            )

    # ── 招生漏斗 ──
    elif "漏斗" in mod:
        st.subheader(f"🔄 {yr} — 招生漏斗分析")
        result = build_dept_stats(p1, p2, p3)
        if result is not None:
            ds, _ = result
            st.dataframe(
                ds.sort_values("一→最終(%)", ascending=False),
                use_container_width=True, hide_index=True
            )
            rc = [c for c in ["一→二階(%)", "一→最終(%)"] if c in ds.columns]
            if rc:
                st.plotly_chart(
                    fig_grouped_bar(ds.sort_values(rc[0], ascending=True),
                                    "科系", rc, "各科系轉換率"),
                    use_container_width=True
                )

            st.subheader("單科系漏斗")
            sel = st.selectbox("選擇科系：", ds["科系"].tolist(),
                               key=f"fun_dept_{yr}")
            row = ds[ds["科系"] == sel].iloc[0]
            fl, fv = ["一階報名"], [int(row["一階人數"])]
            if row["二階人數"] > 0 or p2 is not None:
                fl.append("二階報到")
                fv.append(int(row["二階人數"]))
            if row["最終入學"] > 0 or p3 is not None:
                fl.append("最終入學")
                fv.append(int(row["最終入學"]))
            st.plotly_chart(
                fig_funnel(fl, fv, f"{sel} 漏斗"),
                use_container_width=True
            )

        ss = build_school_stats(p1, p2, p3)
        if ss is not None:
            st.markdown("---")
            st.subheader("各來源學校漏斗")
            mn = st.slider("一階≥", 1, 50, 5, key=f"fun_mn_{yr}")
            sf = ss[ss["一階人數"] >= mn].sort_values("一→最終(%)", ascending=False)
            st.dataframe(sf, use_container_width=True, hide_index=True)
            st.plotly_chart(
                fig_bar_h(sf.head(20), "學校", "一→最終(%)",
                          f"來源學校轉換率 TOP 20（一階≥{mn}）", color="#4CAF50"),
                use_container_width=True
            )

    # ── 入學管道 ──
    elif "管道" in mod:
        st.subheader(f"📈 {yr} — 入學管道分析")
        if p3 is None or not ch_col or ch_col not in (p3.columns if p3 is not None else []):
            st.warning("⚠️ 需要最終入學資料及入學方式欄位。")
            return
        cd = p3[ch_col].value_counts().reset_index()
        cd.columns = ["入學管道", "人數"]
        cd["佔比(%)"] = (cd["人數"] / cd["人數"].sum() * 100).round(1)
        cd["累積(%)"] = cd["佔比(%)"].cumsum().round(1)
        c1, c2 = st.columns(2)
        with c1:
            fig = px.pie(cd, names="入學管道", values="人數",
                         title="管道佔比", hole=.35)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = px.bar(
                cd.sort_values("人數", ascending=True),
                x="人數", y="入學管道", orientation="h", text="人數",
                title="人數排行", color="佔比(%)",
                color_continuous_scale="Viridis"
            )
            fig.update_layout(
                height=max(400, len(cd) * 30),
                yaxis=dict(autorange="reversed")
            )
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(cd, use_container_width=True, hide_index=True)

        dept_col = detect_dept_col(p3)
        if dept_col is None and "科系" in p3.columns:
            dept_col = "科系"
        if dept_col:
            st.markdown("---")
            st.subheader("管道 × 科系")
            valid_p3 = p3[p3[dept_col].notna()] if dept_col == "科系" else p3
            cross = valid_p3.groupby([ch_col, dept_col]).size().reset_index(name="人數")
            fig = px.bar(cross, x=ch_col, y="人數", color=dept_col,
                         barmode="stack", text="人數",
                         title="管道×科系堆疊圖")
            fig.update_layout(height=600, xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
            st.plotly_chart(
                fig_heatmap(cross, dept_col, ch_col, "人數", "管道×科系熱力圖"),
                use_container_width=True
            )

        sc3 = detect_school_col(p3)
        if sc3:
            st.markdown("---")
            st.subheader("管道 × 來源學校多元性")
            dv = p3.groupby(ch_col)[sc3].nunique().reset_index()
            dv.columns = ["入學管道", "來源學校數"]
            fig = px.bar(
                dv.sort_values("來源學校數", ascending=True),
                x="來源學校數", y="入學管道", orientation="h",
                text="來源學校數", color="來源學校數",
                color_continuous_scale="Blues", title="學校多元性"
            )
            fig.update_layout(height=max(400, len(dv) * 28))
            st.plotly_chart(fig, use_container_width=True)

    # ── 地理分布 ──
    elif "地理" in mod:
        st.subheader(f"🗺️ {yr} — 地理分布")
        if geo is None:
            st.warning("⚠️ 一階資料無經緯度欄位，無法繪製地圖。")
            return
        st.caption(f"📍 經緯度資料庫（一階）：{len(geo)} 所學校")
        sc = detect_school_col(p1)
        if not sc:
            st.warning("⚠️ 未偵測到學校欄位。")
            return

        def do_map(src_df, count_label, title_text, phase):
            sc_ = detect_school_col(src_df)
            if sc_ is None:
                return
            agg = src_df.groupby(sc_).size().reset_index(name=count_label)
            agg["_std"] = agg[sc_].apply(norm_school)
            agg = agg.merge(geo, on="_std", how="left").drop(columns=["_std"])
            fig = fig_map(agg, count_label, title_text)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
                ok = agg["lat"].notna().sum()
                total = len(agg)
                pct = ok / total * 100 if total > 0 else 0
                st.caption(f"匹配：{ok}/{total} 校（{pct:.1f}%）")
                miss = agg[agg["lat"].isna()]
                if not miss.empty:
                    with st.expander(f"⚠️ {phase} — {len(miss)} 校未匹配"):
                        st.dataframe(
                            miss[[sc_, count_label]].sort_values(
                                count_label, ascending=False),
                            hide_index=True
                        )
            else:
                st.info(f"{phase} 無匹配結果。")

        st.subheader("一階報名地圖")
        do_map(p1, "報名人數", f"{yr} 一階報名來源", "一階")
        if p2 is not None:
            st.markdown("---")
            st.subheader("二階報到地圖")
            do_map(p2, "報到人數", f"{yr} 二階報到來源", "二階")
        if p3 is not None:
            st.markdown("---")
            st.subheader("最終入學地圖")
            do_map(p3, "入學人數", f"{yr} 最終入學來源", "最終入學")
            if ch_col and ch_col in p3.columns:
                st.markdown("---")
                st.subheader("指定管道地圖")
                chl = p3[ch_col].value_counts().index.tolist()
                sel_ch = st.selectbox("選管道：", chl, key=f"mapch_{yr}")
                if sel_ch:
                    sub = p3[p3[ch_col] == sel_ch]
                    do_map(sub, "入學人數",
                           f'{yr}「{sel_ch}」地圖', f'「{sel_ch}」')

    # ── 科系熱力圖 ──
    elif "熱力圖" in mod:
        st.subheader(f"🏫 {yr} — 科系×學校 熱力圖")
        dc = detect_dept_col(p1)
        sc = detect_school_col(p1)
        if not dc or not sc:
            st.warning("⚠️ 未偵測到科系或學校欄位。")
            return
        mn = st.slider("學校報名≥", 1, 30, 3, key=f"hm_{yr}")
        valid = p1[sc].value_counts()
        valid = valid[valid >= mn].index
        filt = p1[p1[sc].isin(valid)]
        cr = filt.groupby([sc, dc]).size().reset_index(name="人數")
        st.plotly_chart(
            fig_heatmap(cr, dc, sc, "人數", f"一階報名（≥{mn}人）"),
            use_container_width=True
        )
        if p3 is not None:
            dept_col_p3 = detect_dept_col(p3)
            if dept_col_p3 is None and "科系" in p3.columns:
                dept_col_p3 = "科系"
            sc3 = detect_school_col(p3)
            if dept_col_p3 and sc3:
                valid_p3 = p3[p3[dept_col_p3].notna()] if dept_col_p3 == "科系" else p3
                cr3 = valid_p3.groupby([sc3, dept_col_p3]).size().reset_index(name="入學人數")
                cr3 = cr3[cr3[sc3].isin(valid)]
                if not cr3.empty:
                    st.plotly_chart(
                        fig_heatmap(cr3, dept_col_p3, sc3, "入學人數", "最終入學"),
                        use_container_width=True
                    )

    # ── 來源學校 ──
    elif "來源學校" in mod:
        st.subheader(f"🎯 {yr} — 來源學校追蹤")
        ss = build_school_stats(p1, p2, p3)
        if ss is None:
            st.warning("⚠️ 未偵測到學校欄位。")
            return

        def tier(n):
            if n >= 30:
                return "Tier1(≥30)"
            elif n >= 10:
                return "Tier2(10-29)"
            else:
                return "Tier3(<10)"

        ss["分級"] = ss["一階人數"].apply(tier)
        sel_t = st.multiselect(
            "篩選分級：",
            ["Tier1(≥30)", "Tier2(10-29)", "Tier3(<10)"],
            default=["Tier1(≥30)", "Tier2(10-29)"],
            key=f"tier_{yr}"
        )
        disp = ss[ss["分級"].isin(sel_t)].sort_values("一階人數", ascending=False)
        st.dataframe(disp, use_container_width=True, hide_index=True)

        st.subheader("個別學校")
        sel = st.selectbox(
            "選擇學校：",
            ss.sort_values("一階人數", ascending=False)["學校"],
            key=f"sch_{yr}"
        )
        if sel:
            r = ss[ss["學校"] == sel].iloc[0]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("一階", f'{int(r["一階人數"])}')
            c2.metric("二階", f'{int(r["二階人數"])}')
            c3.metric("最終", f'{int(r["最終入學"])}')
            c4.metric("轉換率", f'{r["一→最終(%)"]}%')
            fl, fv = ["一階"], [int(r["一階人數"])]
            if r["二階人數"] > 0 or p2 is not None:
                fl.append("二階")
                fv.append(int(r["二階人數"]))
            if r["最終入學"] > 0 or p3 is not None:
                fl.append("最終")
                fv.append(int(r["最終入學"]))
            if len(fv) > 1:
                st.plotly_chart(
                    fig_funnel(fl, fv, f"{sel} 漏斗"),
                    use_container_width=True
                )

            if p3 is not None and ch_col and ch_col in p3.columns:
                sc3 = detect_school_col(p3)
                if sc3:
                    p3_tmp = p3.copy()
                    p3_tmp["_sch_std"] = p3_tmp[sc3].apply(norm_school)
                    sel_std = norm_school(sel)
                    sub = p3_tmp[p3_tmp["_sch_std"] == sel_std]
                    if not sub.empty:
                        dd = sub[ch_col].value_counts().reset_index()
                        dd.columns = ["管道", "人數"]
                        fig = px.pie(dd, names="管道", values="人數",
                                     title=f"{sel} 入學管道", hole=.35)
                        st.plotly_chart(fig, use_container_width=True)

    # ── 流失預警 ──
    elif "流失" in mod:
        st.subheader(f"⚠️ {yr} — 流失預警分析")
        if p2 is None and p3 is None:
            st.warning("⚠️ 需要至少二階或最終入學資料。")
            return
        ss = build_school_stats(p1, p2, p3)
        if ss is None:
            st.warning("⚠️ 未偵測到學校欄位。")
            return
        has_final = p3 is not None and ss["最終入學"].sum() > 0
        rc = "一→最終(%)" if has_final else "一→二階(%)"
        ll = "最終入學" if has_final else "二階人數"
        sl = "一→最終" if has_final else "一→二階"
        ss["流失人數"] = ss["一階人數"] - ss[ll]

        mn = st.slider("一階≥", 1, 50, 10, key=f"loss_mn_{yr}")
        pool = ss[ss["一階人數"] >= mn]
        avg = pool[rc].mean()
        warn = pool[pool[rc] < avg].sort_values("流失人數", ascending=False)
        if warn.empty:
            st.success("✅ 沒有預警學校！")
        else:
            st.markdown(
                f'<div class="warning-box">⚠️ {len(warn)}所學校低於平均 {avg:.1f}%</div>',
                unsafe_allow_html=True
            )
            st.dataframe(warn, use_container_width=True, hide_index=True)

        st.subheader("IPA 矩陣")
        ana = ss[ss["一階人數"] >= mn].copy()
        if not ana.empty:
            med_x = ana["一階人數"].median()
            med_y = ana[rc].median()
            fig = px.scatter(
                ana, x="一階人數", y=rc, size="流失人數",
                hover_name="學校",
                hover_data={"一階人數": True, ll: True, rc: True},
                title=f"IPA（{sl}）", size_max=40,
                color=rc, color_continuous_scale="RdYlGn"
            )
            fig.add_hline(y=med_y, line_dash="dash", line_color="red",
                          annotation_text=f"中位數 {med_y:.1f}%")
            fig.add_vline(x=med_x, line_dash="dash", line_color="blue",
                          annotation_text=f"中位數 {med_x:.0f}")
            fig.update_layout(height=600)
            st.plotly_chart(fig, use_container_width=True)

        if p3 is not None and ch_col and ch_col in p3.columns:
            st.markdown("---")
            st.subheader("Pareto 分析")
            cs = p3[ch_col].value_counts().reset_index()
            cs.columns = ["入學管道", "人數"]
            cs["佔比(%)"] = (cs["人數"] / cs["人數"].sum() * 100).round(1)
            cs["累積(%)"] = cs["佔比(%)"].cumsum().round(1)
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=cs["入學管道"], y=cs["人數"], name="人數",
                marker_color="#4CAF50", text=cs["人數"],
                textposition="outside"
            ))
            fig.add_trace(go.Scatter(
                x=cs["入學管道"], y=cs["累積(%)"], name="累積%",
                yaxis="y2", line=dict(color="#FF5722", width=3),
                marker=dict(size=8)
            ))
            fig.update_layout(
                title="Pareto圖",
                yaxis=dict(title="人數"),
                yaxis2=dict(title="累積%", overlaying="y",
                            side="right", range=[0, 105]),
                height=500, xaxis_tickangle=-45
            )
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.subheader("科系流失")
        result = build_dept_stats(p1, p2, p3)
        if result is not None:
            ds, _ = result
            has_final_d = p3 is not None and ds["最終入學"].sum() > 0
            drc = "一→最終(%)" if has_final_d else "一→二階(%)"
            dll = "最終入學" if has_final_d else "二階人數"
            ds["流失人數"] = ds["一階人數"] - ds[dll]
            fig = px.scatter(
                ds, x="一階人數", y=drc, size="流失人數",
                hover_name="科系", text="科系",
                title=f"科系 IPA（{sl}）", size_max=50,
                color=drc, color_continuous_scale="RdYlGn"
            )
            fig.update_traces(textposition="top center")
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(
                ds.sort_values("流失人數", ascending=False),
                use_container_width=True, hide_index=True
            )


# ============================================================
# 跨年度比較模組
# ============================================================
def render_cross_year():
    st.header("📊 跨年度比較分析")

    if len(valid_years) < 2:
        st.warning("⚠️ 需要至少 2 個年度才能進行跨年度比較。")
        return

    summaries = []
    for yr in valid_years:
        p1, p2, p3, geo, ch_col = get_year_dfs(yr)
        if p1 is None:
            continue
        n1 = len(p1)
        n2 = len(p2) if p2 is not None else 0
        n3 = len(p3) if p3 is not None else 0
        sc = detect_school_col(p1)
        dc = detect_dept_col(p1)
        n_sch = p1[sc].nunique() if sc else 0
        n_dept = p1[dc].nunique() if dc else 0
        summaries.append({
            "年度": yr, "一階人數": n1, "二階人數": n2, "最終入學": n3,
            "一→二階(%)": round(n2 / n1 * 100, 1) if n1 and n2 else 0,
            "一→最終(%)": round(n3 / n1 * 100, 1) if n1 and n3 else 0,
            "來源學校數": n_sch, "科系數": n_dept
        })
    if not summaries:
        return
    sdf = pd.DataFrame(summaries)

    st.subheader("7-1. 年度總覽")
    st.dataframe(sdf, use_container_width=True, hide_index=True)

    st.subheader("7-2. 招生量趨勢")
    fig = go.Figure()
    fig.add_trace(go.Bar(x=sdf["年度"], y=sdf["一階人數"],
                         name="一階", marker_color="#2196F3"))
    if sdf["二階人數"].sum() > 0:
        fig.add_trace(go.Bar(x=sdf["年度"], y=sdf["二階人數"],
                             name="二階", marker_color="#FF9800"))
    if sdf["最終入學"].sum() > 0:
        fig.add_trace(go.Bar(x=sdf["年度"], y=sdf["最終入學"],
                             name="最終入學", marker_color="#4CAF50"))
    fig.update_layout(barmode="group", title="各年度招生量", height=450)
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("7-3. 轉換率趨勢")
    fig = go.Figure()
    if sdf["一→二階(%)"].sum() > 0:
        fig.add_trace(go.Scatter(
            x=sdf["年度"], y=sdf["一→二階(%)"], name="一→二階",
            mode="lines+markers+text", text=sdf["一→二階(%)"],
            textposition="top center",
            line=dict(width=3, color="#FF9800"), marker=dict(size=12)
        ))
    if sdf["一→最終(%)"].sum() > 0:
        fig.add_trace(go.Scatter(
            x=sdf["年度"], y=sdf["一→最終(%)"], name="一→最終",
            mode="lines+markers+text", text=sdf["一→最終(%)"],
            textposition="top center",
            line=dict(width=3, color="#4CAF50"), marker=dict(size=12)
        ))
    fig.update_layout(title="轉換率趨勢", yaxis_title="轉換率(%)", height=400)
    st.plotly_chart(fig, use_container_width=True)

    # 7-4 科系跨年度
    st.markdown("---")
    st.subheader("7-4. 科系跨年度比較")
    all_depts = set()
    dept_data = {}
    for yr in valid_years:
        p1, p2, p3, _, _ = get_year_dfs(yr)
        if p1 is None:
            continue
        result = build_dept_stats(p1, p2, p3)
        if result is not None:
            ds, _ = result
            ds["_key"] = ds["科系"].apply(norm_dept)
            dept_data[yr] = ds
            all_depts.update(ds["科系"].tolist())
    if dept_data and all_depts:
        sel_dept = st.selectbox("選擇科系：", sorted(all_depts), key="cross_dept")
        sel_key = norm_dept(sel_dept)
        rows = []
        for yr, ds in dept_data.items():
            r = ds[ds["_key"] == sel_key]
            if not r.empty:
                r = r.iloc[0]
                rows.append({
                    "年度": yr,
                    "一階": int(r["一階人數"]),
                    "二階": int(r["二階人數"]),
                    "最終": int(r["最終入學"]),
                    "一→最終(%)": r["一→最終(%)"]
                })
        if rows:
            rdf = pd.DataFrame(rows)
            st.dataframe(rdf, use_container_width=True, hide_index=True)
            fig = go.Figure()
            fig.add_trace(go.Bar(x=rdf["年度"], y=rdf["一階"],
                                 name="一階", marker_color="#2196F3"))
            fig.add_trace(go.Bar(x=rdf["年度"], y=rdf["最終"],
                                 name="最終", marker_color="#4CAF50"))
            fig.add_trace(go.Scatter(
                x=rdf["年度"], y=rdf["一→最終(%)"], name="轉換率",
                yaxis="y2", mode="lines+markers",
                line=dict(color="#E91E63", width=3), marker=dict(size=10)
            ))
            fig.update_layout(
                barmode="group", title=f"「{sel_dept}」跨年度趨勢",
                yaxis=dict(title="人數"),
                yaxis2=dict(title="轉換率(%)", overlaying="y", side="right"),
                height=450
            )
            st.plotly_chart(fig, use_container_width=True)

    # 7-5 來源學校跨年度
    st.markdown("---")
    st.subheader("7-5. 來源學校跨年度比較")
    all_schs = set()
    sch_data = {}
    for yr in valid_years:
        p1, p2, p3, _, _ = get_year_dfs(yr)
        if p1 is None:
            continue
        ss = build_school_stats(p1, p2, p3)
        if ss is not None:
            ss["_key"] = ss["學校"].apply(norm_school)
            sch_data[yr] = ss
            all_schs.update(ss["學校"].tolist())
    if sch_data and all_schs:
        sel_sch = st.selectbox("選擇學校：", sorted(all_schs), key="cross_sch")
        sel_key = norm_school(sel_sch)
        rows = []
        for yr, ss in sch_data.items():
            r = ss[ss["_key"] == sel_key]
            if not r.empty:
                r = r.iloc[0]
                rows.append({
                    "年度": yr,
                    "一階": int(r["一階人數"]),
                    "二階": int(r["二階人數"]),
                    "最終": int(r["最終入學"]),
                    "一→最終(%)": r["一→最終(%)"],
                    "流失": int(r["流失人數"])
                })
        if rows:
            rdf = pd.DataFrame(rows)
            st.dataframe(rdf, use_container_width=True, hide_index=True)
            fig = go.Figure()
            fig.add_trace(go.Bar(x=rdf["年度"], y=rdf["一階"],
                                 name="一階", marker_color="#2196F3"))
            fig.add_trace(go.Bar(x=rdf["年度"], y=rdf["最終"],
                                 name="最終", marker_color="#4CAF50"))
            fig.add_trace(go.Scatter(
                x=rdf["年度"], y=rdf["一→最終(%)"], name="轉換率",
                yaxis="y2", mode="lines+markers",
                line=dict(color="#E91E63", width=3), marker=dict(size=10)
            ))
            fig.update_layout(
                barmode="group", title=f"「{sel_sch}」跨年度趨勢",
                yaxis=dict(title="人數"),
                yaxis2=dict(title="轉換率(%)", overlaying="y", side="right"),
                height=450
            )
            st.plotly_chart(fig, use_container_width=True)

    # 7-6 入學管道跨年度
    st.markdown("---")
    st.subheader("7-6. 入學管道跨年度比較")
    ch_year_data = {}
    all_channels = set()
    for yr in valid_years:
        p1, p2, p3, _, ch_col = get_year_dfs(yr)
        if p3 is None or ch_col is None:
            continue
        if ch_col not in p3.columns:
            continue
        cd = p3[ch_col].value_counts().reset_index()
        cd.columns = ["管道", "人數"]
        cd["年度"] = yr
        ch_year_data[yr] = cd
        all_channels.update(cd["管道"].tolist())

    if ch_year_data and all_channels:
        combined = pd.concat(ch_year_data.values(), ignore_index=True)
        pv = combined.pivot_table(index="管道", columns="年度",
                                  values="人數", aggfunc="sum").fillna(0)
        fig = px.imshow(pv, text_auto=True, aspect="auto",
                        color_continuous_scale="YlGnBu",
                        title="入學管道×年度 人數矩陣")
        fig.update_layout(height=max(400, len(pv) * 30))
        st.plotly_chart(fig, use_container_width=True)

        sel_ch = st.selectbox("選擇管道查看趨勢：", sorted(all_channels),
                              key="cross_ch")
        if sel_ch:
            sub = combined[combined["管道"] == sel_ch].sort_values("年度")
            if not sub.empty:
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=sub["年度"], y=sub["人數"], name="入學人數",
                    marker_color="#4CAF50", text=sub["人數"],
                    textposition="outside"
                ))
                fig.add_trace(go.Scatter(
                    x=sub["年度"], y=sub["人數"], mode="lines+markers",
                    name="趨勢線",
                    line=dict(color="#FF5722", width=3, dash="dash")
                ))
                fig.update_layout(title=f"「{sel_ch}」跨年度趨勢", height=400)
                st.plotly_chart(fig, use_container_width=True)

    # 7-7 年度增減
    st.markdown("---")
    st.subheader("7-7. 年度增減分析")
    if len(sdf) >= 2:
        latest = sdf.iloc[0]
        prev = sdf.iloc[1]
        chg1 = latest["一階人數"] - prev["一階人數"]
        chg3 = latest["最終入學"] - prev["最終入學"]
        chg_r = latest["一→最終(%)"] - prev["一→最終(%)"]

        cols = st.columns(3)
        with cols[0]:
            st.metric(
                f"一階人數變化（{latest['年度']} vs {prev['年度']}）",
                f"{int(latest['一階人數']):,}", f"{int(chg1):+,}"
            )
        with cols[1]:
            st.metric("最終入學變化",
                      f"{int(latest['最終入學']):,}", f"{int(chg3):+,}")
        with cols[2]:
            st.metric("轉換率變化",
                      f"{latest['一→最終(%)']}%", f"{chg_r:+.1f}%")

        dept_data_chg = {}
        for yr in valid_years:
            p1, p2, p3, _, _ = get_year_dfs(yr)
            if p1 is None:
                continue
            result = build_dept_stats(p1, p2, p3)
            if result is not None:
                ds, _ = result
                ds["_key"] = ds["科系"].apply(norm_dept)
                dept_data_chg[yr] = ds

        if len(dept_data_chg) >= 2:
            yrs = list(dept_data_chg.keys())
            d1 = dept_data_chg[yrs[0]].set_index("_key")
            d2 = dept_data_chg[yrs[1]].set_index("_key")
            common = list(set(d1.index) & set(d2.index))
            if common:
                chg_rows = []
                for dk in common:
                    dn = d1.loc[dk, "科系"] if "科系" in d1.columns else dk
                    if isinstance(dn, pd.Series):
                        dn = dn.iloc[0]
                    v1 = d1.loc[dk, "一階人數"]
                    v2 = d2.loc[dk, "一階人數"]
                    f1 = d1.loc[dk, "最終入學"]
                    f2 = d2.loc[dk, "最終入學"]
                    if isinstance(v1, pd.Series):
                        v1 = v1.iloc[0]
                    if isinstance(v2, pd.Series):
                        v2 = v2.iloc[0]
                    if isinstance(f1, pd.Series):
                        f1 = f1.iloc[0]
                    if isinstance(f2, pd.Series):
                        f2 = f2.iloc[0]
                    chg_rows.append({
                        "科系": dn,
                        f"{yrs[0]}一階": int(v1),
                        f"{yrs[1]}一階": int(v2),
                        "一階增減": int(v1 - v2),
                        f"{yrs[0]}最終": int(f1),
                        f"{yrs[1]}最終": int(f2),
                        "最終增減": int(f1 - f2)
                    })
                chg_df = pd.DataFrame(chg_rows).sort_values(
                    "最終增減", ascending=False)
                st.dataframe(chg_df, use_container_width=True, hide_index=True)

                fig = px.bar(
                    chg_df, x="科系", y="最終增減", text="最終增減",
                    title=f"科系最終入學增減（{yrs[0]} vs {yrs[1]}）",
                    color="最終增減", color_continuous_scale="RdYlGn"
                )
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 主畫面
# ============================================================
tab_names = valid_years + (["📊 跨年度比較"] if len(valid_years) >= 2 else [])
tabs = st.tabs(tab_names)

for i, yr in enumerate(valid_years):
    with tabs[i]:
        st.markdown(
            f'<span class="year-tag">📅 {yr}</span>',
            unsafe_allow_html=True
        )
        render_year_analysis(yr)

if len(valid_years) >= 2:
    with tabs[-1]:
        render_cross_year()

# ============================================================
# Footer
# ============================================================
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
ver = st.session_state.get("analysis_version", 0)
st.markdown(
    '<div style="text-align:center;color:#aaa;font-size:.85rem;padding:10px;">'
    '🎓 中華醫事科技大學 招生數據分析系統 v6.2<br>'
    '班級→科系自動映射 ｜ 跨階段名稱正規化 ｜ 多年度標籤式分析<br>'
    '分析版本 #' + str(ver) +
    '</div>',
    unsafe_allow_html=True
)
