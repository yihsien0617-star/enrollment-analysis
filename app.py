# -*- coding: utf-8 -*-
"""
中華醫事科技大學 招生數據分析系統 v6.4
- 新增：跨年度七模組完整分析
- 修正：班級名稱結構化解析（學制+科系+年級+班別）
- 修正：P1科系優先映射引擎
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
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
    .trend-up{color:#4caf50;font-weight:bold;}
    .trend-down{color:#f44336;font-weight:bold;}
    .trend-flat{color:#ff9800;font-weight:bold;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🎓 中華醫事科技大學 招生數據分析系統</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Enrollment Analytics v6.4 ｜ 跨年度七模組完整分析 ｜ 班級結構化解析</div>', unsafe_allow_html=True)
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

# ============================================================
# 常數
# ============================================================
FINAL_CH_CANDIDATES = [
    "入學方式", "入學管道", "錄取管道", "招生管道", "管道",
    "入學途徑", "錄取方式", "報名管道"
]
HWU = {"lat": 22.9340, "lon": 120.2756}

DEPT_ALIAS = {
    "護理": "護理系", "醫技": "醫學檢驗生物技術系",
    "醫檢": "醫學檢驗生物技術系", "醫學檢驗": "醫學檢驗生物技術系",
    "醫學檢驗生物技術": "醫學檢驗生物技術系",
    "藥學": "藥學系", "視光": "視光系",
    "製藥": "製藥工程系", "製藥工程": "製藥工程系",
    "食營": "食品營養系", "食品營養": "食品營養系",
    "職安": "職業安全衛生系", "職業安全衛生": "職業安全衛生系",
    "環安": "環境與安全衛生工程系", "環衛": "環境與安全衛生工程系",
    "環境與安全衛生工程": "環境與安全衛生工程系",
    "資管": "資訊管理系", "資訊管理": "資訊管理系",
    "多媒體": "多媒體設計系", "多媒體設計": "多媒體設計系",
    "幼保": "幼兒保育系", "幼兒保育": "幼兒保育系",
    "運休": "運動健康與休閒系", "運動休閒": "運動健康與休閒系",
    "運動健康與休閒": "運動健康與休閒系",
    "餐旅": "餐旅管理系", "餐旅管理": "餐旅管理系",
    "觀光": "觀光休閒事業管理系", "觀光休閒": "觀光休閒事業管理系",
    "觀光休閒事業管理": "觀光休閒事業管理系",
    "妝管": "化妝品應用與管理系", "美妝": "化妝品應用與管理系",
    "化妝品": "化妝品應用與管理系", "化妝品應用": "化妝品應用與管理系",
    "化妝品應用與管理": "化妝品應用與管理系",
    "調理": "調理保健技術系", "調保": "調理保健技術系",
    "調理保健": "調理保健技術系", "調理保健技術": "調理保健技術系",
    "語治": "語言治療系", "語言治療": "語言治療系",
    "牙技": "牙體技術系", "牙體技術": "牙體技術系",
    "生科": "生物科技系", "生物科技": "生物科技系",
    "寵物": "寵物美容學位學程", "寵物美容": "寵物美容學位學程",
    "寵物照護": "寵物照護學位學程", "寵物經營": "寵物經營學位學程",
    "長照": "長期照護學位學程", "長期照護": "長期照護學位學程",
    "醫管": "醫務暨健康事務管理系", "醫務管理": "醫務暨健康事務管理系",
    "健康事務管理": "醫務暨健康事務管理系",
    "醫務暨健康事務管理": "醫務暨健康事務管理系",
}

# ============================================================
# 班級名稱結構化解析器
# ============================================================
CLASS_PATTERN = re.compile(
    r'^(二技|四技|二|四)?'
    r'(.+?)'
    r'([一二三四五六七1-7])'
    r'([甲乙丙丁戊A-Za-z])'
    r'(?:班)?$'
)
CLASS_PATTERN_SIMPLE = re.compile(
    r'^(二技|四技|二|四)?'
    r'(.+?)'
    r'(?:系|科|學位學程|學程)?$'
)


def parse_class_name(class_name):
    if not isinstance(class_name, str):
        return None
    s = re.sub(r'\s+', '', class_name.strip())
    if not s:
        return None
    m = CLASS_PATTERN.match(s)
    if m:
        prefix = m.group(1) or ""
        dept_kw = m.group(2)
        grade = m.group(3)
        section = m.group(4)
        if prefix in ("二", "二技"):
            program = "二技"
        elif prefix in ("四", "四技"):
            program = "四技"
        else:
            program = "五專"
        return (program, dept_kw, grade, section)
    m2 = CLASS_PATTERN_SIMPLE.match(s)
    if m2:
        prefix = m2.group(1) or ""
        dept_kw = m2.group(2)
        if prefix in ("二", "二技"):
            program = "二技"
        elif prefix in ("四", "四技"):
            program = "四技"
        else:
            program = ""
        return (program, dept_kw, "", "")
    return None


def resolve_dept_from_keyword(dept_kw, p1_depts=None):
    kw = dept_kw.strip()
    if p1_depts:
        p1_cores = {}
        for d in p1_depts:
            core = re.sub(r'(學位學程|學程|系|科)$', '', norm_dept(d))
            p1_cores[core] = d
        if kw in p1_cores:
            return p1_cores[kw]
        for core, dept in sorted(p1_cores.items(), key=lambda x: len(x[0]), reverse=True):
            if len(kw) >= 2 and (kw in core or core in kw):
                return dept
    if kw in DEPT_ALIAS:
        alias_result = DEPT_ALIAS[kw]
        if p1_depts:
            alias_core = re.sub(r'(學位學程|學程|系|科)$', '', norm_dept(alias_result))
            for d in p1_depts:
                d_core = re.sub(r'(學位學程|學程|系|科)$', '', norm_dept(d))
                if alias_core == d_core or alias_core in d_core or d_core in alias_core:
                    return d
        return alias_result
    if p1_depts and len(kw) >= 2:
        short = kw[:2]
        for d in p1_depts:
            if short in d:
                return d
    return None


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
    name = re.sub(r"\s+", "", name.strip())
    name = name.replace("臺", "台").replace("（", "(").replace("）", ")")
    for sfx in ["附設進修學校", "進修學校", "進修部"]:
        name = name.replace(sfx, "")
    return name


def norm_dept(name):
    if not isinstance(name, str):
        return str(name).strip()
    name = re.sub(r"\s+", "", name.strip())
    name = name.replace("臺", "台").replace("（", "(").replace("）", ")").replace("　", "")
    return name


def auto_map_class_to_dept(df, class_col, p1_depts=None):
    df = df.copy()
    classes = df[class_col].fillna("").astype(str).str.strip()
    unique_classes = classes.unique()
    mapping = {}
    match_detail = {}
    for cls in unique_classes:
        if not cls:
            continue
        parsed = parse_class_name(cls)
        if parsed is None:
            match_detail[cls] = "❌ 無法解析格式"
            continue
        program, dept_kw, grade, section = parsed
        dept = resolve_dept_from_keyword(dept_kw, p1_depts)
        if dept:
            mapping[cls] = dept
            match_detail[cls] = f"✅ [{program}] 關鍵字「{dept_kw}」→ {dept}"
        else:
            match_detail[cls] = f"❌ [{program}] 關鍵字「{dept_kw}」→ 找不到對應科系"
    df["_mapped_dept"] = classes.map(mapping)
    return df, mapping, match_detail


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


def trend_arrow(current, previous):
    if previous == 0:
        return "—"
    diff = current - previous
    pct = diff / previous * 100
    if pct > 5:
        return f'<span class="trend-up">▲ +{pct:.1f}%</span>'
    elif pct < -5:
        return f'<span class="trend-down">▼ {pct:.1f}%</span>'
    else:
        return f'<span class="trend-flat">► {pct:+.1f}%</span>'


# ============================================================
# 統計構建
# ============================================================
def get_dept_series(df, p1_depts=None):
    dc = detect_dept_col(df)
    if dc:
        s = df[dc].dropna().apply(norm_dept)
        return s, {"method": "direct", "col": dc, "mapped": len(s), "unmapped": 0}
    cc = detect_class_col(df)
    if cc:
        df_mapped, mapping, match_detail = auto_map_class_to_dept(df, cc, p1_depts)
        s = df_mapped["_mapped_dept"].dropna().apply(norm_dept)
        n_mapped = df_mapped["_mapped_dept"].notna().sum()
        n_unmapped = df_mapped["_mapped_dept"].isna().sum()
        return s, {"method": "class_mapping", "col": cc,
                    "mapped": n_mapped, "unmapped": n_unmapped,
                    "mapping": mapping, "match_detail": match_detail}
    return pd.Series(dtype=str), {"method": "none", "col": None}


def align_dept_name(norm_name, p1_name_map):
    if norm_name in p1_name_map:
        return norm_name
    core = re.sub(r'(學位學程|學程|系|科)$', '', norm_name)
    for p1_norm in p1_name_map:
        p1_core = re.sub(r'(學位學程|學程|系|科)$', '', p1_norm)
        if core == p1_core:
            return p1_norm
        if len(core) >= 3 and (core in p1_core or p1_core in core):
            return p1_norm
    if len(core) >= 2:
        short = core[:2]
        for p1_norm in p1_name_map:
            if short in p1_norm:
                return p1_norm
    return norm_name


def build_dept_stats(p1, p2=None, p3=None):
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
        s2, _ = get_dept_series(p2, p1_depts)
        if not s2.empty:
            s2_aligned = s2.map(lambda x: align_dept_name(x, name_map))
            t2 = s2_aligned.value_counts().reset_index()
            t2.columns = ["_dept_std", "二階人數"]
            s = s.merge(t2, on="_dept_std", how="left")
    if "二階人數" not in s.columns:
        s["二階人數"] = np.nan

    p3_info = None
    if p3 is not None:
        s3, p3_info = get_dept_series(p3, p1_depts)
        if not s3.empty:
            s3_aligned = s3.map(lambda x: align_dept_name(x, name_map))
            t3 = s3_aligned.value_counts().reset_index()
            t3.columns = ["_dept_std", "最終入學"]
            s = s.merge(t3, on="_dept_std", how="left")
            extra = t3[~t3["_dept_std"].isin(s["_dept_std"])]
            if not extra.empty:
                extra_rows = extra.copy()
                extra_rows["一階人數"] = 0
                if "二階人數" not in extra_rows.columns:
                    extra_rows["二階人數"] = 0
                s = pd.concat([s, extra_rows], ignore_index=True)
    if "最終入學" not in s.columns:
        s["最終入學"] = np.nan

    s["二階人數"] = s["二階人數"].fillna(0).astype(int)
    s["最終入學"] = s["最終入學"].fillna(0).astype(int)
    s["一→二階(%)"] = np.where(s["一階人數"] > 0, (s["二階人數"] / s["一階人數"] * 100).round(1), 0)
    s["一→最終(%)"] = np.where(s["一階人數"] > 0, (s["最終入學"] / s["一階人數"] * 100).round(1), 0)
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
    s["一→二階(%)"] = np.where(s["一階人數"] > 0, (s["二階人數"] / s["一階人數"] * 100).round(1), 0)
    s["一→最終(%)"] = np.where(s["一階人數"] > 0, (s["最終入學"] / s["一階人數"] * 100).round(1), 0)
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
    new_year = st.text_input("新增年度標籤：", placeholder="例如：113學年", key="new_year_input")
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
            ydata["p1"] = st.selectbox("🔵 一階（含經緯度）", file_opts, key=f"p1_{yr}",
                index=file_opts.index(ydata["p1"]) if ydata["p1"] in file_opts else 0)
            ydata["p2"] = st.selectbox("🟠 二階", file_opts, key=f"p2_{yr}",
                index=file_opts.index(ydata["p2"]) if ydata["p2"] in file_opts else 0)
            ydata["p3"] = st.selectbox("🟢 最終入學", file_opts, key=f"p3_{yr}",
                index=file_opts.index(ydata["p3"]) if ydata["p3"] in file_opts else 0)

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
                preview_p1_depts = None
                if ydata["p1"] and ydata["p1"] != "-- 未選擇 --":
                    p1df_tmp = st.session_state["all_files"][ydata["p1"]]
                    dc1_tmp = detect_dept_col(p1df_tmp)
                    if dc1_tmp:
                        preview_p1_depts = p1df_tmp[dc1_tmp].dropna().unique().tolist()
                if dc3:
                    st.caption(f"✅ 最終入學有科系欄位：「{dc3}」")
                elif cc3:
                    st.markdown(
                        f'<div class="mapping-box">'
                        f'📋 偵測到班級欄位：「{cc3}」<br>'
                        f'結構化解析：[學制][科系][年級][班別]</div>',
                        unsafe_allow_html=True)
                    sample = p3df[cc3].value_counts().head(8)
                    for cn, cnt in sample.items():
                        parsed = parse_class_name(str(cn))
                        if parsed:
                            prog, dept_kw, grade, sec = parsed
                            dept = resolve_dept_from_keyword(dept_kw, preview_p1_depts)
                            if dept:
                                st.caption(f"　✅ {cn}（{cnt}人）→ [{prog}]{dept_kw} → {dept}")
                            else:
                                st.caption(f"　❌ {cn}（{cnt}人）→ [{prog}]「{dept_kw}」找不到科系")
                        else:
                            st.caption(f"　❌ {cn}（{cnt}人）→ 無法解析格式")
                else:
                    st.caption("⚠️ 未偵測到科系或班級欄位")
                    manual_cc = st.selectbox("手動指定班級欄位：",
                        ["-- 無 --"] + list(p3df.columns), key=f"mcc_{yr}")
                    if manual_cc != "-- 無 --":
                        ydata["class_col_override"] = manual_cc

                ch_col = detect_final_ch_col(p3df)
                if ch_col:
                    st.caption(f"📌 入學方式欄位：「{ch_col}」")
                    vals = p3df[ch_col].fillna("(空白)").astype(str).str.strip().replace("", "(空白)")
                    ch_dist = vals.value_counts()
                    all_chs = ch_dist.index.tolist()
                    for cn, cnt in ch_dist.head(6).items():
                        st.markdown(f'<span class="channel-tag">{cn}</span> {cnt}人', unsafe_allow_html=True)
                    if len(ch_dist) > 6:
                        st.caption(f"... 共 {len(ch_dist)} 種管道")
                    sel_chs = st.multiselect("納入分析的管道：", all_chs, default=all_chs, key=f"chs_{yr}")
                    ydata["channel_col"] = ch_col
                    ydata["selected_channels"] = sel_chs
                else:
                    manual = st.selectbox("手動選擇入學方式欄位：",
                        ["-- 無 --"] + list(p3df.columns), key=f"mch_{yr}")
                    if manual != "-- 無 --":
                        ydata["channel_col"] = manual
                        vals = p3df[manual].fillna("(空白)").value_counts()
                        all_chs = vals.index.tolist()
                        sel_chs = st.multiselect("納入管道：", all_chs, default=all_chs, key=f"mchs_{yr}")
                        ydata["selected_channels"] = sel_chs

            if st.button(f"🗑️ 刪除 {yr}", key=f"del_{yr}"):
                del st.session_state["years"][yr]
                st.rerun()

    st.markdown("---")
    if st.button("🔄 更新分析", type="primary", use_container_width=True):
        st.session_state["analysis_ready"] = True
        st.session_state["analysis_version"] += 1
        st.success("✅ 分析已更新！版本 #" + str(st.session_state["analysis_version"]))

    if st.session_state["analysis_ready"]:
        n_years = len([y for y in st.session_state["years"]
                       if st.session_state["years"][y].get("p1") not in [None, "-- 未選擇 --"]])
        ver = st.session_state["analysis_version"]
        st.markdown(
            '<div class="success-box">✅ 分析就緒　' + str(n_years)
            + ' 個年度<br>版本 #' + str(ver) + '</div>', unsafe_allow_html=True)


# ============================================================
# 取得某年度資料
# ============================================================
def get_year_dfs(yr):
    ydata = st.session_state["years"].get(yr, {})
    p1 = p2 = p3 = None
    geo = None
    s1, s2, s3 = ydata.get("p1"), ydata.get("p2"), ydata.get("p3")
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
                p3, mapping, match_detail = auto_map_class_to_dept(p3, cc3, p1_depts)
                p3["科系"] = p3["_mapped_dept"]
                p3 = p3.drop(columns=["_mapped_dept"], errors="ignore")
    return p1, p2, p3, geo, ydata.get("channel_col")


# ============================================================
# 就緒檢查
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
    </div>""", unsafe_allow_html=True)
    st.stop()

valid_years = [yr for yr in st.session_state["years"]
               if st.session_state["years"][yr].get("p1") not in [None, "-- 未選擇 --"]]
if not valid_years:
    st.warning("⚠️ 沒有任何年度設定了一階資料。")
    st.stop()


# ============================================================
# 欄位診斷
# ============================================================
def show_field_diagnosis(p1, p2, p3, yr_label):
    with st.expander(f"🔍 {yr_label} 欄位偵測與映射診斷", expanded=False):
        diag = []
        for phase, df, label in [("一階", p1, "P1"), ("二階", p2, "P2"), ("最終", p3, "P3")]:
            if df is None:
                diag.append({"階段": phase, "科系欄位": "—", "班級欄位": "—", "學校欄位": "—", "筆數": 0})
                continue
            dc = detect_dept_col(df)
            cc = detect_class_col(df)
            sc = detect_school_col(df)
            diag.append({
                "階段": phase,
                "科系欄位": dc if dc else ("科系(映射)" if "科系" in df.columns else "❌"),
                "班級欄位": cc if cc else "—",
                "學校欄位": sc if sc else "❌",
                "筆數": len(df)
            })
        st.dataframe(pd.DataFrame(diag), use_container_width=True, hide_index=True)

        if p3 is not None and "科系" in p3.columns:
            mapped = p3["科系"].notna().sum()
            unmapped = p3["科系"].isna().sum()
            total = len(p3)
            pct = mapped / total * 100 if total > 0 else 0
            if unmapped == 0:
                st.markdown(f'<div class="success-box">✅ 映射完成：{mapped}/{total}（{pct:.1f}%）全部成功</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="warning-box">⚠️ 映射：{mapped}/{total}（{pct:.1f}%），{unmapped} 筆未映射</div>', unsafe_allow_html=True)
            mapped_depts = p3["科系"].dropna().unique()
            st.caption(f"映射到的科系（{len(mapped_depts)}個）：")
            for d in sorted(mapped_depts, key=str):
                cnt = (p3["科系"] == d).sum()
                st.caption(f"　📌 {d}：{cnt} 人")
            if unmapped > 0:
                cc3 = detect_class_col(p3)
                if cc3 and cc3 in p3.columns:
                    unmapped_classes = p3[p3["科系"].isna()][cc3].value_counts()
                    st.markdown("**未映射的班級名稱：**")
                    for cn, cnt in unmapped_classes.items():
                        parsed = parse_class_name(str(cn))
                        if parsed:
                            _, dept_kw, _, _ = parsed
                            st.caption(f"　❌ {cn}（{cnt}人）→ 關鍵字「{dept_kw}」找不到對應科系")
                        else:
                            st.caption(f"　❌ {cn}（{cnt}人）→ 無法解析格式")


# ============================================================
# 單年度七模組
# ============================================================
def render_year_analysis(yr):
    p1, p2, p3, geo, ch_col = get_year_dfs(yr)
    if p1 is None:
        st.warning(f"⚠️ {yr}：一階資料未指定或無法讀取。")
        return
    mod_opts = ["📊 總覽儀表板", "🔄 招生漏斗", "📈 入學管道",
                "🗺️ 地理分布", "🏫 科系熱力圖", "🎯 來源學校", "⚠️ 流失預警"]
    mod = st.radio("選擇分析模組：", mod_opts, horizontal=True, key=f"mod_{yr}")
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    n1 = len(p1)
    n2 = len(p2) if p2 is not None else None
    n3 = len(p3) if p3 is not None else None

    # ─── 1. 總覽 ───
    if "總覽" in mod:
        st.subheader(f"📊 {yr} — 總覽儀表板")
        cols = st.columns(5)
        with cols[0]:
            st.markdown(f'<div class="metric-card"><h3>一階報名</h3><h1>{n1:,}</h1></div>', unsafe_allow_html=True)
        with cols[1]:
            v = f"{n2:,}" if n2 else "—"
            st.markdown(f'<div class="metric-card metric-orange"><h3>二階報到</h3><h1>{v}</h1></div>', unsafe_allow_html=True)
        with cols[2]:
            v = f"{n3:,}" if n3 else "—"
            st.markdown(f'<div class="metric-card metric-green"><h3>最終入學</h3><h1>{v}</h1></div>', unsafe_allow_html=True)
        with cols[3]:
            r = f"{n2/n1*100:.1f}%" if n2 and n1 else "—"
            st.markdown(f'<div class="metric-card metric-blue"><h3>一→二階</h3><h1>{r}</h1></div>', unsafe_allow_html=True)
        with cols[4]:
            r = f"{n3/n1*100:.1f}%" if n3 and n1 else "—"
            st.markdown(f'<div class="metric-card metric-gold"><h3>一→最終</h3><h1>{r}</h1></div>', unsafe_allow_html=True)

        if p3 is not None and "科系" in p3.columns:
            n_mapped = p3["科系"].notna().sum()
            n_unmapped = p3["科系"].isna().sum()
            if n_unmapped == 0:
                st.markdown(f'<div class="success-box">✅ 班級→科系映射完成：{n_mapped}/{len(p3)} 筆全部成功</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="warning-box">⚠️ 班級→科系映射：{n_mapped}/{len(p3)} 筆成功，{n_unmapped} 筆未映射</div>', unsafe_allow_html=True)

        show_field_diagnosis(p1, p2, p3, yr)

        if p3 is not None and ch_col and ch_col in p3.columns:
            st.markdown("---")
            st.subheader("🟢 最終入學管道分布")
            cd = p3[ch_col].value_counts().reset_index()
            cd.columns = ["入學管道", "人數"]
            cd["佔比(%)"] = (cd["人數"] / cd["人數"].sum() * 100).round(1)
            c1, c2 = st.columns(2)
            with c1:
                fig = px.pie(cd, names="入學管道", values="人數", title="管道佔比", hole=.35)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = px.bar(cd.sort_values("人數", ascending=True),
                             x="人數", y="入學管道", orientation="h", text="人數",
                             title="各管道人數", color="佔比(%)", color_continuous_scale="Viridis")
                fig.update_layout(height=max(400, len(cd)*28))
                st.plotly_chart(fig, use_container_width=True)

        fl, fv = ["一階報名"], [n1]
        if n2: fl.append("二階報到"); fv.append(n2)
        if n3: fl.append("最終入學"); fv.append(n3)
        if len(fv) > 1:
            st.plotly_chart(fig_funnel(fl, fv, f"{yr} 招生漏斗"), use_container_width=True)

        c1, c2 = st.columns(2)
        dc = detect_dept_col(p1); sc = detect_school_col(p1)
        with c1:
            if dc:
                dd = p1[dc].value_counts().reset_index(); dd.columns = ["科系", "人數"]
                fig = px.pie(dd, names="科系", values="人數", title="一階科系分布", hole=.4)
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            if sc:
                sd = p1[sc].value_counts().head(10).reset_index(); sd.columns = ["學校", "人數"]
                fig = px.bar(sd, x="人數", y="學校", orientation="h", title="來源學校 TOP 10", text="人數")
                fig.update_traces(marker_color="#667eea")
                fig.update_layout(yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig, use_container_width=True)

        result = build_dept_stats(p1, p2, p3)
        if result is not None:
            ds, p3_info = result
            st.markdown("---")
            st.subheader("各科系三階段概覽")
            total_final_table = ds["最終入學"].sum()
            if n3 and total_final_table != n3:
                diff = n3 - total_final_table
                st.markdown(f'<div class="warning-box">⚠️ 科系加總（{total_final_table}）≠ 總人數（{n3}），差異 {diff} 人</div>', unsafe_allow_html=True)
            elif n3 and total_final_table == n3:
                st.markdown(f'<div class="success-box">✅ 科系加總（{total_final_table}）= 總人數（{n3}）完全吻合！</div>', unsafe_allow_html=True)
            if p3_info and p3_info.get("method") == "class_mapping":
                mapping = p3_info.get("mapping", {})
                match_detail = p3_info.get("match_detail", {})
                if mapping:
                    with st.expander("🔍 班級→科系映射明細", expanded=False):
                        rows = []
                        for cls in sorted(mapping.keys()):
                            rows.append({"班級名稱": cls, "映射科系": mapping[cls], "映射方法": match_detail.get(cls, "")})
                        for cls, detail in match_detail.items():
                            if cls not in mapping:
                                rows.append({"班級名稱": cls, "映射科系": "❌ 未映射", "映射方法": detail})
                        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            st.dataframe(ds.sort_values("一階人數", ascending=False), use_container_width=True, hide_index=True)

    # ─── 2. 漏斗 ───
    elif "漏斗" in mod:
        st.subheader(f"🔄 {yr} — 招生漏斗分析")
        result = build_dept_stats(p1, p2, p3)
        if result is not None:
            ds, _ = result
            st.dataframe(ds.sort_values("一→最終(%)", ascending=False), use_container_width=True, hide_index=True)
            rc = [c for c in ["一→二階(%)", "一→最終(%)"] if c in ds.columns]
            if rc:
                st.plotly_chart(fig_grouped_bar(ds.sort_values(rc[0], ascending=True), "科系", rc, "各科系轉換率"), use_container_width=True)
            st.subheader("單科系漏斗")
            sel = st.selectbox("選擇科系：", ds["科系"].tolist(), key=f"fun_dept_{yr}")
            row = ds[ds["科系"] == sel].iloc[0]
            fl, fv = ["一階報名"], [int(row["一階人數"])]
            if row["二階人數"] > 0 or p2 is not None: fl.append("二階報到"); fv.append(int(row["二階人數"]))
            if row["最終入學"] > 0 or p3 is not None: fl.append("最終入學"); fv.append(int(row["最終入學"]))
            st.plotly_chart(fig_funnel(fl, fv, f"{sel} 漏斗"), use_container_width=True)
        ss = build_school_stats(p1, p2, p3)
        if ss is not None:
            st.markdown("---"); st.subheader("各來源學校漏斗")
            mn = st.slider("一階≥", 1, 50, 5, key=f"fun_mn_{yr}")
            sf = ss[ss["一階人數"] >= mn].sort_values("一→最終(%)", ascending=False)
            st.dataframe(sf, use_container_width=True, hide_index=True)
            st.plotly_chart(fig_bar_h(sf.head(20), "學校", "一→最終(%)", f"來源學校轉換率 TOP 20（一階≥{mn}）", color="#4CAF50"), use_container_width=True)

    # ─── 3. 管道 ───
    elif "管道" in mod:
        st.subheader(f"📈 {yr} — 入學管道分析")
        if p3 is None or not ch_col or ch_col not in (p3.columns if p3 is not None else []):
            st.warning("⚠️ 需要最終入學資料及入學方式欄位。"); return
        cd = p3[ch_col].value_counts().reset_index(); cd.columns = ["入學管道", "人數"]
        cd["佔比(%)"] = (cd["人數"] / cd["人數"].sum() * 100).round(1)
        cd["累積(%)"] = cd["佔比(%)"].cumsum().round(1)
        c1, c2 = st.columns(2)
        with c1:
            fig = px.pie(cd, names="入學管道", values="人數", title="管道佔比", hole=.35)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = px.bar(cd.sort_values("人數", ascending=True), x="人數", y="入學管道", orientation="h", text="人數", title="人數排行", color="佔比(%)", color_continuous_scale="Viridis")
            fig.update_layout(height=max(400, len(cd)*30))
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(cd, use_container_width=True, hide_index=True)
        dept_col = detect_dept_col(p3)
        if dept_col is None and "科系" in p3.columns: dept_col = "科系"
        if dept_col:
            st.markdown("---"); st.subheader("管道 × 科系")
            valid_p3 = p3[p3[dept_col].notna()] if dept_col == "科系" else p3
            cross = valid_p3.groupby([ch_col, dept_col]).size().reset_index(name="人數")
            fig = px.bar(cross, x=ch_col, y="人數", color=dept_col, barmode="stack", text="人數", title="管道×科系堆疊圖")
            fig.update_layout(height=600, xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
            st.plotly_chart(fig_heatmap(cross, dept_col, ch_col, "人數", "管道×科系熱力圖"), use_container_width=True)

    # ─── 4. 地理 ───
    elif "地理" in mod:
        st.subheader(f"🗺️ {yr} — 地理分布")
        if geo is None:
            st.warning("⚠️ 一階資料無經緯度欄位。"); return
        def do_map(src_df, count_label, title_text, phase):
            sc_ = detect_school_col(src_df)
            if sc_ is None: return
            agg = src_df.groupby(sc_).size().reset_index(name=count_label)
            agg["_std"] = agg[sc_].apply(norm_school)
            agg = agg.merge(geo, on="_std", how="left").drop(columns=["_std"])
            fig = fig_map(agg, count_label, title_text)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
                ok = agg["lat"].notna().sum(); total = len(agg)
                st.caption(f"匹配：{ok}/{total} 校（{ok/total*100:.1f}%）" if total else "")
            else:
                st.info(f"{phase} 無匹配結果。")
        do_map(p1, "報名人數", f"{yr} 一階報名來源", "一階")
        if p2 is not None: st.markdown("---"); do_map(p2, "報到人數", f"{yr} 二階報到來源", "二階")
        if p3 is not None: st.markdown("---"); do_map(p3, "入學人數", f"{yr} 最終入學來源", "最終入學")

    # ─── 5. 熱力圖 ───
    elif "熱力圖" in mod:
        st.subheader(f"🏫 {yr} — 科系×學校 熱力圖")
        dc = detect_dept_col(p1); sc = detect_school_col(p1)
        if not dc or not sc: st.warning("⚠️ 未偵測到科系或學校欄位。"); return
        mn = st.slider("學校報名≥", 1, 30, 3, key=f"hm_{yr}")
        valid = p1[sc].value_counts(); valid = valid[valid >= mn].index
        filt = p1[p1[sc].isin(valid)]
        cr = filt.groupby([sc, dc]).size().reset_index(name="人數")
        st.plotly_chart(fig_heatmap(cr, dc, sc, "人數", f"一階報名（≥{mn}人）"), use_container_width=True)
        if p3 is not None:
            dept_col_p3 = detect_dept_col(p3)
            if dept_col_p3 is None and "科系" in p3.columns: dept_col_p3 = "科系"
            sc3 = detect_school_col(p3)
            if dept_col_p3 and sc3:
                valid_p3 = p3[p3[dept_col_p3].notna()] if dept_col_p3 == "科系" else p3
                cr3 = valid_p3.groupby([sc3, dept_col_p3]).size().reset_index(name="入學人數")
                cr3 = cr3[cr3[sc3].isin(valid)]
                if not cr3.empty:
                    st.plotly_chart(fig_heatmap(cr3, dept_col_p3, sc3, "入學人數", "最終入學"), use_container_width=True)

    # ─── 6. 來源學校 ───
    elif "來源學校" in mod:
        st.subheader(f"🎯 {yr} — 來源學校追蹤")
        ss = build_school_stats(p1, p2, p3)
        if ss is None: st.warning("⚠️ 未偵測到學校欄位。"); return
        def tier(n):
            if n >= 30: return "Tier1(≥30)"
            elif n >= 10: return "Tier2(10-29)"
            else: return "Tier3(<10)"
        ss["分級"] = ss["一階人數"].apply(tier)
        sel_t = st.multiselect("篩選分級：", ["Tier1(≥30)", "Tier2(10-29)", "Tier3(<10)"],
                                default=["Tier1(≥30)", "Tier2(10-29)"], key=f"tier_{yr}")
        disp = ss[ss["分級"].isin(sel_t)].sort_values("一階人數", ascending=False)
        st.dataframe(disp, use_container_width=True, hide_index=True)
        st.subheader("個別學校")
        sel = st.selectbox("選擇學校：", ss.sort_values("一階人數", ascending=False)["學校"], key=f"sch_{yr}")
        if sel:
            r = ss[ss["學校"] == sel].iloc[0]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("一階", f'{int(r["一階人數"])}'); c2.metric("二階", f'{int(r["二階人數"])}')
            c3.metric("最終", f'{int(r["最終入學"])}'); c4.metric("轉換率", f'{r["一→最終(%)"]}%')
            fl, fv = ["一階"], [int(r["一階人數"])]
            if r["二階人數"] > 0 or p2 is not None: fl.append("二階"); fv.append(int(r["二階人數"]))
            if r["最終入學"] > 0 or p3 is not None: fl.append("最終"); fv.append(int(r["最終入學"]))
            if len(fv) > 1: st.plotly_chart(fig_funnel(fl, fv, f"{sel} 漏斗"), use_container_width=True)

    # ─── 7. 流失預警 ───
    elif "流失" in mod:
        st.subheader(f"⚠️ {yr} — 流失預警分析")
        if p2 is None and p3 is None: st.warning("⚠️ 需要至少二階或最終入學資料。"); return
        ss = build_school_stats(p1, p2, p3)
        if ss is None: st.warning("⚠️ 未偵測到學校欄位。"); return
        has_final = p3 is not None and ss["最終入學"].sum() > 0
        rc = "一→最終(%)" if has_final else "一→二階(%)"
        ll = "最終入學" if has_final else "二階人數"
        ss["流失人數"] = ss["一階人數"] - ss[ll]
        mn = st.slider("一階≥", 1, 50, 10, key=f"loss_mn_{yr}")
        pool = ss[ss["一階人數"] >= mn]
        avg = pool[rc].mean()
        warn = pool[pool[rc] < avg].sort_values("流失人數", ascending=False)
        if warn.empty:
            st.success("✅ 沒有預警學校！")
        else:
            st.markdown(f'<div class="warning-box">⚠️ {len(warn)}所學校低於平均 {avg:.1f}%</div>', unsafe_allow_html=True)
            st.dataframe(warn, use_container_width=True, hide_index=True)
        st.subheader("IPA 矩陣")
        ana = pool.copy()
        if not ana.empty:
            med_x = ana["一階人數"].median(); med_y = ana[rc].median()
            fig = px.scatter(ana, x="一階人數", y=rc, size="流失人數", hover_name="學校",
                             title=f"IPA 矩陣", size_max=40, color=rc, color_continuous_scale="RdYlGn")
            fig.add_hline(y=med_y, line_dash="dash", line_color="red", annotation_text=f"中位數 {med_y:.1f}%")
            fig.add_vline(x=med_x, line_dash="dash", line_color="blue", annotation_text=f"中位數 {med_x:.0f}")
            fig.update_layout(height=600); st.plotly_chart(fig, use_container_width=True)
        st.markdown("---"); st.subheader("科系流失")
        result = build_dept_stats(p1, p2, p3)
        if result is not None:
            ds, _ = result
            has_final_d = p3 is not None and ds["最終入學"].sum() > 0
            drc = "一→最終(%)" if has_final_d else "一→二階(%)"
            dll = "最終入學" if has_final_d else "二階人數"
            ds["流失人數"] = ds["一階人數"] - ds[dll]
            fig = px.scatter(ds, x="一階人數", y=drc, size="流失人數", hover_name="科系", text="科系",
                             title=f"科系 IPA 矩陣", size_max=50, color=drc, color_continuous_scale="RdYlGn")
            fig.update_traces(textposition="top center"); fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(ds.sort_values("流失人數", ascending=False), use_container_width=True, hide_index=True)


# ============================================================
# 跨年度七模組完整分析
# ============================================================
def render_cross_year():
    st.header("📊 跨年度比較分析")
    if len(valid_years) < 2:
        st.warning("⚠️ 需要至少 2 個年度。"); return

    mod_opts = ["📊 總覽儀表板", "🔄 招生漏斗", "📈 入學管道",
                "🗺️ 地理分布", "🏫 科系熱力圖", "🎯 來源學校", "⚠️ 流失預警"]
    mod = st.radio("選擇跨年度分析模組：", mod_opts, horizontal=True, key="cross_mod")
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # 預載所有年度資料
    year_cache = {}
    for yr in valid_years:
        p1, p2, p3, geo, ch_col = get_year_dfs(yr)
        if p1 is None:
            continue
        n1 = len(p1)
        n2 = len(p2) if p2 is not None else 0
        n3 = len(p3) if p3 is not None else 0
        result = build_dept_stats(p1, p2, p3)
        ds = result[0] if result else None
        ss = build_school_stats(p1, p2, p3)
        year_cache[yr] = {
            "p1": p1, "p2": p2, "p3": p3, "geo": geo, "ch_col": ch_col,
            "n1": n1, "n2": n2, "n3": n3, "ds": ds, "ss": ss
        }

    if not year_cache:
        st.warning("⚠️ 沒有有效的年度資料。"); return

    # ═══════════════════════════════════════════════
    # 跨年度 1：總覽儀表板
    # ═══════════════════════════════════════════════
    if "總覽" in mod:
        st.subheader("📊 跨年度總覽比較")

        # KPI 摘要表
        summaries = []
        for yr, c in year_cache.items():
            sc = detect_school_col(c["p1"])
            dc = detect_dept_col(c["p1"])
            summaries.append({
                "年度": yr,
                "一階人數": c["n1"], "二階人數": c["n2"], "最終入學": c["n3"],
                "一→二階(%)": round(c["n2"]/c["n1"]*100, 1) if c["n1"] and c["n2"] else 0,
                "一→最終(%)": round(c["n3"]/c["n1"]*100, 1) if c["n1"] and c["n3"] else 0,
                "來源學校數": c["p1"][sc].nunique() if sc else 0,
                "科系數": c["p1"][dc].nunique() if dc else 0
            })
        sdf = pd.DataFrame(summaries)
        st.dataframe(sdf, use_container_width=True, hide_index=True)

        # YoY 趨勢卡片
        if len(summaries) >= 2:
            latest = summaries[-1]; prev = summaries[-2]
            cols = st.columns(4)
            with cols[0]:
                st.markdown(f"**一階人數** {latest['一階人數']:,}", unsafe_allow_html=True)
                st.markdown(trend_arrow(latest['一階人數'], prev['一階人數']), unsafe_allow_html=True)
            with cols[1]:
                st.markdown(f"**最終入學** {latest['最終入學']:,}", unsafe_allow_html=True)
                st.markdown(trend_arrow(latest['最終入學'], prev['最終入學']), unsafe_allow_html=True)
            with cols[2]:
                st.markdown(f"**一→最終(%)** {latest['一→最終(%)']}", unsafe_allow_html=True)
                st.markdown(trend_arrow(latest['一→最終(%)'], prev['一→最終(%)']), unsafe_allow_html=True)
            with cols[3]:
                st.markdown(f"**來源學校數** {latest['來源學校數']}", unsafe_allow_html=True)
                st.markdown(trend_arrow(latest['來源學校數'], prev['來源學校數']), unsafe_allow_html=True)

        # 招生量柱狀圖
        fig = go.Figure()
        fig.add_trace(go.Bar(x=sdf["年度"], y=sdf["一階人數"], name="一階", marker_color="#2196F3"))
        if sdf["二階人數"].sum() > 0:
            fig.add_trace(go.Bar(x=sdf["年度"], y=sdf["二階人數"], name="二階", marker_color="#FF9800"))
        if sdf["最終入學"].sum() > 0:
            fig.add_trace(go.Bar(x=sdf["年度"], y=sdf["最終入學"], name="最終入學", marker_color="#4CAF50"))
        fig.update_layout(barmode="group", title="各年度招生量", height=450)
        st.plotly_chart(fig, use_container_width=True)

        # 轉換率趨勢線
        fig = go.Figure()
        if sdf["一→二階(%)"].sum() > 0:
            fig.add_trace(go.Scatter(x=sdf["年度"], y=sdf["一→二階(%)"], name="一→二階",
                mode="lines+markers+text", text=sdf["一→二階(%)"], textposition="top center",
                line=dict(width=3, color="#FF9800"), marker=dict(size=10)))
        if sdf["一→最終(%)"].sum() > 0:
            fig.add_trace(go.Scatter(x=sdf["年度"], y=sdf["一→最終(%)"], name="一→最終",
                mode="lines+markers+text", text=sdf["一→最終(%)"], textposition="top center",
                line=dict(width=3, color="#4CAF50"), marker=dict(size=12)))
        fig.update_layout(title="轉換率趨勢", yaxis_title="轉換率(%)", height=400)
        st.plotly_chart(fig, use_container_width=True)

        # 科系跨年度彙總
        st.markdown("---"); st.subheader("科系跨年度趨勢")
        all_depts = set()
        for yr, c in year_cache.items():
            if c["ds"] is not None:
                all_depts.update(c["ds"]["科系"].tolist())
        if all_depts:
            sel_dept = st.selectbox("選擇科系：", sorted(all_depts), key="cross_ov_dept")
            sel_key = norm_dept(sel_dept)
            rows = []
            for yr, c in year_cache.items():
                if c["ds"] is not None:
                    r = c["ds"][c["ds"]["科系"].apply(norm_dept) == sel_key]
                    if not r.empty:
                        r = r.iloc[0]
                        rows.append({"年度": yr, "一階": int(r["一階人數"]), "二階": int(r["二階人數"]),
                                     "最終": int(r["最終入學"]), "一→最終(%)": r["一→最終(%)"]})
            if rows:
                rdf = pd.DataFrame(rows)
                st.dataframe(rdf, use_container_width=True, hide_index=True)
                fig = go.Figure()
                fig.add_trace(go.Bar(x=rdf["年度"], y=rdf["一階"], name="一階", marker_color="#2196F3"))
                fig.add_trace(go.Bar(x=rdf["年度"], y=rdf["最終"], name="最終", marker_color="#4CAF50"))
                fig.add_trace(go.Scatter(x=rdf["年度"], y=rdf["一→最終(%)"], name="轉換率%",
                    yaxis="y2", mode="lines+markers+text", text=rdf["一→最終(%)"],
                    textposition="top center", line=dict(width=3, color="#E91E63")))
                fig.update_layout(barmode="group", title=f"「{sel_dept}」跨年度趨勢",
                    yaxis2=dict(title="轉換率(%)", overlaying="y", side="right", range=[0, 100]),
                    height=450)
                st.plotly_chart(fig, use_container_width=True)

    # ═══════════════════════════════════════════════
    # 跨年度 2：招生漏斗
    # ═══════════════════════════════════════════════
    elif "漏斗" in mod:
        st.subheader("🔄 跨年度招生漏斗比較")

        # 全校漏斗比較
        st.markdown("#### 全校三階段漏斗")
        fig = make_subplots(rows=1, cols=len(year_cache),
                            subplot_titles=list(year_cache.keys()),
                            specs=[[{"type": "funnel"}] * len(year_cache)])
        for i, (yr, c) in enumerate(year_cache.items()):
            fl, fv = ["一階"], [c["n1"]]
            if c["n2"]: fl.append("二階"); fv.append(c["n2"])
            if c["n3"]: fl.append("最終"); fv.append(c["n3"])
            fig.add_trace(go.Funnel(y=fl, x=fv, name=yr, textinfo="value+percent initial"), row=1, col=i+1)
        fig.update_layout(height=450, showlegend=True)
        st.plotly_chart(fig, use_container_width=True)

        # 科系轉換率跨年度比較
        st.markdown("---"); st.markdown("#### 各科系轉換率跨年度比較")
        all_depts = set()
        for yr, c in year_cache.items():
            if c["ds"] is not None:
                all_depts.update(c["ds"]["科系"].tolist())
        if all_depts:
            cross_dept_rows = []
            for yr, c in year_cache.items():
                if c["ds"] is not None:
                    for _, row in c["ds"].iterrows():
                        cross_dept_rows.append({"年度": yr, "科系": row["科系"],
                            "一階人數": int(row["一階人數"]), "最終入學": int(row["最終入學"]),
                            "一→最終(%)": row["一→最終(%)"]})
            if cross_dept_rows:
                cdf = pd.DataFrame(cross_dept_rows)
                fig = px.bar(cdf, x="科系", y="一→最終(%)", color="年度", barmode="group",
                             text="一→最終(%)", title="各科系轉換率（跨年度）")
                fig.update_layout(height=500, xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)

                # 轉換率變化表
                pivot = cdf.pivot_table(index="科系", columns="年度", values="一→最終(%)", aggfunc="first")
                st.dataframe(pivot.round(1), use_container_width=True)

        # 學校漏斗跨年度
        st.markdown("---"); st.markdown("#### 來源學校轉換率跨年度")
        mn = st.slider("一階≥", 1, 50, 10, key="cross_fun_mn")
        sch_rows = []
        for yr, c in year_cache.items():
            if c["ss"] is not None:
                for _, row in c["ss"].iterrows():
                    if row["一階人數"] >= mn:
                        sch_rows.append({"年度": yr, "學校": row["學校"],
                            "一階人數": int(row["一階人數"]), "最終入學": int(row["最終入學"]),
                            "一→最終(%)": row["一→最終(%)"]})
        if sch_rows:
            sdf = pd.DataFrame(sch_rows)
            sel_sch = st.selectbox("選擇學校：",
                sorted(sdf["學校"].unique()), key="cross_fun_sch")
            sch_data = sdf[sdf["學校"] == sel_sch]
            if not sch_data.empty:
                st.dataframe(sch_data, use_container_width=True, hide_index=True)
                fig = go.Figure()
                fig.add_trace(go.Bar(x=sch_data["年度"], y=sch_data["一階人數"], name="一階", marker_color="#2196F3"))
                fig.add_trace(go.Bar(x=sch_data["年度"], y=sch_data["最終入學"], name="最終", marker_color="#4CAF50"))
                fig.update_layout(barmode="group", title=f"「{sel_sch}」跨年度漏斗", height=400)
                st.plotly_chart(fig, use_container_width=True)

    # ═══════════════════════════════════════════════
    # 跨年度 3：入學管道
    # ═══════════════════════════════════════════════
    elif "管道" in mod:
        st.subheader("📈 跨年度入學管道比較")

        ch_data_all = []
        for yr, c in year_cache.items():
            if c["p3"] is not None and c["ch_col"] and c["ch_col"] in c["p3"].columns:
                cd = c["p3"][c["ch_col"]].value_counts().reset_index()
                cd.columns = ["入學管道", "人數"]
                cd["年度"] = yr
                cd["佔比(%)"] = (cd["人數"] / cd["人數"].sum() * 100).round(1)
                ch_data_all.append(cd)

        if not ch_data_all:
            st.warning("⚠️ 沒有年度有入學管道資料。"); return

        all_ch = pd.concat(ch_data_all, ignore_index=True)

        # 管道人數跨年度
        fig = px.bar(all_ch, x="入學管道", y="人數", color="年度", barmode="group",
                     text="人數", title="各管道人數（跨年度）")
        fig.update_layout(height=500, xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

        # 管道佔比跨年度
        fig = px.bar(all_ch, x="年度", y="佔比(%)", color="入學管道", barmode="stack",
                     text="佔比(%)", title="管道結構比例（跨年度）")
        fig.update_layout(height=500)
        st.plotly_chart(fig, use_container_width=True)

        # 管道佔比變化表
        pivot = all_ch.pivot_table(index="入學管道", columns="年度", values="人數", aggfunc="sum").fillna(0).astype(int)
        st.subheader("管道人數交叉表")
        st.dataframe(pivot, use_container_width=True)

        # 單一管道趨勢
        st.markdown("---"); st.subheader("單一管道趨勢追蹤")
        all_channels = sorted(all_ch["入學管道"].unique())
        sel_ch = st.selectbox("選擇管道：", all_channels, key="cross_ch_sel")
        ch_trend = all_ch[all_ch["入學管道"] == sel_ch]
        if not ch_trend.empty:
            fig = go.Figure()
            fig.add_trace(go.Bar(x=ch_trend["年度"], y=ch_trend["人數"], name="人數", marker_color="#4CAF50"))
            fig.add_trace(go.Scatter(x=ch_trend["年度"], y=ch_trend["佔比(%)"], name="佔比%",
                yaxis="y2", mode="lines+markers+text", text=ch_trend["佔比(%)"],
                textposition="top center", line=dict(width=3, color="#E91E63")))
            fig.update_layout(title=f"「{sel_ch}」跨年度趨勢",
                yaxis2=dict(title="佔比(%)", overlaying="y", side="right", range=[0, 100]), height=400)
            st.plotly_chart(fig, use_container_width=True)

        # 管道×科系跨年度
        st.markdown("---"); st.subheader("管道×科系 跨年度")
        ch_dept_rows = []
        for yr, c in year_cache.items():
            if c["p3"] is not None and c["ch_col"] and c["ch_col"] in c["p3"].columns:
                dept_col = detect_dept_col(c["p3"])
                if dept_col is None and "科系" in c["p3"].columns:
                    dept_col = "科系"
                if dept_col:
                    valid_p3 = c["p3"][c["p3"][dept_col].notna()] if dept_col == "科系" else c["p3"]
                    cross = valid_p3.groupby([c["ch_col"], dept_col]).size().reset_index(name="人數")
                    cross["年度"] = yr
                    cross.columns = ["入學管道", "科系", "人數", "年度"]
                    ch_dept_rows.append(cross)
        if ch_dept_rows:
            all_cd = pd.concat(ch_dept_rows, ignore_index=True)
            sel_dept_ch = st.selectbox("選擇科系查看管道變化：",
                sorted(all_cd["科系"].unique()), key="cross_ch_dept")
            dept_ch_data = all_cd[all_cd["科系"] == sel_dept_ch]
            if not dept_ch_data.empty:
                fig = px.bar(dept_ch_data, x="入學管道", y="人數", color="年度", barmode="group",
                             text="人數", title=f"「{sel_dept_ch}」各管道跨年度比較")
                fig.update_layout(height=450, xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)

    # ═══════════════════════════════════════════════
    # 跨年度 4：地理分布
    # ═══════════════════════════════════════════════
    elif "地理" in mod:
        st.subheader("🗺️ 跨年度地理分布比較")

        # 合併所有年度的地理資料
        phase_select = st.radio("選擇階段：", ["一階報名", "最終入學"], horizontal=True, key="cross_geo_phase")

        for yr, c in year_cache.items():
            if c["geo"] is None:
                continue
            st.markdown(f"### {yr}")
            if "一階" in phase_select:
                src = c["p1"]
                label = "報名人數"
            else:
                src = c["p3"]
                label = "入學人數"
            if src is None:
                st.info(f"{yr}：無此階段資料"); continue
            sc = detect_school_col(src)
            if sc is None:
                continue
            agg = src.groupby(sc).size().reset_index(name=label)
            agg["_std"] = agg[sc].apply(norm_school)
            agg = agg.merge(c["geo"], on="_std", how="left").drop(columns=["_std"])
            fig = fig_map(agg, label, f"{yr} {phase_select}")
            if fig:
                st.plotly_chart(fig, use_container_width=True)

        # 來源學校地理覆蓋比較
        st.markdown("---"); st.subheader("來源學校數量比較")
        geo_summary = []
        for yr, c in year_cache.items():
            sc = detect_school_col(c["p1"])
            if sc:
                geo_summary.append({
                    "年度": yr,
                    "來源學校數": c["p1"][sc].nunique(),
                    "一階人數": c["n1"],
                    "最終入學": c["n3"]
                })
        if geo_summary:
            gdf = pd.DataFrame(geo_summary)
            fig = go.Figure()
            fig.add_trace(go.Bar(x=gdf["年度"], y=gdf["來源學校數"], name="來源學校數", marker_color="#9C27B0"))
            fig.update_layout(title="來源學校數趨勢", height=400)
            st.plotly_chart(fig, use_container_width=True)

    # ═══════════════════════════════════════════════
    # 跨年度 5：科系熱力圖
    # ═══════════════════════════════════════════════
    elif "熱力圖" in mod:
        st.subheader("🏫 跨年度科系×學校 熱力圖")

        # 科系×年度 招生量熱力圖
        st.markdown("#### 科系×年度 招生量")
        dept_year_rows = []
        for yr, c in year_cache.items():
            if c["ds"] is not None:
                for _, row in c["ds"].iterrows():
                    dept_year_rows.append({"年度": yr, "科系": row["科系"],
                        "一階人數": int(row["一階人數"]), "最終入學": int(row["最終入學"]),
                        "一→最終(%)": row["一→最終(%)"]})
        if dept_year_rows:
            dydf = pd.DataFrame(dept_year_rows)
            metric = st.radio("指標：", ["一階人數", "最終入學", "一→最終(%)"], horizontal=True, key="cross_hm_metric")
            pv = dydf.pivot_table(index="科系", columns="年度", values=metric, aggfunc="first").fillna(0)
            colorscale = "YlOrRd" if "人數" in metric else "RdYlGn"
            fig = px.imshow(pv, text_auto=True, aspect="auto",
                            color_continuous_scale=colorscale, title=f"科系×年度：{metric}")
            fig.update_layout(height=max(400, len(pv) * 30))
            st.plotly_chart(fig, use_container_width=True)

        # 科系×學校 跨年度
        st.markdown("---"); st.markdown("#### 特定科系×學校 跨年度")
        all_depts = set()
        for yr, c in year_cache.items():
            dc = detect_dept_col(c["p1"])
            if dc:
                all_depts.update(c["p1"][dc].unique())
        if all_depts:
            sel = st.selectbox("選擇科系：", sorted(all_depts), key="cross_hm_dept")
            mn = st.slider("學校報名≥", 1, 20, 3, key="cross_hm_mn")
            for yr, c in year_cache.items():
                dc = detect_dept_col(c["p1"]); sc = detect_school_col(c["p1"])
                if dc and sc:
                    filt = c["p1"][c["p1"][dc] == sel]
                    if not filt.empty:
                        sch_cnt = filt[sc].value_counts()
                        sch_cnt = sch_cnt[sch_cnt >= mn]
                        if not sch_cnt.empty:
                            st.markdown(f"**{yr}** — {sel}（≥{mn}人學校）")
                            cdf = sch_cnt.reset_index(); cdf.columns = ["學校", "人數"]
                            fig = px.bar(cdf.sort_values("人數", ascending=True),
                                         x="人數", y="學校", orientation="h", text="人數",
                                         title=f"{yr} {sel}", color="人數", color_continuous_scale="Blues")
                            fig.update_layout(height=max(300, len(cdf)*25))
                            st.plotly_chart(fig, use_container_width=True)

    # ═══════════════════════════════════════════════
    # 跨年度 6：來源學校
    # ═══════════════════════════════════════════════
    elif "來源學校" in mod:
        st.subheader("🎯 跨年度來源學校追蹤")

        # 學校忠誠度分析
        st.markdown("#### 學校出現頻率（跨年度忠誠度）")
        all_schools = {}
        for yr, c in year_cache.items():
            if c["ss"] is not None:
                for _, row in c["ss"].iterrows():
                    sch = norm_school(row["學校"])
                    if sch not in all_schools:
                        all_schools[sch] = {"學校": row["學校"], "出現年度數": 0, "年度明細": []}
                    all_schools[sch]["出現年度數"] += 1
                    all_schools[sch]["年度明細"].append(yr)

        if all_schools:
            loyalty = pd.DataFrame(all_schools.values())
            loyalty["年度明細"] = loyalty["年度明細"].apply(lambda x: ", ".join(x))
            loyalty = loyalty.sort_values("出現年度數", ascending=False)

            n_total = len(valid_years)
            loyal = loyalty[loyalty["出現年度數"] == n_total]
            partial = loyalty[(loyalty["出現年度數"] > 1) & (loyalty["出現年度數"] < n_total)]
            single = loyalty[loyalty["出現年度數"] == 1]

            c1, c2, c3 = st.columns(3)
            c1.metric("🏆 全年度都有", f"{len(loyal)} 校")
            c2.metric("📊 部分年度", f"{len(partial)} 校")
            c3.metric("🆕 僅出現一次", f"{len(single)} 校")

            fig = px.histogram(loyalty, x="出現年度數", nbins=n_total,
                               title="學校出現頻率分布", labels={"出現年度數": "出現年度數"})
            fig.update_layout(height=350)
            st.plotly_chart(fig, use_container_width=True)

            with st.expander("🏆 全年度忠誠學校", expanded=False):
                st.dataframe(loyal, use_container_width=True, hide_index=True)

        # 新增/流失學校追蹤
        st.markdown("---"); st.markdown("#### 新增 & 流失學校追蹤")
        yr_list = list(year_cache.keys())
        if len(yr_list) >= 2:
            for i in range(1, len(yr_list)):
                prev_yr = yr_list[i-1]; curr_yr = yr_list[i]
                prev_ss = year_cache[prev_yr].get("ss")
                curr_ss = year_cache[curr_yr].get("ss")
                if prev_ss is not None and curr_ss is not None:
                    prev_set = set(prev_ss["學校"].apply(norm_school))
                    curr_set = set(curr_ss["學校"].apply(norm_school))
                    new_schools = curr_set - prev_set
                    lost_schools = prev_set - curr_set
                    st.markdown(f"**{prev_yr} → {curr_yr}：** 🆕 新增 {len(new_schools)} 校 ｜ ❌ 流失 {len(lost_schools)} 校")
                    if new_schools:
                        with st.expander(f"🆕 {curr_yr} 新增學校 ({len(new_schools)})", expanded=False):
                            new_data = curr_ss[curr_ss["學校"].apply(norm_school).isin(new_schools)]
                            st.dataframe(new_data.sort_values("一階人數", ascending=False), use_container_width=True, hide_index=True)
                    if lost_schools:
                        with st.expander(f"❌ {curr_yr} 流失學校 ({len(lost_schools)})", expanded=False):
                            lost_data = prev_ss[prev_ss["學校"].apply(norm_school).isin(lost_schools)]
                            st.dataframe(lost_data.sort_values("一階人數", ascending=False), use_container_width=True, hide_index=True)

        # 個別學校跨年度追蹤
        st.markdown("---"); st.markdown("#### 個別學校跨年度追蹤")
        all_sch_names = set()
        for yr, c in year_cache.items():
            if c["ss"] is not None:
                all_sch_names.update(c["ss"]["學校"].tolist())
        if all_sch_names:
            sel_sch = st.selectbox("選擇學校：", sorted(all_sch_names), key="cross_sch_sel")
            sch_rows = []
            for yr, c in year_cache.items():
                if c["ss"] is not None:
                    sel_key = norm_school(sel_sch)
                    r = c["ss"][c["ss"]["學校"].apply(norm_school) == sel_key]
                    if not r.empty:
                        r = r.iloc[0]
                        sch_rows.append({"年度": yr, "一階": int(r["一階人數"]), "二階": int(r["二階人數"]),
                                         "最終": int(r["最終入學"]), "一→最終(%)": r["一→最終(%)"]})
            if sch_rows:
                rdf = pd.DataFrame(sch_rows)
                st.dataframe(rdf, use_container_width=True, hide_index=True)
                fig = go.Figure()
                fig.add_trace(go.Bar(x=rdf["年度"], y=rdf["一階"], name="一階", marker_color="#2196F3"))
                fig.add_trace(go.Bar(x=rdf["年度"], y=rdf["最終"], name="最終", marker_color="#4CAF50"))
                fig.add_trace(go.Scatter(x=rdf["年度"], y=rdf["一→最終(%)"], name="轉換率%",
                    yaxis="y2", mode="lines+markers+text", text=rdf["一→最終(%)"],
                    textposition="top center", line=dict(width=3, color="#E91E63")))
                fig.update_layout(barmode="group", title=f"「{sel_sch}」跨年度趨勢",
                    yaxis2=dict(title="轉換率(%)", overlaying="y", side="right", range=[0, 100]), height=450)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info(f"「{sel_sch}」在選取的年度中無資料。")

    # ═══════════════════════════════════════════════
    # 跨年度 7：流失預警
    # ═══════════════════════════════════════════════
    elif "流失" in mod:
        st.subheader("⚠️ 跨年度流失預警分析")

        # 全校流失趨勢
        st.markdown("#### 全校流失趨勢")
        loss_summary = []
        for yr, c in year_cache.items():
            loss = c["n1"] - c["n3"] if c["n3"] else c["n1"] - c["n2"]
            rate = (c["n3"]/c["n1"]*100) if c["n1"] and c["n3"] else ((c["n2"]/c["n1"]*100) if c["n1"] and c["n2"] else 0)
            loss_summary.append({"年度": yr, "一階": c["n1"], "最終入學": c["n3"],
                                 "流失人數": loss, "轉換率(%)": round(rate, 1),
                                 "流失率(%)": round(100-rate, 1)})
        lsdf = pd.DataFrame(loss_summary)
        st.dataframe(lsdf, use_container_width=True, hide_index=True)

        fig = go.Figure()
        fig.add_trace(go.Bar(x=lsdf["年度"], y=lsdf["流失人數"], name="流失人數", marker_color="#f44336"))
        fig.add_trace(go.Scatter(x=lsdf["年度"], y=lsdf["流失率(%)"], name="流失率%",
            yaxis="y2", mode="lines+markers+text", text=lsdf["流失率(%)"],
            textposition="top center", line=dict(width=3, color="#FF9800")))
        fig.update_layout(title="全校流失趨勢",
            yaxis2=dict(title="流失率(%)", overlaying="y", side="right", range=[0, 100]), height=450)
        st.plotly_chart(fig, use_container_width=True)

        # 科系流失跨年度
        st.markdown("---"); st.markdown("#### 科系流失跨年度比較")
        dept_loss_rows = []
        for yr, c in year_cache.items():
            if c["ds"] is not None:
                for _, row in c["ds"].iterrows():
                    dept_loss_rows.append({
                        "年度": yr, "科系": row["科系"],
                        "一階": int(row["一階人數"]), "最終": int(row["最終入學"]),
                        "流失": int(row["流失人數"]),
                        "一→最終(%)": row["一→最終(%)"]
                    })
        if dept_loss_rows:
            dldf = pd.DataFrame(dept_loss_rows)

            # 流失人數熱力圖
            pv_loss = dldf.pivot_table(index="科系", columns="年度", values="流失", aggfunc="first").fillna(0)
            fig = px.imshow(pv_loss, text_auto=True, aspect="auto",
                            color_continuous_scale="OrRd", title="科系×年度：流失人數")
            fig.update_layout(height=max(400, len(pv_loss)*30))
            st.plotly_chart(fig, use_container_width=True)

            # 轉換率熱力圖
            pv_rate = dldf.pivot_table(index="科系", columns="年度", values="一→最終(%)", aggfunc="first").fillna(0)
            fig = px.imshow(pv_rate, text_auto=True, aspect="auto",
                            color_continuous_scale="RdYlGn", title="科系×年度：轉換率(%)")
            fig.update_layout(height=max(400, len(pv_rate)*30))
            st.plotly_chart(fig, use_container_width=True)

        # 惡化偵測
        st.markdown("---"); st.markdown("#### 🚨 惡化偵測（轉換率下降的科系）")
        yr_list = list(year_cache.keys())
        if len(yr_list) >= 2 and dept_loss_rows:
            dldf_full = pd.DataFrame(dept_loss_rows)
            deteriorating = []
            for dept in dldf_full["科系"].unique():
                dept_data = dldf_full[dldf_full["科系"] == dept].sort_values("年度")
                if len(dept_data) >= 2:
                    latest = dept_data.iloc[-1]["一→最終(%)"]
                    prev = dept_data.iloc[-2]["一→最終(%)"]
                    if prev > 0 and latest < prev:
                        drop = prev - latest
                        deteriorating.append({
                            "科系": dept,
                            f"{yr_list[-2]} 轉換率(%)": prev,
                            f"{yr_list[-1]} 轉換率(%)": latest,
                            "下降幅度": round(drop, 1)
                        })
            if deteriorating:
                det_df = pd.DataFrame(deteriorating).sort_values("下降幅度", ascending=False)
                st.markdown(f'<div class="warning-box">⚠️ {len(det_df)} 個科系轉換率下降</div>', unsafe_allow_html=True)
                st.dataframe(det_df, use_container_width=True, hide_index=True)
            else:
                st.success("✅ 所有科系轉換率持平或上升！")

        # 學校流失惡化
        st.markdown("---"); st.markdown("#### 🚨 學校流失惡化偵測")
        mn = st.slider("一階≥", 1, 50, 10, key="cross_loss_mn")
        if len(yr_list) >= 2:
            prev_yr = yr_list[-2]; curr_yr = yr_list[-1]
            prev_ss = year_cache[prev_yr].get("ss")
            curr_ss = year_cache[curr_yr].get("ss")
            if prev_ss is not None and curr_ss is not None:
                prev_filt = prev_ss[prev_ss["一階人數"] >= mn].copy()
                curr_filt = curr_ss[curr_ss["一階人數"] >= mn].copy()
                prev_filt["_key"] = prev_filt["學校"].apply(norm_school)
                curr_filt["_key"] = curr_filt["學校"].apply(norm_school)
                merged = prev_filt[["_key", "學校", "一→最終(%)"]].merge(
                    curr_filt[["_key", "一→最終(%)"]], on="_key", suffixes=(f"_{prev_yr}", f"_{curr_yr}"))
                merged["變化"] = merged[f"一→最終(%)_{curr_yr}"] - merged[f"一→最終(%)_{prev_yr}"]
                deteriorating_sch = merged[merged["變化"] < -5].sort_values("變化")
                improving_sch = merged[merged["變化"] > 5].sort_values("變化", ascending=False)

                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"**❌ 惡化學校（下降>5%）：{len(deteriorating_sch)} 校**")
                    if not deteriorating_sch.empty:
                        st.dataframe(deteriorating_sch[["學校", f"一→最終(%)_{prev_yr}",
                            f"一→最終(%)_{curr_yr}", "變化"]].head(20),
                            use_container_width=True, hide_index=True)
                    else:
                        st.success("✅ 無嚴重惡化")
                with c2:
                    st.markdown(f"**✅ 改善學校（上升>5%）：{len(improving_sch)} 校**")
                    if not improving_sch.empty:
                        st.dataframe(improving_sch[["學校", f"一→最終(%)_{prev_yr}",
                            f"一→最終(%)_{curr_yr}", "變化"]].head(20),
                            use_container_width=True, hide_index=True)
                    else:
                        st.info("無顯著改善")

        # IPA 矩陣疊合
        st.markdown("---"); st.markdown("#### IPA 矩陣（跨年度疊合）")
        fig = go.Figure()
        colors = px.colors.qualitative.Set2
        for i, (yr, c) in enumerate(year_cache.items()):
            if c["ss"] is not None:
                pool = c["ss"][c["ss"]["一階人數"] >= mn].copy()
                if not pool.empty:
                    fig.add_trace(go.Scatter(
                        x=pool["一階人數"], y=pool["一→最終(%)"],
                        mode="markers", name=yr,
                        marker=dict(size=pool["流失人數"].clip(lower=3)*0.8,
                                    color=colors[i % len(colors)], opacity=0.6),
                        text=pool["學校"], hovertemplate="%{text}<br>一階:%{x}<br>轉換率:%{y}%"
                    ))
        fig.update_layout(title="IPA 矩陣（跨年度）", xaxis_title="一階人數",
                          yaxis_title="一→最終(%)", height=600, showlegend=True)
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 主畫面
# ============================================================
tab_names = valid_years + (["📊 跨年度比較"] if len(valid_years) >= 2 else [])
tabs = st.tabs(tab_names)

for i, yr in enumerate(valid_years):
    with tabs[i]:
        st.markdown(f'<span class="year-tag">📅 {yr}</span>', unsafe_allow_html=True)
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
    '🎓 中華醫事科技大學 招生數據分析系統 v6.4<br>'
    '跨年度七模組完整分析 ｜ 班級結構化解析 ｜ P1科系優先映射<br>'
    '分析版本 #' + str(ver) + '</div>', unsafe_allow_html=True)
