# -*- coding: utf-8 -*-
"""
中華醫事科技大學 招生數據分析系統 v5.4
- 三階段：一階(報名) → 二階(報到確認) → 最終入學(註冊)
- 經緯度資料庫：獨立 Excel 匯入管理
- 最終入學管道從資料中「入學方式」欄位自動讀取
- 統一轉換率分母 = 一階人數
- 「更新分析」按鈕觸發統計重建
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import re
import io

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
        font-size: 2.2rem; font-weight: 800;
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6a4f 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        text-align: center; padding: 10px 0; margin-bottom: 5px;
    }
    .sub-header {
        font-size: 1.0rem; color: #6c757d;
        text-align: center; margin-bottom: 20px;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px; border-radius: 15px; text-align: center;
        color: white; margin: 5px 0;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .metric-card h3 { margin: 0; font-size: 0.85rem; opacity: 0.9; }
    .metric-card h1 { margin: 5px 0 0 0; font-size: 1.8rem; }
    .metric-green { background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%); color: #1a1a2e; }
    .metric-orange { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); }
    .metric-blue { background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); color: #1a1a2e; }
    .metric-gold { background: linear-gradient(135deg, #f6d365 0%, #fda085 100%); color: #1a1a2e; }
    .section-divider {
        height: 3px;
        background: linear-gradient(90deg, #667eea, #764ba2, #f093fb);
        border: none; border-radius: 2px; margin: 25px 0;
    }
    .info-box {
        background: #f0f4ff; border-left: 4px solid #667eea;
        padding: 15px; border-radius: 0 10px 10px 0;
        margin: 10px 0; font-size: 0.9rem;
    }
    .warning-box {
        background: #fff8e1; border-left: 4px solid #ffa726;
        padding: 15px; border-radius: 0 10px 10px 0;
        margin: 10px 0; font-size: 0.9rem;
    }
    .success-box {
        background: #e8f5e9; border-left: 4px solid #4caf50;
        padding: 15px; border-radius: 0 10px 10px 0;
        margin: 10px 0; font-size: 0.9rem;
    }
    .geo-box {
        background: #e8eaf6; border-left: 4px solid #3f51b5;
        padding: 15px; border-radius: 0 10px 10px 0;
        margin: 10px 0; font-size: 0.9rem;
    }
    .channel-tag-final {
        display: inline-block; background: #e8f5e9;
        color: #2e7d32; padding: 2px 10px; border-radius: 12px;
        font-size: 0.8rem; margin: 2px 3px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🎓 中華醫事科技大學 招生數據分析系統</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Enrollment Analytics Platform v5.4 ｜ 經緯度資料庫 Excel 匯入 ｜ 三階段完整分析</div>', unsafe_allow_html=True)
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

# ============================================================
# 常數
# ============================================================
KNOWN_CHANNELS = [
    "聯合免試", "甄選入學", "技優甄審", "運動績優",
    "身障甄試", "單獨招生", "進修部", "產學攜手",
    "運動單獨招生", "四技申請", "繁星計畫"
]

CHANNEL_KEYWORDS = {
    "聯合免試": ["免試", "聯合免試", "聯免"],
    "甄選入學": ["甄選", "甄選入學"],
    "技優甄審": ["技優", "技優甄審"],
    "運動績優": ["運動績優", "運績"],
    "身障甄試": ["身障", "身障甄試"],
    "單獨招生": ["單獨招生", "單招", "單獨"],
    "進修部":   ["進修部", "進修", "夜間"],
    "產學攜手": ["產攜", "產學攜手", "產學"],
    "四技申請": ["四技申請", "申請入學"],
    "繁星計畫": ["繁星"],
}

FINAL_CHANNEL_CANDIDATES = [
    "入學方式", "入學管道", "錄取管道", "招生管道", "管道",
    "入學途徑", "錄取方式", "報名管道"
]

HWU_COORDS = {"lat": 22.9340, "lon": 120.2756}

# ============================================================
# 欄位偵測
# ============================================================
def detect_school_column(df):
    candidates = ["畢業學校", "來源學校", "原就讀學校", "高中職校名",
                  "學校名稱", "校名", "畢業高中職", "畢業學校名稱",
                  "就讀學校", "原學校"]
    for c in candidates:
        for col in df.columns:
            if c in str(col):
                return col
    return None


def detect_department_column(df):
    candidates = ["科系", "系所", "報名科系", "錄取科系", "志願科系",
                  "就讀科系", "科系名稱", "系科", "學系", "錄取系所",
                  "就讀系所", "註冊科系"]
    for c in candidates:
        for col in df.columns:
            if c in str(col):
                return col
    return None


def detect_final_channel_column(df):
    for c in FINAL_CHANNEL_CANDIDATES:
        for col in df.columns:
            if c in str(col):
                return col
    return None


def detect_channel_from_filename(filename):
    if not filename:
        return None
    for ch, kws in CHANNEL_KEYWORDS.items():
        for kw in kws:
            if kw in filename:
                return ch
    return None


def detect_channel_from_columns(df):
    for col_name in ["管道", "入學管道", "招生管道", "報名管道"]:
        if col_name in df.columns:
            top = df[col_name].dropna().value_counts()
            if not top.empty:
                val = str(top.index[0])
                for ch, kws in CHANNEL_KEYWORDS.items():
                    for kw in kws:
                        if kw in val:
                            return ch
                return val
    return None


def normalize_school_name(name):
    if not isinstance(name, str):
        return str(name)
    name = name.strip()
    name = name.replace("臺", "台").replace("（", "(").replace("）", ")")
    name = re.sub(r"\s+", "", name)
    for suffix in ["附設進修學校", "進修學校", "進修部"]:
        name = name.replace(suffix, "")
    return name


def efficiency_stars(rate):
    if rate >= 70:
        return "⭐⭐⭐"
    elif rate >= 40:
        return "⭐⭐"
    else:
        return "⭐"


# ============================================================
# 經緯度資料庫管理（Excel 匯入）
# ============================================================
def detect_geo_columns(df):
    """從經緯度 Excel 中偵測學校名稱、緯度、經度欄位"""
    school_col = None
    lat_col = None
    lon_col = None

    school_kw = ["學校", "校名", "學校名稱", "畢業學校", "school"]
    lat_kw = ["緯度", "lat", "latitude", "Lat", "LAT", "緯"]
    lon_kw = ["經度", "lon", "lng", "longitude", "Lon", "LON", "經"]

    for col in df.columns:
        s = str(col).strip().lower()
        if school_col is None:
            for kw in school_kw:
                if kw.lower() in s:
                    school_col = col
                    break
        if lat_col is None:
            for kw in lat_kw:
                if kw.lower() in s:
                    lat_col = col
                    break
        if lon_col is None:
            for kw in lon_kw:
                if kw.lower() in s:
                    lon_col = col
                    break

    return school_col, lat_col, lon_col


def build_geo_lookup_from_excel(geo_df):
    """從 Excel 建構 geo_lookup 表"""
    school_col, lat_col, lon_col = detect_geo_columns(geo_df)

    if school_col is None or lat_col is None or lon_col is None:
        return None, "❌ 無法偵測欄位。需要：學校名稱、緯度、經度"

    geo = geo_df[[school_col, lat_col, lon_col]].copy()
    geo.columns = ["原始校名", "緯度", "經度"]
    geo["緯度"] = pd.to_numeric(geo["緯度"], errors="coerce")
    geo["經度"] = pd.to_numeric(geo["經度"], errors="coerce")
    geo = geo.dropna(subset=["緯度", "經度"])
    geo = geo[geo["原始校名"].notna()]
    geo["學校名稱_標準"] = geo["原始校名"].apply(normalize_school_name)

    lookup = geo.groupby("學校名稱_標準").agg(
        緯度=("緯度", "mean"), 經度=("經度", "mean")
    ).reset_index()

    msg = f"✅ 成功載入 {len(lookup)} 所學校的經緯度"
    return lookup, msg


def create_geo_template():
    """建立經緯度 Excel 範本"""
    template = pd.DataFrame({
        "學校名稱": [
            "台南高工", "台南女中", "台南一中",
            "台南高商", "長榮中學", "家齊高中",
            "新營高中", "善化高中", "曾文農工",
            "北門農工"
        ],
        "緯度": [
            22.9833, 22.9914, 23.0006,
            22.9836, 22.9647, 22.9825,
            23.3103, 23.1322, 23.2506,
            23.2667
        ],
        "經度": [
            120.2122, 120.2042, 120.2120,
            120.1989, 120.2122, 120.1953,
            120.3164, 120.2967, 120.3461,
            120.1833
        ]
    })
    return template


def enrich_with_geo(df, geo_lookup):
    """將 geo_lookup 的經緯度合併到資料中"""
    school_col = detect_school_column(df)
    if school_col is None or geo_lookup is None or geo_lookup.empty:
        return df
    df = df.copy()
    df["_std"] = df[school_col].apply(normalize_school_name)
    for c in ["緯度", "經度", "lat", "lon", "latitude", "longitude"]:
        if c in df.columns and c != school_col:
            df = df.drop(columns=[c], errors="ignore")
    df = df.merge(
        geo_lookup.rename(columns={"學校名稱_標準": "_std"}),
        on="_std", how="left"
    )
    df = df.drop(columns=["_std"], errors="ignore")
    return df


def get_geo_match_report(df, geo_lookup, phase_name):
    """產生經緯度匹配報告"""
    school_col = detect_school_column(df)
    if school_col is None or geo_lookup is None:
        return None

    schools = df[school_col].dropna().unique()
    std_names = [normalize_school_name(s) for s in schools]
    geo_names = set(geo_lookup["學校名稱_標準"].values)

    matched = [s for s, std in zip(schools, std_names) if std in geo_names]
    unmatched = [s for s, std in zip(schools, std_names) if std not in geo_names]

    return {
        "phase": phase_name,
        "total": len(schools),
        "matched": len(matched),
        "unmatched": len(unmatched),
        "unmatched_list": unmatched,
        "match_rate": len(matched) / len(schools) * 100 if schools.size else 0
    }


# ============================================================
# 三階段統計
# ============================================================
def build_school_stats(p1_df, p2_df=None, p3_df=None):
    sc1 = detect_school_column(p1_df)
    if sc1 is None:
        return None
    stats = p1_df[sc1].value_counts().reset_index()
    stats.columns = ["學校", "一階人數"]

    if p2_df is not None:
        sc2 = detect_school_column(p2_df)
        if sc2:
            t2 = p2_df[sc2].value_counts().reset_index()
            t2.columns = ["學校", "二階人數"]
            stats = stats.merge(t2, on="學校", how="left")
    if "二階人數" not in stats.columns:
        stats["二階人數"] = np.nan

    if p3_df is not None:
        sc3 = detect_school_column(p3_df)
        if sc3:
            t3 = p3_df[sc3].value_counts().reset_index()
            t3.columns = ["學校", "最終入學"]
            stats = stats.merge(t3, on="學校", how="left")
    if "最終入學" not in stats.columns:
        stats["最終入學"] = np.nan

    stats["二階人數"] = stats["二階人數"].fillna(0).astype(int)
    stats["最終入學"] = stats["最終入學"].fillna(0).astype(int)
    stats["一→二階(%)"] = (stats["二階人數"] / stats["一階人數"] * 100).round(1)
    stats["一→最終(%)"] = (stats["最終入學"] / stats["一階人數"] * 100).round(1)
    stats["流失人數"] = stats["一階人數"] - stats["最終入學"]
    stats["效率評等"] = stats["一→最終(%)"].apply(efficiency_stars)
    return stats


def build_dept_stats(p1_df, p2_df=None, p3_df=None):
    dc1 = detect_department_column(p1_df)
    if dc1 is None:
        return None
    stats = p1_df[dc1].value_counts().reset_index()
    stats.columns = ["科系", "一階人數"]

    if p2_df is not None:
        dc2 = detect_department_column(p2_df)
        if dc2:
            t2 = p2_df[dc2].value_counts().reset_index()
            t2.columns = ["科系", "二階人數"]
            stats = stats.merge(t2, on="科系", how="left")
    if "二階人數" not in stats.columns:
        stats["二階人數"] = np.nan

    if p3_df is not None:
        dc3 = detect_department_column(p3_df)
        if dc3:
            t3 = p3_df[dc3].value_counts().reset_index()
            t3.columns = ["科系", "最終入學"]
            stats = stats.merge(t3, on="科系", how="left")
    if "最終入學" not in stats.columns:
        stats["最終入學"] = np.nan

    stats["二階人數"] = stats["二階人數"].fillna(0).astype(int)
    stats["最終入學"] = stats["最終入學"].fillna(0).astype(int)
    stats["一→二階(%)"] = (stats["二階人數"] / stats["一階人數"] * 100).round(1)
    stats["一→最終(%)"] = (stats["最終入學"] / stats["一階人數"] * 100).round(1)
    stats["效率評等"] = stats["一→最終(%)"].apply(efficiency_stars)
    return stats


# ============================================================
# 視覺化
# ============================================================
def create_funnel_chart(labels, values, title="招生漏斗"):
    colors = ["#2196F3", "#FF9800", "#4CAF50", "#E91E63"]
    fig = go.Figure(go.Funnel(
        y=labels, x=values,
        textinfo="value+percent initial",
        marker=dict(color=colors[:len(labels)]),
        connector=dict(line=dict(color="royalblue", width=2))
    ))
    fig.update_layout(title=title, height=420, font=dict(size=14))
    return fig


def create_bar_h(df, y_col, x_col, title, color="#667eea", fmt=".1f"):
    fig = px.bar(
        df.sort_values(x_col, ascending=True),
        x=x_col, y=y_col, orientation="h",
        text=x_col, title=title
    )
    fig.update_traces(marker_color=color,
                      texttemplate=f"%{{text:{fmt}}}%", textposition="outside")
    fig.update_layout(height=max(380, len(df) * 28),
                      xaxis_title="轉換率 (%)", yaxis_title="")
    return fig


def create_grouped_bar(df, y_col, val_cols, title):
    fig = go.Figure()
    colors = ["#2196F3", "#FF9800", "#4CAF50"]
    for i, vc in enumerate(val_cols):
        if vc in df.columns:
            fig.add_trace(go.Bar(
                name=vc, y=df[y_col], x=df[vc],
                orientation="h", marker_color=colors[i % 3],
                text=df[vc], textposition="outside"
            ))
    fig.update_layout(barmode="group", title=title,
                      height=max(400, len(df) * 35),
                      yaxis=dict(autorange="reversed"))
    return fig


def create_map(df, size_col, title, color_col=None):
    if "緯度" not in df.columns or "經度" not in df.columns:
        return None
    mdf = df.dropna(subset=["緯度", "經度"]).copy()
    if mdf.empty:
        return None
    mdf[size_col] = pd.to_numeric(mdf[size_col], errors="coerce").fillna(1)
    school_col = detect_school_column(mdf)
    fig = px.scatter_mapbox(
        mdf, lat="緯度", lon="經度", size=size_col,
        color=color_col if color_col and color_col in mdf.columns else None,
        hover_name=school_col if school_col and school_col in mdf.columns else None,
        hover_data={size_col: True, "緯度": ":.4f", "經度": ":.4f"},
        title=title, size_max=30, zoom=7,
        center={"lat": HWU_COORDS["lat"], "lon": HWU_COORDS["lon"]},
        mapbox_style="carto-positron"
    )
    fig.add_trace(go.Scattermapbox(
        lat=[HWU_COORDS["lat"]], lon=[HWU_COORDS["lon"]],
        mode="markers+text",
        marker=dict(size=18, color="red", symbol="star"),
        text=["中華醫事科技大學"], textposition="top center",
        name="本校", showlegend=True
    ))
    fig.update_layout(height=600, margin=dict(l=0, r=0, t=40, b=0))
    return fig


def create_heatmap(df, x_col, y_col, val_col, title):
    pv = df.pivot_table(index=y_col, columns=x_col, values=val_col,
                        aggfunc="sum").fillna(0)
    fig = px.imshow(pv, text_auto=True, aspect="auto",
                    color_continuous_scale="YlOrRd", title=title)
    fig.update_layout(height=max(400, len(pv) * 25))
    return fig


# ============================================================
# Session State 初始化
# ============================================================
if "datasets" not in st.session_state:
    st.session_state["datasets"] = {}
if "geo_lookup" not in st.session_state:
    st.session_state["geo_lookup"] = None
if "geo_file_name" not in st.session_state:
    st.session_state["geo_file_name"] = None
if "analysis_ready" not in st.session_state:
    st.session_state["analysis_ready"] = False
if "analysis_version" not in st.session_state:
    st.session_state["analysis_version"] = 0

# ============================================================
# Sidebar
# ============================================================
with st.sidebar:
    st.header("📂 資料上傳")

    # ── 招生資料上傳 ──
    uploaded = st.file_uploader(
        "上傳招生資料 (Excel / CSV)", type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="可多檔上傳，再分別指定為一階/二階/最終入學資料"
    )

    if uploaded:
        for uf in uploaded:
            if uf.name not in st.session_state["datasets"]:
                try:
                    df = (pd.read_csv(uf) if uf.name.endswith(".csv")
                          else pd.read_excel(uf))
                    auto_ch = (detect_channel_from_columns(df)
                               or detect_channel_from_filename(uf.name))
                    final_ch_col = detect_final_channel_column(df)
                    st.session_state["datasets"][uf.name] = {
                        "df": df,
                        "channel": auto_ch,
                        "channel_confirmed": auto_ch or "",
                        "filename": uf.name,
                        "has_final_channel": final_ch_col is not None,
                        "final_channel_col": final_ch_col
                    }
                    st.success(f"✅ {uf.name}（{len(df)} 筆）")
                    if final_ch_col:
                        st.info(f"📌 偵測到入學方式欄位：「{final_ch_col}」")
                except Exception as e:
                    st.error(f"❌ {uf.name}: {e}")

    # ── 經緯度資料庫上傳 ──
    st.markdown("---")
    st.header("🌐 經緯度資料庫")

    st.markdown("""
    <div style="font-size:0.8rem; color:#666; margin-bottom:8px;">
    📍 上傳 Excel 檔，包含：<br>
    　　<b>學校名稱</b>、<b>緯度</b>、<b>經度</b> 三欄<br>
    📋 所有階段的地圖均使用此資料庫
    </div>
    """, unsafe_allow_html=True)

    # 範本下載
    template = create_geo_template()
    buffer = io.BytesIO()
    template.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)
    st.download_button(
        label="📥 下載經緯度範本 Excel",
        data=buffer,
        file_name="經緯度資料庫_範本.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="下載範本後填入學校經緯度，再上傳"
    )

    geo_file = st.file_uploader(
        "上傳經緯度 Excel",
        type=["xlsx", "xls"],
        key="geo_uploader",
        help="欄位需包含：學校名稱、緯度、經度"
    )

    if geo_file:
        if geo_file.name != st.session_state.get("geo_file_name"):
            try:
                geo_df = pd.read_excel(geo_file)
                lookup, msg = build_geo_lookup_from_excel(geo_df)
                if lookup is not None:
                    st.session_state["geo_lookup"] = lookup
                    st.session_state["geo_file_name"] = geo_file.name
                    st.success(msg)

                    # 顯示預覽
                    with st.expander(f"📊 已載入 {len(lookup)} 所學校", expanded=False):
                        st.dataframe(lookup, use_container_width=True, hide_index=True)
                else:
                    st.error(msg)
            except Exception as e:
                st.error(f"❌ 讀取失敗：{e}")
    else:
        if st.session_state["geo_lookup"] is not None:
            n = len(st.session_state["geo_lookup"])
            st.markdown(
                f'<div class="geo-box">📍 已載入經緯度資料庫：<b>{n}</b> 所學校<br>'
                f'📄 {st.session_state.get("geo_file_name", "")}</div>',
                unsafe_allow_html=True
            )

    if st.session_state["geo_lookup"] is not None:
        if st.button("🗑️ 清除經緯度資料庫"):
            st.session_state["geo_lookup"] = None
            st.session_state["geo_file_name"] = None
            st.success("已清除經緯度資料庫")
            st.rerun()

    st.markdown("---")

    # ── 已載入資料清單 ──
    if st.session_state["datasets"]:
        st.subheader("📋 已載入招生資料")
        for name, info in st.session_state["datasets"].items():
            ch = info.get("channel_confirmed") or info.get("channel") or "待確認"
            badge = "📋多管道" if info.get("has_final_channel") else "🏷️"
            st.caption(f"📄 **{name}**　{len(info['df'])} 筆　{badge} {ch}")

        if st.button("🗑️ 清除全部招生資料"):
            st.session_state["datasets"] = {}
            st.session_state["analysis_ready"] = False
            st.session_state["analysis_version"] = 0
            st.rerun()

    # ── 三階段指定 ──
    st.markdown("---")
    st.subheader("⚙️ 三階段資料指定")
    st.markdown("""
    <div style="font-size:0.8rem; color:#888; margin-bottom:10px;">
    📊 所有轉換率分母 = 一階人數<br>
    🌐 地圖經緯度 = 經緯度資料庫<br>
    🟢 最終入學管道由資料自動讀取
    </div>
    """, unsafe_allow_html=True)

    ds_names = list(st.session_state["datasets"].keys())
    none_opt = ["-- 未選擇 --"]

    p1_sel = st.selectbox("🔵 一階（報名）", none_opt + ds_names, key="p1_sel")
    p2_sel = st.selectbox("🟠 二階（報到確認）", none_opt + ds_names, key="p2_sel")
    p3_sel = st.selectbox("🟢 最終入學（註冊）", none_opt + ds_names, key="p3_sel")

    # ── 一階/二階 管道確認 ──
    st.markdown("---")
    st.subheader("🏷️ 一階 / 二階 管道確認")

    for sel_key, label in [("p1_sel", "🔵 一階"), ("p2_sel", "🟠 二階")]:
        sel_name = st.session_state.get(sel_key, "-- 未選擇 --")
        if sel_name != "-- 未選擇 --" and sel_name in st.session_state["datasets"]:
            info = st.session_state["datasets"][sel_name]
            with st.expander(f"{label}：{sel_name}"):
                auto = info.get("channel")
                if auto:
                    st.info(f"自動偵測：**{auto}**")
                    keep = st.radio(
                        "使用偵測結果？", ["✅ 是", "❌ 自行選擇"],
                        horizontal=True, key=f"chk_{sel_key}"
                    )
                    if keep == "✅ 是":
                        confirmed = auto
                    else:
                        confirmed = st.selectbox(
                            "選擇管道：", KNOWN_CHANNELS, key=f"chs_{sel_key}"
                        )
                else:
                    st.warning("⚠️ 無法自動偵測管道")
                    confirmed = st.selectbox(
                        "選擇管道：", KNOWN_CHANNELS, key=f"chs_{sel_key}"
                    )
                st.session_state["datasets"][sel_name]["channel_confirmed"] = confirmed

    # ── 最終入學管道確認（自動讀取）──
    st.markdown("---")
    st.subheader("🟢 最終入學 管道確認")

    p3_name = st.session_state.get("p3_sel", "-- 未選擇 --")
    if p3_name != "-- 未選擇 --" and p3_name in st.session_state["datasets"]:
        p3_info = st.session_state["datasets"][p3_name]
        p3_df_preview = p3_info["df"]
        final_ch_col = detect_final_channel_column(p3_df_preview)

        if final_ch_col:
            st.success(f"✅ 偵測到欄位：「**{final_ch_col}**」")
            ch_values = p3_df_preview[final_ch_col].fillna("(空白)").astype(str).str.strip()
            ch_values = ch_values.replace("", "(空白)")
            ch_dist = ch_values.value_counts()

            st.markdown("**📊 各入學管道人數分布：**")
            for ch_name, cnt in ch_dist.items():
                pct = cnt / len(p3_df_preview) * 100
                st.markdown(
                    f'<span class="channel-tag-final">{ch_name}</span> '
                    f'{cnt} 人（{pct:.1f}%）',
                    unsafe_allow_html=True
                )

            st.markdown("---")
            all_channels = ch_dist.index.tolist()
            selected_channels = st.multiselect(
                "📌 選擇要納入分析的入學管道：",
                all_channels,
                default=all_channels,
                key="final_channels_filter",
                help="取消勾選可排除特定管道"
            )
            st.session_state["final_selected_channels"] = selected_channels
            st.session_state["final_channel_col"] = final_ch_col

            n_sel = ch_values.isin(selected_channels).sum()
            st.caption(f"已選管道涵蓋 **{n_sel}** / {len(p3_df_preview)} 人")
        else:
            st.warning("⚠️ 未偵測到「入學方式」欄位")
            manual = st.selectbox(
                "手動選擇欄位：",
                ["-- 無 --"] + list(p3_df_preview.columns),
                key="manual_final_ch"
            )
            if manual != "-- 無 --":
                st.session_state["final_channel_col"] = manual
                ch_dist = p3_df_preview[manual].fillna("(空白)").value_counts()
                all_channels = ch_dist.index.tolist()
                selected_channels = st.multiselect(
                    "選擇管道：", all_channels,
                    default=all_channels, key="final_channels_filter_manual"
                )
                st.session_state["final_selected_channels"] = selected_channels
    else:
        st.info("ℹ️ 請先在上方選擇最終入學資料")

    # ── 更新分析按鈕 ──
    st.markdown("---")
    st.markdown("⬇️ **設定完成後請點擊下方按鈕**")

    update_clicked = st.button(
        "🔄 更新分析",
        type="primary",
        use_container_width=True,
        help="重新刷新所有統計並匹配經緯度"
    )

    if update_clicked:
        geo = st.session_state.get("geo_lookup")

        # 經緯度匹配報告
        if geo is not None and not geo.empty:
            report_lines = []
            for sel_key, label in [("p1_sel", "一階"), ("p2_sel", "二階"), ("p3_sel", "最終入學")]:
                sel_name = st.session_state.get(sel_key, "-- 未選擇 --")
                if sel_name != "-- 未選擇 --" and sel_name in st.session_state["datasets"]:
                    rpt = get_geo_match_report(
                        st.session_state["datasets"][sel_name]["df"],
                        geo, label
                    )
                    if rpt:
                        report_lines.append(
                            f"📍 {label}：{rpt['matched']}/{rpt['total']} "
                            f"校匹配（{rpt['match_rate']:.1f}%）"
                        )
            if report_lines:
                for rl in report_lines:
                    st.info(rl)
        else:
            st.warning("⚠️ 未載入經緯度資料庫，地圖功能將無法使用")

        st.session_state["analysis_ready"] = True
        st.session_state["analysis_version"] += 1
        st.success(f"✅ 分析已更新！（版本 #{st.session_state['analysis_version']}）")

    # 狀態顯示
    st.markdown("---")
    if st.session_state["analysis_ready"]:
        v = st.session_state["analysis_version"]
        geo_status = "✅ 已載入" if st.session_state["geo_lookup"] is not None else "❌ 未載入"
        st.markdown(
            f'<div class="success-box">✅ 分析就緒　版本 #{v}<br>🌐 經緯度：{geo_status}</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown('<div class="warning-box">⏳ 請設定資料後按「更新分析」</div>',
                    unsafe_allow_html=True)


# ============================================================
# 取得三階段資料
# ============================================================
def get_phase_dfs():
    geo = st.session_state.get("geo_lookup")
    p1 = p2 = p3 = None

    sel1 = st.session_state.get("p1_sel", "-- 未選擇 --")
    sel2 = st.session_state.get("p2_sel", "-- 未選擇 --")
    sel3 = st.session_state.get("p3_sel", "-- 未選擇 --")

    if sel1 != "-- 未選擇 --" and sel1 in st.session_state["datasets"]:
        raw = st.session_state["datasets"][sel1]["df"].copy()
        p1 = enrich_with_geo(raw, geo) if geo is not None else raw
    if sel2 != "-- 未選擇 --" and sel2 in st.session_state["datasets"]:
        raw = st.session_state["datasets"][sel2]["df"].copy()
        p2 = enrich_with_geo(raw, geo) if geo is not None else raw
    if sel3 != "-- 未選擇 --" and sel3 in st.session_state["datasets"]:
        raw = st.session_state["datasets"][sel3]["df"].copy()

        # 管道篩選
        ch_col = st.session_state.get("final_channel_col")
        sel_chs = st.session_state.get("final_selected_channels")
        if ch_col and ch_col in raw.columns and sel_chs:
            raw[ch_col] = raw[ch_col].fillna("(空白)").astype(str).str.strip()
            raw.loc[raw[ch_col] == "", ch_col] = "(空白)"
            raw = raw[raw[ch_col].isin(sel_chs)]

        p3 = enrich_with_geo(raw, geo) if geo is not None else raw
    return p1, p2, p3


def get_final_channel_col():
    return st.session_state.get("final_channel_col")


# ============================================================
# 主功能選單
# ============================================================
modules = {
    "📊 Module 0：總覽儀表板":      "mod0",
    "🔄 Module 1：招生漏斗分析":     "mod1",
    "📈 Module 2：入學管道分析":      "mod2",
    "🗺️ Module 3：地理分布（地圖）": "mod3",
    "🏫 Module 4：科系熱力圖":       "mod4",
    "🎯 Module 5：來源學校追蹤":     "mod5",
    "⚠️ Module 6：流失預警分析":     "mod6",
}

selected = st.selectbox("選擇分析模組：", list(modules.keys()))
mod = modules[selected]
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

if not st.session_state["analysis_ready"]:
    st.markdown("""
    <div class="warning-box">
        <h4>⏳ 尚未執行分析</h4>
        <p>請先在左側：</p>
        <ol>
            <li>上傳招生資料檔案</li>
            <li>上傳經緯度資料庫 Excel（地圖功能需要）</li>
            <li>指定三階段（一階/二階/最終入學）</li>
            <li>確認管道設定</li>
            <li>點擊 <strong>🔄 更新分析</strong> 按鈕</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# ============================================================
# Module 0：總覽儀表板
# ============================================================
if mod == "mod0":
    st.header("📊 Module 0：總覽儀表板")
    p1, p2, p3 = get_phase_dfs()

    if p1 is None:
        st.warning("⚠️ 請指定一階資料並按「更新分析」。"); st.stop()

    n1 = len(p1)
    n2 = len(p2) if p2 is not None else None
    n3 = len(p3) if p3 is not None else None

    cols = st.columns(5)
    with cols[0]:
        st.markdown(f'<div class="metric-card"><h3>一階 報名</h3><h1>{n1:,}</h1></div>',
                    unsafe_allow_html=True)
    with cols[1]:
        v = f"{n2:,}" if n2 else "—"
        st.markdown(f'<div class="metric-card metric-orange"><h3>二階 報到</h3><h1>{v}</h1></div>',
                    unsafe_allow_html=True)
    with cols[2]:
        v = f"{n3:,}" if n3 else "—"
        st.markdown(f'<div class="metric-card metric-green"><h3>最終入學</h3><h1>{v}</h1></div>',
                    unsafe_allow_html=True)
    with cols[3]:
        r = f"{n2/n1*100:.1f}%" if n2 and n1 else "—"
        st.markdown(f'<div class="metric-card metric-blue"><h3>一→二階 轉換率</h3><h1>{r}</h1></div>',
                    unsafe_allow_html=True)
    with cols[4]:
        r = f"{n3/n1*100:.1f}%" if n3 and n1 else "—"
        st.markdown(f'<div class="metric-card metric-gold"><h3>一→最終 轉換率</h3><h1>{r}</h1></div>',
                    unsafe_allow_html=True)

    # 經緯度匹配狀態
    geo = st.session_state.get("geo_lookup")
    if geo is not None:
        rpt = get_geo_match_report(p1, geo, "一階")
        if rpt:
            st.markdown(
                f'<div class="geo-box">🌐 經緯度匹配率（一階）：'
                f'<b>{rpt["matched"]}/{rpt["total"]}</b> 所學校 '
                f'（{rpt["match_rate"]:.1f}%）</div>',
                unsafe_allow_html=True
            )
    else:
        st.markdown(
            '<div class="warning-box">🌐 未載入經緯度資料庫 — 地圖功能不可用</div>',
            unsafe_allow_html=True
        )

    # 最終入學管道分布
    if p3 is not None:
        ch_col = get_final_channel_col()
        if ch_col and ch_col in p3.columns:
            st.markdown("---")
            st.subheader("🟢 最終入學 — 各管道人數")
            ch_dist = p3[ch_col].value_counts().reset_index()
            ch_dist.columns = ["入學管道", "人數"]
            ch_dist["佔比(%)"] = (ch_dist["人數"] / ch_dist["人數"].sum() * 100).round(1)

            c1, c2 = st.columns([1, 1])
            with c1:
                fig = px.pie(ch_dist, names="入學管道", values="人數",
                             title="最終入學 管道分布", hole=0.35)
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = px.bar(ch_dist.sort_values("人數", ascending=True),
                             x="人數", y="入學管道", orientation="h",
                             text="人數", title="各管道入學人數排行")
                fig.update_traces(marker_color="#4CAF50")
                fig.update_layout(height=max(400, len(ch_dist) * 28),
                                  yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    f_labels, f_values = ["一階 報名"], [n1]
    if n2 is not None:
        f_labels.append("二階 報到確認"); f_values.append(n2)
    if n3 is not None:
        f_labels.append("最終入學"); f_values.append(n3)
    if len(f_values) > 1:
        st.plotly_chart(create_funnel_chart(f_labels, f_values, "整體招生漏斗"),
                        use_container_width=True)

    c1, c2 = st.columns(2)
    dept_col = detect_department_column(p1)
    school_col = detect_school_column(p1)
    with c1:
        if dept_col:
            dd = p1[dept_col].value_counts().reset_index()
            dd.columns = ["科系", "人數"]
            fig = px.pie(dd, names="科系", values="人數",
                         title="一階報名 科系分布", hole=0.4)
            st.plotly_chart(fig, use_container_width=True)
    with c2:
        if school_col:
            sd = p1[school_col].value_counts().head(10).reset_index()
            sd.columns = ["學校", "人數"]
            fig = px.bar(sd, x="人數", y="學校", orientation="h",
                         title="來源學校 TOP 10", text="人數")
            fig.update_traces(marker_color="#667eea")
            fig.update_layout(yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig, use_container_width=True)

    dstats = build_dept_stats(p1, p2, p3)
    if dstats is not None:
        st.markdown("---")
        st.subheader("各科系三階段概覽")
        st.dataframe(dstats.sort_values("一階人數", ascending=False),
                     use_container_width=True, hide_index=True)


# ============================================================
# Module 1：招生漏斗分析
# ============================================================
elif mod == "mod1":
    st.header("🔄 Module 1：招生漏斗分析")
    p1, p2, p3 = get_phase_dfs()
    if p1 is None:
        st.warning("⚠️ 請先指定一階資料。"); st.stop()

    st.subheader("1-1. 各科系三階段漏斗")
    dstats = build_dept_stats(p1, p2, p3)
    if dstats is not None:
        st.dataframe(dstats.sort_values("一→最終(%)", ascending=False),
                     use_container_width=True, hide_index=True)
        rate_cols = [c for c in ["一→二階(%)", "一→最終(%)"] if c in dstats.columns]
        if rate_cols:
            fig = create_grouped_bar(
                dstats.sort_values(rate_cols[0], ascending=True),
                "科系", rate_cols, "各科系轉換率比較（分母=一階）"
            )
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("1-2. 單科系漏斗圖")
        sel = st.selectbox("選擇科系：", dstats["科系"].tolist())
        row = dstats[dstats["科系"] == sel].iloc[0]
        fl, fv = ["一階報名"], [int(row["一階人數"])]
        if row["二階人數"] > 0 or p2 is not None:
            fl.append("二階報到"); fv.append(int(row["二階人數"]))
        if row["最終入學"] > 0 or p3 is not None:
            fl.append("最終入學"); fv.append(int(row["最終入學"]))
        st.plotly_chart(create_funnel_chart(fl, fv, f"{sel} 招生漏斗"),
                        use_container_width=True)

    if p3 is not None:
        ch_col = get_final_channel_col()
        if ch_col and ch_col in p3.columns:
            st.markdown("---")
            st.subheader("1-3. 各入學管道漏斗")
            ch_list = p3[ch_col].value_counts().index.tolist()
            sel_ch = st.selectbox("選擇入學管道：", ch_list, key="funnel_ch")
            if sel_ch:
                p3_ch = p3[p3[ch_col] == sel_ch]
                n = len(p3_ch)
                st.metric("最終入學人數", f"{n} 人")
                dept_col = detect_department_column(p3_ch)
                if dept_col:
                    dd = p3_ch[dept_col].value_counts().reset_index()
                    dd.columns = ["科系", "人數"]
                    fig = px.bar(dd, x="科系", y="人數", text="人數",
                                 title=f"「{sel_ch}」各科系入學人數",
                                 color="人數", color_continuous_scale="Greens")
                    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("1-4. 各來源學校三階段漏斗")
    sstats = build_school_stats(p1, p2, p3)
    if sstats is not None:
        mn = st.slider("一階報名 ≥", 1, 50, 5, key="m1_sch")
        sf = sstats[sstats["一階人數"] >= mn].sort_values("一→最終(%)", ascending=False)
        st.dataframe(sf, use_container_width=True, hide_index=True)
        fig = create_bar_h(sf.head(20), "學校", "一→最終(%)",
                           f"來源學校 轉換率 TOP 20（一階≥{mn}人）",
                           color="#4CAF50")
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# Module 2：入學管道分析
# ============================================================
elif mod == "mod2":
    st.header("📈 Module 2：入學管道分析")
    p1, p2, p3 = get_phase_dfs()

    if p3 is None:
        st.warning("⚠️ 此模組需要最終入學資料。"); st.stop()

    ch_col = get_final_channel_col()
    if not ch_col or ch_col not in p3.columns:
        st.warning("⚠️ 未偵測到入學方式欄位。"); st.stop()

    st.markdown(f'<div class="info-box">📌 分析依據：最終入學「<b>{ch_col}</b>」欄位</div>',
                unsafe_allow_html=True)

    st.subheader("2-1. 各入學管道人數統計")
    ch_stats = p3[ch_col].value_counts().reset_index()
    ch_stats.columns = ["入學管道", "人數"]
    ch_stats["佔比(%)"] = (ch_stats["人數"] / ch_stats["人數"].sum() * 100).round(1)
    ch_stats["累積佔比(%)"] = ch_stats["佔比(%)"].cumsum().round(1)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.pie(ch_stats, names="入學管道", values="人數",
                     title="入學管道佔比", hole=0.35)
        fig.update_layout(height=500)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(ch_stats.sort_values("人數", ascending=True),
                     x="人數", y="入學管道", orientation="h",
                     text="人數", title="各管道人數排行",
                     color="佔比(%)", color_continuous_scale="Viridis")
        fig.update_layout(height=max(500, len(ch_stats) * 30),
                          yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig, use_container_width=True)

    st.dataframe(ch_stats, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("2-2. 入學管道 × 科系")
    dept_col = detect_department_column(p3)
    if dept_col:
        cross = p3.groupby([ch_col, dept_col]).size().reset_index(name="人數")
        fig = px.bar(cross, x=ch_col, y="人數", color=dept_col,
                     barmode="stack", title="管道×科系 堆疊圖", text="人數")
        fig.update_layout(height=600, xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

        st.plotly_chart(
            create_heatmap(cross, dept_col, ch_col, "人數", "管道×科系 熱力圖"),
            use_container_width=True
        )

        st.subheader("2-3. 各管道科系結構")
        sel_ch = st.selectbox("選擇管道：", ch_stats["入學管道"].tolist(), key="m2_ch")
        if sel_ch:
            sub = p3[p3[ch_col] == sel_ch]
            dd = sub[dept_col].value_counts().reset_index()
            dd.columns = ["科系", "人數"]
            dd["佔比(%)"] = (dd["人數"] / dd["人數"].sum() * 100).round(1)
            c1, c2 = st.columns(2)
            with c1:
                fig = px.pie(dd, names="科系", values="人數",
                             title=f"「{sel_ch}」科系結構", hole=0.35)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                st.dataframe(dd, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("2-4. 入學管道 × 來源學校")
    school_col = detect_school_column(p3)
    if school_col:
        div = p3.groupby(ch_col)[school_col].nunique().reset_index()
        div.columns = ["入學管道", "來源學校數"]
        fig = px.bar(div.sort_values("來源學校數", ascending=True),
                     x="來源學校數", y="入學管道", orientation="h",
                     text="來源學校數", color="來源學校數",
                     color_continuous_scale="Blues",
                     title="各管道來源學校多元性")
        fig.update_layout(height=max(400, len(div) * 28))
        st.plotly_chart(fig, use_container_width=True)

        sel_ch2 = st.selectbox("選擇管道查看學校排行：",
                               ch_stats["入學管道"].tolist(), key="m2_sch")
        if sel_ch2:
            sub = p3[p3[ch_col] == sel_ch2]
            sch = sub[school_col].value_counts().head(15).reset_index()
            sch.columns = [school_col, "人數"]
            fig = px.bar(sch, x="人數", y=school_col, orientation="h",
                         text="人數", title=f"「{sel_ch2}」來源學校 TOP 15")
            fig.update_traces(marker_color="#2196F3")
            fig.update_layout(yaxis=dict(autorange="reversed"),
                              height=max(400, len(sch) * 30))
            st.plotly_chart(fig, use_container_width=True)


# ============================================================
# Module 3：地理分布（地圖）
# ============================================================
elif mod == "mod3":
    st.header("🗺️ Module 3：地理分布（地圖）")
    p1, p2, p3 = get_phase_dfs()
    geo = st.session_state.get("geo_lookup")

    if p1 is None:
        st.warning("⚠️ 請先指定一階資料。"); st.stop()

    if geo is None or geo.empty:
        st.markdown("""
        <div class="warning-box">
            <h4>🌐 未載入經緯度資料庫</h4>
            <p>請在左側「經緯度資料庫」區塊上傳 Excel 檔案。</p>
            <p>Excel 需包含：<b>學校名稱</b>、<b>緯度</b>、<b>經度</b> 三欄。</p>
            <p>可先下載範本 Excel 作為格式參考。</p>
        </div>
        """, unsafe_allow_html=True)
        st.stop()

    st.markdown(
        f'<div class="geo-box">🌐 經緯度資料庫：<b>{len(geo)}</b> 所學校 ｜ '
        f'來源：{st.session_state.get("geo_file_name", "N/A")}</div>',
        unsafe_allow_html=True
    )

    school_col = detect_school_column(p1)
    if not school_col:
        st.warning("⚠️ 未偵測到學校欄位。"); st.stop()

    def make_school_map(source_df, count_label, title_label, phase_name):
        sc = detect_school_column(source_df)
        if sc is None:
            return
        agg = source_df.groupby(sc).size().reset_index(name=count_label)
        agg["_std"] = agg[sc].apply(normalize_school_name)
        agg = agg.merge(
            geo.rename(columns={"學校名稱_標準": "_std"}),
            on="_std", how="left"
        ).drop(columns=["_std"])

        fig = create_map(agg, count_label, title_label)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
            ok = agg["緯度"].notna().sum()
            st.caption(f"📊 經緯度匹配：{ok}/{len(agg)} 所學校 ({ok/len(agg)*100:.1f}%)")
            miss = agg[agg["緯度"].isna()]
            if not miss.empty:
                with st.expander(f"⚠️ {phase_name} — {len(miss)} 所學校未匹配到經緯度"):
                    st.dataframe(miss[[sc, count_label]].sort_values(count_label, ascending=False),
                                 hide_index=True)
                    st.caption("💡 請在經緯度 Excel 中補充這些學校的座標，重新上傳後按「更新分析」")
        else:
            st.warning(f"⚠️ {phase_name}無經緯度匹配結果，無法繪製地圖。")

    st.subheader("3-1. 一階報名 地理分布")
    make_school_map(p1, "報名人數", "一階報名 — 來源學校地圖", "一階")

    if p2 is not None:
        st.markdown("---")
        st.subheader("3-2. 二階報到 地理分布")
        make_school_map(p2, "報到人數", "二階報到 — 來源學校地圖", "二階")

    if p3 is not None:
        st.markdown("---")
        st.subheader("3-3. 最終入學 地理分布")
        make_school_map(p3, "入學人數", "最終入學 — 來源學校地圖", "最終入學")

        ch_col = get_final_channel_col()
        if ch_col and ch_col in p3.columns:
            st.markdown("---")
            st.subheader("3-4. 各入學管道 地理分布")
            ch_list = p3[ch_col].value_counts().index.tolist()
            sel_ch = st.selectbox("選擇管道：", ch_list, key="map_ch")
            if sel_ch:
                sub = p3[p3[ch_col] == sel_ch]
                make_school_map(sub, "入學人數",
                               f"「{sel_ch}」入學地圖", f"「{sel_ch}」")

    if p2 is not None or p3 is not None:
        st.markdown("---")
        st.subheader("3-5. 三階段對比：各校人數變化")
        sstats = build_school_stats(p1, p2, p3)
        if sstats is not None:
            sstats["_std"] = sstats["學校"].apply(normalize_school_name)
            sstats = sstats.merge(
                geo.rename(columns={"學校名稱_標準": "_std"}),
                on="_std", how="left"
            ).drop(columns=["_std"])
            map_df = sstats.dropna(subset=["緯度", "經度"])
            if not map_df.empty:
                fig = px.scatter_mapbox(
                    map_df, lat="緯度", lon="經度",
                    size="一階人數", color="一→最終(%)",
                    hover_name="學校",
                    hover_data={"一階人數": True, "二階人數": True,
                                "最終入學": True, "一→最終(%)": True},
                    color_continuous_scale="RdYlGn",
                    size_max=30, zoom=7,
                    center={"lat": HWU_COORDS["lat"], "lon": HWU_COORDS["lon"]},
                    mapbox_style="carto-positron",
                    title="三階段對比：氣泡=一階人數，顏色=最終轉換率"
                )
                fig.add_trace(go.Scattermapbox(
                    lat=[HWU_COORDS["lat"]], lon=[HWU_COORDS["lon"]],
                    mode="markers+text",
                    marker=dict(size=18, color="red", symbol="star"),
                    text=["中華醫事科技大學"], textposition="top center",
                    name="本校", showlegend=True
                ))
                fig.update_layout(height=600, margin=dict(l=0, r=0, t=40, b=0))
                st.plotly_chart(fig, use_container_width=True)


# ============================================================
# Module 4：科系熱力圖
# ============================================================
elif mod == "mod4":
    st.header("🏫 Module 4：科系 × 來源學校 熱力圖")
    p1, p2, p3 = get_phase_dfs()
    if p1 is None:
        st.warning("⚠️ 請先指定一階資料。"); st.stop()

    dept_col = detect_department_column(p1)
    school_col = detect_school_column(p1)
    if not dept_col or not school_col:
        st.warning("⚠️ 未偵測到科系或學校欄位。"); st.stop()

    mn = st.slider("學校報名人數 ≥", 1, 30, 3, key="hm")
    valid = p1[school_col].value_counts()
    valid = valid[valid >= mn].index
    filt = p1[p1[school_col].isin(valid)]

    st.subheader("4-1. 一階報名 人數熱力圖")
    cr = filt.groupby([school_col, dept_col]).size().reset_index(name="人數")
    st.plotly_chart(create_heatmap(cr, dept_col, school_col, "人數",
                    f"科系×學校 報名人數（≥{mn}人）"), use_container_width=True)

    if p3 is not None:
        st.subheader("4-2. 最終入學 人數熱力圖")
        dc3 = detect_department_column(p3)
        sc3 = detect_school_column(p3)
        if dc3 and sc3:
            cr3 = p3.groupby([sc3, dc3]).size().reset_index(name="入學人數")
            cr3 = cr3[cr3[sc3].isin(valid)]
            st.plotly_chart(create_heatmap(cr3, dc3, sc3, "入學人數",
                            "科系×學校 最終入學人數"), use_container_width=True)

        st.subheader("4-3. 一→最終 轉換率 熱力圖")
        if dc3 and sc3:
            p1c = p1.groupby([school_col, dept_col]).size().reset_index(name="一階")
            p3c = p3.groupby([sc3, dc3]).size().reset_index(name="最終")
            p3c.columns = [school_col, dept_col, "最終"]
            rc = p1c.merge(p3c, on=[school_col, dept_col], how="left").fillna(0)
            rc["轉換率"] = (rc["最終"] / rc["一階"] * 100).round(1)
            rc = rc[rc[school_col].isin(valid)]
            st.plotly_chart(create_heatmap(rc, dept_col, school_col, "轉換率",
                            "一→最終 轉換率 (%)"), use_container_width=True)

        ch_col = get_final_channel_col()
        if ch_col and ch_col in p3.columns and dc3 and sc3:
            st.markdown("---")
            st.subheader("4-4. 指定管道 — 科系×學校")
            ch_list = p3[ch_col].value_counts().index.tolist()
            sel_ch = st.selectbox("選擇管道：", ch_list, key="hm_ch")
            if sel_ch:
                sub = p3[p3[ch_col] == sel_ch]
                cr_ch = sub.groupby([sc3, dc3]).size().reset_index(name="人數")
                if not cr_ch.empty:
                    st.plotly_chart(
                        create_heatmap(cr_ch, dc3, sc3, "人數",
                                      f"「{sel_ch}」科系×學校"),
                        use_container_width=True
                    )

    elif p2 is not None:
        st.subheader("4-2. 一→二階 轉換率 熱力圖")
        dc2 = detect_department_column(p2)
        sc2 = detect_school_column(p2)
        if dc2 and sc2:
            p1c = p1.groupby([school_col, dept_col]).size().reset_index(name="一階")
            p2c = p2.groupby([sc2, dc2]).size().reset_index(name="二階")
            p2c.columns = [school_col, dept_col, "二階"]
            rc = p1c.merge(p2c, on=[school_col, dept_col], how="left").fillna(0)
            rc["轉換率"] = (rc["二階"] / rc["一階"] * 100).round(1)
            rc = rc[rc[school_col].isin(valid)]
            st.plotly_chart(create_heatmap(rc, dept_col, school_col, "轉換率",
                            "一→二階 轉換率 (%)"), use_container_width=True)


# ============================================================
# Module 5：來源學校追蹤
# ============================================================
elif mod == "mod5":
    st.header("🎯 Module 5：來源學校精準追蹤")
    p1, p2, p3 = get_phase_dfs()
    if p1 is None:
        st.warning("⚠️ 請先指定一階資料。"); st.stop()

    sstats = build_school_stats(p1, p2, p3)
    if sstats is None:
        st.warning("⚠️ 未偵測到學校欄位。"); st.stop()

    def tier(n):
        if n >= 30: return "Tier 1 (≥30人)"
        elif n >= 10: return "Tier 2 (10-29人)"
        else: return "Tier 3 (<10人)"

    sstats["經營分級"] = sstats["一階人數"].apply(tier)

    st.subheader("5-1. 學校總覽")
    sel_tiers = st.multiselect(
        "篩選分級：",
        ["Tier 1 (≥30人)", "Tier 2 (10-29人)", "Tier 3 (<10人)"],
        default=["Tier 1 (≥30人)", "Tier 2 (10-29人)"]
    )
    disp = sstats[sstats["經營分級"].isin(sel_tiers)].sort_values("一階人數", ascending=False)
    st.dataframe(disp, use_container_width=True, hide_index=True)

    st.subheader("5-2. 分級彙總")
    ts = sstats.groupby("經營分級").agg(
        學校數=("學校", "count"),
        一階合計=("一階人數", "sum"),
        二階合計=("二階人數", "sum"),
        最終合計=("最終入學", "sum"),
        平均最終轉換率=("一→最終(%)", "mean")
    ).round(1).reset_index()
    st.dataframe(ts, use_container_width=True, hide_index=True)

    st.subheader("5-3. 個別學校分析")
    sel = st.selectbox("選擇學校：",
                       sstats.sort_values("一階人數", ascending=False)["學校"])
    if sel:
        r = sstats[sstats["學校"] == sel].iloc[0]
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("一階", f'{int(r["一階人數"])}人')
        c2.metric("二階", f'{int(r["二階人數"])}人')
        c3.metric("最終入學", f'{int(r["最終入學"])}人')
        c4.metric("一→最終", f'{r["一→最終(%)"]}%')
        c5.metric("經營分級", r["經營分級"])

        fl = ["一階報名"]; fv = [int(r["一階人數"])]
        if r["二階人數"] > 0 or p2 is not None:
            fl.append("二階報到"); fv.append(int(r["二階人數"]))
        if r["最終入學"] > 0 or p3 is not None:
            fl.append("最終入學"); fv.append(int(r["最終入學"]))
        if len(fv) > 1:
            st.plotly_chart(create_funnel_chart(fl, fv, f"{sel} 招生漏斗"),
                            use_container_width=True)

        if p3 is not None:
            ch_col = get_final_channel_col()
            school_col_p3 = detect_school_column(p3)
            if ch_col and ch_col in p3.columns and school_col_p3:
                sub = p3[p3[school_col_p3] == sel]
                if not sub.empty:
                    st.markdown("---")
                    st.markdown(f"**🟢 「{sel}」最終入學管道分布：**")
                    dd = sub[ch_col].value_counts().reset_index()
                    dd.columns = ["入學管道", "人數"]
                    c1, c2 = st.columns(2)
                    with c1:
                        fig = px.pie(dd, names="入學管道", values="人數",
                                     title=f"{sel} — 入學管道", hole=0.35)
                        st.plotly_chart(fig, use_container_width=True)
                    with c2:
                        st.dataframe(dd, hide_index=True)

        dept_col = detect_department_column(p1)
        school_col = detect_school_column(p1)
        if dept_col and school_col:
            dd = p1[p1[school_col] == sel][dept_col].value_counts().reset_index()
            dd.columns = ["科系", "人數"]
            fig = px.pie(dd, names="科系", values="人數",
                         title=f"{sel} — 報名科系分布", hole=0.35)
            st.plotly_chart(fig, use_container_width=True)


# ============================================================
# Module 6：流失預警分析
# ============================================================
elif mod == "mod6":
    st.header("⚠️ Module 6：流失預警分析")
    p1, p2, p3 = get_phase_dfs()

    if p1 is None:
        st.warning("⚠️ 請先指定一階資料。"); st.stop()
    if p2 is None and p3 is None:
        st.warning("⚠️ 此模組需要至少二階或最終入學資料。"); st.stop()

    sstats = build_school_stats(p1, p2, p3)
    if sstats is None:
        st.warning("⚠️ 未偵測到學校欄位。"); st.stop()

    has_final = p3 is not None
    rate_col = "一→最終(%)" if has_final else "一→二階(%)"
    loss_label = "最終入學" if has_final else "二階人數"
    stage_label = "一→最終" if has_final else "一→二階"
    sstats["流失人數"] = sstats["一階人數"] - sstats[loss_label]

    st.subheader("6-1. 高報名低轉換 預警學校")
    mn = st.slider("一階報名 ≥", 1, 50, 10, key="m6")
    pool = sstats[sstats["一階人數"] >= mn]
    avg = pool[rate_col].mean()
    warn = pool[pool[rate_col] < avg].sort_values("流失人數", ascending=False)

    if warn.empty:
        st.success("✅ 沒有預警學校！")
    else:
        st.markdown(
            f'<div class="warning-box">⚠️ 以下學校一階≥{mn}人，但 {stage_label} '
            f'轉換率低於平均 ({avg:.1f}%)</div>', unsafe_allow_html=True
        )
        st.dataframe(warn, use_container_width=True, hide_index=True)

    st.subheader("6-2. IPA 四象限矩陣")
    ana = sstats[sstats["一階人數"] >= mn].copy()
    if ana.empty:
        st.info("篩選後無資料"); st.stop()

    med_x = ana["一階人數"].median()
    med_y = ana[rate_col].median()
    fig = px.scatter(
        ana, x="一階人數", y=rate_col,
        size="流失人數", hover_name="學校",
        hover_data={"一階人數": True, "二階人數": True,
                    "最終入學": True, rate_col: True},
        title=f"IPA 矩陣（氣泡=流失人數，基準={stage_label}）",
        size_max=40, color=rate_col, color_continuous_scale="RdYlGn"
    )
    fig.add_hline(y=med_y, line_dash="dash", line_color="red",
                  annotation_text=f"轉換率中位數 {med_y:.1f}%")
    fig.add_vline(x=med_x, line_dash="dash", line_color="blue",
                  annotation_text=f"報名中位數 {med_x:.0f}人")
    x_max = ana["一階人數"].max(); x_min = ana["一階人數"].min()
    y_max = ana[rate_col].max(); y_min = ana[rate_col].min()
    fig.add_annotation(x=x_max*0.85, y=y_max*0.95,
                       text="🌟 高量能高效率<br>持續維護", showarrow=False,
                       font=dict(size=11, color="green"))
    fig.add_annotation(x=x_max*0.85, y=y_min+(y_max-y_min)*0.05,
                       text="⚠️ 高量能低效率<br>重點改善", showarrow=False,
                       font=dict(size=11, color="red"))
    fig.add_annotation(x=x_min+(x_max-x_min)*0.05, y=y_max*0.95,
                       text="📈 低量能高效率<br>擴大招生", showarrow=False,
                       font=dict(size=11, color="blue"))
    fig.add_annotation(x=x_min+(x_max-x_min)*0.05, y=y_min+(y_max-y_min)*0.05,
                       text="🔍 低量能低效率<br>評估投入", showarrow=False,
                       font=dict(size=11, color="gray"))
    fig.update_layout(height=600)
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("6-3. 流失人數排行 TOP 15")
    top15 = sstats.sort_values("流失人數", ascending=False).head(15)
    fig = px.bar(
        top15, x="流失人數", y="學校", orientation="h",
        text="流失人數", color=rate_col,
        color_continuous_scale="RdYlGn",
        title=f"流失人數排行（{stage_label}）"
    )
    fig.update_layout(yaxis=dict(autorange="reversed"), height=500)
    st.plotly_chart(fig, use_container_width=True)

    if p3 is not None:
        ch_col = get_final_channel_col()
        if ch_col and ch_col in p3.columns:
            st.markdown("---")
            st.subheader("6-4. 各入學管道貢獻度 Pareto 分析")
            ch_summary = p3[ch_col].value_counts().reset_index()
            ch_summary.columns = ["入學管道", "最終入學人數"]
            ch_summary["佔比(%)"] = (ch_summary["最終入學人數"] / ch_summary["最終入學人數"].sum() * 100).round(1)
            ch_summary["累積佔比(%)"] = ch_summary["佔比(%)"].cumsum().round(1)

            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=ch_summary["入學管道"], y=ch_summary["最終入學人數"],
                name="入學人數", marker_color="#4CAF50",
                text=ch_summary["最終入學人數"], textposition="outside"
            ))
            fig.add_trace(go.Scatter(
                x=ch_summary["入學管道"], y=ch_summary["累積佔比(%)"],
                name="累積佔比 (%)", yaxis="y2",
                line=dict(color="#FF5722", width=3),
                marker=dict(size=8)
            ))
            fig.update_layout(
                title="入學管道 Pareto 圖",
                yaxis=dict(title="入學人數"),
                yaxis2=dict(title="累積佔比 (%)", overlaying="y", side="right",
                           range=[0, 105]),
                height=500, xaxis_tickangle=-45
            )
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(ch_summary, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("6-5. 科系維度流失分析")
    dstats = build_dept_stats(p1, p2, p3)
    if dstats is not None:
        d_rate_col = "一→最終(%)" if has_final else "一→二階(%)"
        d_loss_label = "最終入學" if has_final else "二階人數"
        dstats["流失人數"] = dstats["一階人數"] - dstats[d_loss_label]
        fig = px.scatter(
            dstats, x="一階人數", y=d_rate_col,
            size="流失人數", hover_name="科系",
            text="科系", title=f"科系 IPA（{stage_label}）",
            size_max=50, color=d_rate_col, color_continuous_scale="RdYlGn"
        )
        fig.update_traces(textposition="top center")
        fig.add_hline(y=dstats[d_rate_col].median(), line_dash="dash", line_color="red")
        fig.add_vline(x=dstats["一階人數"].median(), line_dash="dash", line_color="blue")
        fig.update_layout(height=500)
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(dstats.sort_values("流失人數", ascending=False),
                     use_container_width=True, hide_index=True)


# ============================================================
# Footer
# ============================================================
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
geo_n = len(st.session_state["geo_lookup"]) if st.session_state.get("geo_lookup") is not None else 0
st.markdown(f"""
<div style="text-align: center; color: #aaa; font-size: 0.85rem; padding: 10px;">
    🎓 中華醫事科技大學 招生數據分析系統 v5.4<br>
    三階段：一階報名 → 二階報到 → 最終入學 ｜ 轉換率統一以一階為分母<br>
    🌐 經緯度資料庫：Excel 匯入（{geo_n} 所學校）<br>
    🟢 最終入學管道由資料「入學方式」欄位自動讀取<br>
    分析版本 #{st.session_state.get('analysis_version', 0)}<br>
    Built with Streamlit + Plotly ｜ © 2024 HWU Admissions Office
</div>
""", unsafe_allow_html=True)
