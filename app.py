# -*- coding: utf-8 -*-
"""
中華醫事科技大學 招生數據分析系統 v5.2
- 三階段：一階(報名) → 二階(報到確認) → 最終入學(註冊)
- 統一轉換率分母 = 一階人數
- 二階＋最終入學經緯度統一從一階讀取
- 所有資料皆需管道確認
- 「更新分析」按鈕觸發統計重建
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
    .metric-green { background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%); }
    .metric-orange { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); }
    .metric-blue { background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); }
    .metric-gold { background: linear-gradient(135deg, #f6d365 0%, #fda085 100%); }
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
    .update-btn-area {
        background: linear-gradient(135deg, #667eea22, #764ba222);
        border: 2px dashed #667eea; border-radius: 12px;
        padding: 15px; margin: 15px 0; text-align: center;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🎓 中華醫事科技大學 招生數據分析系統</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Enrollment Analytics Platform v5.2 ｜ 一階報名 → 二階報到 → 最終入學 三階段完整分析</div>', unsafe_allow_html=True)
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
                  "就讀科系", "科系名稱", "系科", "學系", "錄取系所"]
    for c in candidates:
        for col in df.columns:
            if c in str(col):
                return col
    return None


def detect_lat_lon_columns(df):
    lat_kw = ["緯度", "lat", "latitude", "Lat", "LAT"]
    lon_kw = ["經度", "lon", "lng", "longitude", "Lon", "LON"]
    lat_col = lon_col = None
    for col in df.columns:
        s = str(col).strip()
        if lat_col is None:
            for kw in lat_kw:
                if kw in s:
                    lat_col = col; break
        if lon_col is None:
            for kw in lon_kw:
                if kw in s:
                    lon_col = col; break
    return lat_col, lon_col


def detect_channel_from_filename(filename):
    if not filename:
        return None
    for ch, kws in CHANNEL_KEYWORDS.items():
        for kw in kws:
            if kw in filename:
                return ch
    return None


def detect_channel_from_columns(df):
    for col in ["管道", "入學管道", "招生管道", "報名管道"]:
        if col in df.columns:
            top = df[col].dropna().value_counts()
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
# geo_lookup 建構與合併
# ============================================================
def build_school_geo_lookup(phase1_df):
    school_col = detect_school_column(phase1_df)
    lat_col, lon_col = detect_lat_lon_columns(phase1_df)
    if school_col is None:
        return pd.DataFrame(columns=["學校名稱_標準", "緯度", "經度"])
    if lat_col and lon_col:
        geo = phase1_df[[school_col, lat_col, lon_col]].copy()
        geo.columns = ["原始", "緯度", "經度"]
        geo["緯度"] = pd.to_numeric(geo["緯度"], errors="coerce")
        geo["經度"] = pd.to_numeric(geo["經度"], errors="coerce")
        geo = geo.dropna(subset=["緯度", "經度"])
        geo["學校名稱_標準"] = geo["原始"].apply(normalize_school_name)
        lookup = geo.groupby("學校名稱_標準").agg(
            緯度=("緯度", "mean"), 經度=("經度", "mean")
        ).reset_index()
        return lookup
    else:
        schools = phase1_df[school_col].dropna().unique()
        return pd.DataFrame({
            "學校名稱_標準": [normalize_school_name(s) for s in schools],
            "緯度": np.nan, "經度": np.nan
        }).drop_duplicates(subset=["學校名稱_標準"])


def enrich_with_geo(df, geo_lookup):
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
if "analysis_ready" not in st.session_state:
    st.session_state["analysis_ready"] = False
if "analysis_version" not in st.session_state:
    st.session_state["analysis_version"] = 0

# ============================================================
# Sidebar
# ============================================================
with st.sidebar:
    st.header("📂 資料上傳")

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
                    st.session_state["datasets"][uf.name] = {
                        "df": df,
                        "channel": auto_ch,
                        "channel_confirmed": auto_ch or "",
                        "filename": uf.name
                    }
                    st.success(f"✅ {uf.name}（{len(df)} 筆）")
                except Exception as e:
                    st.error(f"❌ {uf.name}: {e}")

    st.markdown("---")

    # 已載入清單
    if st.session_state["datasets"]:
        st.subheader("📋 已載入資料")
        for name, info in st.session_state["datasets"].items():
            ch = info.get("channel_confirmed") or info.get("channel") or "待確認"
            st.caption(f"📄 **{name}**　{len(info['df'])} 筆　管道：{ch}")

        if st.button("🗑️ 清除全部資料"):
            st.session_state["datasets"] = {}
            st.session_state["geo_lookup"] = None
            st.session_state["analysis_ready"] = False
            st.session_state["analysis_version"] = 0
            st.rerun()

    # ── 三階段指定 ──
    st.markdown("---")
    st.subheader("⚙️ 三階段資料指定")
    st.markdown("""
    <div style="font-size:0.8rem; color:#888; margin-bottom:10px;">
    📍 一階 = 經緯度唯一來源<br>
    📊 所有轉換率分母 = 一階人數
    </div>
    """, unsafe_allow_html=True)

    ds_names = list(st.session_state["datasets"].keys())
    none_opt = ["-- 未選擇 --"]

    p1_sel = st.selectbox("🔵 一階（報名）", none_opt + ds_names, key="p1_sel")
    p2_sel = st.selectbox("🟠 二階（報到確認）", none_opt + ds_names, key="p2_sel")
    p3_sel = st.selectbox("🟢 最終入學（註冊）", none_opt + ds_names, key="p3_sel")

    # ── 各檔管道確認 ──
    st.markdown("---")
    st.subheader("🏷️ 管道確認（各檔獨立）")

    for name, info in st.session_state["datasets"].items():
        # 判斷此檔被指定為哪個階段
        phase_tag = ""
        if name == st.session_state.get("p1_sel"):
            phase_tag = " 🔵一階"
        elif name == st.session_state.get("p2_sel"):
            phase_tag = " 🟠二階"
        elif name == st.session_state.get("p3_sel"):
            phase_tag = " 🟢最終入學"

        with st.expander(f"📄 {name}{phase_tag}"):
            auto = info.get("channel")
            if auto:
                st.info(f"自動偵測：**{auto}**")
                keep = st.radio(
                    "使用偵測結果？", ["✅ 是", "❌ 自行選擇"],
                    horizontal=True, key=f"chk_{name}"
                )
                if keep == "✅ 是":
                    confirmed = auto
                else:
                    confirmed = st.selectbox(
                        "選擇管道：", KNOWN_CHANNELS, key=f"chs_{name}"
                    )
            else:
                st.warning("⚠️ 無法自動偵測管道")
                confirmed = st.selectbox(
                    "選擇管道：", KNOWN_CHANNELS, key=f"chs_{name}"
                )
            st.session_state["datasets"][name]["channel_confirmed"] = confirmed

    # ── 更新分析按鈕 ──
    st.markdown("---")
    st.markdown('<div class="update-btn-area">', unsafe_allow_html=True)
    st.markdown("⬇️ 設定完成後請點擊下方按鈕")

    update_clicked = st.button(
        "🔄 更新分析",
        type="primary",
        use_container_width=True,
        help="重新建構經緯度對照表並刷新所有統計"
    )

    if update_clicked:
        # 重建 geo_lookup
        if p1_sel != "-- 未選擇 --":
            p1_raw = st.session_state["datasets"][p1_sel]["df"]
            geo_lookup = build_school_geo_lookup(p1_raw)
            st.session_state["geo_lookup"] = geo_lookup
            geo_ok = geo_lookup["緯度"].notna().sum() if not geo_lookup.empty else 0
            if geo_ok > 0:
                st.success(f"📍 經緯度對照：{geo_ok} 所學校")
            else:
                st.warning("⚠️ 一階資料無經緯度欄位")
        else:
            st.session_state["geo_lookup"] = None

        st.session_state["analysis_ready"] = True
        st.session_state["analysis_version"] += 1
        st.success(f"✅ 分析已更新！（版本 #{st.session_state['analysis_version']}）")

    st.markdown('</div>', unsafe_allow_html=True)

    # 狀態顯示
    st.markdown("---")
    if st.session_state["analysis_ready"]:
        v = st.session_state["analysis_version"]
        st.markdown(f'<div class="success-box">✅ 分析就緒　版本 #{v}</div>',
                    unsafe_allow_html=True)
        # 摘要
        items = []
        if p1_sel != "-- 未選擇 --":
            ch = st.session_state["datasets"][p1_sel].get("channel_confirmed", "")
            items.append(f"🔵 一階：{p1_sel}（{ch}）")
        if p2_sel != "-- 未選擇 --":
            ch = st.session_state["datasets"][p2_sel].get("channel_confirmed", "")
            items.append(f"🟠 二階：{p2_sel}（{ch}）")
        if p3_sel != "-- 未選擇 --":
            ch = st.session_state["datasets"][p3_sel].get("channel_confirmed", "")
            items.append(f"🟢 最終：{p3_sel}（{ch}）")
        for it in items:
            st.caption(it)
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
        p1 = st.session_state["datasets"][sel1]["df"].copy()
    if sel2 != "-- 未選擇 --" and sel2 in st.session_state["datasets"]:
        raw = st.session_state["datasets"][sel2]["df"].copy()
        p2 = enrich_with_geo(raw, geo) if geo is not None else raw
    if sel3 != "-- 未選擇 --" and sel3 in st.session_state["datasets"]:
        raw = st.session_state["datasets"][sel3]["df"].copy()
        p3 = enrich_with_geo(raw, geo) if geo is not None else raw
    return p1, p2, p3


# ============================================================
# 主功能選單
# ============================================================
modules = {
    "📊 Module 0：總覽儀表板":      "mod0",
    "🔄 Module 1：招生漏斗分析":     "mod1",
    "📈 Module 2：管道比較分析":      "mod2",
    "🗺️ Module 3：地理分布（地圖）": "mod3",
    "🏫 Module 4：科系熱力圖":       "mod4",
    "🎯 Module 5：來源學校追蹤":     "mod5",
    "⚠️ Module 6：流失預警分析":     "mod6",
}

selected = st.selectbox("選擇分析模組：", list(modules.keys()))
mod = modules[selected]
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

# ── 未更新提示 ──
if not st.session_state["analysis_ready"]:
    st.markdown("""
    <div class="warning-box">
        <h4>⏳ 尚未執行分析</h4>
        <p>請先在左側：</p>
        <ol>
            <li>上傳資料檔案</li>
            <li>指定三階段（一階/二階/最終入學）</li>
            <li>確認各檔管道</li>
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
        st.warning("⚠️ 請在側邊欄指定【一階（報名）】資料並按「更新分析」。")
        st.stop()

    n1 = len(p1)
    n2 = len(p2) if p2 is not None else None
    n3 = len(p3) if p3 is not None else None
    school_col = detect_school_column(p1)
    dept_col = detect_department_column(p1)

    # KPI
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

    # 管道資訊
    st.markdown("---")
    info_parts = []
    for label, sel_key in [("一階", "p1_sel"), ("二階", "p2_sel"), ("最終入學", "p3_sel")]:
        sel = st.session_state.get(sel_key, "-- 未選擇 --")
        if sel != "-- 未選擇 --" and sel in st.session_state["datasets"]:
            ch = st.session_state["datasets"][sel].get("channel_confirmed", "未確認")
            info_parts.append(f"**{label}**：{ch}")
    if info_parts:
        st.markdown(f'<div class="info-box">📌 管道設定：{"　｜　".join(info_parts)}</div>',
                    unsafe_allow_html=True)

    st.markdown("---")

    # 漏斗
    f_labels, f_values = ["一階 報名"], [n1]
    if n2 is not None:
        f_labels.append("二階 報到確認"); f_values.append(n2)
    if n3 is not None:
        f_labels.append("最終入學"); f_values.append(n3)
    if len(f_values) > 1:
        st.plotly_chart(create_funnel_chart(f_labels, f_values, "整體招生漏斗"),
                        use_container_width=True)

    c1, c2 = st.columns(2)
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
        st.warning("⚠️ 請先指定一階資料並按「更新分析」。"); st.stop()

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
    else:
        st.warning("⚠️ 未偵測到科系欄位。")

    st.markdown("---")
    st.subheader("1-3. 各來源學校三階段漏斗")
    sstats = build_school_stats(p1, p2, p3)
    if sstats is not None:
        mn = st.slider("一階報名 ≥", 1, 50, 5, key="m1_sch")
        sf = sstats[sstats["一階人數"] >= mn].sort_values("一→最終(%)", ascending=False)
        st.dataframe(sf, use_container_width=True, hide_index=True)
        fig = create_bar_h(sf.head(20), "學校", "一→最終(%)",
                           f"來源學校 一→最終 轉換率 TOP 20（一階≥{mn}人）",
                           color="#4CAF50")
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# Module 2：管道比較分析
# ============================================================
elif mod == "mod2":
    st.header("📈 Module 2：管道比較分析")

    if not st.session_state["datasets"]:
        st.warning("⚠️ 請先上傳資料。"); st.stop()

    frames = []
    for name, info in st.session_state["datasets"].items():
        df = info["df"].copy()
        df["_管道"] = info.get("channel_confirmed") or info.get("channel") or "未分類"
        df["_來源"] = name
        # 標記階段
        phase = "其他"
        if name == st.session_state.get("p1_sel"):
            phase = "一階"
        elif name == st.session_state.get("p2_sel"):
            phase = "二階"
        elif name == st.session_state.get("p3_sel"):
            phase = "最終入學"
        df["_階段"] = phase
        frames.append(df)
    combined = pd.concat(frames, ignore_index=True)

    st.subheader("2-1. 各管道 × 各階段人數")
    cd = combined.groupby(["_管道", "_階段"]).size().reset_index(name="人數")
    fig = px.bar(cd, x="_管道", y="人數", color="_階段",
                 barmode="group", text="人數",
                 title="各管道各階段資料筆數",
                 color_discrete_map={"一階": "#2196F3", "二階": "#FF9800",
                                     "最終入學": "#4CAF50", "其他": "#9E9E9E"})
    st.plotly_chart(fig, use_container_width=True)

    dept_col = detect_department_column(combined)
    if dept_col:
        st.subheader("2-2. 管道 × 科系")
        cross = combined.groupby(["_管道", dept_col]).size().reset_index(name="人數")
        fig = px.bar(cross, x="_管道", y="人數", color=dept_col,
                     barmode="group", title="管道 × 科系分布")
        st.plotly_chart(fig, use_container_width=True)

    school_col = detect_school_column(combined)
    if school_col:
        st.subheader("2-3. 管道來源學校多元性")
        div = combined.groupby("_管道")[school_col].nunique().reset_index()
        div.columns = ["管道", "來源學校數"]
        fig = px.bar(div, x="管道", y="來源學校數", text="來源學校數",
                     color="管道", title="各管道來源學校多元性")
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# Module 3：地理分布（地圖）
# ============================================================
elif mod == "mod3":
    st.header("🗺️ Module 3：地理分布（地圖）")
    p1, p2, p3 = get_phase_dfs()
    geo = st.session_state.get("geo_lookup")

    if p1 is None:
        st.warning("⚠️ 請先指定一階資料並按「更新分析」。"); st.stop()

    school_col = detect_school_column(p1)
    if not school_col:
        st.warning("⚠️ 未偵測到學校欄位。"); st.stop()

    def make_school_map(source_df, count_label, title_label, phase_name):
        sc = detect_school_column(source_df)
        if sc is None:
            return
        enriched = enrich_with_geo(source_df, geo) if geo is not None else source_df
        agg = enriched.groupby(sc).size().reset_index(name=count_label)
        if geo is not None and not geo.empty:
            agg["_std"] = agg[sc].apply(normalize_school_name)
            agg = agg.merge(
                geo.rename(columns={"學校名稱_標準": "_std"}),
                on="_std", how="left"
            ).drop(columns=["_std"])
        fig = create_map(agg, count_label, title_label)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
            if "緯度" in agg.columns:
                ok = agg["緯度"].notna().sum()
                st.caption(f"📊 經緯度匹配：{ok}/{len(agg)} 所學校")
                miss = agg[agg["緯度"].isna()]
                if not miss.empty:
                    with st.expander(f"⚠️ {phase_name}中 {len(miss)} 所學校無經緯度"):
                        st.dataframe(miss[[sc, count_label]], hide_index=True)
        else:
            st.warning(f"⚠️ {phase_name}無法繪製地圖（缺經緯度）。")

    st.subheader("3-1. 一階報名 地理分布")
    make_school_map(p1, "報名人數", "一階報名 — 來源學校地圖", "一階")

    if p2 is not None:
        st.markdown("---")
        st.subheader("3-2. 二階報到 地理分布")
        st.markdown('<div class="info-box">📍 經緯度來源：一階資料</div>',
                    unsafe_allow_html=True)
        make_school_map(p2, "報到人數", "二階報到 — 來源學校地圖（經緯度←一階）", "二階")

    if p3 is not None:
        st.markdown("---")
        st.subheader("3-3. 最終入學 地理分布")
        st.markdown('<div class="info-box">📍 經緯度來源：一階資料</div>',
                    unsafe_allow_html=True)
        make_school_map(p3, "入學人數", "最終入學 — 來源學校地圖（經緯度←一階）", "最終入學")

    if p2 is not None or p3 is not None:
        st.markdown("---")
        st.subheader("3-4. 三階段對比：各校人數變化")
        sstats = build_school_stats(p1, p2, p3)
        if sstats is not None and geo is not None:
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
        st.warning("⚠️ 請先指定一階資料並按「更新分析」。"); st.stop()

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

    if p3 is not None:
        st.subheader("4-3. 一→最終 轉換率 熱力圖")
        dc3 = detect_department_column(p3)
        sc3 = detect_school_column(p3)
        if dc3 and sc3:
            p1c = p1.groupby([school_col, dept_col]).size().reset_index(name="一階")
            p3c = p3.groupby([sc3, dc3]).size().reset_index(name="最終")
            p3c.columns = [school_col, dept_col, "最終"]
            rc = p1c.merge(p3c, on=[school_col, dept_col], how="left").fillna(0)
            rc["轉換率"] = (rc["最終"] / rc["一階"] * 100).round(1)
            rc = rc[rc[school_col].isin(valid)]
            st.plotly_chart(create_heatmap(rc, dept_col, school_col, "轉換率",
                            "一→最終 轉換率 (%)"), use_container_width=True)
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
        st.warning("⚠️ 請先指定一階資料並按「更新分析」。"); st.stop()

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
        st.warning("⚠️ 請先指定一階資料並按「更新分析」。"); st.stop()
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

    st.markdown("---")
    st.subheader("6-4. 科系維度流失分析")
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
st.markdown(f"""
<div style="text-align: center; color: #aaa; font-size: 0.85rem; padding: 10px;">
    🎓 中華醫事科技大學 招生數據分析系統 v5.2<br>
    三階段：一階報名 → 二階報到 → 最終入學 ｜ 轉換率統一以一階為分母<br>
    📍 經緯度統一由一階提供 ｜ 分析版本 #{st.session_state.get('analysis_version', 0)}<br>
    Built with Streamlit + Plotly ｜ © 2024 HWU Admissions Office
</div>
""", unsafe_allow_html=True)
