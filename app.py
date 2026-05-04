# -*- coding: utf-8 -*-
"""
中華醫事科技大學 招生數據分析系統 v5.0
更新重點：
- 二階報到資料的畢業學校經緯度統一從一階報名資料讀取
- 統一轉換率分母為一階人數
- Module 0-7 完整功能
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import json, re, os
from collections import Counter

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
# 自訂 CSS
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
    .metric-card h3 { margin: 0; font-size: 0.9rem; opacity: 0.9; }
    .metric-card h1 { margin: 5px 0 0 0; font-size: 2.0rem; }
    .channel-badge {
        display: inline-block; padding: 5px 15px; border-radius: 20px;
        font-weight: bold; font-size: 0.85rem; margin: 2px;
    }
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
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🎓 中華醫事科技大學 招生數據分析系統</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Chung Hwa University of Medical Technology — Enrollment Analytics Platform v5.0</div>', unsafe_allow_html=True)
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

# ============================================================
# 常數與輔助函式
# ============================================================

KNOWN_CHANNELS = [
    "聯合免試", "甄選入學", "技優甄審", "運動績優",
    "身障甄試", "單獨招生", "進修部", "產學攜手",
    "運動單獨招生", "四技申請", "繁星計畫",
    "設有專班", "外國學生", "陸生招生", "僑生招生"
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

DEPARTMENT_ALIASES = {
    "護理": "護理系", "護理系": "護理系",
    "醫技": "醫學檢驗生物技術系", "醫檢": "醫學檢驗生物技術系",
    "醫學檢驗": "醫學檢驗生物技術系", "醫學檢驗生物技術系": "醫學檢驗生物技術系",
    "藥學": "藥學系", "藥學系": "藥學系",
    "食營": "食品營養系", "食品營養": "食品營養系", "食品營養系": "食品營養系",
    "化妝品": "化妝品應用與管理系", "化妝品應用": "化妝品應用與管理系",
    "職安": "職業安全衛生系", "職業安全": "職業安全衛生系",
    "長照": "長期照顧系", "長期照顧": "長期照顧系",
    "幼保": "幼兒保育系", "幼兒保育": "幼兒保育系",
    "寵美": "寵物美容學位學程", "寵物美容": "寵物美容學位學程",
    "資管": "資訊管理系", "資訊管理": "資訊管理系",
    "醫管": "醫務暨健康事業管理系", "醫務管理": "醫務暨健康事業管理系",
    "視光": "視光系", "視光系": "視光系",
}

# 台灣主要城市座標（作為縣市層級備用）
TAIWAN_CITY_COORDS = {
    "台北市": (25.0330, 121.5654), "新北市": (25.0120, 121.4650),
    "桃園市": (24.9936, 121.3010), "台中市": (24.1477, 120.6736),
    "台南市": (22.9998, 120.2270), "高雄市": (22.6273, 120.3014),
    "基隆市": (25.1276, 121.7392), "新竹市": (24.8138, 120.9675),
    "新竹縣": (24.8390, 121.0042), "苗栗縣": (24.5602, 120.8214),
    "彰化縣": (24.0518, 120.5161), "南投縣": (23.9610, 120.6718),
    "雲林縣": (23.7092, 120.4313), "嘉義市": (23.4801, 120.4491),
    "嘉義縣": (23.4518, 120.2551), "屏東縣": (22.5519, 120.5487),
    "宜蘭縣": (24.7570, 121.7533), "花蓮縣": (23.9871, 121.6016),
    "台東縣": (22.7583, 121.1444), "澎湖縣": (23.5711, 119.5793),
    "金門縣": (24.4493, 118.3767), "連江縣": (26.1505, 119.9499),
    "臺北市": (25.0330, 121.5654), "臺中市": (24.1477, 120.6736),
    "臺南市": (22.9998, 120.2270), "臺東縣": (22.7583, 121.1444),
}

HWU_COORDS = {"lat": 22.9340, "lon": 120.2756}


def detect_channel_from_filename(filename: str) -> str | None:
    """從檔案名稱偵測招生管道"""
    if not filename:
        return None
    for channel, keywords in CHANNEL_KEYWORDS.items():
        for kw in keywords:
            if kw in filename:
                return channel
    return None


def detect_channel_from_columns(df: pd.DataFrame) -> str | None:
    """從資料欄位內容偵測招生管道"""
    channel_col_candidates = ["管道", "入學管道", "招生管道", "報名管道", "升學管道"]
    for col in channel_col_candidates:
        if col in df.columns:
            top_val = df[col].dropna().value_counts()
            if not top_val.empty:
                val = str(top_val.index[0])
                for channel, keywords in CHANNEL_KEYWORDS.items():
                    for kw in keywords:
                        if kw in val:
                            return channel
                return val
    return None


def resolve_channel(auto_detected: str | None, uploaded_filename: str) -> str:
    """整合自動偵測與使用者確認，回傳最終管道名稱"""
    filename_channel = detect_channel_from_filename(uploaded_filename)
    detected = auto_detected or filename_channel
    if detected:
        st.info(f"🔍 系統自動偵測到招生管道：**{detected}**")
        confirm = st.radio("請確認此招生管道是否正確：", ["✅ 正確", "❌ 我要自己選"], horizontal=True, key=f"ch_confirm_{uploaded_filename}")
        if confirm == "✅ 正確":
            return detected
    return st.selectbox("請選擇此資料的招生管道：", KNOWN_CHANNELS, key=f"ch_select_{uploaded_filename}")


def normalize_school_name(name: str) -> str:
    """標準化學校名稱"""
    if not isinstance(name, str):
        return str(name)
    name = name.strip()
    name = name.replace("臺", "台").replace("（", "(").replace("）", ")")
    name = re.sub(r"\s+", "", name)
    for suffix in ["附設進修學校", "進修學校", "進修部"]:
        name = name.replace(suffix, "")
    return name


def detect_phase_columns(df: pd.DataFrame) -> dict:
    """
    偵測資料中的階段欄位。
    回傳如 { 'phase1': '報名人數', 'phase2': '到考人數', 'final': '報到人數' }
    """
    mapping = {"phase1": None, "phase2": None, "final": None}
    phase1_kw = ["報名", "登記", "一階", "第一階段", "Phase1", "phase1", "填志願", "通過"]
    phase2_kw = ["到考", "面試", "二階", "第二階段", "Phase2", "phase2", "甄試", "考試", "複試"]
    final_kw  = ["報到", "註冊", "錄取確認", "最終", "入學", "Final", "final", "就讀"]
    for col in df.columns:
        col_str = str(col)
        for kw in phase1_kw:
            if kw in col_str and mapping["phase1"] is None:
                mapping["phase1"] = col
        for kw in phase2_kw:
            if kw in col_str and mapping["phase2"] is None:
                mapping["phase2"] = col
        for kw in final_kw:
            if kw in col_str and mapping["final"] is None:
                mapping["final"] = col
    return mapping


def detect_school_column(df: pd.DataFrame) -> str | None:
    """偵測畢業學校欄位"""
    candidates = ["畢業學校", "來源學校", "原就讀學校", "高中職校名",
                  "學校名稱", "校名", "畢業高中職", "畢業學校名稱",
                  "就讀學校", "原學校"]
    for c in candidates:
        for col in df.columns:
            if c in str(col):
                return col
    return None


def detect_department_column(df: pd.DataFrame) -> str | None:
    """偵測報名科系欄位"""
    candidates = ["科系", "系所", "報名科系", "錄取科系", "志願科系",
                  "就讀科系", "科系名稱", "系科", "學系"]
    for c in candidates:
        for col in df.columns:
            if c in str(col):
                return col
    return None


def detect_lat_lon_columns(df: pd.DataFrame) -> tuple:
    """偵測經緯度欄位"""
    lat_candidates = ["緯度", "lat", "latitude", "Lat", "LAT", "Latitude"]
    lon_candidates = ["經度", "lon", "lng", "longitude", "Lon", "LON", "Longitude"]
    lat_col, lon_col = None, None
    for col in df.columns:
        col_str = str(col).strip()
        if lat_col is None:
            for kw in lat_candidates:
                if kw in col_str:
                    lat_col = col
                    break
        if lon_col is None:
            for kw in lon_candidates:
                if kw in col_str:
                    lon_col = col
                    break
    return lat_col, lon_col


def detect_name_id_column(df: pd.DataFrame) -> str | None:
    """偵測學生姓名或編號欄位（用於一二階串聯）"""
    candidates = ["准考證號", "准考證", "學號", "報名序號", "編號",
                  "考生編號", "序號", "身分證字號", "ID",
                  "姓名", "考生姓名", "學生姓名", "name"]
    for c in candidates:
        for col in df.columns:
            if c in str(col):
                return col
    return None


# ============================================================
# 核心：建構一階學校經緯度對照表
# ============================================================
def build_school_geo_lookup(phase1_df: pd.DataFrame) -> pd.DataFrame:
    """
    從一階（報名）資料中，建構 畢業學校 → (緯度, 經度) 的對照表。
    此對照表作為系統中所有畢業學校地理資訊的唯一來源。
    """
    school_col = detect_school_column(phase1_df)
    lat_col, lon_col = detect_lat_lon_columns(phase1_df)

    if school_col is None:
        return pd.DataFrame(columns=["學校名稱_標準", "緯度", "經度"])

    # 建構基礎 lookup
    if lat_col and lon_col:
        geo_df = phase1_df[[school_col, lat_col, lon_col]].copy()
        geo_df.columns = ["學校名稱_原始", "緯度", "經度"]
        geo_df["緯度"] = pd.to_numeric(geo_df["緯度"], errors="coerce")
        geo_df["經度"] = pd.to_numeric(geo_df["經度"], errors="coerce")
        geo_df = geo_df.dropna(subset=["緯度", "經度"])
        geo_df["學校名稱_標準"] = geo_df["學校名稱_原始"].apply(normalize_school_name)
        # 取每所學校的平均經緯度（防止同校有微小差異）
        lookup = geo_df.groupby("學校名稱_標準").agg(
            緯度=("緯度", "mean"),
            經度=("經度", "mean")
        ).reset_index()
        return lookup
    else:
        # 無經緯度欄位，回傳空的但保留學校名稱
        schools = phase1_df[school_col].dropna().unique()
        lookup = pd.DataFrame({
            "學校名稱_標準": [normalize_school_name(s) for s in schools],
            "緯度": np.nan,
            "經度": np.nan
        })
        return lookup.drop_duplicates(subset=["學校名稱_標準"])


def enrich_with_geo(df: pd.DataFrame, geo_lookup: pd.DataFrame) -> pd.DataFrame:
    """
    將任意 DataFrame 透過畢業學校名稱，從 geo_lookup 中補上經緯度。
    用於二階資料、合併分析資料等。
    """
    school_col = detect_school_column(df)
    if school_col is None or geo_lookup.empty:
        return df

    df = df.copy()
    df["_school_std"] = df[school_col].apply(normalize_school_name)

    # 移除可能已存在的經緯度欄位（改用一階的）
    for col in ["緯度", "經度", "lat", "lon", "latitude", "longitude"]:
        if col in df.columns and col != school_col:
            df = df.drop(columns=[col], errors="ignore")

    # 合併
    df = df.merge(
        geo_lookup.rename(columns={"學校名稱_標準": "_school_std"}),
        on="_school_std",
        how="left"
    )
    df = df.drop(columns=["_school_std"], errors="ignore")
    return df


# ============================================================
# 核心：統一轉換率計算（分母 = 一階人數）
# ============================================================
def compute_conversion_metrics(phase1_count: int, phase2_count: int = None, final_count: int = None) -> dict:
    """
    以一階為統一分母計算所有轉換率。
    """
    metrics = {
        "phase1_count": phase1_count,
        "phase2_count": phase2_count,
        "final_count": final_count,
        "p1_to_p2_rate": None,
        "p1_to_final_rate": None,
        "p2_to_final_rate": None,
    }
    if phase1_count and phase1_count > 0:
        if phase2_count is not None:
            metrics["p1_to_p2_rate"] = round(phase2_count / phase1_count * 100, 1)
        if final_count is not None:
            metrics["p1_to_final_rate"] = round(final_count / phase1_count * 100, 1)
    if phase2_count and phase2_count > 0 and final_count is not None:
        metrics["p2_to_final_rate"] = round(final_count / phase2_count * 100, 1)
    return metrics


def build_funnel_data(summary_df: pd.DataFrame, phase_cols: dict) -> pd.DataFrame:
    """建構漏斗資料，所有轉換率以一階為分母"""
    records = []
    p1_col = phase_cols.get("phase1")
    p2_col = phase_cols.get("phase2")
    fn_col = phase_cols.get("final")

    if p1_col is None:
        return pd.DataFrame()

    for _, row in summary_df.iterrows():
        p1 = int(row.get(p1_col, 0)) if pd.notna(row.get(p1_col)) else 0
        p2 = int(row.get(p2_col, 0)) if p2_col and pd.notna(row.get(p2_col)) else None
        fn = int(row.get(fn_col, 0)) if fn_col and pd.notna(row.get(fn_col)) else None
        m = compute_conversion_metrics(p1, p2, fn)
        dept_col = detect_department_column(summary_df)
        label = row.get(dept_col, "整體") if dept_col else "整體"
        records.append({"項目": label, **m})
    return pd.DataFrame(records)


# ============================================================
# 視覺化函式
# ============================================================
def create_funnel_chart(labels: list, values: list, title: str = "招生漏斗"):
    fig = go.Figure(go.Funnel(
        y=labels, x=values,
        textinfo="value+percent initial",
        marker=dict(color=["#2196F3", "#FF9800", "#4CAF50", "#E91E63"][:len(labels)]),
        connector=dict(line=dict(color="royalblue", width=2))
    ))
    fig.update_layout(title=title, height=400, font=dict(size=14))
    return fig


def create_conversion_bar(df: pd.DataFrame, x_col: str, y_col: str, title: str, color: str = "#667eea"):
    fig = px.bar(
        df.sort_values(y_col, ascending=True),
        x=y_col, y=x_col, orientation="h",
        text=y_col, title=title
    )
    fig.update_traces(marker_color=color, texttemplate="%{text:.1f}%", textposition="outside")
    fig.update_layout(height=max(350, len(df) * 30), xaxis_title="轉換率 (%)", yaxis_title="")
    return fig


def create_map(df: pd.DataFrame, size_col: str, title: str, color_col: str = None):
    """建立地圖，經緯度來自 geo_lookup 合併後的欄位"""
    if "緯度" not in df.columns or "經度" not in df.columns:
        st.warning("⚠️ 資料中缺少經緯度欄位，無法繪製地圖。")
        return None

    map_df = df.dropna(subset=["緯度", "經度"]).copy()
    if map_df.empty:
        st.warning("⚠️ 無有效的經緯度資料可繪製地圖。")
        return None

    map_df[size_col] = pd.to_numeric(map_df[size_col], errors="coerce").fillna(1)
    school_col = detect_school_column(map_df) or "學校"

    fig = px.scatter_mapbox(
        map_df, lat="緯度", lon="經度",
        size=size_col,
        color=color_col if color_col and color_col in map_df.columns else None,
        hover_name=school_col if school_col in map_df.columns else None,
        hover_data={size_col: True, "緯度": ":.4f", "經度": ":.4f"},
        title=title,
        size_max=30,
        zoom=7,
        center={"lat": HWU_COORDS["lat"], "lon": HWU_COORDS["lon"]},
        mapbox_style="carto-positron"
    )
    # 加上本校標記
    fig.add_trace(go.Scattermapbox(
        lat=[HWU_COORDS["lat"]], lon=[HWU_COORDS["lon"]],
        mode="markers+text",
        marker=dict(size=18, color="red", symbol="star"),
        text=["中華醫事科技大學"], textposition="top center",
        name="本校位置", showlegend=True
    ))
    fig.update_layout(height=600, margin=dict(l=0, r=0, t=40, b=0))
    return fig


def create_heatmap(df: pd.DataFrame, x_col: str, y_col: str, value_col: str, title: str):
    pivot = df.pivot_table(index=y_col, columns=x_col, values=value_col, aggfunc="sum").fillna(0)
    fig = px.imshow(
        pivot, text_auto=True, aspect="auto",
        color_continuous_scale="YlOrRd", title=title
    )
    fig.update_layout(height=max(400, len(pivot) * 25))
    return fig


def create_radar(categories: list, values: list, title: str):
    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=values + [values[0]],
        theta=categories + [categories[0]],
        fill="toself",
        marker=dict(color="#667eea")
    ))
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, max(values) * 1.2 if values else 5])),
        showlegend=False, title=title, height=400
    )
    return fig


def efficiency_stars(rate: float) -> str:
    """轉換率的三星評等"""
    if rate >= 70:
        return "⭐⭐⭐"
    elif rate >= 40:
        return "⭐⭐"
    else:
        return "⭐"


# ============================================================
# 資料上傳與管理（Session State）
# ============================================================
def init_session():
    defaults = {
        "datasets": {},        # {name: {"df": df, "channel": str, "phase": str, "filename": str}}
        "geo_lookup": None,    # 從一階資料建構的學校經緯度對照表
        "survey_data": {},     # 問卷資料
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_session()


# ============================================================
# Sidebar — 資料上傳
# ============================================================
with st.sidebar:
    st.header("📂 資料上傳")

    uploaded_files = st.file_uploader(
        "上傳招生資料 (Excel/CSV)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="支援多檔上傳，系統會自動偵測管道與階段"
    )

    if uploaded_files:
        for uf in uploaded_files:
            if uf.name not in st.session_state["datasets"]:
                try:
                    if uf.name.endswith(".csv"):
                        df = pd.read_csv(uf)
                    else:
                        df = pd.read_excel(uf)

                    # 偵測管道
                    auto_ch = detect_channel_from_columns(df)
                    # 偵測階段
                    phase_map = detect_phase_columns(df)

                    st.session_state["datasets"][uf.name] = {
                        "df": df,
                        "channel": auto_ch,
                        "phase_map": phase_map,
                        "filename": uf.name,
                    }
                    st.success(f"✅ 已載入：{uf.name}（{len(df)} 筆）")
                except Exception as e:
                    st.error(f"❌ 載入失敗 {uf.name}: {e}")

    st.markdown("---")

    # 顯示已載入資料
    if st.session_state["datasets"]:
        st.subheader("📋 已載入資料集")
        for name, info in st.session_state["datasets"].items():
            df = info["df"]
            ch = info.get("channel", "未偵測")
            st.markdown(f"**{name}**")
            st.caption(f"　{len(df)} 筆 ｜ 管道：{ch or '待確認'}")

        if st.button("🗑️ 清除所有資料", type="secondary"):
            st.session_state["datasets"] = {}
            st.session_state["geo_lookup"] = None
            st.rerun()

    st.markdown("---")
    st.subheader("⚙️ 資料階段設定")

    # 選擇一階資料（經緯度來源）
    if st.session_state["datasets"]:
        st.markdown('<div class="info-box">📍 <strong>一階（報名）資料</strong>是畢業學校經緯度的唯一來源。<br>二階資料中的學生將自動從一階查找對應的學校位置。</div>', unsafe_allow_html=True)

        dataset_names = list(st.session_state["datasets"].keys())
        phase1_source = st.selectbox(
            "🔵 請選擇【一階（報名）】資料：",
            ["-- 請選擇 --"] + dataset_names,
            key="phase1_select"
        )
        phase2_source = st.selectbox(
            "🟠 請選擇【二階（報到確認）】資料（可選）：",
            ["-- 無 --"] + dataset_names,
            key="phase2_select"
        )

        # 建構 geo_lookup
        if phase1_source != "-- 請選擇 --":
            p1_df = st.session_state["datasets"][phase1_source]["df"]
            geo_lookup = build_school_geo_lookup(p1_df)
            st.session_state["geo_lookup"] = geo_lookup

            if not geo_lookup.empty and geo_lookup["緯度"].notna().any():
                st.success(f"📍 已從一階資料建立 {len(geo_lookup)} 所學校的經緯度對照表")
            else:
                st.warning("⚠️ 一階資料中未找到經緯度欄位，地圖功能可能受限")

            # 如果有二階資料，自動補上經緯度
            if phase2_source != "-- 無 --":
                p2_df = st.session_state["datasets"][phase2_source]["df"]
                p2_enriched = enrich_with_geo(p2_df, geo_lookup)
                st.session_state["datasets"][phase2_source]["df_enriched"] = p2_enriched
                matched = p2_enriched["緯度"].notna().sum() if "緯度" in p2_enriched.columns else 0
                total = len(p2_enriched)
                st.info(f"🔗 二階資料已自動對應經緯度：{matched}/{total} 筆成功匹配")

    # 管道確認
    st.markdown("---")
    st.subheader("🏷️ 管道確認")
    for name, info in st.session_state["datasets"].items():
        with st.expander(f"📄 {name}"):
            current = info.get("channel")
            confirmed = resolve_channel(current, name)
            st.session_state["datasets"][name]["channel"] = confirmed


# ============================================================
# 主功能選單
# ============================================================
st.markdown("---")

modules = {
    "📊 Module 0：總覽儀表板":       "mod0",
    "🔄 Module 1：招生漏斗分析":      "mod1",
    "📈 Module 2：管道比較分析":       "mod2",
    "🗺️ Module 3：地理分布（地圖）":  "mod3",
    "🏫 Module 4：科系熱力圖":        "mod4",
    "🎯 Module 5：來源學校精準追蹤":  "mod5",
    "⚠️ Module 6：流失預警分析":      "mod6",
    "📋 Module 7：前瞻意向調查分析":  "mod7",
}

selected_module = st.selectbox("選擇分析模組：", list(modules.keys()))
module_key = modules[selected_module]

st.markdown('<hr class="section-divider">', unsafe_allow_html=True)


# ============================================================
# 輔助：取得合併資料
# ============================================================
def get_combined_data() -> pd.DataFrame | None:
    """合併所有已載入的資料集"""
    if not st.session_state["datasets"]:
        st.warning("⚠️ 請先上傳招生資料。")
        return None

    frames = []
    for name, info in st.session_state["datasets"].items():
        df = info.get("df_enriched", info["df"]).copy()
        df["_資料來源"] = name
        df["_招生管道"] = info.get("channel", "未分類")
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)
    return combined


def get_phase_data():
    """分別取得一階和二階資料"""
    p1_key = st.session_state.get("phase1_select")
    p2_key = st.session_state.get("phase2_select")

    p1_df = None
    p2_df = None

    if p1_key and p1_key != "-- 請選擇 --":
        p1_df = st.session_state["datasets"][p1_key]["df"].copy()

    if p2_key and p2_key != "-- 無 --":
        info = st.session_state["datasets"][p2_key]
        p2_df = info.get("df_enriched", info["df"]).copy()

    return p1_df, p2_df


# ============================================================
# Module 0：總覽儀表板
# ============================================================
if module_key == "mod0":
    st.header("📊 Module 0：總覽儀表板")

    p1_df, p2_df = get_phase_data()

    if p1_df is None:
        st.warning("⚠️ 請先在側邊欄選擇一階（報名）資料。")
    else:
        school_col = detect_school_column(p1_df)
        dept_col = detect_department_column(p1_df)

        p1_count = len(p1_df)
        p2_count = len(p2_df) if p2_df is not None else None
        final_count = None  # 若有三階段資料可再擴充

        metrics = compute_conversion_metrics(p1_count, p2_count, final_count)

        # KPI 卡片
        cols = st.columns(4)
        with cols[0]:
            st.markdown(f"""
            <div class="metric-card">
                <h3>一階 報名人數</h3>
                <h1>{p1_count:,}</h1>
            </div>""", unsafe_allow_html=True)
        with cols[1]:
            p2_display = f"{p2_count:,}" if p2_count else "N/A"
            st.markdown(f"""
            <div class="metric-card" style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);">
                <h3>二階 報到人數</h3>
                <h1>{p2_display}</h1>
            </div>""", unsafe_allow_html=True)
        with cols[2]:
            rate_display = f'{metrics["p1_to_p2_rate"]}%' if metrics["p1_to_p2_rate"] is not None else "N/A"
            st.markdown(f"""
            <div class="metric-card" style="background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);">
                <h3>一→二階轉換率</h3>
                <h1>{rate_display}</h1>
            </div>""", unsafe_allow_html=True)
        with cols[3]:
            school_count = p1_df[school_col].nunique() if school_col else "N/A"
            st.markdown(f"""
            <div class="metric-card" style="background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%);">
                <h3>來源學校數</h3>
                <h1>{school_count}</h1>
            </div>""", unsafe_allow_html=True)

        st.markdown("---")

        # 漏斗圖
        funnel_labels = ["一階 報名"]
        funnel_values = [p1_count]
        if p2_count is not None:
            funnel_labels.append("二階 報到確認")
            funnel_values.append(p2_count)

        if len(funnel_values) > 1:
            fig_funnel = create_funnel_chart(funnel_labels, funnel_values, "整體招生漏斗")
            st.plotly_chart(fig_funnel, use_container_width=True)

        col_a, col_b = st.columns(2)
        with col_a:
            # 科系分布
            if dept_col and dept_col in p1_df.columns:
                dept_dist = p1_df[dept_col].value_counts().reset_index()
                dept_dist.columns = ["科系", "報名人數"]
                fig_dept = px.pie(dept_dist, names="科系", values="報名人數",
                                  title="一階報名 — 科系分布", hole=0.4)
                st.plotly_chart(fig_dept, use_container_width=True)

        with col_b:
            # 來源學校 TOP 10
            if school_col and school_col in p1_df.columns:
                top_schools = p1_df[school_col].value_counts().head(10).reset_index()
                top_schools.columns = ["學校", "人數"]
                fig_school = px.bar(top_schools, x="人數", y="學校", orientation="h",
                                    title="一階報名 — 來源學校 TOP 10",
                                    text="人數")
                fig_school.update_traces(marker_color="#667eea")
                fig_school.update_layout(yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig_school, use_container_width=True)


# ============================================================
# Module 1：招生漏斗分析
# ============================================================
elif module_key == "mod1":
    st.header("🔄 Module 1：招生漏斗分析")

    p1_df, p2_df = get_phase_data()

    if p1_df is None:
        st.warning("⚠️ 請先在側邊欄選擇一階（報名）資料。")
    else:
        dept_col = detect_department_column(p1_df)
        school_col = detect_school_column(p1_df)

        st.subheader("1-1. 依科系的漏斗分析")

        if dept_col and p2_df is not None:
            dept_col_p2 = detect_department_column(p2_df)

            p1_by_dept = p1_df[dept_col].value_counts().reset_index()
            p1_by_dept.columns = ["科系", "一階人數"]

            if dept_col_p2:
                p2_by_dept = p2_df[dept_col_p2].value_counts().reset_index()
                p2_by_dept.columns = ["科系", "二階人數"]
                dept_summary = p1_by_dept.merge(p2_by_dept, on="科系", how="left").fillna(0)
            else:
                dept_summary = p1_by_dept.copy()
                dept_summary["二階人數"] = 0

            dept_summary["二階人數"] = dept_summary["二階人數"].astype(int)
            dept_summary["轉換率(%)"] = (dept_summary["二階人數"] / dept_summary["一階人數"] * 100).round(1)
            dept_summary["效率評等"] = dept_summary["轉換率(%)"].apply(efficiency_stars)

            st.dataframe(
                dept_summary.sort_values("轉換率(%)", ascending=False),
                use_container_width=True, hide_index=True
            )

            fig = create_conversion_bar(dept_summary, "科系", "轉換率(%)",
                                        "各科系 一→二階 轉換率（分母=一階報名）")
            st.plotly_chart(fig, use_container_width=True)

            # 各科系漏斗
            st.subheader("1-2. 各科系漏斗圖")
            sel_dept = st.selectbox("選擇科系：", dept_summary["科系"].tolist())
            row = dept_summary[dept_summary["科系"] == sel_dept].iloc[0]
            fig_f = create_funnel_chart(
                ["一階報名", "二階報到"],
                [int(row["一階人數"]), int(row["二階人數"])],
                f"{sel_dept} 招生漏斗"
            )
            st.plotly_chart(fig_f, use_container_width=True)

        elif dept_col:
            st.info("ℹ️ 僅有一階資料，顯示各科系報名分布。")
            dist = p1_df[dept_col].value_counts().reset_index()
            dist.columns = ["科系", "報名人數"]
            fig = px.bar(dist, x="報名人數", y="科系", orientation="h",
                         title="各科系報名人數", text="報名人數")
            fig.update_traces(marker_color="#667eea")
            fig.update_layout(yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("⚠️ 未偵測到科系欄位。")

        # 依學校的漏斗
        st.markdown("---")
        st.subheader("1-3. 依來源學校的漏斗分析")

        if school_col and p2_df is not None:
            school_col_p2 = detect_school_column(p2_df)

            p1_by_sch = p1_df[school_col].value_counts().reset_index()
            p1_by_sch.columns = ["學校", "一階人數"]

            if school_col_p2:
                p2_by_sch = p2_df[school_col_p2].value_counts().reset_index()
                p2_by_sch.columns = ["學校", "二階人數"]
                sch_summary = p1_by_sch.merge(p2_by_sch, on="學校", how="left").fillna(0)
            else:
                sch_summary = p1_by_sch.copy()
                sch_summary["二階人數"] = 0

            sch_summary["二階人數"] = sch_summary["二階人數"].astype(int)
            sch_summary["轉換率(%)"] = (sch_summary["二階人數"] / sch_summary["一階人數"] * 100).round(1)
            sch_summary["效率評等"] = sch_summary["轉換率(%)"].apply(efficiency_stars)

            min_p1 = st.slider("篩選：一階報名人數 ≥", 1, 50, 5, key="sch_filter")
            filtered = sch_summary[sch_summary["一階人數"] >= min_p1].sort_values("轉換率(%)", ascending=False)

            st.dataframe(filtered, use_container_width=True, hide_index=True)

            fig_sch = create_conversion_bar(filtered.head(20), "學校", "轉換率(%)",
                                            f"來源學校轉換率 TOP 20（一階≥{min_p1}人）")
            st.plotly_chart(fig_sch, use_container_width=True)


# ============================================================
# Module 2：管道比較分析
# ============================================================
elif module_key == "mod2":
    st.header("📈 Module 2：管道比較分析")

    combined = get_combined_data()
    if combined is not None:
        channels = combined["_招生管道"].unique()

        if len(channels) < 2:
            st.info("ℹ️ 目前僅有一個管道的資料，上傳多個管道的資料後可進行比較分析。")

        st.subheader("2-1. 各管道報名人數概覽")
        ch_dist = combined["_招生管道"].value_counts().reset_index()
        ch_dist.columns = ["管道", "人數"]

        fig = px.bar(ch_dist, x="管道", y="人數", text="人數",
                     title="各管道報名人數", color="管道")
        st.plotly_chart(fig, use_container_width=True)

        # 各管道科系分布
        st.subheader("2-2. 各管道 × 科系分布")
        dept_col = detect_department_column(combined)
        if dept_col:
            cross = combined.groupby(["_招生管道", dept_col]).size().reset_index(name="人數")
            fig_cross = px.bar(cross, x="_招生管道", y="人數", color=dept_col,
                               title="各管道科系分布", barmode="group")
            st.plotly_chart(fig_cross, use_container_width=True)

        # 各管道來源學校多元性
        st.subheader("2-3. 各管道來源學校多元性")
        school_col = detect_school_column(combined)
        if school_col:
            diversity = combined.groupby("_招生管道")[school_col].nunique().reset_index()
            diversity.columns = ["管道", "來源學校數"]
            fig_div = px.bar(diversity, x="管道", y="來源學校數", text="來源學校數",
                             title="各管道來源學校多元性", color="管道")
            st.plotly_chart(fig_div, use_container_width=True)


# ============================================================
# Module 3：地理分布（地圖）
# ============================================================
elif module_key == "mod3":
    st.header("🗺️ Module 3：地理分布（地圖）")

    p1_df, p2_df = get_phase_data()
    geo_lookup = st.session_state.get("geo_lookup")

    if p1_df is None:
        st.warning("⚠️ 請先在側邊欄選擇一階（報名）資料。")
    else:
        school_col = detect_school_column(p1_df)

        if school_col is None:
            st.warning("⚠️ 未偵測到畢業學校欄位。")
        else:
            st.subheader("3-1. 一階報名 — 來源學校地圖")

            # 一階資料使用自身的經緯度（因為 geo_lookup 就是從一階建的）
            p1_enriched = enrich_with_geo(p1_df, geo_lookup) if geo_lookup is not None else p1_df

            p1_school_agg = p1_enriched.groupby(school_col).agg(
                報名人數=(school_col, "size"),
            ).reset_index()

            # 合併經緯度
            if geo_lookup is not None and not geo_lookup.empty:
                p1_school_agg["_std"] = p1_school_agg[school_col].apply(normalize_school_name)
                p1_school_agg = p1_school_agg.merge(
                    geo_lookup.rename(columns={"學校名稱_標準": "_std"}),
                    on="_std", how="left"
                ).drop(columns=["_std"])

            fig_map = create_map(p1_school_agg, "報名人數", "一階報名 — 來源學校地理分布")
            if fig_map:
                st.plotly_chart(fig_map, use_container_width=True)

            # 二階地圖（經緯度從一階查）
            if p2_df is not None:
                st.markdown("---")
                st.subheader("3-2. 二階報到 — 來源學校地圖")
                st.markdown('<div class="info-box">📍 二階資料的經緯度由一階資料自動匹配</div>', unsafe_allow_html=True)

                p2_enriched = enrich_with_geo(p2_df, geo_lookup) if geo_lookup is not None else p2_df
                school_col_p2 = detect_school_column(p2_enriched)

                if school_col_p2:
                    p2_school_agg = p2_enriched.groupby(school_col_p2).agg(
                        報到人數=(school_col_p2, "size"),
                    ).reset_index()

                    if geo_lookup is not None and not geo_lookup.empty:
                        p2_school_agg["_std"] = p2_school_agg[school_col_p2].apply(normalize_school_name)
                        p2_school_agg = p2_school_agg.merge(
                            geo_lookup.rename(columns={"學校名稱_標準": "_std"}),
                            on="_std", how="left"
                        ).drop(columns=["_std"])

                    fig_map2 = create_map(p2_school_agg, "報到人數",
                                          "二階報到 — 來源學校地理分布（經緯度來源：一階資料）")
                    if fig_map2:
                        st.plotly_chart(fig_map2, use_container_width=True)

                    # 匹配統計
                    if "緯度" in p2_school_agg.columns:
                        matched = p2_school_agg["緯度"].notna().sum()
                        total = len(p2_school_agg)
                        unmatched = p2_school_agg[p2_school_agg["緯度"].isna()]
                        st.info(f"📊 學校經緯度匹配：{matched}/{total} 所成功 ({matched/total*100:.0f}%)")
                        if not unmatched.empty:
                            st.warning(f"⚠️ 以下 {len(unmatched)} 所學校在一階資料中找不到經緯度：")
                            st.dataframe(unmatched[[school_col_p2, "報到人數"]],
                                         use_container_width=True, hide_index=True)


# ============================================================
# Module 4：科系熱力圖
# ============================================================
elif module_key == "mod4":
    st.header("🏫 Module 4：科系 × 來源學校 熱力圖")

    p1_df, p2_df = get_phase_data()

    if p1_df is None:
        st.warning("⚠️ 請先在側邊欄選擇一階（報名）資料。")
    else:
        dept_col = detect_department_column(p1_df)
        school_col = detect_school_column(p1_df)

        if dept_col and school_col:
            st.subheader("4-1. 一階報名熱力圖")

            min_count = st.slider("篩選：學校報名人數 ≥", 1, 30, 3, key="hm_filter")
            school_counts = p1_df[school_col].value_counts()
            valid_schools = school_counts[school_counts >= min_count].index

            filtered = p1_df[p1_df[school_col].isin(valid_schools)]
            cross = filtered.groupby([school_col, dept_col]).size().reset_index(name="人數")

            fig_hm = create_heatmap(cross, dept_col, school_col, "人數",
                                    f"科系 × 來源學校 熱力圖（學校報名≥{min_count}人）")
            st.plotly_chart(fig_hm, use_container_width=True)

            # 轉換率熱力圖（如有二階）
            if p2_df is not None:
                st.subheader("4-2. 轉換率熱力圖")
                dept_col_p2 = detect_department_column(p2_df)
                school_col_p2 = detect_school_column(p2_df)

                if dept_col_p2 and school_col_p2:
                    p1_cross = p1_df.groupby([school_col, dept_col]).size().reset_index(name="一階")
                    p2_cross = p2_df.groupby([school_col_p2, dept_col_p2]).size().reset_index(name="二階")
                    p2_cross.columns = [school_col, dept_col, "二階"]

                    rate_cross = p1_cross.merge(p2_cross, on=[school_col, dept_col], how="left").fillna(0)
                    rate_cross["轉換率"] = (rate_cross["二階"] / rate_cross["一階"] * 100).round(1)
                    rate_cross = rate_cross[rate_cross[school_col].isin(valid_schools)]

                    fig_hm2 = create_heatmap(rate_cross, dept_col, school_col, "轉換率",
                                             "一→二階 轉換率 熱力圖 (%)")
                    st.plotly_chart(fig_hm2, use_container_width=True)
        else:
            st.warning("⚠️ 未偵測到科系或學校欄位。")


# ============================================================
# Module 5：來源學校精準追蹤
# ============================================================
elif module_key == "mod5":
    st.header("🎯 Module 5：來源學校精準追蹤")

    p1_df, p2_df = get_phase_data()

    if p1_df is None:
        st.warning("⚠️ 請先在側邊欄選擇一階（報名）資料。")
    else:
        school_col = detect_school_column(p1_df)
        if school_col is None:
            st.warning("⚠️ 未偵測到畢業學校欄位。")
        else:
            # 學校層級統計
            p1_by_sch = p1_df[school_col].value_counts().reset_index()
            p1_by_sch.columns = ["學校", "一階人數"]

            if p2_df is not None:
                school_col_p2 = detect_school_column(p2_df)
                if school_col_p2:
                    p2_by_sch = p2_df[school_col_p2].value_counts().reset_index()
                    p2_by_sch.columns = ["學校", "二階人數"]
                    sch_stats = p1_by_sch.merge(p2_by_sch, on="學校", how="left").fillna(0)
                    sch_stats["二階人數"] = sch_stats["二階人數"].astype(int)
                else:
                    sch_stats = p1_by_sch.copy()
                    sch_stats["二階人數"] = 0
            else:
                sch_stats = p1_by_sch.copy()
                sch_stats["二階人數"] = 0

            sch_stats["轉換率(%)"] = (sch_stats["二階人數"] / sch_stats["一階人數"] * 100).round(1)
            sch_stats["流失人數"] = sch_stats["一階人數"] - sch_stats["二階人數"]
            sch_stats["效率評等"] = sch_stats["轉換率(%)"].apply(efficiency_stars)

            # Tier 分級
            def assign_tier(count):
                if count >= 30:
                    return "⭐⭐⭐ Tier 1"
                elif count >= 10:
                    return "⭐⭐ Tier 2"
                else:
                    return "⭐ Tier 3"

            sch_stats["經營分級"] = sch_stats["一階人數"].apply(assign_tier)

            # 顯示
            st.subheader("5-1. 來源學校總覽")
            tier_filter = st.multiselect(
                "篩選經營分級：",
                ["⭐⭐⭐ Tier 1", "⭐⭐ Tier 2", "⭐ Tier 3"],
                default=["⭐⭐⭐ Tier 1", "⭐⭐ Tier 2"],
                key="tier_filter"
            )
            display_df = sch_stats[sch_stats["經營分級"].isin(tier_filter)].sort_values("一階人數", ascending=False)
            st.dataframe(display_df, use_container_width=True, hide_index=True)

            # Tier 統計
            st.subheader("5-2. 各分級統計")
            tier_summary = sch_stats.groupby("經營分級").agg(
                學校數=("學校", "count"),
                總一階人數=("一階人數", "sum"),
                總二階人數=("二階人數", "sum"),
                平均轉換率=("轉換率(%)", "mean")
            ).round(1).reset_index()
            st.dataframe(tier_summary, use_container_width=True, hide_index=True)

            # 個別學校深入
            st.subheader("5-3. 個別學校深入分析")
            sel_school = st.selectbox("選擇學校：", sch_stats.sort_values("一階人數", ascending=False)["學校"].tolist())

            if sel_school:
                row = sch_stats[sch_stats["學校"] == sel_school].iloc[0]
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("一階報名", f'{int(row["一階人數"])} 人')
                c2.metric("二階報到", f'{int(row["二階人數"])} 人')
                c3.metric("轉換率", f'{row["轉換率(%)"]}%')
                c4.metric("經營分級", row["經營分級"])

                # 該校報名科系分布
                dept_col = detect_department_column(p1_df)
                if dept_col:
                    sch_dept = p1_df[p1_df[school_col] == sel_school][dept_col].value_counts().reset_index()
                    sch_dept.columns = ["科系", "報名人數"]
                    fig = px.pie(sch_dept, names="科系", values="報名人數",
                                 title=f"{sel_school} — 報名科系分布")
                    st.plotly_chart(fig, use_container_width=True)


# ============================================================
# Module 6：流失預警分析
# ============================================================
elif module_key == "mod6":
    st.header("⚠️ Module 6：流失預警分析")

    p1_df, p2_df = get_phase_data()

    if p1_df is None or p2_df is None:
        st.warning("⚠️ 此模組需要同時上傳一階（報名）和二階（報到確認）資料。")
    else:
        school_col = detect_school_column(p1_df)
        school_col_p2 = detect_school_column(p2_df)

        if school_col and school_col_p2:
            # 建構流失分析資料
            p1_by_sch = p1_df[school_col].value_counts().reset_index()
            p1_by_sch.columns = ["學校", "一階人數"]

            p2_by_sch = p2_df[school_col_p2].value_counts().reset_index()
            p2_by_sch.columns = ["學校", "二階人數"]

            loss_df = p1_by_sch.merge(p2_by_sch, on="學校", how="left").fillna(0)
            loss_df["二階人數"] = loss_df["二階人數"].astype(int)
            loss_df["轉換率(%)"] = (loss_df["二階人數"] / loss_df["一階人數"] * 100).round(1)
            loss_df["流失人數"] = loss_df["一階人數"] - loss_df["二階人數"]
            loss_df["流失率(%)"] = (100 - loss_df["轉換率(%)"]).round(1)

            # 6-1. 高流失學校預警
            st.subheader("6-1. 高報名低轉換預警學校")

            min_reg = st.slider("篩選：一階報名 ≥", 1, 50, 10, key="loss_filter")
            avg_rate = loss_df[loss_df["一階人數"] >= min_reg]["轉換率(%)"].mean()

            warning_schools = loss_df[
                (loss_df["一階人數"] >= min_reg) &
                (loss_df["轉換率(%)"] < avg_rate)
            ].sort_values("流失人數", ascending=False)

            if not warning_schools.empty:
                st.markdown(f'<div class="warning-box">⚠️ 以下學校一階報名≥{min_reg}人，但轉換率低於平均 ({avg_rate:.1f}%)，建議重點關注</div>', unsafe_allow_html=True)
                st.dataframe(warning_schools, use_container_width=True, hide_index=True)

                # IPA 四象限分析
                st.subheader("6-2. IPA 矩陣分析")
                st.markdown("**X軸**：一階報名人數（量能）　**Y軸**：轉換率（效率）")

                analysis_df = loss_df[loss_df["一階人數"] >= min_reg].copy()
                median_x = analysis_df["一階人數"].median()
                median_y = analysis_df["轉換率(%)"].median()

                fig_ipa = px.scatter(
                    analysis_df, x="一階人數", y="轉換率(%)",
                    size="流失人數", hover_name="學校",
                    title="來源學校 IPA 矩陣（氣泡大小 = 流失人數）",
                    size_max=40
                )
                fig_ipa.add_hline(y=median_y, line_dash="dash", line_color="red",
                                  annotation_text=f"轉換率中位數 {median_y:.1f}%")
                fig_ipa.add_vline(x=median_x, line_dash="dash", line_color="blue",
                                  annotation_text=f"報名中位數 {median_x:.0f}人")

                # 象限標注
                fig_ipa.add_annotation(x=analysis_df["一階人數"].max()*0.9, y=analysis_df["轉換率(%)"].max()*0.95,
                                       text="🌟 高量能高效率<br>持續維護", showarrow=False,
                                       font=dict(size=11, color="green"))
                fig_ipa.add_annotation(x=analysis_df["一階人數"].max()*0.9, y=analysis_df["轉換率(%)"].min()*1.1,
                                       text="⚠️ 高量能低效率<br>重點改善", showarrow=False,
                                       font=dict(size=11, color="red"))
                fig_ipa.add_annotation(x=analysis_df["一階人數"].min()*1.1, y=analysis_df["轉換率(%)"].max()*0.95,
                                       text="📈 低量能高效率<br>擴大招生", showarrow=False,
                                       font=dict(size=11, color="blue"))
                fig_ipa.add_annotation(x=analysis_df["一階人數"].min()*1.1, y=analysis_df["轉換率(%)"].min()*1.1,
                                       text="🔍 低量能低效率<br>評估投入", showarrow=False,
                                       font=dict(size=11, color="gray"))

                fig_ipa.update_layout(height=600)
                st.plotly_chart(fig_ipa, use_container_width=True)

                # 流失人數排行
                st.subheader("6-3. 流失人數排行 TOP 15")
                top_loss = loss_df.sort_values("流失人數", ascending=False).head(15)
                fig_loss = px.bar(
                    top_loss, x="流失人數", y="學校", orientation="h",
                    text="流失人數", title="流失人數排行",
                    color="轉換率(%)", color_continuous_scale="RdYlGn"
                )
                fig_loss.update_layout(yaxis=dict(autorange="reversed"), height=500)
                st.plotly_chart(fig_loss, use_container_width=True)
            else:
                st.success("✅ 目前沒有符合預警條件的學校。")
        else:
            st.warning("⚠️ 未偵測到畢業學校欄位。")


# ============================================================
# Module 7：前瞻意向調查分析
# ============================================================
elif module_key == "mod7":
    st.header("📋 Module 7：前瞻意向調查分析")

    st.markdown("""
    <div class="info-box">
    📌 <strong>本模組分析高二（問卷A）及高三（問卷B）的升學意向調查結果。</strong><br>
    各次施測為獨立分析，以群體趨勢取代個人追蹤。<br>
    支援 SurveyCake / Google Forms 匯出的 Excel 或 CSV 格式。
    </div>
    """, unsafe_allow_html=True)

    survey_tab1, survey_tab2, survey_tab3 = st.tabs(["📤 資料上傳", "📊 問卷A分析", "📊 問卷B分析"])

    with survey_tab1:
        st.subheader("上傳意向調查資料")

        survey_type = st.radio("問卷類型：", ["問卷A（高二探索意向）", "問卷B（高三選校決策）"], horizontal=True)

        survey_file = st.file_uploader(
            "上傳問卷結果 (Excel/CSV)",
            type=["xlsx", "xls", "csv"],
            key="survey_upload"
        )

        if survey_file:
            try:
                if survey_file.name.endswith(".csv"):
                    survey_df = pd.read_csv(survey_file)
                else:
                    survey_df = pd.read_excel(survey_file)

                stype = "A" if "問卷A" in survey_type else "B"
                academic_year = st.text_input("學年度（如：113）", "113")
                key_name = f"{stype}_{academic_year}"

                st.session_state["survey_data"][key_name] = survey_df
                st.success(f"✅ 已載入 {survey_type}（{academic_year}學年度）：{len(survey_df)} 筆回覆")
                st.dataframe(survey_df.head(), use_container_width=True)
            except Exception as e:
                st.error(f"❌ 載入失敗：{e}")

        # 顯示已載入問卷
        if st.session_state["survey_data"]:
            st.markdown("---")
            st.subheader("已載入的調查資料")
            for k, v in st.session_state["survey_data"].items():
                stype, yr = k.split("_")
                label = "問卷A（高二）" if stype == "A" else "問卷B（高三）"
                st.markdown(f"- **{label} {yr}學年度**：{len(v)} 筆")

    with survey_tab2:
        st.subheader("📊 問卷A 分析（高二探索意向）")

        # 找出所有問卷A的資料
        a_data = {k: v for k, v in st.session_state["survey_data"].items() if k.startswith("A_")}

        if not a_data:
            st.info("ℹ️ 尚未上傳問卷A資料。請在「資料上傳」頁籤中上傳。")
        else:
            sel_a = st.selectbox("選擇分析的資料：", list(a_data.keys()),
                                 format_func=lambda x: f"問卷A — {x.split('_')[1]}學年度")
            df_a = a_data[sel_a]

            st.markdown("---")
            st.markdown("#### 以下為通用分析框架，將根據實際欄位名稱自動適配")

            # 自動尋找可能的欄位
            cols_a = df_a.columns.tolist()
            st.caption(f"偵測到的欄位：{', '.join(cols_a[:20])}{'...' if len(cols_a)>20 else ''}")

            # 嘗試分析：各欄位的值分布
            analysis_cols = st.multiselect(
                "選擇要分析的欄位：", cols_a,
                default=cols_a[:5] if len(cols_a) >= 5 else cols_a
            )

            for col in analysis_cols:
                st.markdown(f"##### 📌 {col}")
                val_counts = df_a[col].dropna().astype(str).value_counts()

                if len(val_counts) <= 15:
                    fig = px.pie(values=val_counts.values, names=val_counts.index,
                                 title=f"{col} — 回覆分布")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    fig = px.bar(x=val_counts.index[:20], y=val_counts.values[:20],
                                 title=f"{col} — 回覆分布 (TOP 20)")
                    st.plotly_chart(fig, use_container_width=True)

    with survey_tab3:
        st.subheader("📊 問卷B 分析（高三選校決策）")

        b_data = {k: v for k, v in st.session_state["survey_data"].items() if k.startswith("B_")}

        if not b_data:
            st.info("ℹ️ 尚未上傳問卷B資料。請在「資料上傳」頁籤中上傳。")
        else:
            sel_b = st.selectbox("選擇分析的資料：", list(b_data.keys()),
                                 format_func=lambda x: f"問卷B — {x.split('_')[1]}學年度")
            df_b = b_data[sel_b]

            st.markdown("---")
            cols_b = df_b.columns.tolist()
            st.caption(f"偵測到的欄位：{', '.join(cols_b[:20])}{'...' if len(cols_b)>20 else ''}")

            analysis_cols_b = st.multiselect(
                "選擇要分析的欄位：", cols_b,
                default=cols_b[:5] if len(cols_b) >= 5 else cols_b,
                key="b_cols"
            )

            for col in analysis_cols_b:
                st.markdown(f"##### 📌 {col}")
                val_counts = df_b[col].dropna().astype(str).value_counts()

                if len(val_counts) <= 15:
                    fig = px.pie(values=val_counts.values, names=val_counts.index,
                                 title=f"{col} — 回覆分布")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    fig = px.bar(x=val_counts.index[:20], y=val_counts.values[:20],
                                 title=f"{col} — 回覆分布 (TOP 20)")
                    st.plotly_chart(fig, use_container_width=True)

            # 本校考量分析（嘗試自動找欄位）
            st.markdown("---")
            st.subheader("🎯 本校定位分析")
            hwu_col_candidates = ["中華醫事", "本校考量", "考慮學校", "B11"]
            hwu_col = None
            for c in hwu_col_candidates:
                for col in cols_b:
                    if c in str(col):
                        hwu_col = col
                        break

            if hwu_col:
                hwu_dist = df_b[hwu_col].dropna().value_counts()
                fig_hwu = px.pie(values=hwu_dist.values, names=hwu_dist.index,
                                 title="本校在學生考量中的定位", hole=0.4,
                                 color_discrete_sequence=px.colors.qualitative.Set2)
                st.plotly_chart(fig_hwu, use_container_width=True)
            else:
                st.info("ℹ️ 未自動偵測到本校考量相關欄位，請在上方手動選擇分析。")


# ============================================================
# Footer
# ============================================================
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
st.markdown("""
<div style="text-align: center; color: #aaa; font-size: 0.85rem; padding: 10px;">
    🎓 中華醫事科技大學 招生數據分析系統 v5.0<br>
    📍 二階資料經緯度統一由一階資料提供 ｜ 📊 轉換率統一以一階為分母<br>
    Built with Streamlit + Plotly ｜ © 2024 HWU Admissions Office
</div>
""", unsafe_allow_html=True)
