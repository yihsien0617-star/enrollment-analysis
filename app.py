# -*- coding: utf-8 -*-
"""
中華醫事科技大學 招生數據分析系統 v6.8
- 新增：各階段「實際人頭數 vs 志願次數」分析（解決一人填多科系重複計算問題）
- 新增：自動偵測學生ID欄位（准考證號、學號、身分證字號等）
- 新增：科系 × 人頭數交叉分析（去重後的真實科系分布）
- 新增：總覽6 KPI → 9 KPI，含人頭數/志願次/重複率
- 保留：全年度惡化偵測、三段轉換率、縮寫展開引擎
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
from itertools import combinations

# ============================================================
# 頁面設定
# ============================================================
st.set_page_config(
    page_title="HWAI 招生數據分析系統",
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
    .metric-purple{background:linear-gradient(135deg,#a18cd1 0%,#fbc2eb 100%);color:#1a1a2e;}
    .metric-teal{background:linear-gradient(135deg,#2af598 0%,#009efd 100%);color:#1a1a2e;}
    .metric-red{background:linear-gradient(135deg,#ff6b6b 0%,#ee5a24 100%);}
    .metric-cyan{background:linear-gradient(135deg,#74b9ff 0%,#0984e3 100%);}
    .metric-lime{background:linear-gradient(135deg,#a8e063 0%,#56ab2f 100%);color:#1a1a2e;}
    .section-divider{
        height:3px;background:linear-gradient(90deg,#667eea,#764ba2,#f093fb);
        border:none;border-radius:2px;margin:25px 0;
    }
    .info-box{background:#f0f4ff;border-left:4px solid #667eea;padding:15px;border-radius:0 10px 10px 0;margin:10px 0;font-size:.9rem;}
    .warning-box{background:#fff8e1;border-left:4px solid #ffa726;padding:15px;border-radius:0 10px 10px 0;margin:10px 0;font-size:.9rem;}
    .success-box{background:#e8f5e9;border-left:4px solid #4caf50;padding:15px;border-radius:0 10px 10px 0;margin:10px 0;font-size:.9rem;}
    .danger-box{background:#ffebee;border-left:4px solid #f44336;padding:15px;border-radius:0 10px 10px 0;margin:10px 0;font-size:.9rem;}
    .mapping-box{background:#f3e5f5;border-left:4px solid #9c27b0;padding:15px;border-radius:0 10px 10px 0;margin:10px 0;font-size:.9rem;}
    .dedup-box{background:#e3f2fd;border-left:4px solid #1976d2;padding:15px;border-radius:0 10px 10px 0;margin:10px 0;font-size:.9rem;}
    .year-tag{display:inline-block;background:#e3f2fd;color:#1565c0;padding:3px 12px;border-radius:12px;font-size:.85rem;margin:2px 3px;font-weight:600;}
    .channel-tag{display:inline-block;background:#e8f5e9;color:#2e7d32;padding:2px 10px;border-radius:12px;font-size:.8rem;margin:2px 3px;}
    .trend-up{color:#4caf50;font-weight:bold;}
    .trend-down{color:#f44336;font-weight:bold;}
    .trend-flat{color:#ff9800;font-weight:bold;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🎓 中華醫事科技大學 招生數據分析系統</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Enrollment Analytics v6.8 ｜ 人數 vs 志願次數 ｜ 全年度惡化偵測 ｜ 三段轉換率</div>', unsafe_allow_html=True)
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

# ============================================================
# 常數
# ============================================================
FINAL_CH_CANDIDATES = [
    "入學方式", "入學管道", "錄取管道", "招生管道", "管道",
    "入學途徑", "錄取方式", "報名管道"
]
HWU = {"lat": 22.9340, "lon": 120.2756}

# 學生ID欄位候選關鍵字
STUDENT_ID_KEYWORDS = [
    "准考證", "准考證號", "准考證號碼", "考生號碼", "報名序號",
    "學號", "身分證", "身分證字號", "身份證", "身份證字號",
    "考生編號", "編號", "報名編號", "序號", "統測准考證",
    "ID", "id", "student_id", "StudentID", "exam_id",
]

ABBREV_EXPAND = {
    "寵護": "寵物照護與美容",
    "寵美": "寵物美容",
    "寵經": "寵物經營",
    "食營": "食品營養",
    "環安": "環境與安全衛生工程",
    "職安": "職業安全衛生",
    "幼保": "幼兒保育",
    "運休": "運動健康與休閒",
    "餐旅": "餐旅管理",
    "調保": "調理保健技術",
    "調理": "調理保健技術",
    "語治": "語言治療",
    "醫技": "醫學檢驗生物技術",
    "醫檢": "醫學檢驗生物技術",
    "醫管": "醫務暨健康事務管理",
    "製藥": "製藥工程",
    "長照": "長期照護",
    "護理": "護理",
    "視光": "視光",
}

DEPT_ALIAS = {
    "護理": "護理系",
    "醫技": "醫學檢驗生物技術系", "醫檢": "醫學檢驗生物技術系",
    "醫學檢驗": "醫學檢驗生物技術系", "醫學檢驗生物技術": "醫學檢驗生物技術系",
    "藥學": "藥學系", "視光": "視光系",
    "製藥": "製藥工程系", "製藥工程": "製藥工程系",
    "食營": "食品營養系", "食品營養": "食品營養系",
    "職安": "職業安全衛生系", "職業安全衛生": "職業安全衛生系",
    "環安": "環境與安全衛生工程系", "環衛": "環境與安全衛生工程系",
    "環境與安全衛生工程": "環境與安全衛生工程系",
    "資管": "資訊管理系", "資訊管理": "資訊管理系",
    "多媒體": "多媒體設計系", "多媒體設計": "多媒體設計系", "多媒": "多媒體設計系",
    "幼保": "幼兒保育系", "幼兒保育": "幼兒保育系",
    "運休": "運動健康與休閒系", "運動休閒": "運動健康與休閒系",
    "運動健康與休閒": "運動健康與休閒系",
    "餐旅": "餐旅管理系", "餐旅管理": "餐旅管理系",
    "觀光": "觀光休閒事業管理系", "觀光休閒": "觀光休閒事業管理系",
    "觀休": "觀光休閒事業管理系", "觀光休閒事業管理": "觀光休閒事業管理系",
    "妝管": "化妝品應用與管理系", "美妝": "化妝品應用與管理系",
    "化妝品": "化妝品應用與管理系", "化妝品應用": "化妝品應用與管理系",
    "化妝品應用與管理": "化妝品應用與管理系", "化妝": "化妝品應用與管理系",
    "調理": "調理保健技術系", "調保": "調理保健技術系",
    "調理保健": "調理保健技術系", "調理保健技術": "調理保健技術系",
    "語治": "語言治療系", "語言治療": "語言治療系",
    "牙技": "牙體技術系", "牙體技術": "牙體技術系",
    "生科": "生物科技系", "生物科技": "生物科技系",
    "寵物": "寵物美容學位學程", "寵物美容": "寵物美容學位學程",
    "寵美": "寵物美容學位學程",
    "寵物照護": "寵物照護學位學程", "寵護": "寵物照護學位學程",
    "寵物經營": "寵物經營學位學程", "寵經": "寵物經營學位學程",
    "長照": "長期照護學位學程", "長期照護": "長期照護學位學程",
    "醫管": "醫務暨健康事務管理系", "醫務管理": "醫務暨健康事務管理系",
    "健康事務管理": "醫務暨健康事務管理系", "健管": "醫務暨健康事務管理系",
    "醫務暨健康事務管理": "醫務暨健康事務管理系",
}

# ============================================================
# v6.8 學生ID偵測引擎
# ============================================================
def detect_student_id_col(df):
    """自動偵測學生唯一識別欄位"""
    for kw in STUDENT_ID_KEYWORDS:
        for c in df.columns:
            if kw in str(c):
                nuniq = df[c].dropna().nunique()
                total = df[c].dropna().shape[0]
                if nuniq > 0 and nuniq / total >= 0.5:
                    return c
    for c in df.columns:
        s = str(c).strip().lower()
        if any(k in s for k in ["准考", "學號", "身分證", "身份證", "考生", "student", "exam"]):
            nuniq = df[c].dropna().nunique()
            total = df[c].dropna().shape[0]
            if nuniq > 0 and nuniq / total >= 0.5:
                return c
    return None


def compute_headcount_stats(df, id_col, dept_col=None):
    """
    計算人頭數統計：
    - 總列數（志願次數）
    - 不重複人頭數
    - 重複率
    - 每人平均填報科系數
    - 各科系的人頭數（去重後）
    """
    result = {}
    total_rows = len(df)
    if id_col and id_col in df.columns:
        ids = df[id_col].dropna().astype(str).str.strip()
        ids = ids[ids != ""]
        unique_ids = ids.nunique()
        total_with_id = len(ids)
        no_id = total_rows - total_with_id
        dup_count = total_with_id - unique_ids
        dup_rate = round(dup_count / total_with_id * 100, 1) if total_with_id > 0 else 0
        avg_entries = round(total_with_id / unique_ids, 2) if unique_ids > 0 else 0

        result = {
            "total_rows": total_rows,
            "total_with_id": total_with_id,
            "unique_headcount": unique_ids,
            "no_id_rows": no_id,
            "dup_count": dup_count,
            "dup_rate": dup_rate,
            "avg_entries_per_person": avg_entries,
            "id_col": id_col,
        }

        # 各科系人頭數（去重）
        if dept_col and dept_col in df.columns:
            df_with_id = df[df[id_col].notna()].copy()
            df_with_id["_id_clean"] = df_with_id[id_col].astype(str).str.strip()

            # 志願次數（不去重）
            dept_wish = df_with_id.groupby(dept_col).size().reset_index(name="志願次數")

            # 人頭數（去重）
            dept_head = df_with_id.groupby(dept_col)["_id_clean"].nunique().reset_index(name="人頭數")

            dept_stats = dept_wish.merge(dept_head, on=dept_col, how="outer").fillna(0)
            dept_stats["志願次數"] = dept_stats["志願次數"].astype(int)
            dept_stats["人頭數"] = dept_stats["人頭數"].astype(int)
            dept_stats["重複率(%)"] = np.where(
                dept_stats["志願次數"] > 0,
                ((dept_stats["志願次數"] - dept_stats["人頭數"]) / dept_stats["志願次數"] * 100).round(1),
                0
            )
            result["dept_stats"] = dept_stats
    else:
        result = {
            "total_rows": total_rows,
            "total_with_id": 0,
            "unique_headcount": None,
            "no_id_rows": total_rows,
            "dup_count": None,
            "dup_rate": None,
            "avg_entries_per_person": None,
            "id_col": None,
        }
    return result


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


def expand_abbrev(dept_kw):
    if dept_kw in ABBREV_EXPAND:
        return ABBREV_EXPAND[dept_kw]
    if len(dept_kw) >= 2 and dept_kw[:2] in ABBREV_EXPAND:
        return ABBREV_EXPAND[dept_kw[:2]]
    return dept_kw


def parse_class_name(class_name):
    if not isinstance(class_name, str):
        return None
    s = re.sub(r'\s+', '', class_name.strip())
    if not s:
        return None
    m = CLASS_PATTERN.match(s)
    if m:
        prefix = m.group(1) or ""
        dept_kw_raw = m.group(2)
        grade = m.group(3)
        section = m.group(4)
        program = "二技" if prefix in ("二", "二技") else ("四技" if prefix in ("四", "四技") else "五專")
        dept_kw = expand_abbrev(dept_kw_raw)
        return (program, dept_kw, grade, section, dept_kw_raw)
    m2 = CLASS_PATTERN_SIMPLE.match(s)
    if m2:
        prefix = m2.group(1) or ""
        dept_kw_raw = m2.group(2)
        program = "二技" if prefix in ("二", "二技") else ("四技" if prefix in ("四", "四技") else "")
        dept_kw = expand_abbrev(dept_kw_raw)
        return (program, dept_kw, "", "", dept_kw_raw)
    return None


def norm_dept(name):
    if not isinstance(name, str):
        return str(name).strip()
    name = re.sub(r"\s+", "", name.strip())
    name = name.replace("臺", "台").replace("（", "(").replace("）", ")").replace("　", "")
    return name


def resolve_dept_from_keyword(dept_kw, p1_depts=None, dept_kw_raw=None):
    kw = dept_kw.strip()
    raw = (dept_kw_raw or kw).strip()
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
    for try_kw in [raw, kw]:
        if try_kw in DEPT_ALIAS:
            alias_result = DEPT_ALIAS[try_kw]
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
    if p1_depts and len(raw) >= 2 and raw != kw:
        short = raw[:2]
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
        if len(parsed) == 5:
            program, dept_kw, grade, section, dept_kw_raw = parsed
        else:
            program, dept_kw, grade, section = parsed
            dept_kw_raw = dept_kw
        dept = resolve_dept_from_keyword(dept_kw, p1_depts, dept_kw_raw)
        if dept:
            mapping[cls] = dept
            abbrev_note = f"（縮寫「{dept_kw_raw}」→「{dept_kw}」）" if dept_kw_raw != dept_kw else ""
            match_detail[cls] = f"✅ [{program}] 關鍵字「{dept_kw}」{abbrev_note} → {dept}"
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
    if r >= 70: return "⭐⭐⭐"
    elif r >= 40: return "⭐⭐"
    else: return "⭐"


def safe_pct(num, den):
    if den and den > 0:
        return round(num / den * 100, 1)
    return 0.0


def trend_arrow(current, previous):
    if previous == 0: return "—"
    diff = current - previous
    pct = diff / previous * 100
    if pct > 5: return f'<span class="trend-up">▲ +{pct:.1f}%</span>'
    elif pct < -5: return f'<span class="trend-down">▼ {pct:.1f}%</span>'
    else: return f'<span class="trend-flat">► {pct:+.1f}%</span>'


# ============================================================
# 統計構建（含人頭數）
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
        if core == p1_core: return p1_norm
        if len(core) >= 3 and (core in p1_core or p1_core in core): return p1_norm
    if len(core) >= 2:
        short = core[:2]
        for p1_norm in p1_name_map:
            if short in p1_norm: return p1_norm
    return norm_name


def build_dept_stats(p1, p2=None, p3=None):
    dc1 = detect_dept_col(p1)
    if dc1 is None:
        return None, None
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
    s["二→最終(%)"] = np.where(s["二階人數"] > 0, (s["最終入學"] / s["二階人數"] * 100).round(1), 0)
    s["一→最終(%)"] = np.where(s["一階人數"] > 0, (s["最終入學"] / s["一階人數"] * 100).round(1), 0)
    s["流失人數"] = s["一階人數"] - s["最終入學"]
    s["二階流失"] = s["二階人數"] - s["最終入學"]
    s["效率評等"] = s["一→最終(%)"].apply(eff_stars)
    s["科系"] = s["_dept_std"].map(name_map).fillna(s["_dept_std"])
    s = s.drop(columns=["_dept_std"])
    col_order = ["科系", "一階人數", "二階人數", "最終入學",
                 "一→二階(%)", "二→最終(%)", "一→最終(%)",
                 "流失人數", "二階流失", "效率評等"]
    s = s[[c for c in col_order if c in s.columns]]
    return s, p3_info


def build_dept_stats_headcount(p1, p2=None, p3=None, id_cols=None):
    """
    建立科系統計（含人頭數版本）
    id_cols = {"p1": id_col_name, "p2": ..., "p3": ...}
    """
    if id_cols is None:
        id_cols = {}

    dc1 = detect_dept_col(p1)
    if dc1 is None:
        return None

    p1_depts = p1[dc1].dropna().unique().tolist()
    name_map = {}
    for raw in p1[dc1].dropna().unique():
        name_map[norm_dept(raw)] = raw

    rows = []
    # 一階
    id1 = id_cols.get("p1")
    p1c = p1.copy()
    p1c["_dept_norm"] = p1c[dc1].apply(norm_dept)
    for dept_norm, grp in p1c.groupby("_dept_norm"):
        wish_count = len(grp)
        if id1 and id1 in grp.columns:
            head_count = grp[id1].dropna().astype(str).str.strip().nunique()
        else:
            head_count = None
        rows.append({
            "_dept_std": dept_norm,
            "一階志願次數": wish_count,
            "一階人頭數": head_count,
        })

    base = pd.DataFrame(rows)

    # 二階
    if p2 is not None:
        id2 = id_cols.get("p2")
        s2, _ = get_dept_series(p2, p1_depts)
        if not s2.empty:
            s2_aligned = s2.map(lambda x: align_dept_name(x, name_map))
            # 志願次數
            t2w = s2_aligned.value_counts().reset_index()
            t2w.columns = ["_dept_std", "二階志願次數"]
            base = base.merge(t2w, on="_dept_std", how="left")
            # 人頭數
            if id2 and id2 in p2.columns:
                dc2 = detect_dept_col(p2)
                if dc2:
                    p2c = p2.copy()
                    p2c["_dept_norm"] = p2c[dc2].apply(norm_dept).map(lambda x: align_dept_name(x, name_map))
                    t2h = p2c.groupby("_dept_norm")[id2].apply(lambda x: x.dropna().astype(str).str.strip().nunique()).reset_index()
                    t2h.columns = ["_dept_std", "二階人頭數"]
                    base = base.merge(t2h, on="_dept_std", how="left")

    if "二階志願次數" not in base.columns: base["二階志願次數"] = np.nan
    if "二階人頭數" not in base.columns: base["二階人頭數"] = np.nan

    # 最終
    if p3 is not None:
        id3 = id_cols.get("p3")
        s3, _ = get_dept_series(p3, p1_depts)
        if not s3.empty:
            s3_aligned = s3.map(lambda x: align_dept_name(x, name_map))
            t3w = s3_aligned.value_counts().reset_index()
            t3w.columns = ["_dept_std", "最終志願次數"]
            base = base.merge(t3w, on="_dept_std", how="left")
            if id3 and id3 in p3.columns:
                dc3 = detect_dept_col(p3) or ("科系" if "科系" in p3.columns else None)
                if dc3:
                    p3c = p3.copy()
                    p3c["_dept_norm"] = p3c[dc3].apply(norm_dept).map(lambda x: align_dept_name(x, name_map))
                    t3h = p3c.groupby("_dept_norm")[id3].apply(lambda x: x.dropna().astype(str).str.strip().nunique()).reset_index()
                    t3h.columns = ["_dept_std", "最終人頭數"]
                    base = base.merge(t3h, on="_dept_std", how="left")

    if "最終志願次數" not in base.columns: base["最終志願次數"] = np.nan
    if "最終人頭數" not in base.columns: base["最終人頭數"] = np.nan

    # 填充
    for c in ["二階志願次數", "二階人頭數", "最終志願次數", "最終人頭數"]:
        base[c] = base[c].fillna(0).astype(int)

    # 重複率
    base["一階重複率(%)"] = np.where(
        (base["一階人頭數"].notna()) & (base["一階志願次數"] > 0) & (base["一階人頭數"] > 0),
        ((base["一階志願次數"] - base["一階人頭數"]) / base["一階志願次數"] * 100).round(1), None)
    base["二階重複率(%)"] = np.where(
        (base["二階人頭數"] > 0) & (base["二階志願次數"] > 0),
        ((base["二階志願次數"] - base["二階人頭數"]) / base["二階志願次數"] * 100).round(1), None)

    base["科系"] = base["_dept_std"].map(name_map).fillna(base["_dept_std"])
    base = base.drop(columns=["_dept_std"])

    col_order = ["科系",
                 "一階志願次數", "一階人頭數", "一階重複率(%)",
                 "二階志願次數", "二階人頭數", "二階重複率(%)",
                 "最終志願次數", "最終人頭數"]
    base = base[[c for c in col_order if c in base.columns]]
    return base


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
    if "二階人數" not in s.columns: s["二階人數"] = np.nan

    if p3 is not None:
        sc3 = detect_school_col(p3)
        if sc3:
            tmp3 = p3[[sc3]].copy()
            tmp3["_sch_std"] = tmp3[sc3].apply(norm_school)
            t3 = tmp3["_sch_std"].value_counts().reset_index()
            t3.columns = ["_sch_std", "最終入學"]
            s = s.merge(t3, on="_sch_std", how="left")
    if "最終入學" not in s.columns: s["最終入學"] = np.nan

    s["二階人數"] = s["二階人數"].fillna(0).astype(int)
    s["最終入學"] = s["最終入學"].fillna(0).astype(int)
    s["一→二階(%)"] = np.where(s["一階人數"] > 0, (s["二階人數"] / s["一階人數"] * 100).round(1), 0)
    s["二→最終(%)"] = np.where(s["二階人數"] > 0, (s["最終入學"] / s["二階人數"] * 100).round(1), 0)
    s["一→最終(%)"] = np.where(s["一階人數"] > 0, (s["最終入學"] / s["一階人數"] * 100).round(1), 0)
    s["流失人數"] = s["一階人數"] - s["最終入學"]
    s["二階流失"] = s["二階人數"] - s["最終入學"]
    s["效率評等"] = s["一→最終(%)"].apply(eff_stars)
    name_map = tmp1.drop_duplicates("_sch_std").set_index("_sch_std")[sc1]
    s["學校"] = s["_sch_std"].map(name_map).fillna(s["_sch_std"])
    s = s.drop(columns=["_sch_std"])
    col_order = ["學校", "一階人數", "二階人數", "最終入學",
                 "一→二階(%)", "二→最終(%)", "一→最終(%)",
                 "流失人數", "二階流失", "效率評等"]
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


def fig_grouped_bar(df, y, vals, title):
    fig = go.Figure()
    colors = ["#2196F3", "#9C27B0", "#4CAF50", "#FF9800"]
    for i, v in enumerate(vals):
        if v in df.columns:
            fig.add_trace(go.Bar(
                name=v, y=df[y], x=df[v], orientation="h",
                marker_color=colors[i % len(colors)], text=df[v], textposition="outside"
            ))
    fig.update_layout(barmode="group", title=title,
                      height=max(400, len(df) * 35),
                      yaxis=dict(autorange="reversed"))
    return fig


def fig_map(df, size_col, title, color_col=None):
    if "lat" not in df.columns or "lon" not in df.columns:
        return None
    m = df.dropna(subset=["lat", "lon"]).copy()
    if m.empty: return None
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


def fig_three_rates_bar(df, y_col, title):
    rate_cols = [c for c in ["一→二階(%)", "二→最終(%)", "一→最終(%)"] if c in df.columns]
    if not rate_cols: return None
    return fig_grouped_bar(df.sort_values(rate_cols[-1], ascending=True), y_col, rate_cols, title)


def fig_three_rates_trend(sdf, title):
    fig = go.Figure()
    rate_info = [
        ("一→二階(%)", "#2196F3", "一→二階"),
        ("二→最終(%)", "#9C27B0", "二→最終"),
        ("一→最終(%)", "#4CAF50", "一→最終"),
    ]
    for col, color, label in rate_info:
        if col in sdf.columns and sdf[col].sum() > 0:
            fig.add_trace(go.Scatter(
                x=sdf["年度"], y=sdf[col], name=label,
                mode="lines+markers+text", text=sdf[col],
                textposition="top center",
                line=dict(width=3, color=color), marker=dict(size=10)
            ))
    fig.update_layout(title=title, yaxis_title="轉換率(%)", height=420)
    return fig


# ============================================================
# v6.7 全年度惡化偵測引擎（保留）
# ============================================================
def detect_deterioration_full(data_by_year, metric, entity_col, min_count=0, min_p2=False):
    yr_list = sorted(data_by_year.keys())
    if len(yr_list) < 2: return [], []
    entity_series = {}
    for yr in yr_list:
        df = data_by_year[yr]
        if df is None or df.empty: continue
        if min_p2 and "二階人數" in df.columns: df = df[df["二階人數"] > 0]
        if min_count > 0 and "一階人數" in df.columns: df = df[df["一階人數"] >= min_count]
        for _, row in df.iterrows():
            ent = row[entity_col]
            if ent not in entity_series: entity_series[ent] = {}
            if metric in row.index: entity_series[ent][yr] = row[metric]

    pair_results = []
    for i in range(len(yr_list) - 1):
        prev_yr, curr_yr = yr_list[i], yr_list[i + 1]
        for ent, vals in entity_series.items():
            if prev_yr in vals and curr_yr in vals:
                pv, cv = vals[prev_yr], vals[curr_yr]
                if pv > 0 and cv < pv:
                    drop = round(pv - cv, 1)
                    sev = "🔴 嚴重" if drop > 10 else ("🟠 注意" if drop > 5 else "🟡 輕微")
                    pair_results.append({
                        entity_col: ent, "比較區間": f"{prev_yr} → {curr_yr}",
                        f"{prev_yr}": pv, f"{curr_yr}": cv, "下降幅度": drop, "嚴重程度": sev})

    consecutive_results = []
    for ent, vals in entity_series.items():
        sorted_yrs = [y for y in yr_list if y in vals]
        if len(sorted_yrs) < 2: continue
        consec_count = 0; drops = []
        for i in range(len(sorted_yrs) - 1):
            if vals[sorted_yrs[i + 1]] < vals[sorted_yrs[i]]:
                consec_count += 1
                drops.append(round(vals[sorted_yrs[i]] - vals[sorted_yrs[i + 1]], 1))
            else:
                consec_count = 0; drops = []
        if consec_count >= 2:
            first_val = vals[sorted_yrs[-(consec_count + 1)]]
            last_val = vals[sorted_yrs[-1]]
            consecutive_results.append({
                entity_col: ent, "連續下降年數": consec_count,
                "起始年度": sorted_yrs[-(consec_count + 1)],
                "最新年度": sorted_yrs[-1],
                "累積下降": round(first_val - last_val, 1),
                "各年度值": {y: vals[y] for y in sorted_yrs},
                "逐年跌幅": drops[-consec_count:],
                "加速中": drops[-1] > drops[-2] if len(drops) >= 2 else False})

    pair_results.sort(key=lambda x: x["下降幅度"], reverse=True)
    consecutive_results.sort(key=lambda x: x["累積下降"], reverse=True)
    return pair_results, consecutive_results


def detect_school_deterioration_full(year_cache, metric, min_count=2):
    data_by_year = {}
    name_lookup = {}
    for yr, c in year_cache.items():
        ss = c.get("ss")
        if ss is not None:
            data_by_year[yr] = ss.rename(columns={"學校": "_entity"}).copy()
            data_by_year[yr]["_entity"] = data_by_year[yr]["_entity"].apply(norm_school)
            for _, row in ss.iterrows():
                name_lookup[norm_school(row["學校"])] = row["學校"]
    if not data_by_year: return [], []

    yr_list = sorted(data_by_year.keys())
    entity_series = {}
    for yr in yr_list:
        df = data_by_year[yr]
        if min_count > 0 and "一階人數" in df.columns: df = df[df["一階人數"] >= min_count]
        if "二→最終" in metric and "二階人數" in df.columns: df = df[df["二階人數"] > 0]
        for _, row in df.iterrows():
            ent = row["_entity"]
            if ent not in entity_series: entity_series[ent] = {}
            if metric in row.index: entity_series[ent][yr] = row[metric]

    pair_results = []
    for i in range(len(yr_list) - 1):
        prev_yr, curr_yr = yr_list[i], yr_list[i + 1]
        for ent, vals in entity_series.items():
            if prev_yr in vals and curr_yr in vals:
                pv, cv = vals[prev_yr], vals[curr_yr]
                if pv > 0 and cv < pv:
                    drop = round(pv - cv, 1)
                    sev = "🔴 嚴重" if drop > 10 else ("🟠 注意" if drop > 5 else "🟡 輕微")
                    pair_results.append({
                        "學校": name_lookup.get(ent, ent),
                        "比較區間": f"{prev_yr} → {curr_yr}",
                        f"{prev_yr}": pv, f"{curr_yr}": cv, "下降幅度": drop, "嚴重程度": sev})

    consecutive_results = []
    for ent, vals in entity_series.items():
        sorted_yrs = [y for y in yr_list if y in vals]
        if len(sorted_yrs) < 2: continue
        consec_count = 0; drops = []
        for i in range(len(sorted_yrs) - 1):
            vc = vals.get(sorted_yrs[i + 1]); vp = vals.get(sorted_yrs[i])
            if vc is not None and vp is not None and vc < vp:
                consec_count += 1
                drops.append(round(vp - vc, 1))
            else:
                consec_count = 0; drops = []
        if consec_count >= 2:
            consecutive_results.append({
                "學校": name_lookup.get(ent, ent),
                "連續下降年數": consec_count,
                "起始年度": sorted_yrs[-(consec_count + 1)],
                "最新年度": sorted_yrs[-1],
                "累積下降": round(vals.get(sorted_yrs[-(consec_count + 1)], 0) - vals.get(sorted_yrs[-1], 0), 1),
                "加速中": drops[-1] > drops[-2] if len(drops) >= 2 else False})

    pair_results.sort(key=lambda x: x["下降幅度"], reverse=True)
    consecutive_results.sort(key=lambda x: x["累積下降"], reverse=True)
    return pair_results, consecutive_results


# ============================================================
# Session State
# ============================================================
if "years" not in st.session_state: st.session_state["years"] = {}
if "all_files" not in st.session_state: st.session_state["all_files"] = {}
if "analysis_ready" not in st.session_state: st.session_state["analysis_ready"] = False
if "analysis_version" not in st.session_state: st.session_state["analysis_version"] = 0

# ============================================================
# Sidebar
# ============================================================
with st.sidebar:
    st.header("📂 資料管理")
    uploaded = st.file_uploader(
        "上傳所有招生資料 (Excel/CSV)", type=["xlsx", "xls", "csv"],
        accept_multiple_files=True, help="可一次上傳多個年度的檔案")
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
                "id_col_p1": None, "id_col_p2": None, "id_col_p3": None,
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

            # === v6.8 學生ID欄位偵測 ===
            for phase_key, phase_label in [("p1", "一階"), ("p2", "二階"), ("p3", "最終")]:
                fname = ydata.get(phase_key)
                if fname and fname != "-- 未選擇 --" and fname in st.session_state["all_files"]:
                    df_tmp = st.session_state["all_files"][fname]
                    auto_id = detect_student_id_col(df_tmp)
                    id_key = f"id_col_{phase_key}"

                    if auto_id:
                        nuniq = df_tmp[auto_id].dropna().nunique()
                        total = len(df_tmp)
                        st.caption(f"🆔 {phase_label} 自動偵測：「{auto_id}」（{nuniq}唯一/{total}列）")
                        ydata[id_key] = auto_id
                    else:
                        manual_id = st.selectbox(
                            f"🆔 {phase_label} 學生ID欄位：",
                            ["-- 無 --"] + list(df_tmp.columns),
                            key=f"mid_{phase_key}_{yr}")
                        if manual_id != "-- 無 --":
                            ydata[id_key] = manual_id
                        else:
                            ydata[id_key] = None

            # 一階經緯度診斷
            if ydata["p1"] and ydata["p1"] != "-- 未選擇 --":
                p1df = st.session_state["all_files"][ydata["p1"]]
                lat_c, lon_c = detect_lat_lon_cols(p1df)
                if lat_c and lon_c:
                    n_valid = p1df[[lat_c, lon_c]].dropna().shape[0]
                    st.caption(f"📍 經緯度：{lat_c}/{lon_c}（{n_valid}筆有效）")
                else:
                    st.caption("⚠️ 一階未偵測到經緯度欄位")

            # 最終入學科系/班級/管道偵測
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
                        f'<div class="mapping-box">📋 偵測到班級欄位：「{cc3}」<br>'
                        f'v6.8 解析：[學制][科系縮寫→展開][年級][班別]</div>',
                        unsafe_allow_html=True)
                    sample = p3df[cc3].value_counts().head(8)
                    for cn, cnt in sample.items():
                        parsed = parse_class_name(str(cn))
                        if parsed:
                            prog, dept_kw = parsed[0], parsed[1]
                            dept_kw_raw = parsed[4] if len(parsed) == 5 else dept_kw
                            dept = resolve_dept_from_keyword(dept_kw, preview_p1_depts, dept_kw_raw)
                            abbrev_note = f"「{dept_kw_raw}」→「{dept_kw}」" if dept_kw_raw != dept_kw else f"「{dept_kw}」"
                            if dept:
                                st.caption(f"　✅ {cn}（{cnt}人）→ [{prog}]{abbrev_note} → {dept}")
                            else:
                                st.caption(f"　❌ {cn}（{cnt}人）→ [{prog}]{abbrev_note} → 找不到")
                        else:
                            st.caption(f"　❌ {cn}（{cnt}人）→ 無法解析")
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
        if geo is not None: p2 = enrich_geo(p2, geo)
    if s3 and s3 != "-- 未選擇 --" and s3 in st.session_state["all_files"]:
        p3 = st.session_state["all_files"][s3].copy()
        ch_col = ydata.get("channel_col")
        sel_chs = ydata.get("selected_channels")
        if ch_col and ch_col in p3.columns and sel_chs:
            p3[ch_col] = p3[ch_col].fillna("(空白)").astype(str).str.strip()
            p3.loc[p3[ch_col] == "", ch_col] = "(空白)"
            p3 = p3[p3[ch_col].isin(sel_chs)]
        if geo is not None: p3 = enrich_geo(p3, geo)
        dc3 = detect_dept_col(p3)
        if dc3 is None:
            cc3 = detect_class_col(p3) or ydata.get("class_col_override")
            if cc3 and cc3 in p3.columns:
                p1_depts = None
                if p1 is not None:
                    dc1 = detect_dept_col(p1)
                    if dc1: p1_depts = p1[dc1].dropna().unique().tolist()
                p3, mapping, match_detail = auto_map_class_to_dept(p3, cc3, p1_depts)
                p3["科系"] = p3["_mapped_dept"]
                p3 = p3.drop(columns=["_mapped_dept"], errors="ignore")

    id_cols = {
        "p1": ydata.get("id_col_p1"),
        "p2": ydata.get("id_col_p2"),
        "p3": ydata.get("id_col_p3"),
    }
    return p1, p2, p3, geo, ydata.get("channel_col"), id_cols


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
# v6.8 人頭數分析模組
# ============================================================
def render_headcount_analysis(yr, p1, p2, p3, id_cols):
    """各階段人頭數 vs 志願次數分析"""
    st.subheader(f"👤 {yr} — 各階段人頭數 vs 志願次數分析")

    st.markdown(
        '<div class="dedup-box">'
        '📌 <b>為什麼需要這個分析？</b><br>'
        '在一階和二階中，一位學生可能填報多個科系志願，導致「科系人數加總 > 實際人數」。<br>'
        '此模組透過<b>學生ID</b>去重，顯示每個階段的<b>真實人頭數</b>。'
        '</div>', unsafe_allow_html=True)

    phases = [("一階", p1, id_cols.get("p1")),
              ("二階", p2, id_cols.get("p2")),
              ("最終", p3, id_cols.get("p3"))]

    # === 全校摘要 ===
    st.markdown("#### 📊 全校各階段人頭數摘要")

    summary_rows = []
    phase_hc = {}
    for label, df, id_col in phases:
        if df is None:
            summary_rows.append({
                "階段": label, "總列數(志願次數)": "—", "學生ID欄位": "—",
                "不重複人頭數": "—", "重複列數": "—", "重複率(%)": "—",
                "平均每人填報數": "—"
            })
            continue
        dc = detect_dept_col(df) or ("科系" if "科系" in df.columns else None)
        hc = compute_headcount_stats(df, id_col, dc)
        phase_hc[label] = hc

        if hc["unique_headcount"] is not None:
            summary_rows.append({
                "階段": label,
                "總列數(志願次數)": hc["total_rows"],
                "學生ID欄位": hc["id_col"],
                "不重複人頭數": hc["unique_headcount"],
                "重複列數": hc["dup_count"],
                "重複率(%)": hc["dup_rate"],
                "平均每人填報數": hc["avg_entries_per_person"],
            })
        else:
            summary_rows.append({
                "階段": label,
                "總列數(志願次數)": hc["total_rows"],
                "學生ID欄位": "⚠️ 無ID欄位",
                "不重複人頭數": f"⚠️ 無法計算（{hc['total_rows']}列）",
                "重複列數": "—", "重複率(%)": "—", "平均每人填報數": "—",
            })

    sdf = pd.DataFrame(summary_rows)
    st.dataframe(sdf, use_container_width=True, hide_index=True)

    # KPI 卡片
    cols = st.columns(3)
    for i, (label, df, id_col) in enumerate(phases):
        hc = phase_hc.get(label)
        if hc and hc.get("unique_headcount") is not None:
            card_class = ["metric-blue", "metric-orange", "metric-green"][i]
            with cols[i]:
                st.markdown(
                    f'<div class="metric-card {card_class}">'
                    f'<h3>{label}</h3>'
                    f'<h1>{hc["unique_headcount"]:,} 人</h1>'
                    f'<h3>列數 {hc["total_rows"]:,} ｜ 重複率 {hc["dup_rate"]}%</h3>'
                    f'</div>', unsafe_allow_html=True)

    # === 人頭數 vs 志願次數漏斗 ===
    has_headcount = any(phase_hc.get(l, {}).get("unique_headcount") is not None for l in ["一階", "二階", "最終"])
    if has_headcount:
        st.markdown("---")
        st.markdown("#### 🔄 人頭數漏斗 vs 志願次數漏斗")

        c1, c2 = st.columns(2)
        with c1:
            fl, fv = [], []
            for label in ["一階", "二階", "最終"]:
                hc = phase_hc.get(label)
                if hc and hc.get("unique_headcount") is not None:
                    fl.append(f"{label}人頭")
                    fv.append(hc["unique_headcount"])
            if len(fv) > 1:
                st.plotly_chart(fig_funnel(fl, fv, "人頭數漏斗（去重）"), use_container_width=True)

        with c2:
            fl2, fv2 = [], []
            for label in ["一階", "二階", "最終"]:
                hc = phase_hc.get(label)
                if hc:
                    fl2.append(f"{label}列數")
                    fv2.append(hc["total_rows"])
            if len(fv2) > 1:
                st.plotly_chart(fig_funnel(fl2, fv2, "志願次數漏斗（含重複）"), use_container_width=True)

        # 人頭數轉換率
        st.markdown("#### 📈 人頭數轉換率 vs 列數轉換率")
        tr_rows = []
        p_labels = ["一階", "二階", "最終"]
        for i in range(len(p_labels) - 1):
            hc_from = phase_hc.get(p_labels[i])
            hc_to = phase_hc.get(p_labels[i + 1])
            if hc_from and hc_to:
                row_from = hc_from["total_rows"]
                row_to = hc_to["total_rows"]
                head_from = hc_from.get("unique_headcount")
                head_to = hc_to.get("unique_headcount")

                r = {"轉換": f"{p_labels[i]} → {p_labels[i+1]}",
                     "列數轉換率(%)": safe_pct(row_to, row_from)}
                if head_from and head_to:
                    r["人頭數轉換率(%)"] = safe_pct(head_to, head_from)
                    r["差異"] = round(r["人頭數轉換率(%)"] - r["列數轉換率(%)"], 1)
                    r["說明"] = "人頭數轉換率較高→列數因重複被稀釋" if r["差異"] > 0 else "列數轉換率較高→下一階重複數增加"
                else:
                    r["人頭數轉換率(%)"] = "—"
                    r["差異"] = "—"
                    r["說明"] = "無法計算（缺少ID欄位）"
                tr_rows.append(r)

        if tr_rows:
            st.dataframe(pd.DataFrame(tr_rows), use_container_width=True, hide_index=True)

    # === 各科系人頭數 ===
    st.markdown("---")
    st.markdown("#### 🏫 各科系人頭數 vs 志願次數")

    dc1 = detect_dept_col(p1)
    if dc1 and p1 is not None:
        hc_dept = build_dept_stats_headcount(
            p1, p2, p3,
            id_cols=id_cols
        )
        if hc_dept is not None and not hc_dept.empty:
            st.dataframe(hc_dept.sort_values(
                "一階志願次數" if "一階志願次數" in hc_dept.columns else hc_dept.columns[1],
                ascending=False), use_container_width=True, hide_index=True)

            # 一階：志願次數 vs 人頭數比較
            if "一階人頭數" in hc_dept.columns and hc_dept["一階人頭數"].notna().any():
                valid_dept = hc_dept[hc_dept["一階人頭數"].notna() & (hc_dept["一階人頭數"] > 0)].copy()
                if not valid_dept.empty:
                    fig = go.Figure()
                    fig.add_trace(go.Bar(
                        name="志願次數", y=valid_dept["科系"], x=valid_dept["一階志願次數"],
                        orientation="h", marker_color="#2196F3"))
                    fig.add_trace(go.Bar(
                        name="人頭數", y=valid_dept["科系"], x=valid_dept["一階人頭數"],
                        orientation="h", marker_color="#4CAF50"))
                    fig.update_layout(barmode="group",
                                      title="一階：志願次數 vs 人頭數（各科系）",
                                      height=max(400, len(valid_dept) * 30),
                                      yaxis=dict(autorange="reversed"))
                    st.plotly_chart(fig, use_container_width=True)

                    # 重複率 bar
                    if "一階重複率(%)" in valid_dept.columns:
                        valid_dup = valid_dept[valid_dept["一階重複率(%)"].notna()].copy()
                        if not valid_dup.empty:
                            fig2 = px.bar(
                                valid_dup.sort_values("一階重複率(%)", ascending=True),
                                x="一階重複率(%)", y="科系", orientation="h",
                                text="一階重複率(%)",
                                title="一階各科系重複率（同一人跨科填報比例）",
                                color="一階重複率(%)", color_continuous_scale="OrRd")
                            fig2.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
                            fig2.update_layout(height=max(380, len(valid_dup) * 28))
                            st.plotly_chart(fig2, use_container_width=True)

            # 全階段對照
            if all(c in hc_dept.columns for c in ["一階人頭數", "最終人頭數"]):
                valid_all = hc_dept[(hc_dept["一階人頭數"].notna()) & (hc_dept["一階人頭數"] > 0)].copy()
                if not valid_all.empty and valid_all["最終人頭數"].sum() > 0:
                    st.markdown("---")
                    st.markdown("#### 🎯 人頭數三段轉換率（去重後）")
                    valid_all["人頭一→最終(%)"] = np.where(
                        valid_all["一階人頭數"] > 0,
                        (valid_all["最終人頭數"] / valid_all["一階人頭數"] * 100).round(1), 0)
                    if "二階人頭數" in valid_all.columns:
                        valid_all["人頭一→二(%)"] = np.where(
                            valid_all["一階人頭數"] > 0,
                            (valid_all["二階人頭數"] / valid_all["一階人頭數"] * 100).round(1), 0)
                        valid_all["人頭二→最終(%)"] = np.where(
                            valid_all["二階人頭數"] > 0,
                            (valid_all["最終人頭數"] / valid_all["二階人頭數"] * 100).round(1), 0)
                    display_cols = ["科系", "一階人頭數"]
                    if "二階人頭數" in valid_all.columns:
                        display_cols.extend(["二階人頭數", "人頭一→二(%)", "人頭二→最終(%)"])
                    display_cols.extend(["最終人頭數", "人頭一→最終(%)"])
                    st.dataframe(valid_all[display_cols].sort_values("人頭一→最終(%)", ascending=False),
                                 use_container_width=True, hide_index=True)
    else:
        st.info("⚠️ 一階未偵測到科系欄位，無法進行科系人頭數分析。")

    # === 跨科填報學生分析 ===
    st.markdown("---")
    st.markdown("#### 🔀 跨科填報學生分析（一人填多個科系）")
    hc1 = phase_hc.get("一階")
    if hc1 and hc1.get("unique_headcount") is not None and id_cols.get("p1"):
        id1 = id_cols["p1"]
        dc1 = detect_dept_col(p1)
        if dc1 and id1 in p1.columns:
            p1_clean = p1[[id1, dc1]].dropna().copy()
            p1_clean[id1] = p1_clean[id1].astype(str).str.strip()
            dept_per_student = p1_clean.groupby(id1)[dc1].nunique().reset_index()
            dept_per_student.columns = ["學生ID", "填報科系數"]

            multi = dept_per_student[dept_per_student["填報科系數"] > 1]
            single = dept_per_student[dept_per_student["填報科系數"] == 1]

            c1, c2, c3 = st.columns(3)
            c1.metric("📊 總學生數", f"{len(dept_per_student):,}")
            c2.metric("1️⃣ 只填一科", f"{len(single):,}（{len(single)/len(dept_per_student)*100:.1f}%）")
            c3.metric("🔀 填多科", f"{len(multi):,}（{len(multi)/len(dept_per_student)*100:.1f}%）")

            dist = dept_per_student["填報科系數"].value_counts().sort_index().reset_index()
            dist.columns = ["填報科系數", "學生數"]
            fig = px.bar(dist, x="填報科系數", y="學生數", text="學生數",
                         title="每位學生填報科系數分布")
            fig.update_traces(marker_color="#667eea", textposition="outside")
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)

            # 跨科填報的科系組合 TOP 10
            if not multi.empty:
                multi_ids = multi["學生ID"].tolist()
                multi_students = p1_clean[p1_clean[id1].isin(multi_ids)]
                combos = multi_students.groupby(id1)[dc1].apply(
                    lambda x: " + ".join(sorted(x.unique()))
                ).value_counts().head(10).reset_index()
                combos.columns = ["科系組合", "學生數"]
                st.markdown("**跨科填報 TOP 10 組合：**")
                st.dataframe(combos, use_container_width=True, hide_index=True)
    else:
        st.info("⚠️ 需要一階的學生ID欄位才能分析跨科填報。")


# ============================================================
# 欄位診斷
# ============================================================
def show_field_diagnosis(p1, p2, p3, yr_label, id_cols=None):
    with st.expander(f"🔍 {yr_label} 欄位偵測與映射診斷", expanded=False):
        diag = []
        for phase, df, label in [("一階", p1, "P1"), ("二階", p2, "P2"), ("最終", p3, "P3")]:
            if df is None:
                diag.append({"階段": phase, "科系欄位": "—", "班級欄位": "—",
                             "學校欄位": "—", "ID欄位": "—", "筆數": 0})
                continue
            dc = detect_dept_col(df)
            cc = detect_class_col(df)
            sc = detect_school_col(df)
            id_c = id_cols.get({"一階": "p1", "二階": "p2", "最終": "p3"}.get(phase)) if id_cols else None
            diag.append({
                "階段": phase,
                "科系欄位": dc if dc else ("科系(映射)" if "科系" in df.columns else "❌"),
                "班級欄位": cc if cc else "—",
                "學校欄位": sc if sc else "❌",
                "ID欄位": id_c if id_c else "❌ 無",
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


# ============================================================
# 單年度模組
# ============================================================
def render_year_analysis(yr):
    p1, p2, p3, geo, ch_col, id_cols = get_year_dfs(yr)
    if p1 is None:
        st.warning(f"⚠️ {yr}：一階資料未指定或無法讀取。")
        return
    mod_opts = ["📊 總覽儀表板", "👤 人頭數分析", "🔄 招生漏斗", "📈 入學管道",
                "🗺️ 地理分布", "🏫 科系熱力圖", "🎯 來源學校", "⚠️ 流失預警"]
    mod = st.radio("選擇分析模組：", mod_opts, horizontal=True, key=f"mod_{yr}")
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    n1 = len(p1)
    n2 = len(p2) if p2 is not None else None
    n3 = len(p3) if p3 is not None else None

    # ─── v6.8 人頭數分析 ───
    if "人頭數" in mod:
        render_headcount_analysis(yr, p1, p2, p3, id_cols)
        return

    # ─── 1. 總覽 ───
    if "總覽" in mod:
        st.subheader(f"📊 {yr} — 總覽儀表板")

        # 計算人頭數
        hc_summary = {}
        for phase_key, df_phase in [("p1", p1), ("p2", p2), ("p3", p3)]:
            if df_phase is not None:
                id_c = id_cols.get(phase_key)
                hc = compute_headcount_stats(df_phase, id_c)
                hc_summary[phase_key] = hc

        # 第一排：列數（志願次數）
        cols = st.columns(6)
        with cols[0]:
            st.markdown(f'<div class="metric-card"><h3>一階列數(志願次)</h3><h1>{n1:,}</h1></div>', unsafe_allow_html=True)
        with cols[1]:
            v = f"{n2:,}" if n2 else "—"
            st.markdown(f'<div class="metric-card metric-orange"><h3>二階列數(志願次)</h3><h1>{v}</h1></div>', unsafe_allow_html=True)
        with cols[2]:
            v = f"{n3:,}" if n3 else "—"
            st.markdown(f'<div class="metric-card metric-green"><h3>最終入學</h3><h1>{v}</h1></div>', unsafe_allow_html=True)
        with cols[3]:
            r = f"{n2/n1*100:.1f}%" if n2 and n1 else "—"
            st.markdown(f'<div class="metric-card metric-blue"><h3>一→二階(列)</h3><h1>{r}</h1></div>', unsafe_allow_html=True)
        with cols[4]:
            r = f"{n3/n2*100:.1f}%" if n3 and n2 else "—"
            st.markdown(f'<div class="metric-card metric-purple"><h3>二→最終(列)</h3><h1>{r}</h1></div>', unsafe_allow_html=True)
        with cols[5]:
            r = f"{n3/n1*100:.1f}%" if n3 and n1 else "—"
            st.markdown(f'<div class="metric-card metric-gold"><h3>一→最終(列)</h3><h1>{r}</h1></div>', unsafe_allow_html=True)

        # 第二排：人頭數
        any_hc = any(hc_summary.get(k, {}).get("unique_headcount") is not None for k in ["p1", "p2", "p3"])
        if any_hc:
            st.markdown("##### 👤 人頭數（去重後）")
            cols2 = st.columns(6)
            hc1_n = hc_summary.get("p1", {}).get("unique_headcount")
            hc2_n = hc_summary.get("p2", {}).get("unique_headcount")
            hc3_n = hc_summary.get("p3", {}).get("unique_headcount")
            with cols2[0]:
                v = f"{hc1_n:,}" if hc1_n else "—"
                st.markdown(f'<div class="metric-card metric-cyan"><h3>一階人頭數</h3><h1>{v}</h1></div>', unsafe_allow_html=True)
            with cols2[1]:
                v = f"{hc2_n:,}" if hc2_n else "—"
                st.markdown(f'<div class="metric-card metric-red"><h3>二階人頭數</h3><h1>{v}</h1></div>', unsafe_allow_html=True)
            with cols2[2]:
                v = f"{hc3_n:,}" if hc3_n else "—"
                st.markdown(f'<div class="metric-card metric-lime"><h3>最終人頭數</h3><h1>{v}</h1></div>', unsafe_allow_html=True)
            with cols2[3]:
                r = f"{hc2_n/hc1_n*100:.1f}%" if hc1_n and hc2_n else "—"
                st.markdown(f'<div class="metric-card metric-teal"><h3>一→二階(人頭)</h3><h1>{r}</h1></div>', unsafe_allow_html=True)
            with cols2[4]:
                r = f"{hc3_n/hc2_n*100:.1f}%" if hc2_n and hc3_n else "—"
                st.markdown(f'<div class="metric-card metric-purple"><h3>二→最終(人頭)</h3><h1>{r}</h1></div>', unsafe_allow_html=True)
            with cols2[5]:
                r = f"{hc3_n/hc1_n*100:.1f}%" if hc1_n and hc3_n else "—"
                st.markdown(f'<div class="metric-card metric-gold"><h3>一→最終(人頭)</h3><h1>{r}</h1></div>', unsafe_allow_html=True)

            # 列數 vs 人頭數比較提示
            if hc1_n and n1 and hc1_n != n1:
                dup = n1 - hc1_n
                st.markdown(
                    f'<div class="dedup-box">'
                    f'📌 一階：{n1} 筆志願 ÷ {hc1_n} 位學生 = 平均每人填 <b>{n1/hc1_n:.1f}</b> 個志願<br>'
                    f'重複列：<b>{dup}</b> 筆（{dup/n1*100:.1f}%）'
                    f'</div>', unsafe_allow_html=True)

        show_field_diagnosis(p1, p2, p3, yr, id_cols)

        if n2 and n3:
            loss_12 = n1 - n2; loss_23 = n2 - n3; loss_total = n1 - n3
            st.markdown(
                f'<div class="info-box">'
                f'📊 <b>三段流失分析（列數）</b><br>'
                f'一→二階流失：<b>{loss_12}</b> ｜二→最終流失：<b>{loss_23}</b> ｜總流失：<b>{loss_total}</b>'
                f'</div>', unsafe_allow_html=True)

        if p3 is not None and ch_col and ch_col in p3.columns:
            st.markdown("---"); st.subheader("🟢 最終入學管道分布")
            cd = p3[ch_col].value_counts().reset_index(); cd.columns = ["入學管道", "人數"]
            cd["佔比(%)"] = (cd["人數"] / cd["人數"].sum() * 100).round(1)
            c1, c2 = st.columns(2)
            with c1:
                fig = px.pie(cd, names="入學管道", values="人數", title="管道佔比", hole=.35)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = px.bar(cd.sort_values("人數", ascending=True), x="人數", y="入學管道", orientation="h",
                             text="人數", title="各管道人數", color="佔比(%)", color_continuous_scale="Viridis")
                fig.update_layout(height=max(400, len(cd)*28))
                st.plotly_chart(fig, use_container_width=True)

        fl, fv = ["一階報名"], [n1]
        if n2: fl.append("二階報到"); fv.append(n2)
        if n3: fl.append("最終入學"); fv.append(n3)
        if len(fv) > 1:
            st.plotly_chart(fig_funnel(fl, fv, f"{yr} 招生漏斗（列數）"), use_container_width=True)

        result = build_dept_stats(p1, p2, p3)
        if result and result[0] is not None:
            ds, p3_info = result
            st.markdown("---"); st.subheader("各科系三階段概覽（列數統計）")
            total_final_table = ds["最終入學"].sum()
            if n3 and total_final_table != n3:
                diff = n3 - total_final_table
                st.markdown(f'<div class="warning-box">⚠️ 科系加總（{total_final_table}）≠ 總列數（{n3}），差異 {diff}</div>', unsafe_allow_html=True)
            elif n3 and total_final_table == n3:
                st.markdown(f'<div class="success-box">✅ 科系加總（{total_final_table}）= 總列數（{n3}）完全吻合！</div>', unsafe_allow_html=True)
            st.dataframe(ds.sort_values("一階人數", ascending=False), use_container_width=True, hide_index=True)
            rate_fig = fig_three_rates_bar(ds, "科系", f"{yr} 各科系三段轉換率（列數）")
            if rate_fig: st.plotly_chart(rate_fig, use_container_width=True)

    # ─── 2. 漏斗 ───
    elif "漏斗" in mod:
        st.subheader(f"🔄 {yr} — 招生漏斗分析")
        result = build_dept_stats(p1, p2, p3)
        if result and result[0] is not None:
            ds, _ = result
            st.dataframe(ds.sort_values("一→最終(%)", ascending=False), use_container_width=True, hide_index=True)
            rate_fig = fig_three_rates_bar(ds, "科系", "各科系三段轉換率比較")
            if rate_fig: st.plotly_chart(rate_fig, use_container_width=True)
            if "二→最終(%)" in ds.columns and ds["二→最終(%)"].sum() > 0:
                st.markdown("---"); st.subheader("🟣 二→最終 轉換率分析")
                ds_p2 = ds[ds["二階人數"] > 0].copy()
                if not ds_p2.empty:
                    c1, c2 = st.columns(2)
                    with c1:
                        fig = px.bar(ds_p2.sort_values("二→最終(%)", ascending=True),
                                     x="二→最終(%)", y="科系", orientation="h", text="二→最終(%)",
                                     title="二→最終 轉換率排行", color="二→最終(%)", color_continuous_scale="RdYlGn")
                        fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
                        fig.update_layout(height=max(380, len(ds_p2)*28))
                        st.plotly_chart(fig, use_container_width=True)
                    with c2:
                        fig = px.bar(ds_p2.sort_values("二階流失", ascending=False),
                                     x="二階流失", y="科系", orientation="h", text="二階流失",
                                     title="二階報到後流失人數", color="二階流失", color_continuous_scale="OrRd")
                        fig.update_traces(textposition="outside")
                        fig.update_layout(height=max(380, len(ds_p2)*28))
                        st.plotly_chart(fig, use_container_width=True)
            st.markdown("---"); st.subheader("單科系漏斗")
            sel = st.selectbox("選擇科系：", ds["科系"].tolist(), key=f"fun_dept_{yr}")
            row = ds[ds["科系"] == sel].iloc[0]
            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("一階", f'{int(row["一階人數"])}')
            c2.metric("二階", f'{int(row["二階人數"])}')
            c3.metric("最終", f'{int(row["最終入學"])}')
            c4.metric("一→二階", f'{row["一→二階(%)"]:.1f}%')
            c5.metric("二→最終", f'{row["二→最終(%)"]:.1f}%')
            c6.metric("一→最終", f'{row["一→最終(%)"]:.1f}%')
            fl, fv = ["一階"], [int(row["一階人數"])]
            if row["二階人數"] > 0 or p2 is not None: fl.append("二階"); fv.append(int(row["二階人數"]))
            if row["最終入學"] > 0 or p3 is not None: fl.append("最終"); fv.append(int(row["最終入學"]))
            st.plotly_chart(fig_funnel(fl, fv, f"{sel} 漏斗"), use_container_width=True)

        ss = build_school_stats(p1, p2, p3)
        if ss is not None:
            st.markdown("---"); st.subheader("各來源學校漏斗")
            mn = st.slider("一階≥", 1, 50, 5, key=f"fun_mn_{yr}")
            sf = ss[ss["一階人數"] >= mn].sort_values("一→最終(%)", ascending=False)
            st.dataframe(sf, use_container_width=True, hide_index=True)

    # ─── 3. 管道 ───
    elif "管道" in mod:
        st.subheader(f"📈 {yr} — 入學管道分析")
        if p3 is None or not ch_col or ch_col not in (p3.columns if p3 is not None else []):
            st.warning("⚠️ 需要最終入學資料及入學方式欄位。"); return
        cd = p3[ch_col].value_counts().reset_index(); cd.columns = ["入學管道", "人數"]
        cd["佔比(%)"] = (cd["人數"] / cd["人數"].sum() * 100).round(1)
        c1, c2 = st.columns(2)
        with c1:
            fig = px.pie(cd, names="入學管道", values="人數", title="管道佔比", hole=.35)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = px.bar(cd.sort_values("人數", ascending=True), x="人數", y="入學管道", orientation="h",
                         text="人數", title="人數排行", color="佔比(%)", color_continuous_scale="Viridis")
            fig.update_layout(height=max(400, len(cd)*30))
            st.plotly_chart(fig, use_container_width=True)
        dept_col = detect_dept_col(p3) or ("科系" if "科系" in p3.columns else None)
        if dept_col:
            st.markdown("---"); st.subheader("管道 × 科系")
            valid_p3 = p3[p3[dept_col].notna()] if dept_col == "科系" else p3
            cross = valid_p3.groupby([ch_col, dept_col]).size().reset_index(name="人數")
            fig = px.bar(cross, x=ch_col, y="人數", color=dept_col, barmode="stack",
                         text="人數", title="管道×科系堆疊圖")
            fig.update_layout(height=600, xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)

    # ─── 4. 地理 ───
    elif "地理" in mod:
        st.subheader(f"🗺️ {yr} — 地理分布")
        if geo is None: st.warning("⚠️ 一階資料無經緯度欄位。"); return
        def do_map(src_df, count_label, title_text):
            sc_ = detect_school_col(src_df)
            if sc_ is None: return
            agg = src_df.groupby(sc_).size().reset_index(name=count_label)
            agg["_std"] = agg[sc_].apply(norm_school)
            agg = agg.merge(geo, on="_std", how="left").drop(columns=["_std"])
            fig = fig_map(agg, count_label, title_text)
            if fig: st.plotly_chart(fig, use_container_width=True)
        do_map(p1, "報名人數", f"{yr} 一階報名來源")
        if p2 is not None: st.markdown("---"); do_map(p2, "報到人數", f"{yr} 二階報到來源")
        if p3 is not None: st.markdown("---"); do_map(p3, "入學人數", f"{yr} 最終入學來源")

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
            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("一階", f'{int(r["一階人數"])}')
            c2.metric("二階", f'{int(r["二階人數"])}')
            c3.metric("最終", f'{int(r["最終入學"])}')
            c4.metric("一→二階", f'{r["一→二階(%)"]:.1f}%')
            c5.metric("二→最終", f'{r["二→最終(%)"]:.1f}%')
            c6.metric("一→最終", f'{r["一→最終(%)"]:.1f}%')
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

        if n2 and n3:
            st.markdown("#### 全校三段流失概覽")
            loss_data = {
                "階段": ["一→二階", "二→最終", "一→最終（總計）"],
                "起始人數": [n1, n2, n1], "終點人數": [n2, n3, n3],
                "流失人數": [n1-n2, n2-n3, n1-n3],
                "轉換率(%)": [round(n2/n1*100,1), round(n3/n2*100,1), round(n3/n1*100,1)],
                "流失率(%)": [round((n1-n2)/n1*100,1), round((n2-n3)/n2*100,1), round((n1-n3)/n1*100,1)]
            }
            st.dataframe(pd.DataFrame(loss_data), use_container_width=True, hide_index=True)

        rate_select = st.radio("分析維度：", ["一→最終(%)", "一→二階(%)", "二→最終(%)"],
                               horizontal=True, key=f"loss_rate_{yr}")
        mn = st.slider("一階≥", 1, 50, 10, key=f"loss_mn_{yr}")
        pool = ss[ss["一階人數"] >= mn].copy()
        if rate_select == "二→最終(%)": pool = pool[pool["二階人數"] > 0]; loss_col = "二階流失"
        else: loss_col = "流失人數"
        avg = pool[rate_select].mean()
        warn = pool[pool[rate_select] < avg].sort_values(loss_col, ascending=False)
        if warn.empty:
            st.success("✅ 沒有預警學校！")
        else:
            st.markdown(f'<div class="warning-box">⚠️ {len(warn)} 所學校的 {rate_select} 低於平均 {avg:.1f}%</div>', unsafe_allow_html=True)
            st.dataframe(warn, use_container_width=True, hide_index=True)

        st.markdown("---"); st.subheader("科系流失（三段）")
        result = build_dept_stats(p1, p2, p3)
        if result and result[0] is not None:
            ds, _ = result
            st.dataframe(ds.sort_values("流失人數", ascending=False), use_container_width=True, hide_index=True)


# ============================================================
# 跨年度模組
# ============================================================
def render_cross_year():
    st.header("📊 跨年度比較分析")
    if len(valid_years) < 2: st.warning("⚠️ 需要至少 2 個年度。"); return

    mod_opts = ["📊 總覽儀表板", "👤 人頭數趨勢", "🔄 招生漏斗", "📈 入學管道",
                "🗺️ 地理分布", "🏫 科系熱力圖", "🎯 來源學校", "⚠️ 流失預警"]
    mod = st.radio("選擇跨年度分析模組：", mod_opts, horizontal=True, key="cross_mod")
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    year_cache = {}
    for yr in valid_years:
        p1, p2, p3, geo, ch_col, id_cols = get_year_dfs(yr)
        if p1 is None: continue
        n1 = len(p1); n2 = len(p2) if p2 is not None else 0; n3 = len(p3) if p3 is not None else 0
        result = build_dept_stats(p1, p2, p3)
        ds = result[0] if result and result[0] is not None else None
        ss = build_school_stats(p1, p2, p3)
        # 人頭數
        hc_info = {}
        for pk, dfp in [("p1", p1), ("p2", p2), ("p3", p3)]:
            if dfp is not None:
                hc_info[pk] = compute_headcount_stats(dfp, id_cols.get(pk))
        year_cache[yr] = {
            "p1": p1, "p2": p2, "p3": p3, "geo": geo, "ch_col": ch_col, "id_cols": id_cols,
            "n1": n1, "n2": n2, "n3": n3, "ds": ds, "ss": ss, "hc": hc_info
        }

    if not year_cache: st.warning("⚠️ 沒有有效的年度資料。"); return
    yr_list = list(year_cache.keys())

    # ═══ 跨年度 ：總覽 ═══
    if "總覽" in mod and "人頭" not in mod:
        st.subheader("📊 跨年度總覽比較")
        summaries = []
        for yr, c in year_cache.items():
            row = {
                "年度": yr,
                "一階列數": c["n1"], "二階列數": c["n2"], "最終入學": c["n3"],
                "一→二階(%)(列)": safe_pct(c["n2"], c["n1"]),
                "二→最終(%)(列)": safe_pct(c["n3"], c["n2"]),
                "一→最終(%)(列)": safe_pct(c["n3"], c["n1"]),
            }
            # 人頭數
            hc = c.get("hc", {})
            for pk, label in [("p1", "一階"), ("p2", "二階"), ("p3", "最終")]:
                hci = hc.get(pk, {})
                n = hci.get("unique_headcount")
                row[f"{label}人頭數"] = n if n else "—"
            hc1 = hc.get("p1", {}).get("unique_headcount")
            hc2 = hc.get("p2", {}).get("unique_headcount")
            hc3 = hc.get("p3", {}).get("unique_headcount")
            row["一→二階(%)(人頭)"] = safe_pct(hc2, hc1) if hc1 and hc2 else "—"
            row["一→最終(%)(人頭)"] = safe_pct(hc3, hc1) if hc1 and hc3 else "—"
            summaries.append(row)
        sdf = pd.DataFrame(summaries)
        st.dataframe(sdf, use_container_width=True, hide_index=True)

        fig = go.Figure()
        fig.add_trace(go.Bar(x=sdf["年度"], y=sdf["一階列數"], name="一階列數", marker_color="#2196F3"))
        if sdf["二階列數"].sum() > 0:
            fig.add_trace(go.Bar(x=sdf["年度"], y=sdf["二階列數"], name="二階列數", marker_color="#FF9800"))
        if sdf["最終入學"].sum() > 0:
            fig.add_trace(go.Bar(x=sdf["年度"], y=sdf["最終入學"], name="最終入學", marker_color="#4CAF50"))
        # 人頭數線
        hc1_vals = [s.get("一階人頭數") for s in summaries]
        if any(isinstance(v, int) for v in hc1_vals):
            fig.add_trace(go.Scatter(
                x=sdf["年度"], y=[v if isinstance(v, int) else None for v in hc1_vals],
                name="一階人頭數", mode="lines+markers", line=dict(dash="dot", width=3, color="#E91E63"),
                marker=dict(size=12, symbol="diamond")))
        fig.update_layout(barmode="group", title="各年度招生量（列數 vs 人頭數）", height=450)
        st.plotly_chart(fig, use_container_width=True)

        rate_fig = fig_three_rates_trend(
            pd.DataFrame([{"年度": s["年度"],
                           "一→二階(%)": s["一→二階(%)(列)"],
                           "二→最終(%)": s["二→最終(%)(列)"],
                           "一→最終(%)": s["一→最終(%)(列)"]} for s in summaries]),
            "全校三段轉換率趨勢（列數）")
        if rate_fig: st.plotly_chart(rate_fig, use_container_width=True)

    # ═══ 跨年度：人頭數趨勢 ═══
    elif "人頭" in mod:
        st.subheader("👤 跨年度人頭數趨勢分析")
        st.markdown(
            '<div class="dedup-box">'
            '📌 此模組比較各年度各階段的「真實人頭數」變化，'
            '排除一人多志願造成的列數膨脹。'
            '</div>', unsafe_allow_html=True)

        hc_rows = []
        for yr, c in year_cache.items():
            hc = c.get("hc", {})
            row = {"年度": yr, "一階列數": c["n1"], "二階列數": c["n2"], "最終列數": c["n3"]}
            for pk, label in [("p1", "一階"), ("p2", "二階"), ("p3", "最終")]:
                hci = hc.get(pk, {})
                row[f"{label}人頭數"] = hci.get("unique_headcount")
                row[f"{label}重複率(%)"] = hci.get("dup_rate")
                row[f"{label}平均填報數"] = hci.get("avg_entries_per_person")
            hc_rows.append(row)
        hcdf = pd.DataFrame(hc_rows)
        st.dataframe(hcdf, use_container_width=True, hide_index=True)

        # 趨勢圖
        fig = go.Figure()
        for label, color, dash in [
            ("一階列數", "#2196F3", "solid"), ("一階人頭數", "#1565C0", "dot"),
            ("最終列數", "#4CAF50", "solid"), ("最終人頭數", "#2E7D32", "dot")]:
            vals = hcdf[label].tolist() if label in hcdf.columns else []
            if any(v is not None and v > 0 for v in vals):
                fig.add_trace(go.Scatter(
                    x=hcdf["年度"], y=vals, name=label,
                    mode="lines+markers", line=dict(dash=dash, width=3, color=color),
                    marker=dict(size=10)))
        fig.update_layout(title="列數 vs 人頭數趨勢", height=450, yaxis_title="人數")
        st.plotly_chart(fig, use_container_width=True)

        # 重複率趨勢
        has_dup = any(r.get("一階重複率(%)") is not None for r in hc_rows)
        if has_dup:
            st.markdown("---"); st.markdown("#### 📊 各階段重複率趨勢")
            fig = go.Figure()
            for label, color in [("一階重複率(%)", "#2196F3"), ("二階重複率(%)", "#FF9800"), ("最終重複率(%)", "#4CAF50")]:
                vals = hcdf[label].tolist() if label in hcdf.columns else []
                if any(v is not None and v > 0 for v in vals):
                    fig.add_trace(go.Scatter(
                        x=hcdf["年度"], y=vals, name=label,
                        mode="lines+markers+text", text=[f"{v}%" if v else "" for v in vals],
                        textposition="top center",
                        line=dict(width=3, color=color), marker=dict(size=10)))
            fig.update_layout(title="各階段重複率趨勢", height=400, yaxis_title="重複率(%)")
            st.plotly_chart(fig, use_container_width=True)

        # 人頭數轉換率
        st.markdown("---"); st.markdown("#### 🎯 人頭數轉換率 vs 列數轉換率（跨年度）")
        compare_rows = []
        for yr, c in year_cache.items():
            hc = c.get("hc", {})
            hc1 = hc.get("p1", {}).get("unique_headcount")
            hc3 = hc.get("p3", {}).get("unique_headcount")
            compare_rows.append({
                "年度": yr,
                "列數一→最終(%)": safe_pct(c["n3"], c["n1"]),
                "人頭一→最終(%)": safe_pct(hc3, hc1) if hc1 and hc3 else None,
                "差異(人頭-列數)": round(safe_pct(hc3, hc1) - safe_pct(c["n3"], c["n1"]), 1) if hc1 and hc3 else None,
            })
        cdf = pd.DataFrame(compare_rows)
        st.dataframe(cdf, use_container_width=True, hide_index=True)

        fig = go.Figure()
        fig.add_trace(go.Bar(x=cdf["年度"], y=cdf["列數一→最終(%)"], name="列數轉換率", marker_color="#2196F3"))
        hv = cdf["人頭一→最終(%)"].tolist()
        if any(v is not None for v in hv):
            fig.add_trace(go.Bar(x=cdf["年度"], y=hv, name="人頭數轉換率", marker_color="#4CAF50"))
        fig.update_layout(barmode="group", title="一→最終 轉換率比較（列數 vs 人頭數）", height=400)
        st.plotly_chart(fig, use_container_width=True)

    # ═══ 跨年度：漏斗 ═══
    elif "漏斗" in mod:
        st.subheader("🔄 跨年度招生漏斗比較")
        n_cols = len(year_cache)
        fig = make_subplots(rows=1, cols=n_cols,
                            subplot_titles=list(year_cache.keys()),
                            specs=[[{"type": "funnel"}] * n_cols])
        for i, (yr, c) in enumerate(year_cache.items()):
            fl, fv = ["一階"], [c["n1"]]
            if c["n2"]: fl.append("二階"); fv.append(c["n2"])
            if c["n3"]: fl.append("最終"); fv.append(c["n3"])
            fig.add_trace(go.Funnel(y=fl, x=fv, name=yr, textinfo="value+percent initial"), row=1, col=i+1)
        fig.update_layout(height=450, showlegend=True)
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("---"); st.markdown("#### 各科系轉換率跨年度比較")
        all_depts = set()
        for c in year_cache.values():
            if c["ds"] is not None: all_depts.update(c["ds"]["科系"].tolist())
        if all_depts:
            rate_metric = st.radio("轉換率指標：", ["一→二階(%)", "二→最終(%)", "一→最終(%)"],
                                   horizontal=True, key="cross_fun_rate")
            cross_dept_rows = []
            for yr, c in year_cache.items():
                if c["ds"] is not None:
                    for _, row in c["ds"].iterrows():
                        cross_dept_rows.append({"年度": yr, "科系": row["科系"],
                            rate_metric: row[rate_metric] if rate_metric in row.index else 0})
            if cross_dept_rows:
                cdf = pd.DataFrame(cross_dept_rows)
                fig = px.bar(cdf, x="科系", y=rate_metric, color="年度", barmode="group",
                             text=rate_metric, title=f"各科系 {rate_metric}（跨年度）")
                fig.update_layout(height=500, xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)

    # ═══ 跨年度：管道 ═══
    elif "管道" in mod:
        st.subheader("📈 跨年度入學管道比較")
        ch_data_all = []
        for yr, c in year_cache.items():
            if c["p3"] is not None and c["ch_col"] and c["ch_col"] in c["p3"].columns:
                cd = c["p3"][c["ch_col"]].value_counts().reset_index()
                cd.columns = ["入學管道", "人數"]; cd["年度"] = yr
                ch_data_all.append(cd)
        if not ch_data_all: st.warning("⚠️ 沒有管道資料。"); return
        all_ch = pd.concat(ch_data_all, ignore_index=True)
        fig = px.bar(all_ch, x="入學管道", y="人數", color="年度", barmode="group",
                     text="人數", title="各管道人數（跨年度）")
        fig.update_layout(height=500, xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

    # ═══ 跨年度：地理 ═══
    elif "地理" in mod:
        st.subheader("🗺️ 跨年度地理分布比較")
        phase_select = st.radio("選擇階段：", ["一階報名", "最終入學"], horizontal=True, key="cross_geo_phase")
        for yr, c in year_cache.items():
            if c["geo"] is None: continue
            st.markdown(f"### {yr}")
            src = c["p1"] if "一階" in phase_select else c["p3"]
            label = "報名人數" if "一階" in phase_select else "入學人數"
            if src is None: st.info(f"{yr}：無此階段資料"); continue
            sc = detect_school_col(src)
            if sc is None: continue
            agg = src.groupby(sc).size().reset_index(name=label)
            agg["_std"] = agg[sc].apply(norm_school)
            agg = agg.merge(c["geo"], on="_std", how="left").drop(columns=["_std"])
            fig = fig_map(agg, label, f"{yr} {phase_select}")
            if fig: st.plotly_chart(fig, use_container_width=True)

    # ═══ 跨年度：熱力圖 ═══
    elif "熱力圖" in mod:
        st.subheader("🏫 跨年度科系×年度 熱力圖")
        dept_year_rows = []
        for yr, c in year_cache.items():
            if c["ds"] is not None:
                for _, row in c["ds"].iterrows():
                    dept_year_rows.append({"年度": yr, "科系": row["科系"],
                        "一階人數": int(row["一階人數"]),
                        "最終入學": int(row["最終入學"]),
                        "一→最終(%)": row["一→最終(%)"]})
        if dept_year_rows:
            dydf = pd.DataFrame(dept_year_rows)
            metric = st.radio("指標：", ["一階人數", "最終入學", "一→最終(%)"],
                              horizontal=True, key="cross_hm_metric")
            pv = dydf.pivot_table(index="科系", columns="年度", values=metric, aggfunc="first").fillna(0)
            colorscale = "RdYlGn" if "%" in metric else "YlOrRd"
            fig = px.imshow(pv, text_auto=True, aspect="auto",
                            color_continuous_scale=colorscale, title=f"科系×年度：{metric}")
            fig.update_layout(height=max(400, len(pv) * 30))
            st.plotly_chart(fig, use_container_width=True)

    # ═══ 跨年度：來源學校 ═══
    elif "來源學校" in mod:
        st.subheader("🎯 跨年度來源學校追蹤")
        all_sch_names = set()
        for c in year_cache.values():
            if c["ss"] is not None: all_sch_names.update(c["ss"]["學校"].tolist())
        if all_sch_names:
            sel_sch = st.selectbox("選擇學校：", sorted(all_sch_names), key="cross_sch_sel")
            sch_rows = []
            for yr, c in year_cache.items():
                if c["ss"] is not None:
                    r = c["ss"][c["ss"]["學校"].apply(norm_school) == norm_school(sel_sch)]
                    if not r.empty:
                        r = r.iloc[0]
                        sch_rows.append({"年度": yr, "一階": int(r["一階人數"]),
                                         "二階": int(r["二階人數"]), "最終": int(r["最終入學"]),
                                         "一→二階(%)": r["一→二階(%)"],
                                         "二→最終(%)": r["二→最終(%)"],
                                         "一→最終(%)": r["一→最終(%)"]})
            if sch_rows:
                rdf = pd.DataFrame(sch_rows)
                st.dataframe(rdf, use_container_width=True, hide_index=True)
                rate_fig = fig_three_rates_trend(rdf, f"「{sel_sch}」三段轉換率趨勢")
                if rate_fig: st.plotly_chart(rate_fig, use_container_width=True)

    # ═══ 跨年度：流失預警（v6.7 全年度偵測）═══
    elif "流失" in mod:
        st.subheader("⚠️ 跨年度流失預警分析（全年度偵測）")
        st.markdown("#### 📊 全校三段流失趨勢")
        loss_summary = []
        for yr, c in year_cache.items():
            loss_summary.append({
                "年度": yr, "一階": c["n1"], "二階": c["n2"], "最終": c["n3"],
                "一→二階(%)": safe_pct(c["n2"], c["n1"]),
                "二→最終(%)": safe_pct(c["n3"], c["n2"]),
                "一→最終(%)": safe_pct(c["n3"], c["n1"]),
            })
        lsdf = pd.DataFrame(loss_summary)
        st.dataframe(lsdf, use_container_width=True, hide_index=True)
        rate_fig = fig_three_rates_trend(lsdf, "全校三段轉換率趨勢")
        if rate_fig: st.plotly_chart(rate_fig, use_container_width=True)

        st.markdown("---"); st.markdown("#### 🚨 科系惡化偵測")
        det_metric = st.radio("偵測指標：", ["一→二階(%)", "二→最終(%)", "一→最終(%)"],
                              horizontal=True, key="cross_det_metric")
        dept_data_by_year = {yr: c["ds"] for yr, c in year_cache.items() if c["ds"] is not None}
        if len(dept_data_by_year) >= 2:
            col_a, col_b = st.columns(2)
            yr_opts = list(dept_data_by_year.keys())
            with col_a: yr_from = st.selectbox("起始年度：", yr_opts, index=0, key="det_yr_from")
            with col_b:
                yr_to_opts = [y for y in yr_opts if y != yr_from]
                yr_to = st.selectbox("比較年度：", yr_to_opts,
                    index=len(yr_to_opts)-1 if yr_to_opts else 0, key="det_yr_to")

            if yr_from in dept_data_by_year and yr_to in dept_data_by_year:
                df_from = dept_data_by_year[yr_from]
                df_to = dept_data_by_year[yr_to]
                min_p2 = "二→最終" in det_metric
                if min_p2:
                    df_from = df_from[df_from["二階人數"] > 0]
                    df_to = df_to[df_to["二階人數"] > 0]
                from_map = dict(zip(df_from["科系"].apply(norm_dept), df_from[det_metric]))
                to_map = dict(zip(df_to["科系"].apply(norm_dept), df_to[det_metric]))
                dept_name_map = {}
                for d in df_from["科系"]: dept_name_map[norm_dept(d)] = d
                for d in df_to["科系"]: dept_name_map[norm_dept(d)] = d
                compare_rows = []
                for dk in set(from_map) | set(to_map):
                    vf, vt = from_map.get(dk), to_map.get(dk)
                    change = round(vt - vf, 1) if vf is not None and vt is not None else None
                    sev = ""
                    if change is not None and change < 0:
                        sev = "🔴 嚴重" if abs(change) > 10 else ("🟠 注意" if abs(change) > 5 else "🟡 輕微")
                    elif change is not None and change > 0: sev = "✅ 改善"
                    compare_rows.append({
                        "科系": dept_name_map.get(dk, dk),
                        f"{yr_from}": vf if vf is not None else "—",
                        f"{yr_to}": vt if vt is not None else "—",
                        "變化": change if change is not None else "—",
                        "狀態": sev if sev else ("🆕" if vf is None else ("❌ 消失" if vt is None else "—"))
                    })
                compare_df = pd.DataFrame(compare_rows)
                def sk(x): return x if isinstance(x, (int, float)) else 999
                compare_df["_s"] = compare_df["變化"].apply(sk)
                compare_df = compare_df.sort_values("_s").drop(columns=["_s"])
                deteriorated = [r for r in compare_rows if isinstance(r["變化"], (int, float)) and r["變化"] < 0]
                if deteriorated:
                    severe = sum(1 for r in deteriorated if "嚴重" in r.get("狀態", ""))
                    st.markdown(
                        f'<div class="warning-box">⚠️ {yr_from}→{yr_to}：{len(deteriorated)} 科系下降（🔴{severe}）</div>',
                        unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="success-box">✅ 所有科系持平或上升！</div>', unsafe_allow_html=True)
                st.dataframe(compare_df, use_container_width=True, hide_index=True)

                viz = [r for r in compare_rows if isinstance(r["變化"], (int, float))]
                if viz:
                    vizdf = pd.DataFrame(viz).sort_values("變化")
                    colors = ["#f44336" if v < -10 else "#FF9800" if v < -5 else "#FFC107" if v < 0 else "#4CAF50" for v in vizdf["變化"]]
                    fig = go.Figure(go.Bar(x=vizdf["變化"], y=vizdf["科系"], orientation="h", text=vizdf["變化"], marker_color=colors))
                    fig.update_traces(texttemplate="%{text:+.1f}", textposition="outside")
                    fig.add_vline(x=0, line_color="black", line_width=2)
                    fig.update_layout(title=f"{yr_from}→{yr_to}：{det_metric} 變化",
                                      height=max(400, len(vizdf)*30))
                    st.plotly_chart(fig, use_container_width=True)

            # 連續下降
            pair_results, consec_results = detect_deterioration_full(
                dept_data_by_year, det_metric, "科系", min_p2="二→最終" in det_metric)
            if consec_results:
                st.markdown("---"); st.markdown("##### 🔻 連續下降偵測")
                for r in consec_results:
                    accel = "⚡ 加速惡化" if r["加速中"] else "📉 持續下降"
                    st.markdown(
                        f'<div class="danger-box">🔻 <b>{r["科系"]}</b>：連續 {r["連續下降年數"]} 年下降 '
                        f'（{r["起始年度"]}→{r["最新年度"]}，累積 {r["累積下降"]}）{accel}</div>',
                        unsafe_allow_html=True)
                    yr_vals = r["各年度值"]
                    trend_df = pd.DataFrame([{"年度": y, det_metric: v} for y, v in yr_vals.items()])
                    fig = go.Figure(go.Scatter(x=trend_df["年度"], y=trend_df[det_metric],
                        mode="lines+markers+text", text=trend_df[det_metric], textposition="top center",
                        line=dict(width=3, color="#f44336"), marker=dict(size=12)))
                    fig.update_layout(title=f"「{r['科系']}」{det_metric} 趨勢", height=350)
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
    '🎓 中華醫事科技大學 招生數據分析系統 v6.8<br>'
    '人頭數 vs 志願次數引擎 ｜ 全年度惡化偵測 ｜ 三段轉換率 ｜ 縮寫展開引擎<br>'
    '分析版本 #' + str(ver) + '</div>', unsafe_allow_html=True)
