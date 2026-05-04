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
# 深色模式相容 CSS
# ============================================================
st.markdown("""
<style>
    .sidebar-card {
        background: rgba(255, 152, 56, 0.15);
        border: 1px solid rgba(255, 152, 56, 0.4);
        padding: 12px;
        border-radius: 8px;
        margin-bottom: 12px;
        color: inherit;
    }
    .sidebar-card small { color: inherit; opacity: 0.9; }
    .sidebar-card b { color: #FFB74D; }
    .guide-container { text-align: center; padding: 40px 20px; }
    .guide-cards {
        display: flex; justify-content: center; gap: 24px;
        margin-top: 30px; flex-wrap: wrap;
    }
    .guide-card {
        background: rgba(255, 152, 56, 0.12);
        border: 1px solid rgba(255, 152, 56, 0.35);
        padding: 24px 18px; border-radius: 14px; width: 220px;
        transition: transform 0.2s, border-color 0.2s; color: inherit;
    }
    .guide-card:hover { transform: translateY(-4px); border-color: rgba(255, 152, 56, 0.7); }
    .guide-card h3 { margin: 0 0 8px 0; font-size: 1.15em; color: #FFB74D; }
    .guide-card p { font-size: 0.85em; margin: 0; opacity: 0.8; color: inherit; }
    .guide-arrow { font-size: 2em; margin-top: 10px; opacity: 0.6; }
    .channel-badge {
        display: inline-block; padding: 3px 10px; border-radius: 12px;
        font-size: 0.82em; font-weight: bold; margin: 2px;
    }
    .channel-badge.orange { background: rgba(255,152,56,0.25); color: #FFB74D; border: 1px solid rgba(255,152,56,0.5); }
    .channel-badge.blue { background: rgba(66,165,245,0.25); color: #42A5F5; border: 1px solid rgba(66,165,245,0.5); }
    .channel-badge.green { background: rgba(102,187,106,0.25); color: #66BB6A; border: 1px solid rgba(102,187,106,0.5); }
    .channel-badge.purple { background: rgba(171,130,255,0.25); color: #AB82FF; border: 1px solid rgba(171,130,255,0.5); }
    .js-plotly-plot .plotly .main-svg { background: transparent !important; }
</style>
""", unsafe_allow_html=True)

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
    df = deduplicate_columns(df)
    col_mapping = {}
    for col in df.columns:
        c = str(col).strip().replace(" ", "")
        if "學年" in c and "學年度" not in col_mapping.values():
            col_mapping[col] = "學年度"
        elif any(k in c for k in ["管道", "入學管道", "招生管道", "錄取管道"]) and "招生管道" not in col_mapping.values():
            col_mapping[col] = "招生管道"
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
        elif any(k in c for k in ["錄取", "報到", "入學", "狀態"]):
            if "入學狀態" not in col_mapping.values():
                col_mapping[col] = "入學狀態"
    df = df.rename(columns=col_mapping)
    df = deduplicate_columns(df)
    return df

def safe_get_series(df, col_name):
    col_data = df[col_name]
    if isinstance(col_data, pd.DataFrame):
        col_data = col_data.iloc[:, 0]
    return col_data

def standardize_data(df):
    df = df.copy()
    if "學年度" in df.columns:
        col_data = safe_get_series(df, "學年度")
        df["學年度"] = col_data.apply(lambda x: str(int(float(x))) if pd.notna(x) else None)
    for col_name in ["報考科系", "畢業學校", "姓名", "身分證字號", "招生管道"]:
        if col_name in df.columns:
            col_data = safe_get_series(df, col_name)
            df[col_name] = col_data.astype(str).str.strip()
            df[col_name] = df[col_name].replace({"nan": None, "None": None, "": None})
    return df

def parse_coordinates(df):
    df = df.copy()
    if "緯度" in df.columns and "經度" in df.columns:
        df["緯度"] = pd.to_numeric(safe_get_series(df, "緯度"), errors="coerce")
        df["經度"] = pd.to_numeric(safe_get_series(df, "經度"), errors="coerce")
        return df
    if "經緯度" not in df.columns:
        return df
    coord_col = safe_get_series(df, "經緯度")
    lats, lons = [], []
    for val in coord_col:
        try:
            val_str = str(val).strip()
            if val_str in ("", "nan", "None", "NaN"):
                lats.append(None); lons.append(None); continue
            for ch in ["(", ")", "（", "）", "「", "」"]:
                val_str = val_str.replace(ch, "")
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

def detect_channel_from_filename(filename):
    fn = filename.lower()
    patterns = {
        "聯合免試": ["聯合免試", "聯免", "免試"],
        "甄選入學": ["甄選入學", "甄選"],
        "技優甄審": ["技優", "甄審"],
        "運動績優": ["運動", "績優", "體育"],
        "身障甄試": ["身障", "身心障礙"],
        "單獨招生": ["單獨招生", "單招"],
        "進修部": ["進修部", "進修"],
    }
    for channel, keywords in patterns.items():
        for kw in keywords:
            if kw in fn:
                return channel
    return None

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
            p2_extra = [c for c in p2.columns if c not in p1.columns]
            p2_merge = p2[available_keys_p2 + p2_extra].drop_duplicates(subset=available_keys_p2, keep="last").copy()
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
            pf_extra = [c for c in pf.columns if c not in p1.columns]
            pf_merge = pf[available_keys_pf + pf_extra].drop_duplicates(subset=available_keys_pf, keep="last").copy()
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


# ────────────────────────────────────────
# 核心轉換率計算（分母統一為一階報名人數）
# ────────────────────────────────────────
def compute_conversion(df, group_col, group_label="群組"):
    """以一階報名為分母，計算各群組三階段轉換率"""
    if group_col not in df.columns:
        return pd.DataFrame()
    groups = df[group_col].dropna().unique()
    results = []
    for g in sorted(groups):
        gd = df[df[group_col] == g]
        n1 = len(gd)
        n2 = len(gd[gd["二階甄試"] == "✅"]) if "二階甄試" in gd.columns else 0
        nf = len(gd[gd["最終入學"] == "✅"]) if "最終入學" in gd.columns else 0
        # 所有轉換率分母 = 一階報名人數
        rate_1to2 = n2 / n1 * 100 if n1 > 0 else 0
        rate_1tof = nf / n1 * 100 if n1 > 0 else 0
        # 階段內轉換（輔助參考）
        rate_2tof_internal = nf / n2 * 100 if n2 > 0 else 0
        results.append({
            group_label: g,
            "①一階報名": n1,
            "②二階甄試": n2,
            "③最終入學": nf,
            "進入二階率(以一階為母體)": rate_1to2,
            "最終入學率(以一階為母體)": rate_1tof,
            "二階→入學(階段內)": rate_2tof_internal,
            "一→二流失": n1 - n2,
            "二→入學流失": n2 - nf,
            "總流失": n1 - nf,
        })
    result_df = pd.DataFrame(results)
    if not result_df.empty:
        result_df = result_df.sort_values("①一階報名", ascending=False)
    return result_df


def format_pct(val):
    return "{:.1f}%".format(val)


def dark_friendly_plotly(fig):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color="rgba(255,255,255,0.85)",
        xaxis=dict(gridcolor="rgba(255,255,255,0.1)"),
        yaxis=dict(gridcolor="rgba(255,255,255,0.1)"),
    )
    return fig


# 管道顏色
CHANNEL_COLORS = {
    "聯合免試": "#FF8A65",
    "甄選入學": "#42A5F5",
    "技優甄審": "#66BB6A",
    "運動績優": "#AB82FF",
    "身障甄試": "#FFD54F",
    "單獨招生": "#4DB6AC",
    "進修部": "#F06292",
}
DEFAULT_CHANNEL_COLOR = "#90A4AE"

def get_channel_color(ch):
    return CHANNEL_COLORS.get(ch, DEFAULT_CHANNEL_COLOR)


# ============================================================
# 側邊欄
# ============================================================
with st.sidebar:
    st.markdown("## 🎓 HWU 招生分析系統")
    st.markdown(
        '<div class="sidebar-card">'
        '<small>📌 <b>多管道三階段匯入</b><br>'
        '① 一階報名：學年度、科系、姓名、學校、座標<br>'
        '② 二階甄試：姓名 + 學年度 + 成績<br>'
        '③ 最終入學：姓名（+學年度）<br><br>'
        '💡 <b>招生管道</b>辨識方式：<br>'
        '• 資料含「招生管道」欄位 → 自動辨識<br>'
        '• 或由<b>檔名</b>自動判斷（如：聯合免試_113.xlsx）<br>'
        '• 或上傳後手動指定<br><br>'
        '📐 所有轉換率<b>分母 = 一階報名人數</b></small>'
        '</div>',
        unsafe_allow_html=True
    )
    st.markdown("---")

    # ── 管道設定 ──
    st.markdown("### 🔀 招生管道設定")
    default_channel = st.selectbox(
        "本次上傳的預設管道",
        ["自動偵測（由檔名/欄位判斷）", "聯合免試", "甄選入學", "技優甄審",
         "運動績優", "身障甄試", "單獨招生", "進修部", "其他"],
        key="default_channel"
    )
    if default_channel == "其他":
        custom_channel = st.text_input("自訂管道名稱", key="custom_channel")
    else:
        custom_channel = ""
    st.markdown("---")

    st.markdown("### 📋 ① 一階報名資料")
    st.caption("需含：學年度、報考科系、姓名、畢業學校、經緯度")
    phase1_files = st.file_uploader("上傳一階 Excel", type=["xlsx", "xls"], accept_multiple_files=True, key="p1_upload")
    st.markdown("### 📝 ② 二階甄試資料")
    st.caption("需含：姓名（+學年度），可含成績")
    phase2_files = st.file_uploader("上傳二階 Excel", type=["xlsx", "xls"], accept_multiple_files=True, key="p2_upload")
    st.markdown("### 🎓 ③ 最終入學資料")
    st.caption("最少僅需：姓名（+學年度）")
    final_files = st.file_uploader("上傳入學 Excel", type=["xlsx", "xls"], accept_multiple_files=True, key="pf_upload")

    def resolve_channel(filename, df_columns):
        if default_channel == "其他" and custom_channel:
            return custom_channel
        if default_channel != "自動偵測（由檔名/欄位判斷）":
            return default_channel
        # 嘗試從檔名判斷
        ch = detect_channel_from_filename(filename)
        if ch:
            return ch
        return None  # 交由資料中的欄位處理

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
                original_cols = list(new_df.columns)
                new_df = clean_column_names(new_df, phase=phase_key)
                new_df = standardize_data(new_df)
                if phase_key == "phase1":
                    new_df = parse_coordinates(new_df)
                if "姓名" not in new_df.columns:
                    msg = "❌ {}（{}）：缺少「姓名」欄位。原始欄位：{}".format(
                        uf.name, phase_label, ", ".join(str(c) for c in original_cols[:10]))
                    st.session_state.upload_log.append(msg)
                    continue
                # 管道處理
                ch = resolve_channel(uf.name, new_df.columns)
                if "招生管道" not in new_df.columns and ch:
                    new_df["招生管道"] = ch
                elif "招生管道" not in new_df.columns:
                    new_df["招生管道"] = "未分類"
                # 若欄位中有管道但也有手動指定，手動優先
                if ch and "招生管道" in new_df.columns:
                    mask_empty = new_df["招生管道"].isna() | (new_df["招生管道"] == "") | (new_df["招生管道"] == "未分類")
                    if mask_empty.any():
                        new_df.loc[mask_empty, "招生管道"] = ch

                existing = st.session_state["{}_data".format(phase_key)]
                if existing.empty:
                    st.session_state["{}_data".format(phase_key)] = new_df
                else:
                    combined = pd.concat([existing, new_df], ignore_index=True, sort=False)
                    combined = combined.drop_duplicates(keep="last")
                    st.session_state["{}_data".format(phase_key)] = combined
                row_count = len(new_df)
                year_info = ""
                if "學年度" in new_df.columns:
                    yrs = safe_sort_years(new_df["學年度"])
                    year_info = "（{}）".format(", ".join(yrs))
                ch_info = ""
                if "招生管道" in new_df.columns:
                    chs = new_df["招生管道"].dropna().unique().tolist()
                    ch_info = " 📌{}".format("/".join(chs[:3]))
                st.session_state.uploaded_hashes[fhash] = uf.name
                msg = "✅ {}【{}】：{} 筆 {}{}".format(uf.name, phase_label, row_count, year_info, ch_info)
                st.session_state.upload_log.append(msg)
            except Exception as e:
                msg = "❌ {}（{}）：{}".format(uf.name, phase_label, str(e))
                st.session_state.upload_log.append(msg)

    process_upload(phase1_files, "phase1", "一階")
    process_upload(phase2_files, "phase2", "二階")
    process_upload(final_files, "final", "入學")
    merge_all_phases()

    if st.session_state.upload_log:
        st.markdown("---")
        st.markdown("### 📋 匯入紀錄")
        for log in st.session_state.upload_log[-12:]:
            st.caption(log)

    st.markdown("---")
    st.markdown("### 📊 資料狀態")
    p1_n = len(st.session_state.phase1_data)
    p2_n = len(st.session_state.phase2_data)
    pf_n = len(st.session_state.final_data)
    mg_n = len(st.session_state.merged_data)
    st.markdown(
        "| 階段 | 筆數 |\n|------|------|\n"
        "| ① 一階報名 | **{:,}** |\n| ② 二階甄試 | **{:,}** |\n"
        "| ③ 最終入學 | **{:,}** |\n| 🔗 合併後 | **{:,}** |".format(p1_n, p2_n, pf_n, mg_n)
    )
    # 管道分布
    if mg_n > 0 and "招生管道" in st.session_state.merged_data.columns:
        ch_counts = st.session_state.merged_data["招生管道"].value_counts()
        st.markdown("### 🔀 管道分布")
        for ch_name, ch_cnt in ch_counts.items():
            st.caption("• {}：{:,}".format(ch_name, ch_cnt))

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
        '<div class="guide-container">'
        '<h2>👈 請先從左側匯入各階段資料</h2>'
        '<div class="guide-cards">'
        '  <div class="guide-card"><h3>📋 一階報名</h3><p>學年度、科系、姓名<br>學校、座標</p></div>'
        '  <div class="guide-arrow">→</div>'
        '  <div class="guide-card"><h3>📝 二階甄試</h3><p>姓名、學年度<br>+ 成績（選填）</p></div>'
        '  <div class="guide-arrow">→</div>'
        '  <div class="guide-card"><h3>🎓 最終入學</h3><p>僅需：姓名<br>（+學年度）</p></div>'
        '</div>'
        '<p style="margin-top:30px; opacity:0.5;">'
        '🔀 支援多管道分析（聯合免試/甄選入學/...）<br>'
        '📐 所有轉換率分母 = <b>一階報名人數</b>，真實反映管道效益</p>'
        '</div>',
        unsafe_allow_html=True
    )
    st.stop()

# ── 全域篩選器 ──
has_channel = "招生管道" in data.columns and data["招生管道"].nunique() > 0
has_year = "學年度" in data.columns

# ============================================================
# 分頁
# ============================================================
tab_labels = ["🔄 招生漏斗", "🔀 管道比較分析", "🗺️ 地圖視覺化",
              "📊 科系轉換分析", "🏫 來源學校分析", "📈 跨年度比較", "🔍 資料檢視與匯出"]
tab0, tab_ch, tab1, tab2, tab3, tab4, tab5 = st.tabs(tab_labels)

# ============================================================
# TAB 0: 招生漏斗
# ============================================================
with tab0:
    st.header("🔄 招生漏斗分析")
    st.caption("📐 所有轉換率分母 = 一階報名人數")

    col_f_yr, col_f_ch = st.columns(2)
    with col_f_yr:
        if has_year:
            yr_opts_t0 = ["全部"] + safe_sort_years(data["學年度"])
            sel_yr_t0 = st.selectbox("學年度", yr_opts_t0, key="t0_yr")
        else:
            sel_yr_t0 = "全部"
    with col_f_ch:
        if has_channel:
            ch_opts_t0 = ["全部（含所有管道）"] + sorted(data["招生管道"].dropna().unique().tolist())
            sel_ch_t0 = st.selectbox("招生管道", ch_opts_t0, key="t0_ch")
        else:
            sel_ch_t0 = "全部（含所有管道）"

    fd = data.copy()
    if sel_yr_t0 != "全部" and has_year:
        fd = fd[fd["學年度"] == sel_yr_t0]
    if sel_ch_t0 != "全部（含所有管道）" and has_channel:
        fd = fd[fd["招生管道"] == sel_ch_t0]

    n1 = len(fd[fd["一階報名"] == "✅"]) if "一階報名" in fd.columns else len(fd)
    n2 = len(fd[fd["二階甄試"] == "✅"]) if "二階甄試" in fd.columns else 0
    nf = len(fd[fd["最終入學"] == "✅"]) if "最終入學" in fd.columns else 0

    rate_12 = n2 / n1 * 100 if n1 > 0 else 0
    rate_1f = nf / n1 * 100 if n1 > 0 else 0

    kpi = st.columns(5)
    with kpi[0]:
        st.metric("📋 一階報名（母體）", "{:,} 人".format(n1))
    with kpi[1]:
        st.metric("📝 進入二階", "{:,} 人".format(n2), "佔一階 {:.1f}%".format(rate_12))
    with kpi[2]:
        st.metric("🎓 最終入學", "{:,} 人".format(nf), "佔一階 {:.1f}%".format(rate_1f))
    with kpi[3]:
        st.metric("🎯 管道轉換率", "{:.1f}%".format(rate_1f))
    with kpi[4]:
        st.metric("📉 總流失", "{:,} 人".format(n1 - nf))

    col_fn1, col_fn2 = st.columns([2, 1])
    with col_fn1:
        fig_funnel = go.Figure(go.Funnel(
            y=["一階報名（母體）", "二階甄試", "最終入學"],
            x=[n1, n2, nf],
            textinfo="value+percent initial",
            marker=dict(color=["#FDD7B4", "#E8792F", "#B84500"]),
            connector=dict(line=dict(color="#E8792F", width=2)),
        ))
        fig_funnel.update_layout(title="招生漏斗（% = 佔一階報名比例）", height=400)
        dark_friendly_plotly(fig_funnel)
        st.plotly_chart(fig_funnel, use_container_width=True)

    with col_fn2:
        lost_12 = n1 - n2
        lost_2f = n2 - nf
        loss_pct_12 = lost_12 / n1 * 100 if n1 > 0 else 0
        loss_pct_2f = lost_2f / n1 * 100 if n1 > 0 else 0  # 以一階為母體
        st.markdown("**📉 流失明細（分母=一階報名）**")
        st.markdown(
            "| 階段 | 流失人數 | 佔一階% |\n"
            "|------|---------|--------|\n"
            "| 一階→二階 | {:,} | {:.1f}% |\n"
            "| 二階→入學 | {:,} | {:.1f}% |\n"
            "| **總流失** | **{:,}** | **{:.1f}%** |".format(
                lost_12, loss_pct_12, lost_2f, loss_pct_2f,
                n1 - nf, (n1 - nf) / n1 * 100 if n1 > 0 else 0
            )
        )
        if n1 > 0:
            st.markdown("---")
            st.markdown("**🔍 流失瓶頸判斷**")
            if loss_pct_12 > loss_pct_2f and loss_pct_12 > 30:
                st.error("⚠️ 主要瓶頸在「一階→二階」（佔一階 {:.1f}% 流失），建議加強報名後追蹤".format(loss_pct_12))
            elif loss_pct_2f > loss_pct_12 and loss_pct_2f > 20:
                st.error("⚠️ 主要瓶頸在「二階→入學」（佔一階 {:.1f}% 流失），建議優化錄取流程".format(loss_pct_2f))
            else:
                st.success("✅ 各階段流失比例相近（一→二 {:.1f}%, 二→入 {:.1f}%）".format(loss_pct_12, loss_pct_2f))

    # 各科系分母=一階
    if "報考科系" in fd.columns:
        st.subheader("📊 各科系轉換率（分母=一階報名）")
        dept_conv = compute_conversion(fd, "報考科系", "科系")
        if not dept_conv.empty:
            disp_dc = dept_conv.copy()
            disp_dc["進入二階率(以一階為母體)"] = disp_dc["進入二階率(以一階為母體)"].apply(format_pct)
            disp_dc["最終入學率(以一階為母體)"] = disp_dc["最終入學率(以一階為母體)"].apply(format_pct)
            disp_dc["二階→入學(階段內)"] = disp_dc["二階→入學(階段內)"].apply(format_pct)
            st.dataframe(disp_dc, use_container_width=True, hide_index=True)

# ============================================================
# TAB CH: 管道比較分析（新增）
# ============================================================
with tab_ch:
    st.header("🔀 招生管道比較分析")
    st.caption("📐 各管道轉換率分母 = 該管道一階報名人數，真實反映各管道效益")

    if not has_channel or data["招生管道"].nunique() < 1:
        st.info("💡 尚未偵測到「招生管道」資訊。請確認：\n"
                "1. 資料中包含「招生管道」欄位，或\n"
                "2. 在側邊欄指定管道，或\n"
                "3. 以管道名稱命名檔案（如：聯合免試_113.xlsx）")
    else:
        if has_year:
            yr_opts_ch = ["全部"] + safe_sort_years(data["學年度"])
            sel_yr_ch = st.selectbox("學年度", yr_opts_ch, key="ch_yr")
        else:
            sel_yr_ch = "全部"

        ch_data = data.copy()
        if sel_yr_ch != "全部" and has_year:
            ch_data = ch_data[ch_data["學年度"] == sel_yr_ch]

        channel_conv = compute_conversion(ch_data, "招生管道", "招生管道")

        if channel_conv.empty:
            st.info("無管道資料可分析")
        else:
            # KPI 各管道一覽
            st.subheader("📊 各管道轉換率總覽")
            disp_ch = channel_conv.copy()
            disp_ch["進入二階率(以一階為母體)"] = disp_ch["進入二階率(以一階為母體)"].apply(format_pct)
            disp_ch["最終入學率(以一階為母體)"] = disp_ch["最終入學率(以一階為母體)"].apply(format_pct)
            disp_ch["二階→入學(階段內)"] = disp_ch["二階→入學(階段內)"].apply(format_pct)
            st.dataframe(disp_ch, use_container_width=True, hide_index=True)

            # 管道 KPI 卡片
            kpi_ch_cols = st.columns(min(len(channel_conv), 5))
            for idx, (_, row) in enumerate(channel_conv.iterrows()):
                if idx >= len(kpi_ch_cols):
                    break
                with kpi_ch_cols[idx]:
                    ch_name = row["招生管道"]
                    st.metric(
                        "📌 {}".format(ch_name),
                        "{:,} → {:,} 人".format(int(row["①一階報名"]), int(row["③最終入學"])),
                        "入學率 {:.1f}%".format(row["最終入學率(以一階為母體)"])
                    )

            col_ch1, col_ch2 = st.columns(2)
            with col_ch1:
                # 各管道人數比較
                fig_ch_bar = go.Figure()
                fig_ch_bar.add_trace(go.Bar(
                    name="一階報名", x=channel_conv["招生管道"], y=channel_conv["①一階報名"],
                    marker_color="#FDD7B4", text=channel_conv["①一階報名"], textposition="outside"
                ))
                fig_ch_bar.add_trace(go.Bar(
                    name="二階甄試", x=channel_conv["招生管道"], y=channel_conv["②二階甄試"],
                    marker_color="#E8792F", text=channel_conv["②二階甄試"], textposition="outside"
                ))
                fig_ch_bar.add_trace(go.Bar(
                    name="最終入學", x=channel_conv["招生管道"], y=channel_conv["③最終入學"],
                    marker_color="#B84500", text=channel_conv["③最終入學"], textposition="outside"
                ))
                fig_ch_bar.update_layout(barmode="group", title="各管道三階段人數比較", height=450)
                dark_friendly_plotly(fig_ch_bar)
                st.plotly_chart(fig_ch_bar, use_container_width=True)

            with col_ch2:
                # 各管道轉換率
                fig_ch_rate = go.Figure()
                fig_ch_rate.add_trace(go.Bar(
                    name="進入二階率", x=channel_conv["招生管道"],
                    y=channel_conv["進入二階率(以一階為母體)"],
                    marker_color="#E8792F",
                    text=channel_conv["進入二階率(以一階為母體)"].apply(lambda v: "{:.1f}%".format(v)),
                    textposition="outside"
                ))
                fig_ch_rate.add_trace(go.Bar(
                    name="最終入學率", x=channel_conv["招生管道"],
                    y=channel_conv["最終入學率(以一階為母體)"],
                    marker_color="#B84500",
                    text=channel_conv["最終入學率(以一階為母體)"].apply(lambda v: "{:.1f}%".format(v)),
                    textposition="outside"
                ))
                fig_ch_rate.update_layout(
                    barmode="group", title="各管道轉換率比較（分母=一階報名）",
                    yaxis_title="轉換率 (%)", height=450
                )
                dark_friendly_plotly(fig_ch_rate)
                st.plotly_chart(fig_ch_rate, use_container_width=True)

            # 各管道漏斗圖
            st.subheader("🔄 各管道漏斗對比")
            n_channels = len(channel_conv)
            funnel_cols = st.columns(min(n_channels, 4))
            for idx, (_, row) in enumerate(channel_conv.iterrows()):
                col_idx = idx % len(funnel_cols)
                with funnel_cols[col_idx]:
                    ch_name = row["招生管道"]
                    ch_color = get_channel_color(ch_name)
                    fig_mini = go.Figure(go.Funnel(
                        y=["一階報名", "二階甄試", "最終入學"],
                        x=[int(row["①一階報名"]), int(row["②二階甄試"]), int(row["③最終入學"])],
                        textinfo="value+percent initial",
                        marker=dict(color=[ch_color, ch_color, ch_color]),
                        opacity=0.85,
                    ))
                    fig_mini.update_layout(
                        title=ch_name, height=300,
                        margin=dict(l=10, r=10, t=40, b=10),
                        showlegend=False
                    )
                    dark_friendly_plotly(fig_mini)
                    st.plotly_chart(fig_mini, use_container_width=True)

            # 管道 × 科系 交叉分析
            if "報考科系" in ch_data.columns and ch_data["招生管道"].nunique() > 1:
                st.subheader("🔀 管道 × 科系 入學人數交叉分析")
                enrolled_ch = ch_data[ch_data.get("最終入學", "") == "✅"]
                if len(enrolled_ch) > 0:
                    cross = pd.crosstab(enrolled_ch["招生管道"], enrolled_ch["報考科系"])
                    fig_heat_ch = px.imshow(
                        cross, title="管道 × 科系 入學人數",
                        color_continuous_scale="Oranges", aspect="auto"
                    )
                    fig_heat_ch.update_layout(height=max(350, n_channels * 50))
                    dark_friendly_plotly(fig_heat_ch)
                    st.plotly_chart(fig_heat_ch, use_container_width=True)
                else:
                    st.info("無入學資料可顯示交叉分析")

            # 管道效益評估
            st.subheader("⭐ 管道效益評估")
            eval_rows = []
            for _, row in channel_conv.iterrows():
                ch_name = row["招生管道"]
                n1_ch = int(row["①一階報名"])
                nf_ch = int(row["③最終入學"])
                rate_ch = row["最終入學率(以一階為母體)"]
                if rate_ch >= 50:
                    grade = "⭐⭐⭐ 高效管道"
                    suggestion = "轉換率佳，建議擴大此管道招生規模"
                elif rate_ch >= 25:
                    grade = "⭐⭐ 中效管道"
                    suggestion = "轉換率中等，可針對瓶頸階段優化"
                elif rate_ch >= 10:
                    grade = "⭐ 低效管道"
                    suggestion = "轉換率偏低，需深入分析流失原因"
                else:
                    grade = "⚠️ 待改善"
                    suggestion = "轉換率極低，建議評估投入產出比"
                # 判斷主要流失階段
                loss_12_pct = row["一→二流失"] / n1_ch * 100 if n1_ch > 0 else 0
                loss_2f_pct = row["二→入學流失"] / n1_ch * 100 if n1_ch > 0 else 0
                if loss_12_pct > loss_2f_pct:
                    bottleneck = "一→二（{:.0f}人/{:.1f}%）".format(row["一→二流失"], loss_12_pct)
                else:
                    bottleneck = "二→入學（{:.0f}人/{:.1f}%）".format(row["二→入學流失"], loss_2f_pct)
                eval_rows.append({
                    "管道": ch_name,
                    "一階報名": n1_ch,
                    "最終入學": nf_ch,
                    "管道轉換率": format_pct(rate_ch),
                    "主要流失階段": bottleneck,
                    "效益等級": grade,
                    "建議": suggestion,
                })
            eval_df = pd.DataFrame(eval_rows)
            st.dataframe(eval_df, use_container_width=True, hide_index=True)

            # 歷年管道趨勢
            if has_year and data["學年度"].nunique() > 1 and sel_yr_ch == "全部":
                st.subheader("📈 各管道歷年趨勢")
                yearly_ch = []
                for yr in safe_sort_years(data["學年度"]):
                    yd = data[data["學年度"] == yr]
                    for ch in data["招生管道"].dropna().unique():
                        cd = yd[yd["招生管道"] == ch]
                        y1 = len(cd)
                        yf = len(cd[cd["最終入學"] == "✅"]) if "最終入學" in cd.columns else 0
                        yearly_ch.append({
                            "學年度": yr, "管道": ch,
                            "報名": y1, "入學": yf,
                            "轉換率": yf / y1 * 100 if y1 > 0 else 0,
                        })
                ych_df = pd.DataFrame(yearly_ch)
                col_ych1, col_ych2 = st.columns(2)
                with col_ych1:
                    fig_ych_n = px.line(
                        ych_df, x="學年度", y="報名", color="管道",
                        markers=True, title="各管道歷年報名人數"
                    )
                    fig_ych_n.update_layout(height=400)
                    dark_friendly_plotly(fig_ych_n)
                    st.plotly_chart(fig_ych_n, use_container_width=True)
                with col_ych2:
                    fig_ych_r = px.line(
                        ych_df, x="學年度", y="轉換率", color="管道",
                        markers=True, title="各管道歷年轉換率（%）"
                    )
                    fig_ych_r.update_layout(height=400, yaxis_title="%")
                    dark_friendly_plotly(fig_ych_r)
                    st.plotly_chart(fig_ych_r, use_container_width=True)

# ============================================================
# TAB 1: 地圖
# ============================================================
with tab1:
    st.header("🗺️ 報考生分布地圖")
    col_mf1, col_mf2, col_mf3, col_mf4 = st.columns(4)
    with col_mf1:
        if has_year:
            yr_opts_m = ["全部"] + safe_sort_years(data["學年度"])
            sel_yr_m = st.selectbox("學年度", yr_opts_m, key="map_yr")
        else:
            sel_yr_m = "全部"
    with col_mf2:
        if "報考科系" in data.columns:
            dept_opts_m = ["全部"] + sorted(data["報考科系"].dropna().unique().tolist())
            sel_dept_m = st.selectbox("科系", dept_opts_m, key="map_dept")
        else:
            sel_dept_m = "全部"
    with col_mf3:
        if "目前狀態" in data.columns:
            stg_opts_m = ["全部"] + sorted(data["目前狀態"].dropna().unique().tolist())
            sel_stg_m = st.selectbox("階段", stg_opts_m, key="map_stg")
        else:
            sel_stg_m = "全部"
    with col_mf4:
        if has_channel:
            ch_opts_m = ["全部"] + sorted(data["招生管道"].dropna().unique().tolist())
            sel_ch_m = st.selectbox("管道", ch_opts_m, key="map_ch")
        else:
            sel_ch_m = "全部"

    md = data.copy()
    if sel_yr_m != "全部" and has_year:
        md = md[md["學年度"] == sel_yr_m]
    if sel_dept_m != "全部" and "報考科系" in md.columns:
        md = md[md["報考科系"] == sel_dept_m]
    if sel_stg_m != "全部" and "目前狀態" in md.columns:
        md = md[md["目前狀態"] == sel_stg_m]
    if sel_ch_m != "全部" and has_channel:
        md = md[md["招生管道"] == sel_ch_m]

    has_coords = "緯度" in md.columns and "經度" in md.columns
    if has_coords:
        vc = md.dropna(subset=["緯度", "經度"]).copy()
        vc["緯度"] = pd.to_numeric(vc["緯度"], errors="coerce")
        vc["經度"] = pd.to_numeric(vc["經度"], errors="coerce")
        vc = vc.dropna(subset=["緯度", "經度"])
        if len(vc) > 0:
            col_mp, col_ms = st.columns([3, 1])
            with col_mp:
                m = folium.Map(location=[23.5, 120.5], zoom_start=8, tiles="CartoDB dark_matter")
                folium.Marker(
                    location=[22.9908, 120.2133], popup="中華醫事科技大學",
                    icon=folium.Icon(color="red", icon="star", prefix="fa"),
                ).add_to(m)
                color_map = {"已入學": "#FF6B35", "二階未入學": "#FFB347", "僅一階": "#87CEEB"}
                display_vc = vc.sample(min(1000, len(vc)), random_state=42) if len(vc) > 1000 else vc
                if len(vc) > 1000:
                    st.caption("⚡ 地圖抽樣 1,000 筆（共 {:,}）".format(len(vc)))
                for _, row in display_vc.iterrows():
                    status = row.get("目前狀態", "僅一階")
                    color = color_map.get(status, "#FFB347")
                    popup_parts = []
                    for f in ["畢業學校", "報考科系", "學年度", "目前狀態", "招生管道"]:
                        if f in row.index and pd.notna(row.get(f)):
                            popup_parts.append("{}：{}".format(f, row[f]))
                    popup_obj = folium.Popup("<br>".join(popup_parts), max_width=200) if popup_parts else None
                    folium.CircleMarker(
                        location=[float(row["緯度"]), float(row["經度"])],
                        radius=5, color=color, fill=True, fill_color=color, fill_opacity=0.7,
                        popup=popup_obj,
                    ).add_to(m)
                st_folium(m, width=None, height=550, use_container_width=True)
            with col_ms:
                st.metric("篩選結果", "{:,} 人".format(len(md)))
                st.metric("有座標", "{:,} 人".format(len(vc)))
                cov = len(vc) / len(md) * 100 if len(md) > 0 else 0
                st.metric("座標涵蓋率", "{:.1f}%".format(cov))
                if "目前狀態" in vc.columns:
                    st.markdown("**📊 階段分布**")
                    for sv, sc in vc["目前狀態"].value_counts().items():
                        icons = {"已入學": "🟠", "二階未入學": "🟡", "僅一階": "🔵"}
                        st.caption("{} {}：{}".format(icons.get(sv, "⚪"), sv, sc))
                if has_channel and "招生管道" in vc.columns:
                    st.markdown("**🔀 管道分布**")
                    for cv, cc in vc["招生管道"].value_counts().items():
                        st.caption("• {}：{}".format(cv, cc))
        else:
            st.warning("⚠️ 篩選後無有效座標")
    else:
        st.warning("⚠️ 無座標欄位")

# ============================================================
# TAB 2: 科系轉換分析
# ============================================================
with tab2:
    st.header("📊 科系三階段轉換分析")
    st.caption("📐 所有轉換率分母 = 一階報名人數")

    if "報考科系" not in data.columns:
        st.warning("資料中未包含「報考科系」欄位")
    else:
        col_t2f1, col_t2f2 = st.columns(2)
        with col_t2f1:
            if has_year:
                yr_opts_t2 = ["全部"] + safe_sort_years(data["學年度"])
                sel_yr_t2 = st.selectbox("學年度", yr_opts_t2, key="t2_yr")
            else:
                sel_yr_t2 = "全部"
        with col_t2f2:
            if has_channel:
                ch_opts_t2 = ["全部（含所有管道）"] + sorted(data["招生管道"].dropna().unique().tolist())
                sel_ch_t2 = st.selectbox("管道", ch_opts_t2, key="t2_ch")
            else:
                sel_ch_t2 = "全部（含所有管道）"

        t2d = data.copy()
        if sel_yr_t2 != "全部" and has_year:
            t2d = t2d[t2d["學年度"] == sel_yr_t2]
        if sel_ch_t2 != "全部（含所有管道）" and has_channel:
            t2d = t2d[t2d["招生管道"] == sel_ch_t2]

        dept_conv = compute_conversion(t2d, "報考科系", "科系")
        if not dept_conv.empty:
            st.subheader("📋 各科系轉換率（分母=一階報名）")
            disp_dc = dept_conv.copy()
            disp_dc["進入二階率(以一階為母體)"] = disp_dc["進入二階率(以一階為母體)"].apply(format_pct)
            disp_dc["最終入學率(以一階為母體)"] = disp_dc["最終入學率(以一階為母體)"].apply(format_pct)
            disp_dc["二階→入學(階段內)"] = disp_dc["二階→入學(階段內)"].apply(format_pct)
            st.dataframe(disp_dc, use_container_width=True, hide_index=True)

            kpi_d = st.columns(4)
            with kpi_d[0]:
                st.metric("科系數", "{}".format(len(dept_conv)))
            with kpi_d[1]:
                st.metric("平均進入二階率", "{:.1f}%".format(dept_conv["進入二階率(以一階為母體)"].mean()))
            with kpi_d[2]:
                st.metric("平均入學率", "{:.1f}%".format(dept_conv["最終入學率(以一階為母體)"].mean()))
            with kpi_d[3]:
                best_dept = dept_conv.iloc[0]["科系"] if not dept_conv.empty else "-"
                st.metric("最大報名量科系", best_dept)

            col_dc1, col_dc2 = st.columns(2)
            with col_dc1:
                fig_dc = go.Figure()
                fig_dc.add_trace(go.Bar(name="一階報名", x=dept_conv["科系"], y=dept_conv["①一階報名"], marker_color="#FDD7B4"))
                fig_dc.add_trace(go.Bar(name="二階甄試", x=dept_conv["科系"], y=dept_conv["②二階甄試"], marker_color="#E8792F"))
                fig_dc.add_trace(go.Bar(name="最終入學", x=dept_conv["科系"], y=dept_conv["③最終入學"], marker_color="#B84500"))
                fig_dc.update_layout(barmode="group", title="各科系三階段人數", height=450)
                dark_friendly_plotly(fig_dc)
                st.plotly_chart(fig_dc, use_container_width=True)
            with col_dc2:
                fig_dcr = go.Figure()
                fig_dcr.add_trace(go.Scatter(
                    x=dept_conv["科系"], y=dept_conv["進入二階率(以一階為母體)"],
                    mode="lines+markers+text", name="進入二階率",
                    line=dict(color="#E8792F", width=2),
                    text=dept_conv["進入二階率(以一階為母體)"].apply(lambda v: "{:.1f}%".format(v)),
                    textposition="top center"
                ))
                fig_dcr.add_trace(go.Scatter(
                    x=dept_conv["科系"], y=dept_conv["最終入學率(以一階為母體)"],
                    mode="lines+markers+text", name="最終入學率",
                    line=dict(color="#FFFFFF", width=3),
                    text=dept_conv["最終入學率(以一階為母體)"].apply(lambda v: "{:.1f}%".format(v)),
                    textposition="top center"
                ))
                fig_dcr.update_layout(title="各科系轉換率（分母=一階報名）", yaxis_title="%", height=450)
                dark_friendly_plotly(fig_dcr)
                st.plotly_chart(fig_dcr, use_container_width=True)

            st.subheader("🔍 科系瓶頸分析")
            for _, row in dept_conv.iterrows():
                n1_d = row["①一階報名"]
                r_enter = row["進入二階率(以一階為母體)"]
                r_final = row["最終入學率(以一階為母體)"]
                loss_12_d = row["一→二流失"]
                loss_2f_d = row["二→入學流失"]
                loss12_pct = loss_12_d / n1_d * 100 if n1_d > 0 else 0
                loss2f_pct = loss_2f_d / n1_d * 100 if n1_d > 0 else 0
                if r_enter < 30:
                    tag = "🔴 一→二嚴重流失"
                elif r_enter < 50:
                    tag = "🟡 一→二中度流失"
                elif loss2f_pct > 30:
                    tag = "🟠 二→入學流失偏高"
                else:
                    tag = "🟢 轉換良好"
                st.caption(
                    "**{}**：進入二階 {:.1f}%（流失 {}人/{:.1f}%）→ 入學 {:.1f}%（流失 {}人/{:.1f}%）→ {}".format(
                        row["科系"], r_enter, int(loss_12_d), loss12_pct,
                        r_final, int(loss_2f_d), loss2f_pct, tag
                    )
                )

            # 管道 × 科系
            if has_channel and sel_ch_t2 == "全部（含所有管道）" and t2d["招生管道"].nunique() > 1:
                st.subheader("🔀 管道 × 科系 入學率交叉分析")
                cross_data = []
                for ch in t2d["招生管道"].dropna().unique():
                    for dept in t2d["報考科系"].dropna().unique():
                        sub = t2d[(t2d["招生管道"] == ch) & (t2d["報考科系"] == dept)]
                        n1_x = len(sub)
                        nf_x = len(sub[sub["最終入學"] == "✅"]) if "最終入學" in sub.columns else 0
                        rate_x = nf_x / n1_x * 100 if n1_x > 0 else 0
                        cross_data.append({"管道": ch, "科系": dept, "入學率%": rate_x, "報名": n1_x})
                if cross_data:
                    cross_df = pd.DataFrame(cross_data)
                    pivot = cross_df.pivot_table(index="管道", columns="科系", values="入學率%", fill_value=0)
                    fig_hm = px.imshow(
                        pivot, title="管道 × 科系 入學率（%，分母=一階報名）",
                        color_continuous_scale="RdYlGn", aspect="auto",
                        labels=dict(color="入學率%")
                    )
                    fig_hm.update_layout(height=max(300, t2d["招生管道"].nunique() * 50))
                    dark_friendly_plotly(fig_hm)
                    st.plotly_chart(fig_hm, use_container_width=True)

# ============================================================
# TAB 3: 來源學校分析
# ============================================================
with tab3:
    st.header("🏫 來源學校三階段精準轉換分析")
    st.caption("📐 轉換率分母 = 該校一階報名人數")

    if "畢業學校" not in data.columns:
        st.warning("資料中未包含「畢業學校」欄位")
    else:
        col_t3a, col_t3b, col_t3c, col_t3d = st.columns(4)
        with col_t3a:
            if has_year:
                yr_opts_t3 = ["全部"] + safe_sort_years(data["學年度"])
                sel_yr_t3 = st.selectbox("學年度", yr_opts_t3, key="t3_yr")
            else:
                sel_yr_t3 = "全部"
        with col_t3b:
            top_n = st.slider("Top N 所學校", 5, 50, 20, key="t3_topn")
        with col_t3c:
            if "報考科系" in data.columns:
                dept_opts_t3 = ["全部"] + sorted(data["報考科系"].dropna().unique().tolist())
                sel_dept_t3 = st.selectbox("科系", dept_opts_t3, key="t3_dept")
            else:
                sel_dept_t3 = "全部"
        with col_t3d:
            if has_channel:
                ch_opts_t3 = ["全部"] + sorted(data["招生管道"].dropna().unique().tolist())
                sel_ch_t3 = st.selectbox("管道", ch_opts_t3, key="t3_ch")
            else:
                sel_ch_t3 = "全部"

        t3d = data.copy()
        if sel_yr_t3 != "全部" and has_year:
            t3d = t3d[t3d["學年度"] == sel_yr_t3]
        if sel_dept_t3 != "全部" and "報考科系" in t3d.columns:
            t3d = t3d[t3d["報考科系"] == sel_dept_t3]
        if sel_ch_t3 != "全部" and has_channel:
            t3d = t3d[t3d["招生管道"] == sel_ch_t3]

        school_conv = compute_conversion(t3d, "畢業學校", "學校")
        if school_conv.empty:
            st.info("無符合條件的資料")
        else:
            school_top = school_conv.head(top_n).copy()

            total_s1 = int(school_conv["①一階報名"].sum())
            total_sf = int(school_conv["③最終入學"].sum())
            kpi_s = st.columns(5)
            with kpi_s[0]:
                st.metric("來源學校數", "{:,}".format(len(school_conv)))
            with kpi_s[1]:
                st.metric("一階報名總數", "{:,}".format(total_s1))
            with kpi_s[2]:
                st.metric("最終入學總數", "{:,}".format(total_sf))
            with kpi_s[3]:
                st.metric("整體管道轉換率", "{:.1f}%".format(total_sf / total_s1 * 100 if total_s1 > 0 else 0))
            with kpi_s[4]:
                st.metric("Top {} 校佔比".format(top_n),
                          "{:.1f}%".format(school_top["①一階報名"].sum() / total_s1 * 100 if total_s1 > 0 else 0))

            st.subheader("📋 Top {} 學校轉換率（分母=一階報名）".format(top_n))
            disp_sc = school_top.copy()
            disp_sc["進入二階率(以一階為母體)"] = disp_sc["進入二階率(以一階為母體)"].apply(format_pct)
            disp_sc["最終入學率(以一階為母體)"] = disp_sc["最終入學率(以一階為母體)"].apply(format_pct)
            disp_sc["二階→入學(階段內)"] = disp_sc["二階→入學(階段內)"].apply(format_pct)
            st.dataframe(disp_sc, use_container_width=True, hide_index=True)

            # 水平三階段長條圖
            fig_s3 = go.Figure()
            fig_s3.add_trace(go.Bar(name="一階報名", y=school_top["學校"], x=school_top["①一階報名"],
                                    orientation="h", marker_color="#FDD7B4", text=school_top["①一階報名"], textposition="auto"))
            fig_s3.add_trace(go.Bar(name="二階甄試", y=school_top["學校"], x=school_top["②二階甄試"],
                                    orientation="h", marker_color="#E8792F", text=school_top["②二階甄試"], textposition="auto"))
            fig_s3.add_trace(go.Bar(name="最終入學", y=school_top["學校"], x=school_top["③最終入學"],
                                    orientation="h", marker_color="#B84500", text=school_top["③最終入學"], textposition="auto"))
            fig_s3.update_layout(barmode="group", height=max(500, top_n * 35),
                                 yaxis={"categoryorder": "total ascending"}, title="Top {} 學校三階段人數".format(top_n))
            dark_friendly_plotly(fig_s3)
            st.plotly_chart(fig_s3, use_container_width=True)

            col_sr1, col_sr2 = st.columns(2)
            with col_sr1:
                sorted_r = school_top.sort_values("進入二階率(以一階為母體)", ascending=True)
                fig_r1 = px.bar(sorted_r, x="進入二階率(以一階為母體)", y="學校", orientation="h",
                                title="進入二階率排名（分母=一階報名）",
                                color="進入二階率(以一階為母體)", color_continuous_scale=["#FF6666", "#FFB347", "#66BB6A"],
                                text=sorted_r["進入二階率(以一階為母體)"].apply(lambda v: "{:.1f}%".format(v)))
                fig_r1.update_layout(height=max(450, top_n * 28), showlegend=False, xaxis_title="%")
                dark_friendly_plotly(fig_r1)
                st.plotly_chart(fig_r1, use_container_width=True)
            with col_sr2:
                sorted_rf = school_top.sort_values("最終入學率(以一階為母體)", ascending=True)
                fig_rf = px.bar(sorted_rf, x="最終入學率(以一階為母體)", y="學校", orientation="h",
                                title="最終入學率排名（分母=一階報名）",
                                color="最終入學率(以一階為母體)", color_continuous_scale=["#FF6666", "#FFB347", "#66BB6A"],
                                text=sorted_rf["最終入學率(以一階為母體)"].apply(lambda v: "{:.1f}%".format(v)))
                fig_rf.update_layout(height=max(450, top_n * 28), showlegend=False, xaxis_title="%")
                dark_friendly_plotly(fig_rf)
                st.plotly_chart(fig_rf, use_container_width=True)

            # 流失
            fig_loss = go.Figure()
            fig_loss.add_trace(go.Bar(name="一→二流失", y=school_top["學校"], x=school_top["一→二流失"],
                                      orientation="h", marker_color="#FF8A80", text=school_top["一→二流失"], textposition="auto"))
            fig_loss.add_trace(go.Bar(name="二→入學流失", y=school_top["學校"], x=school_top["二→入學流失"],
                                      orientation="h", marker_color="#EF5350", text=school_top["二→入學流失"], textposition="auto"))
            fig_loss.update_layout(barmode="stack", height=max(450, top_n * 30),
                                   yaxis={"categoryorder": "total ascending"}, title="各學校流失人數（堆疊）", xaxis_title="流失人數")
            dark_friendly_plotly(fig_loss)
            st.plotly_chart(fig_loss, use_container_width=True)

            # 散布圖
            fig_scat = px.scatter(
                school_top, x="①一階報名", y="最終入學率(以一階為母體)",
                size="③最終入學", color="進入二階率(以一階為母體)", hover_name="學校",
                color_continuous_scale=["#FF6666", "#FFB347", "#66BB6A"],
                title="報名量 vs 入學率（氣泡=入學人數）",
                labels={"①一階報名": "一階報名人數", "最終入學率(以一階為母體)": "入學率(%)"}
            )
            fig_scat.update_layout(height=500)
            dark_friendly_plotly(fig_scat)
            st.plotly_chart(fig_scat, use_container_width=True)

            # 經營建議
            st.subheader("⭐ 來源學校經營建議")
            all_sc = school_conv.copy()
            cumulative = 0
            recs = []
            for _, row in all_sc.iterrows():
                cumulative += row["①一階報名"]
                ratio = row["①一階報名"] / total_s1 * 100 if total_s1 > 0 else 0
                cum_ratio = cumulative / total_s1 * 100 if total_s1 > 0 else 0
                r_total = row["最終入學率(以一階為母體)"]
                if cum_ratio <= 50 and r_total >= 30:
                    level = "⭐⭐⭐ 重點深耕"
                    action = "高量高轉換，維持密切合作"
                elif cum_ratio <= 50 and r_total < 30:
                    level = "⭐⭐⭐ 重點改善"
                    action = "報名量大但轉換低，分析流失原因"
                elif cum_ratio <= 80 and r_total >= 50:
                    level = "⭐⭐ 潛力培養"
                    action = "轉換率佳，可增加招生投入"
                elif cum_ratio <= 80:
                    level = "⭐⭐ 持續關注"
                    action = "中等貢獻，定期維護"
                else:
                    level = "⭐ 一般維護"
                    action = "少量貢獻，基本聯繫"
                recs.append({
                    "學校": row["學校"], "報名": int(row["①一階報名"]),
                    "入學": int(row["③最終入學"]),
                    "入學率(以一階為母體)": format_pct(r_total),
                    "佔比": "{:.1f}%".format(ratio), "累積佔比": "{:.1f}%".format(cum_ratio),
                    "建議等級": level, "行動建議": action,
                })
            st.dataframe(pd.DataFrame(recs).head(top_n), use_container_width=True, hide_index=True)

# ============================================================
# TAB 4: 跨年度比較
# ============================================================
with tab4:
    st.header("📈 跨年度比較分析")
    st.caption("📐 轉換率分母 = 一階報名人數")

    if not has_year or data["學年度"].nunique() < 2:
        st.info("💡 需要至少兩個學年度的資料")
    else:
        if has_channel:
            ch_opts_t4 = ["全部（含所有管道）"] + sorted(data["招生管道"].dropna().unique().tolist())
            sel_ch_t4 = st.selectbox("篩選管道", ch_opts_t4, key="t4_ch")
        else:
            sel_ch_t4 = "全部（含所有管道）"

        t4d = data.copy()
        if sel_ch_t4 != "全部（含所有管道）" and has_channel:
            t4d = t4d[t4d["招生管道"] == sel_ch_t4]

        years_sorted = safe_sort_years(t4d["學年度"])
        yearly_stats = []
        for yr in years_sorted:
            yd = t4d[t4d["學年度"] == yr]
            y1 = len(yd)
            y2 = len(yd[yd["二階甄試"] == "✅"]) if "二階甄試" in yd.columns else 0
            yf = len(yd[yd["最終入學"] == "✅"]) if "最終入學" in yd.columns else 0
            yearly_stats.append({
                "學年度": yr, "一階報名": y1, "二階甄試": y2, "最終入學": yf,
                "進入二階率%": y2 / y1 * 100 if y1 > 0 else 0,
                "入學率%": yf / y1 * 100 if y1 > 0 else 0,
            })
        ys_df = pd.DataFrame(yearly_stats)

        disp_ys = ys_df.copy()
        disp_ys["進入二階率%"] = disp_ys["進入二階率%"].apply(format_pct)
        disp_ys["入學率%"] = disp_ys["入學率%"].apply(format_pct)
        st.dataframe(disp_ys, use_container_width=True, hide_index=True)

        col_y1, col_y2 = st.columns(2)
        with col_y1:
            fig_yt = go.Figure()
            for cn, cc in [("一階報名", "#FDD7B4"), ("二階甄試", "#E8792F"), ("最終入學", "#B84500")]:
                fig_yt.add_trace(go.Bar(name=cn, x=ys_df["學年度"], y=ys_df[cn], marker_color=cc,
                                        text=ys_df[cn], textposition="outside"))
            fig_yt.update_layout(barmode="group", title="歷年各階段人數", height=400)
            dark_friendly_plotly(fig_yt)
            st.plotly_chart(fig_yt, use_container_width=True)
        with col_y2:
            raw_ys = pd.DataFrame(yearly_stats)
            fig_yr = go.Figure()
            fig_yr.add_trace(go.Scatter(
                x=raw_ys["學年度"], y=raw_ys["進入二階率%"],
                mode="lines+markers+text", name="進入二階率",
                line=dict(color="#E8792F", width=2),
                text=raw_ys["進入二階率%"].apply(lambda v: "{:.1f}%".format(v)), textposition="top left"
            ))
            fig_yr.add_trace(go.Scatter(
                x=raw_ys["學年度"], y=raw_ys["入學率%"],
                mode="lines+markers+text", name="入學率(分母=一階)",
                line=dict(color="#FFFFFF", width=3),
                text=raw_ys["入學率%"].apply(lambda v: "{:.1f}%".format(v)), textposition="bottom center"
            ))
            fig_yr.update_layout(title="歷年轉換率趨勢（分母=一階報名）", yaxis_title="%", height=400)
            dark_friendly_plotly(fig_yr)
            st.plotly_chart(fig_yr, use_container_width=True)

        st.subheader("📊 年度增減")
        for i in range(1, len(yearly_stats)):
            prev_n = yearly_stats[i-1]["一階報名"]
            curr_n = yearly_stats[i]["一階報名"]
            diff = curr_n - prev_n
            pct = diff / prev_n * 100 if prev_n > 0 else 0
            prev_f = yearly_stats[i-1]["最終入學"]
            curr_f = yearly_stats[i]["最終入學"]
            diff_f = curr_f - prev_f
            col_ig1, col_ig2 = st.columns(2)
            with col_ig1:
                st.metric("{} 一階報名".format(yearly_stats[i]["學年度"]),
                          "{:,}".format(curr_n), "{:+d}（{:+.1f}%）".format(diff, pct))
            with col_ig2:
                st.metric("{} 最終入學".format(yearly_stats[i]["學年度"]),
                          "{:,}".format(curr_f), "{:+d}".format(diff_f))

        if "報考科系" in t4d.columns and len(years_sorted) >= 2:
            st.subheader("📊 科系歷年增減")
            ya, yb = years_sorted[-2], years_sorted[-1]
            da = t4d[t4d["學年度"] == ya]["報考科系"].value_counts()
            db = t4d[t4d["學年度"] == yb]["報考科系"].value_counts()
            comp = []
            for dept in sorted(set(da.index) | set(db.index)):
                a_v, b_v = int(da.get(dept, 0)), int(db.get(dept, 0))
                d_v = b_v - a_v
                icon = "📈" if d_v > 0 else ("📉" if d_v < 0 else "➡️")
                comp.append({"科系": dept, ya: a_v, yb: b_v, "增減": d_v, "趨勢": icon})
            st.dataframe(pd.DataFrame(comp).sort_values("增減", ascending=False), use_container_width=True, hide_index=True)

# ============================================================
# TAB 5: 資料檢視與匯出
# ============================================================
with tab5:
    st.header("🔍 資料檢視與匯出")

    with st.expander("📋 各階段原始資料", expanded=False):
        stabs = st.tabs(["一階報名", "二階甄試", "最終入學"])
        with stabs[0]:
            if not st.session_state.phase1_data.empty:
                st.caption("共 {:,} 筆".format(len(st.session_state.phase1_data)))
                st.dataframe(st.session_state.phase1_data.head(100), use_container_width=True, hide_index=True)
            else:
                st.info("尚無一階資料")
        with stabs[1]:
            if not st.session_state.phase2_data.empty:
                st.caption("共 {:,} 筆".format(len(st.session_state.phase2_data)))
                st.dataframe(st.session_state.phase2_data.head(100), use_container_width=True, hide_index=True)
            else:
                st.info("尚無二階資料")
        with stabs[2]:
            if not st.session_state.final_data.empty:
                st.caption("共 {:,} 筆".format(len(st.session_state.final_data)))
                st.dataframe(st.session_state.final_data.head(100), use_container_width=True, hide_index=True)
            else:
                st.info("尚無入學資料")

    st.subheader("🔗 合併後資料")
    col_e1, col_e2, col_e3, col_e4, col_e5 = st.columns(5)
    export_data = data.copy()
    with col_e1:
        if has_year:
            yr_opts_e = ["全部"] + safe_sort_years(data["學年度"])
            sel_yr_e = st.selectbox("學年度", yr_opts_e, key="e_yr")
            if sel_yr_e != "全部":
                export_data = export_data[export_data["學年度"] == sel_yr_e]
    with col_e2:
        if "報考科系" in data.columns:
            dept_opts_e = ["全部"] + sorted(data["報考科系"].dropna().unique().tolist())
            sel_dept_e = st.selectbox("科系", dept_opts_e, key="e_dept")
            if sel_dept_e != "全部":
                export_data = export_data[export_data["報考科系"] == sel_dept_e]
    with col_e3:
        if "目前狀態" in data.columns:
            stg_opts_e = ["全部"] + sorted(data["目前狀態"].dropna().unique().tolist())
            sel_stg_e = st.selectbox("階段", stg_opts_e, key="e_stg")
            if sel_stg_e != "全部":
                export_data = export_data[export_data["目前狀態"] == sel_stg_e]
    with col_e4:
        if has_channel:
            ch_opts_e = ["全部"] + sorted(data["招生管道"].dropna().unique().tolist())
            sel_ch_e = st.selectbox("管道", ch_opts_e, key="e_ch")
            if sel_ch_e != "全部":
                export_data = export_data[export_data["招生管道"] == sel_ch_e]
    with col_e5:
        if "畢業學校" in data.columns:
            search_sch = st.text_input("搜尋學校", key="e_sch")
            if search_sch:
                export_data = export_data[export_data["畢業學校"].astype(str).str.contains(search_sch, na=False)]

    st.caption("篩選結果：{:,} 筆".format(len(export_data)))
    hide_cols = ["身分證字號"]
    display_cols = [c for c in export_data.columns if c not in hide_cols]
    st.dataframe(export_data[display_cols], use_container_width=True, hide_index=True, height=400)

    st.subheader("📥 匯出資料")
    col_dl1, col_dl2, col_dl3 = st.columns(3)
    with col_dl1:
        csv_buf = io.BytesIO()
        export_data.to_csv(csv_buf, index=False, encoding="utf-8-sig")
        st.download_button("⬇️ CSV", csv_buf.getvalue(), "招生分析.csv", "text/csv", use_container_width=True)
    with col_dl2:
        xls_buf = io.BytesIO()
        with pd.ExcelWriter(xls_buf, engine="openpyxl") as writer:
            export_data.to_excel(writer, index=False, sheet_name="合併資料")
            if "報考科系" in data.columns:
                dc = compute_conversion(export_data, "報考科系", "科系")
                if not dc.empty:
                    dc.to_excel(writer, index=False, sheet_name="科系轉換率")
            if "畢業學校" in data.columns:
                sc = compute_conversion(export_data, "畢業學校", "學校")
                if not sc.empty:
                    sc.to_excel(writer, index=False, sheet_name="學校轉換率")
            if has_channel:
                cc = compute_conversion(export_data, "招生管道", "招生管道")
                if not cc.empty:
                    cc.to_excel(writer, index=False, sheet_name="管道轉換率")
        st.download_button("⬇️ Excel（含轉換率）", xls_buf.getvalue(), "招生分析_含轉換率.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with col_dl3:
        if has_channel:
            ch_csv = io.BytesIO()
            ch_full = compute_conversion(export_data, "招生管道", "招生管道")
            if not ch_full.empty:
                ch_full.to_csv(ch_csv, index=False, encoding="utf-8-sig")
                st.download_button("⬇️ 管道轉換率 CSV", ch_csv.getvalue(), "管道轉換率.csv", "text/csv", use_container_width=True)

    st.subheader("📋 資料品質報告")
    quality = []
    for col_q in data.columns:
        cs = safe_get_series(data, col_q)
        nn = int(cs.notna().sum())
        tr = len(data)
        quality.append({
            "欄位": col_q, "非空筆數": nn, "總筆數": tr,
            "完整率": "{:.1f}%".format(nn / tr * 100) if tr > 0 else "0%",
            "空值數": tr - nn,
        })
    st.dataframe(pd.DataFrame(quality), use_container_width=True, hide_index=True)
