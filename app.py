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
            pass
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
        df["學年度"] = col_data.apply(
            lambda x: str(int(float(x))) if pd.notna(x) else None
        )
    for col_name in ["報考科系", "畢業學校", "姓名", "身分證字號"]:
        if col_name in df.columns:
            col_data = safe_get_series(df, col_name)
            df[col_name] = col_data.astype(str).str.strip()
            df[col_name] = df[col_name].replace({"nan": None, "None": None, "": None})
    return df

def parse_coordinates(df):
    df = df.copy()
    if "緯度" in df.columns and "經度" in df.columns:
        lat_col = safe_get_series(df, "緯度")
        lon_col = safe_get_series(df, "經度")
        df["緯度"] = pd.to_numeric(lat_col, errors="coerce")
        df["經度"] = pd.to_numeric(lon_col, errors="coerce")
        return df
    if "經緯度" not in df.columns:
        return df
    coord_col = safe_get_series(df, "經緯度")
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


def compute_school_conversion(df, school_col="畢業學校"):
    """計算每所學校的三階段精準轉換率"""
    if school_col not in df.columns:
        return pd.DataFrame()

    schools = df[school_col].dropna().unique()
    results = []
    for school in schools:
        sd = df[df[school_col] == school]
        n1 = len(sd)
        n2 = len(sd[sd["二階甄試"] == "✅"]) if "二階甄試" in sd.columns else 0
        nf = len(sd[sd["最終入學"] == "✅"]) if "最終入學" in sd.columns else 0

        rate_1to2 = n2 / n1 * 100 if n1 > 0 else 0
        rate_2tof = nf / n2 * 100 if n2 > 0 else 0
        rate_1tof = nf / n1 * 100 if n1 > 0 else 0
        lost_1to2 = n1 - n2
        lost_2tof = n2 - nf

        results.append({
            "學校": school,
            "①一階報名": n1,
            "②二階甄試": n2,
            "③最終入學": nf,
            "一→二轉換率": rate_1to2,
            "二→入學轉換率": rate_2tof,
            "總轉換率(一→入學)": rate_1tof,
            "一→二流失": lost_1to2,
            "二→入學流失": lost_2tof,
        })
    result_df = pd.DataFrame(results)
    if not result_df.empty:
        result_df = result_df.sort_values("①一階報名", ascending=False)
    return result_df


def compute_dept_conversion(df, dept_col="報考科系"):
    """計算每個科系的三階段精準轉換率"""
    if dept_col not in df.columns:
        return pd.DataFrame()

    depts = df[dept_col].dropna().unique()
    results = []
    for dept in sorted(depts):
        dd = df[df[dept_col] == dept]
        n1 = len(dd)
        n2 = len(dd[dd["二階甄試"] == "✅"]) if "二階甄試" in dd.columns else 0
        nf = len(dd[dd["最終入學"] == "✅"]) if "最終入學" in dd.columns else 0

        rate_1to2 = n2 / n1 * 100 if n1 > 0 else 0
        rate_2tof = nf / n2 * 100 if n2 > 0 else 0
        rate_1tof = nf / n1 * 100 if n1 > 0 else 0

        results.append({
            "科系": dept,
            "①一階報名": n1,
            "②二階甄試": n2,
            "③最終入學": nf,
            "一→二轉換率": rate_1to2,
            "二→入學轉換率": rate_2tof,
            "總轉換率(一→入學)": rate_1tof,
            "一→二流失": n1 - n2,
            "二→入學流失": n2 - nf,
        })
    return pd.DataFrame(results)


def format_pct(val):
    return "{:.1f}%".format(val)


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
        "<p style='font-size:13px; color:gray;'>學年度、科系、姓名<br>學校、座標、身分證</p></div>"
        "<div style='background:#FFF3E8; padding:20px; border-radius:12px; width:220px;'>"
        "<h3>📝 二階甄試</h3>"
        "<p style='font-size:13px; color:gray;'>姓名、學年度<br>+ 成績欄位（選填）</p></div>"
        "<div style='background:#FFF3E8; padding:20px; border-radius:12px; width:220px;'>"
        "<h3>🎓 最終入學</h3>"
        "<p style='font-size:13px; color:gray;'>僅需：姓名<br>（+學年度更精準）</p></div>"
        "</div></div>",
        unsafe_allow_html=True
    )
    st.stop()

# ============================================================
# 分頁
# ============================================================
tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🔄 招生漏斗",
    "🗺️ 地圖視覺化",
    "📊 科系轉換分析",
    "🏫 來源學校轉換分析",
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

    n_phase1 = len(funnel_data[funnel_data["一階報名"] == "✅"]) if "一階報名" in funnel_data.columns else len(funnel_data)
    n_phase2 = len(funnel_data[funnel_data["二階甄試"] == "✅"]) if "二階甄試" in funnel_data.columns else 0
    n_final = len(funnel_data[funnel_data["最終入學"] == "✅"]) if "最終入學" in funnel_data.columns else 0

    kpi_cols = st.columns(5)
    with kpi_cols[0]:
        st.metric("📋 一階報名", "{:,} 人".format(n_phase1))
    with kpi_cols[1]:
        rate_12 = n_phase2 / n_phase1 * 100 if n_phase1 > 0 else 0
        st.metric("📝 二階甄試", "{:,} 人".format(n_phase2), "一→二 {:.1f}%".format(rate_12))
    with kpi_cols[2]:
        rate_2f = n_final / n_phase2 * 100 if n_phase2 > 0 else 0
        st.metric("🎓 最終入學", "{:,} 人".format(n_final), "二→入學 {:.1f}%".format(rate_2f))
    with kpi_cols[3]:
        rate_1f = n_final / n_phase1 * 100 if n_phase1 > 0 else 0
        st.metric("🎯 總轉換率", "{:.1f}%".format(rate_1f))
    with kpi_cols[4]:
        total_lost = n_phase1 - n_final
        st.metric("📉 總流失", "{:,} 人".format(total_lost))

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
        st.markdown("**📉 各階段流失明細**")
        lost_12 = n_phase1 - n_phase2
        lost_2f = n_phase2 - n_final
        loss_rate_12 = lost_12 / n_phase1 * 100 if n_phase1 > 0 else 0
        loss_rate_2f = lost_2f / n_phase2 * 100 if n_phase2 > 0 else 0
        st.markdown(
            "| 階段轉換 | 流失人數 | 流失率 | 通過率 |\n"
            "|----------|---------|--------|--------|\n"
            "| 一階→二階 | {:,} | {:.1f}% | {:.1f}% |\n"
            "| 二階→入學 | {:,} | {:.1f}% | {:.1f}% |\n"
            "| **一階→入學** | **{:,}** | **{:.1f}%** | **{:.1f}%** |".format(
                lost_12, loss_rate_12, rate_12,
                lost_2f, loss_rate_2f, rate_2f,
                total_lost, total_lost / n_phase1 * 100 if n_phase1 > 0 else 0, rate_1f
            )
        )

        if n_phase1 > 0:
            st.markdown("---")
            st.markdown("**🔍 流失瓶頸判斷**")
            if loss_rate_12 > loss_rate_2f:
                st.error("⚠️ 主要瓶頸在「一階→二階」（流失率 {:.1f}%），建議加強報名後的追蹤聯繫".format(loss_rate_12))
            elif loss_rate_2f > loss_rate_12:
                st.error("⚠️ 主要瓶頸在「二階→入學」（流失率 {:.1f}%），建議優化甄試流程與錄取通知".format(loss_rate_2f))
            else:
                st.success("✅ 各階段流失率相近，建議全面均衡優化")

    if "報考科系" in funnel_data.columns:
        st.subheader("📊 各科系三階段轉換率")
        dept_conv_df = compute_dept_conversion(funnel_data)
        if not dept_conv_df.empty:
            display_dept = dept_conv_df.copy()
            display_dept["一→二轉換率"] = display_dept["一→二轉換率"].apply(format_pct)
            display_dept["二→入學轉換率"] = display_dept["二→入學轉換率"].apply(format_pct)
            display_dept["總轉換率(一→入學)"] = display_dept["總轉換率(一→入學)"].apply(format_pct)
            st.dataframe(display_dept, use_container_width=True, hide_index=True)

            fig_dept_conv = go.Figure()
            fig_dept_conv.add_trace(go.Bar(
                name="一→二轉換率", x=dept_conv_df["科系"],
                y=dept_conv_df["一→二轉換率"], marker_color="#FDD7B4",
                text=dept_conv_df["一→二轉換率"].apply(lambda v: "{:.1f}%".format(v)),
                textposition="outside"
            ))
            fig_dept_conv.add_trace(go.Bar(
                name="二→入學轉換率", x=dept_conv_df["科系"],
                y=dept_conv_df["二→入學轉換率"], marker_color="#E8792F",
                text=dept_conv_df["二→入學轉換率"].apply(lambda v: "{:.1f}%".format(v)),
                textposition="outside"
            ))
            fig_dept_conv.add_trace(go.Bar(
                name="總轉換率(一→入學)", x=dept_conv_df["科系"],
                y=dept_conv_df["總轉換率(一→入學)"], marker_color="#B84500",
                text=dept_conv_df["總轉換率(一→入學)"].apply(lambda v: "{:.1f}%".format(v)),
                textposition="outside"
            ))
            fig_dept_conv.update_layout(
                barmode="group", title="各科系階段轉換率比較（%）",
                yaxis_title="轉換率 (%)", height=500
            )
            st.plotly_chart(fig_dept_conv, use_container_width=True)

    if "學年度" in data.columns and data["學年度"].nunique() > 1 and sel_yr_funnel == "全部":
        st.subheader("📈 歷年漏斗趨勢")
        yearly_funnel = []
        for yr in safe_sort_years(data["學年度"]):
            yd = data[data["學年度"] == yr]
            y1 = len(yd)
            y2 = len(yd[yd["二階甄試"] == "✅"]) if "二階甄試" in yd.columns else 0
            yf = len(yd[yd["最終入學"] == "✅"]) if "最終入學" in yd.columns else 0
            yearly_funnel.append({
                "學年度": yr, "一階報名": y1, "二階甄試": y2, "最終入學": yf,
                "一→二%": y2 / y1 * 100 if y1 > 0 else 0,
                "二→入學%": yf / y2 * 100 if y2 > 0 else 0,
                "總轉換%": yf / y1 * 100 if y1 > 0 else 0,
            })
        yf_df = pd.DataFrame(yearly_funnel)

        col_yf1, col_yf2 = st.columns(2)
        with col_yf1:
            fig_yf = go.Figure()
            fig_yf.add_trace(go.Scatter(
                x=yf_df["學年度"], y=yf_df["一階報名"],
                mode="lines+markers+text", name="一階報名",
                line=dict(color="#FDD7B4", width=3),
                text=yf_df["一階報名"], textposition="top center"
            ))
            fig_yf.add_trace(go.Scatter(
                x=yf_df["學年度"], y=yf_df["二階甄試"],
                mode="lines+markers+text", name="二階甄試",
                line=dict(color="#E8792F", width=3),
                text=yf_df["二階甄試"], textposition="top center"
            ))
            fig_yf.add_trace(go.Scatter(
                x=yf_df["學年度"], y=yf_df["最終入學"],
                mode="lines+markers+text", name="最終入學",
                line=dict(color="#B84500", width=3),
                text=yf_df["最終入學"], textposition="top center"
            ))
            fig_yf.update_layout(title="歷年各階段人數", height=400)
            st.plotly_chart(fig_yf, use_container_width=True)

        with col_yf2:
            fig_rate = go.Figure()
            fig_rate.add_trace(go.Scatter(
                x=yf_df["學年度"], y=yf_df["一→二%"],
                mode="lines+markers+text", name="一→二%",
                line=dict(color="#E8792F", width=2, dash="dot"),
                text=yf_df["一→二%"].apply(lambda v: "{:.1f}%".format(v)),
                textposition="top center"
            ))
            fig_rate.add_trace(go.Scatter(
                x=yf_df["學年度"], y=yf_df["二→入學%"],
                mode="lines+markers+text", name="二→入學%",
                line=dict(color="#B84500", width=2, dash="dot"),
                text=yf_df["二→入學%"].apply(lambda v: "{:.1f}%".format(v)),
                textposition="top center"
            ))
            fig_rate.add_trace(go.Scatter(
                x=yf_df["學年度"], y=yf_df["總轉換%"],
                mode="lines+markers+text", name="總轉換%",
                line=dict(color="#333333", width=3),
                text=yf_df["總轉換%"].apply(lambda v: "{:.1f}%".format(v)),
                textposition="top center"
            ))
            fig_rate.update_layout(title="歷年各階段轉換率趨勢（%）", yaxis_title="%", height=400)
            st.plotly_chart(fig_rate, use_container_width=True)

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
# TAB 2: 科系轉換分析
# ============================================================
with tab2:
    st.header("📊 科系三階段轉換分析")

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

        dept_conv = compute_dept_conversion(t2_data)

        if not dept_conv.empty:
            st.subheader("📋 各科系轉換率總表")
            display_dc = dept_conv.copy()
            display_dc["一→二轉換率"] = display_dc["一→二轉換率"].apply(format_pct)
            display_dc["二→入學轉換率"] = display_dc["二→入學轉換率"].apply(format_pct)
            display_dc["總轉換率(一→入學)"] = display_dc["總轉換率(一→入學)"].apply(format_pct)
            st.dataframe(display_dc, use_container_width=True, hide_index=True)

            kpi_dept_cols = st.columns(4)
            with kpi_dept_cols[0]:
                st.metric("科系數", "{}".format(len(dept_conv)))
            with kpi_dept_cols[1]:
                avg_12 = dept_conv["一→二轉換率"].mean()
                st.metric("平均一→二轉換率", "{:.1f}%".format(avg_12))
            with kpi_dept_cols[2]:
                avg_2f = dept_conv["二→入學轉換率"].mean()
                st.metric("平均二→入學轉換率", "{:.1f}%".format(avg_2f))
            with kpi_dept_cols[3]:
                avg_1f = dept_conv["總轉換率(一→入學)"].mean()
                st.metric("平均總轉換率", "{:.1f}%".format(avg_1f))

            col_dc1, col_dc2 = st.columns(2)
            with col_dc1:
                fig_dc_bar = go.Figure()
                fig_dc_bar.add_trace(go.Bar(
                    name="一階報名", x=dept_conv["科系"], y=dept_conv["①一階報名"],
                    marker_color="#FDD7B4"
                ))
                fig_dc_bar.add_trace(go.Bar(
                    name="二階甄試", x=dept_conv["科系"], y=dept_conv["②二階甄試"],
                    marker_color="#E8792F"
                ))
                fig_dc_bar.add_trace(go.Bar(
                    name="最終入學", x=dept_conv["科系"], y=dept_conv["③最終入學"],
                    marker_color="#B84500"
                ))
                fig_dc_bar.update_layout(barmode="group", title="各科系三階段人數比較", height=450)
                st.plotly_chart(fig_dc_bar, use_container_width=True)

            with col_dc2:
                fig_dc_rate = go.Figure()
                fig_dc_rate.add_trace(go.Scatter(
                    x=dept_conv["科系"], y=dept_conv["一→二轉換率"],
                    mode="lines+markers+text", name="一→二",
                    line=dict(color="#E8792F", width=2),
                    text=dept_conv["一→二轉換率"].apply(lambda v: "{:.1f}%".format(v)),
                    textposition="top center"
                ))
                fig_dc_rate.add_trace(go.Scatter(
                    x=dept_conv["科系"], y=dept_conv["二→入學轉換率"],
                    mode="lines+markers+text", name="二→入學",
                    line=dict(color="#B84500", width=2),
                    text=dept_conv["二→入學轉換率"].apply(lambda v: "{:.1f}%".format(v)),
                    textposition="top center"
                ))
                fig_dc_rate.add_trace(go.Scatter(
                    x=dept_conv["科系"], y=dept_conv["總轉換率(一→入學)"],
                    mode="lines+markers+text", name="總轉換率",
                    line=dict(color="#333", width=3),
                    text=dept_conv["總轉換率(一→入學)"].apply(lambda v: "{:.1f}%".format(v)),
                    textposition="top center"
                ))
                fig_dc_rate.update_layout(title="各科系轉換率比較（%）", yaxis_title="%", height=450)
                st.plotly_chart(fig_dc_rate, use_container_width=True)

            st.subheader("🔍 科系瓶頸分析")
            for _, row in dept_conv.iterrows():
                dept_name = row["科系"]
                r12 = row["一→二轉換率"]
                r2f = row["二→入學轉換率"]
                loss12 = row["一→二流失"]
                loss2f = row["二→入學流失"]
                if r12 < 30:
                    bottleneck = "🔴 一→二嚴重流失"
                elif r12 < 50:
                    bottleneck = "🟡 一→二中度流失"
                elif r2f < 50:
                    bottleneck = "🟠 二→入學流失偏高"
                else:
                    bottleneck = "🟢 轉換良好"
                st.caption(
                    "**{}**：一→二 {:.1f}%（流失 {}人）｜二→入學 {:.1f}%（流失 {}人）→ {}".format(
                        dept_name, r12, loss12, r2f, loss2f, bottleneck
                    )
                )

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

# ============================================================
# TAB 3: 來源學校轉換分析
# ============================================================
with tab3:
    st.header("🏫 來源學校三階段精準轉換分析")

    if "畢業學校" not in data.columns:
        st.warning("資料中未包含「畢業學校」欄位")
    else:
        col_t3f1, col_t3f2, col_t3f3 = st.columns(3)
        with col_t3f1:
            if "學年度" in data.columns:
                yr_opts_t3 = ["全部"] + safe_sort_years(data["學年度"])
                sel_year_t3 = st.selectbox("選擇學年度", yr_opts_t3, key="t3_year")
            else:
                sel_year_t3 = "全部"
        with col_t3f2:
            top_n = st.slider("顯示前 N 所學校", 5, 50, 20, key="t3_topn")
        with col_t3f3:
            if "報考科系" in data.columns:
                dept_opts_t3 = ["全部"] + sorted(data["報考科系"].dropna().unique().tolist())
                sel_dept_t3 = st.selectbox("篩選科系", dept_opts_t3, key="t3_dept")
            else:
                sel_dept_t3 = "全部"

        t3_data = data.copy()
        if sel_year_t3 != "全部" and "學年度" in t3_data.columns:
            t3_data = t3_data[t3_data["學年度"] == sel_year_t3]
        if sel_dept_t3 != "全部" and "報考科系" in t3_data.columns:
            t3_data = t3_data[t3_data["報考科系"] == sel_dept_t3]

        school_conv = compute_school_conversion(t3_data)
        if school_conv.empty:
            st.info("無符合條件的資料")
        else:
            school_top = school_conv.head(top_n).copy()

            # ---- KPI ----
            st.subheader("📊 整體轉換率概覽")
            total_s1 = int(school_conv["①一階報名"].sum())
            total_s2 = int(school_conv["②二階甄試"].sum())
            total_sf = int(school_conv["③最終入學"].sum())

            kpi_s = st.columns(6)
            with kpi_s[0]:
                st.metric("來源學校數", "{:,}".format(len(school_conv)))
            with kpi_s[1]:
                st.metric("一階報名總數", "{:,}".format(total_s1))
            with kpi_s[2]:
                st.metric("二階甄試總數", "{:,}".format(total_s2))
            with kpi_s[3]:
                st.metric("最終入學總數", "{:,}".format(total_sf))
            with kpi_s[4]:
                overall_12 = total_s2 / total_s1 * 100 if total_s1 > 0 else 0
                st.metric("整體一→二%", "{:.1f}%".format(overall_12))
            with kpi_s[5]:
                overall_1f = total_sf / total_s1 * 100 if total_s1 > 0 else 0
                st.metric("整體總轉換%", "{:.1f}%".format(overall_1f))

            # ---- 精準轉換率表 ----
            st.subheader("📋 Top {} 學校三階段精準轉換率".format(top_n))
            display_sc = school_top.copy()
            display_sc["一→二轉換率"] = display_sc["一→二轉換率"].apply(format_pct)
            display_sc["二→入學轉換率"] = display_sc["二→入學轉換率"].apply(format_pct)
            display_sc["總轉換率(一→入學)"] = display_sc["總轉換率(一→入學)"].apply(format_pct)
            st.dataframe(display_sc, use_container_width=True, hide_index=True)

            # ---- 三階段人數堆疊圖 ----
            st.subheader("📊 各學校三階段人數比較")
            fig_s3 = go.Figure()
            fig_s3.add_trace(go.Bar(
                name="一階報名", y=school_top["學校"], x=school_top["①一階報名"],
                orientation="h", marker_color="#FDD7B4",
                text=school_top["①一階報名"], textposition="auto"
            ))
            fig_s3.add_trace(go.Bar(
                name="二階甄試", y=school_top["學校"], x=school_top["②二階甄試"],
                orientation="h", marker_color="#E8792F",
                text=school_top["②二階甄試"], textposition="auto"
            ))
            fig_s3.add_trace(go.Bar(
                name="最終入學", y=school_top["學校"], x=school_top["③最終入學"],
                orientation="h", marker_color="#B84500",
                text=school_top["③最終入學"], textposition="auto"
            ))
            fig_s3.update_layout(
                barmode="group", height=max(500, top_n * 35),
                yaxis={"categoryorder": "total ascending"},
                title="Top {} 學校三階段人數".format(top_n)
            )
            st.plotly_chart(fig_s3, use_container_width=True)

            # ---- 轉換率比較圖 ----
            st.subheader("📈 各學校轉換率比較")
            col_sc1, col_sc2 = st.columns(2)

            with col_sc1:
                fig_r12 = px.bar(
                    school_top.sort_values("一→二轉換率", ascending=True),
                    x="一→二轉換率", y="學校", orientation="h",
                    title="一階→二階 轉換率排名",
                    color="一→二轉換率",
                    color_continuous_scale=["#FFCCCC", "#FF6600", "#006600"],
                    text=school_top.sort_values("一→二轉換率", ascending=True)["一→二轉換率"].apply(
                        lambda v: "{:.1f}%".format(v)
                    )
                )
                fig_r12.update_layout(
                    height=max(450, top_n * 28), showlegend=False,
                    xaxis_title="轉換率（%）"
                )
                st.plotly_chart(fig_r12, use_container_width=True)

            with col_sc2:
                fig_r2f = px.bar(
                    school_top.sort_values("總轉換率(一→入學)", ascending=True),
                    x="總轉換率(一→入學)", y="學校", orientation="h",
                    title="一階→入學 總轉換率排名",
                    color="總轉換率(一→入學)",
                    color_continuous_scale=["#FFCCCC", "#FF6600", "#006600"],
                    text=school_top.sort_values("總轉換率(一→入學)", ascending=True)["總轉換率(一→入學)"].apply(
                        lambda v: "{:.1f}%".format(v)
                    )
                )
                fig_r2f.update_layout(
                    height=max(450, top_n * 28), showlegend=False,
                    xaxis_title="轉換率（%）"
                )
                st.plotly_chart(fig_r2f, use_container_width=True)

            # ---- 流失分析 ----
            st.subheader("📉 各學校流失分析")
            fig_loss = go.Figure()
            fig_loss.add_trace(go.Bar(
                name="一→二流失", y=school_top["學校"], x=school_top["一→二流失"],
                orientation="h", marker_color="#FF9999",
                text=school_top["一→二流失"], textposition="auto"
            ))
            fig_loss.add_trace(go.Bar(
                name="二→入學流失", y=school_top["學校"], x=school_top["二→入學流失"],
                orientation="h", marker_color="#CC3333",
                text=school_top["二→入學流失"], textposition="auto"
            ))
            fig_loss.update_layout(
                barmode="stack", height=max(450, top_n * 30),
                yaxis={"categoryorder": "total ascending"},
                title="各學校各階段流失人數（堆疊）",
                xaxis_title="流失人數"
            )
            st.plotly_chart(fig_loss, use_container_width=True)

            # ---- 散布圖 ----
            st.subheader("🔍 報名量 vs 總轉換率 散布圖")
            fig_scatter = px.scatter(
                school_top, x="①一階報名", y="總轉換率(一→入學)",
                size="③最終入學", color="一→二轉換率",
                hover_name="學校",
                color_continuous_scale=["#FF6666", "#FFCC00", "#00AA00"],
                title="報名量 vs 總轉換率（氣泡大小＝入學人數）",
                labels={"①一階報名": "一階報名人數", "總轉換率(一→入學)": "總轉換率(%)"}
            )
            fig_scatter.update_layout(height=500)
            st.plotly_chart(fig_scatter, use_container_width=True)

            # ---- 學校 × 科系交叉分析 ----
            if "報考科系" in t3_data.columns and sel_dept_t3 == "全部":
                st.subheader("🔀 來源學校 × 科系 入學交叉分析")
                top_school_names = school_top["學校"].tolist()
                enrolled_data = t3_data[
                    (t3_data["畢業學校"].isin(top_school_names)) &
                    (t3_data.get("最終入學", "") == "✅")
                ]
                if len(enrolled_data) > 0:
                    cross_table = pd.crosstab(enrolled_data["畢業學校"], enrolled_data["報考科系"])
                    fig_heat = px.imshow(
                        cross_table, title="來源學校 vs 報考科系（入學人數）",
                        color_continuous_scale="Oranges", aspect="auto"
                    )
                    fig_heat.update_layout(height=max(400, len(top_school_names) * 28))
                    st.plotly_chart(fig_heat, use_container_width=True)
                else:
                    st.info("無入學交叉資料可顯示")

            # ---- 管理建議 ----
            st.subheader("⭐ 來源學校經營建議")
            all_sc = school_conv.copy()
            total_all = int(all_sc["①一階報名"].sum())
            cumulative = 0
            recs = []
            for _, row in all_sc.iterrows():
                cumulative += row["①一階報名"]
                ratio = row["①一階報名"] / total_all * 100 if total_all > 0 else 0
                cum_ratio = cumulative / total_all * 100 if total_all > 0 else 0
                r_total = row["總轉換率(一→入學)"]
                r_12 = row["一→二轉換率"]

                if cum_ratio <= 50 and r_total >= 30:
                    level = "⭐⭐⭐ 重點深耕"
                    action = "高量高轉換，維持密切合作"
                elif cum_ratio <= 50 and r_total < 30:
                    level = "⭐⭐⭐ 重點改善"
                    action = "報名量大但轉換低，需分析流失原因"
                elif cum_ratio <= 80 and r_total >= 50:
                    level = "⭐⭐ 潛力培養"
                    action = "轉換率佳，可增加招生投入"
                elif cum_ratio <= 80:
                    level = "⭐⭐ 持續關注"
                    action = "中等貢獻，定期維護關係"
                else:
                    level = "⭐ 一般維護"
                    action = "少量貢獻，基本聯繫即可"

                recs.append({
                    "學校": row["學校"],
                    "報名": int(row["①一階報名"]),
                    "入學": int(row["③最終入學"]),
                    "總轉換率": format_pct(r_total),
                    "佔比": "{:.1f}%".format(ratio),
                    "累積佔比": "{:.1f}%".format(cum_ratio),
                    "建議等級": level,
                    "行動建議": action,
                })

            recs_df = pd.DataFrame(recs).head(top_n)
            st.dataframe(recs_df, use_container_width=True, hide_index=True)

            # ---- 歷年學校趨勢 ----
            if "學年度" in data.columns and data["學年度"].nunique() > 1 and sel_year_t3 == "全部":
                st.subheader("📈 Top 10 學校歷年入學趨勢")
                top10_schools = school_conv.head(10)["學校"].tolist()
                yearly_school = []
                for yr in safe_sort_years(data["學年度"]):
                    yd = data[data["學年度"] == yr]
                    for sch in top10_schools:
                        sd = yd[yd["畢業學校"] == sch]
                        n1_ys = len(sd)
                        nf_ys = len(sd[sd["最終入學"] == "✅"]) if "最終入學" in sd.columns else 0
                        rate_ys = nf_ys / n1_ys * 100 if n1_ys > 0 else 0
                        yearly_school.append({
                            "學年度": yr, "學校": sch,
                            "報名": n1_ys, "入學": nf_ys, "轉換率": rate_ys
                        })
                ys_school_df = pd.DataFrame(yearly_school)

                col_yst1, col_yst2 = st.columns(2)
                with col_yst1:
                    fig_ys_n = px.line(
                        ys_school_df, x="學年度", y="報名", color="學校",
                        markers=True, title="Top 10 學校歷年報名人數"
                    )
                    fig_ys_n.update_layout(height=450)
                    st.plotly_chart(fig_ys_n, use_container_width=True)

                with col_yst2:
                    fig_ys_r = px.line(
                        ys_school_df, x="學年度", y="轉換率", color="學校",
                        markers=True, title="Top 10 學校歷年總轉換率（%）"
                    )
                    fig_ys_r.update_layout(height=450, yaxis_title="%")
                    st.plotly_chart(fig_ys_r, use_container_width=True)

# ============================================================
# TAB 4: 跨年度比較
# ============================================================
with tab4:
    st.header("📈 跨年度比較分析")

    if "學年度" not in data.columns or data["學年度"].nunique() < 2:
        st.info("💡 需要至少兩個學年度的資料才能進行跨年比較。")
    else:
        years_sorted = safe_sort_years(data["學年度"])

        st.subheader("📊 歷年招生各階段趨勢")
        yearly_stats = []
        for yr in years_sorted:
            yd = data[data["學年度"] == yr]
            y1 = len(yd)
            y2 = len(yd[yd["二階甄試"] == "✅"]) if "二階甄試" in yd.columns else 0
            yf = len(yd[yd["最終入學"] == "✅"]) if "最終入學" in yd.columns else 0
            yearly_stats.append({
                "學年度": yr, "一階報名": y1, "二階甄試": y2, "最終入學": yf,
                "一→二%": y2 / y1 * 100 if y1 > 0 else 0,
                "二→入學%": yf / y2 * 100 if y2 > 0 else 0,
                "總轉換%": yf / y1 * 100 if y1 > 0 else 0,
            })
        ys_df = pd.DataFrame(yearly_stats)

        display_ys = ys_df.copy()
        display_ys["一→二%"] = display_ys["一→二%"].apply(format_pct)
        display_ys["二→入學%"] = display_ys["二→入學%"].apply(format_pct)
        display_ys["總轉換%"] = display_ys["總轉換%"].apply(format_pct)
        st.dataframe(display_ys, use_container_width=True, hide_index=True)

        col_y1, col_y2 = st.columns(2)
        with col_y1:
            fig_trend = go.Figure()
            for col_name, color in [("一階報名", "#FDD7B4"), ("二階甄試", "#E8792F"), ("最終入學", "#B84500")]:
                fig_trend.add_trace(go.Bar(
                    name=col_name, x=ys_df["學年度"], y=ys_df[col_name],
                    marker_color=color, text=ys_df[col_name], textposition="outside"
                ))
            fig_trend.update_layout(barmode="group", title="歷年各階段人數", height=400)
            st.plotly_chart(fig_trend, use_container_width=True)

        with col_y2:
            fig_rate_y = go.Figure()
            fig_rate_y.add_trace(go.Scatter(
                x=ys_df["學年度"], y=ys_df["一→二%"].apply(lambda s: float(s) if isinstance(s, (int, float)) else float(str(s).replace('%',''))),
                mode="lines+markers+text", name="一→二%",
                line=dict(color="#E8792F", width=2, dash="dot"),
                text=ys_df["一→二%"].apply(lambda v: "{:.1f}%".format(v) if isinstance(v, (int, float)) else v),
                textposition="top center"
            ))
            # 用原始數值重繪
            fig_rate_y2 = go.Figure()
            raw_ys = pd.DataFrame(yearly_stats)
            fig_rate_y2.add_trace(go.Scatter(
                x=raw_ys["學年度"], y=raw_ys["一→二%"],
                mode="lines+markers+text", name="一→二%",
                line=dict(color="#E8792F", width=2),
                text=raw_ys["一→二%"].apply(lambda v: "{:.1f}%".format(v)),
                textposition="top left"
            ))
            fig_rate_y2.add_trace(go.Scatter(
                x=raw_ys["學年度"], y=raw_ys["二→入學%"],
                mode="lines+markers+text", name="二→入學%",
                line=dict(color="#B84500", width=2),
                text=raw_ys["二→入學%"].apply(lambda v: "{:.1f}%".format(v)),
                textposition="top right"
            ))
            fig_rate_y2.add_trace(go.Scatter(
                x=raw_ys["學年度"], y=raw_ys["總轉換%"],
                mode="lines+markers+text", name="總轉換%",
                line=dict(color="#333", width=3),
                text=raw_ys["總轉換%"].apply(lambda v: "{:.1f}%".format(v)),
                textposition="bottom center"
            ))
            fig_rate_y2.update_layout(title="歷年轉換率趨勢（%）", yaxis_title="%", height=400)
            st.plotly_chart(fig_rate_y2, use_container_width=True)

        st.subheader("📊 年度增減分析")
        col_inc1, col_inc2 = st.columns([1, 2])
        with col_inc1:
            for i in range(1, len(ys_df)):
                prev_n = yearly_stats[i - 1]["一階報名"]
                curr_n = yearly_stats[i]["一階報名"]
                diff = curr_n - prev_n
                pct = diff / prev_n * 100 if prev_n > 0 else 0
                st.metric(
                    "{} 學年".format(yearly_stats[i]["學年度"]),
                    "{:,} 人".format(curr_n),
                    "{:+d}（{:+.1f}%）".format(diff, pct)
                )

        with col_inc2:
            if "報考科系" in data.columns and len(years_sorted) >= 2:
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
                st.markdown("**{} → {} 各科系增減**".format(year_a, year_b))
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
    col_dl1, col_dl2, col_dl3 = st.columns(3)
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
            export_data.to_excel(writer, index=False, sheet_name="合併資料")
            if "報考科系" in data.columns:
                dept_c = compute_dept_conversion(export_data)
                if not dept_c.empty:
                    dept_c.to_excel(writer, index=False, sheet_name="科系轉換率")
            if "畢業學校" in data.columns:
                sch_c = compute_school_conversion(export_data)
                if not sch_c.empty:
                    sch_c.to_excel(writer, index=False, sheet_name="學校轉換率")
        st.download_button(
            "⬇️ 下載 Excel（含轉換率）", excel_buf.getvalue(),
            "招生分析資料_含轉換率.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_dl3:
        if "畢業學校" in data.columns:
            sch_csv = io.BytesIO()
            sch_full = compute_school_conversion(export_data)
            if not sch_full.empty:
                sch_full.to_csv(sch_csv, index=False, encoding="utf-8-sig")
                st.download_button(
                    "⬇️ 下載學校轉換率 CSV", sch_csv.getvalue(),
                    "學校三階段轉換率.csv", "text/csv", use_container_width=True
                )

    st.subheader("📋 資料品質報告")
    quality = []
    for col_q in data.columns:
        col_series = safe_get_series(data, col_q)
        non_null = int(col_series.notna().sum())
        total_rows = len(data)
        rate_q = "{:.1f}%".format(non_null / total_rows * 100) if total_rows > 0 else "0%"
        quality.append({
            "欄位": col_q,
            "非空筆數": non_null,
            "總筆數": total_rows,
            "完整率": rate_q,
            "空值數": total_rows - non_null,
        })
    st.dataframe(pd.DataFrame(quality), use_container_width=True, hide_index=True)
