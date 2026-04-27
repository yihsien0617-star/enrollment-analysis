import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import folium
from streamlit_folium import st_folium
import hashlib
import io
from collections import defaultdict

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
# 初始化 Session State
# ============================================================
if "main_data" not in st.session_state:
    st.session_state.main_data = pd.DataFrame()
if "uploaded_hashes" not in st.session_state:
    st.session_state.uploaded_hashes = {}
if "upload_log" not in st.session_state:
    st.session_state.upload_log = []

# ============================================================
# 工具函式
# ============================================================
def compute_file_hash(file_bytes, filename):
    """用檔案內容 + 檔名產生唯一 hash"""
    content = filename.encode() + file_bytes
    return hashlib.md5(content).hexdigest()

def clean_column_names(df):
    """清理並標準化欄位名稱"""
    col_mapping = {}
    for col in df.columns:
        c = col.strip().replace(" ", "")
        if "學年" in c:
            col_mapping[col] = "學年度"
        elif "科系" in c or "報考" in c or "系所" in c or "系別" in c:
            col_mapping[col] = "報考科系"
        elif "姓名" in c:
            col_mapping[col] = "姓名"
        elif "畢業" in c or "來源" in c or "學校" in c or "高中" in c:
            col_mapping[col] = "畢業學校"
        elif "經緯" in c or "座標" in c or "坐標" in c:
            col_mapping[col] = "經緯度"
        elif "身分證" in c or "身份證" in c or "ID" in c.upper():
            col_mapping[col] = "身分證字號"
        elif "緯度" in c or "lat" in c.lower():
            col_mapping[col] = "緯度"
        elif "經度" in c or "lon" in c.lower() or "lng" in c.lower():
            col_mapping[col] = "經度"
    df = df.rename(columns=col_mapping)
    return df

def parse_coordinates(df):
    """解析經緯度欄位，支援多種格式"""
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
            # 格式: "25.033, 121.565" 或 "(25.033, 121.565)"
            val_str = val_str.replace("(", "").replace(")", "")
            parts = val_str.replace("，", ",").split(",")
            if len(parts) == 2:
                lat = float(parts[0].strip())
                lon = float(parts[1].strip())
                # 台灣緯度約 21.9~25.3, 經度約 120~122
                if 21 < lat < 26 and 119 < lon < 123:
                    lats.append(lat)
                    lons.append(lon)
                elif 21 < lon < 26 and 119 < lat < 123:
                    # 經緯度反了
                    lats.append(lon)
                    lons.append(lat)
                else:
                    lats.append(None)
                    lons.append(None)
            else:
                lats.append(None)
                lons.append(None)
        except:
            lats.append(None)
            lons.append(None)

    df["緯度"] = lats
    df["經度"] = lons
    return df

def merge_data(existing_df, new_df):
    """合併資料，以身分證字號 + 學年度為唯一鍵"""
    if existing_df.empty:
        return new_df
    combined = pd.concat([existing_df, new_df], ignore_index=True)
    if "身分證字號" in combined.columns and "學年度" in combined.columns:
        combined = combined.drop_duplicates(
            subset=["身分證字號", "學年度"],
            keep="last"
        )
    else:
        combined = combined.drop_duplicates(keep="last")
    return combined.reset_index(drop=True)

def generate_sample_data():
    """產生示範資料"""
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
        ("雲林縣立斗六高中", 23.707, 120.543),
        ("高雄市立鳳山高中", 22.626, 120.357),
        ("台南市私立崑山高中", 22.950, 120.260),
        ("嘉義縣立東石高中", 23.460, 120.155),
        ("澎湖縣立馬公高中", 23.566, 119.577),
    ]

    departments = [
        "護理系", "醫學檢驗暨生物技術系", "職業安全衛生系",
        "食品營養系", "幼兒保育系", "長期照護系",
        "視光系", "藥學系", "運動健康與休閒系"
    ]

    surnames = "陳林黃張李王吳劉蔡楊許鄭謝洪曾"
    names = "雅婷志明家豪淑芬建宏美玲俊傑怡君宗翰佩珊"

    records = []
    id_counter = 0
    for year in [112, 113, 114]:
        n = random.randint(180, 250)
        for _ in range(n):
            school_name, base_lat, base_lon = random.choice(schools)
            dept = random.choice(departments)
            name = random.choice(surnames) + random.choice(names) + random.choice(names)
            lat = base_lat + random.uniform(-0.02, 0.02)
            lon = base_lon + random.uniform(-0.02, 0.02)
            id_counter += 1
            fake_id = f"A{random.randint(100000000, 299999999)}"
            records.append({
                "學年度": year,
                "報考科系": dept,
                "姓名": name,
                "畢業學校": school_name,
                "經緯度": f"{lat:.4f}, {lon:.4f}",
                "身分證字號": fake_id
            })

    return pd.DataFrame(records)

# ============================================================
# 側邊欄 - 資料管理
# ============================================================
with st.sidebar:
    st.image("https://via.placeholder.com/280x60/E8792F/FFFFFF?text=HWU+招生分析系統", use_container_width=True)
    st.markdown("## 📁 資料匯入")

    uploaded_files = st.file_uploader(
        "上傳 Excel 檔案（可多檔）",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="file_uploader"
    )

    if uploaded_files:
        for uf in uploaded_files:
            file_bytes = uf.read()
            uf.seek(0)
            fhash = compute_file_hash(file_bytes, uf.name)

            if fhash not in st.session_state.uploaded_hashes:
                try:
                    new_df = pd.read_excel(uf)
                    new_df = clean_column_names(new_df)
                    new_df = parse_coordinates(new_df)
                    row_count = len(new_df)

                    st.session_state.main_data = merge_data(
                        st.session_state.main_data, new_df
                    )

                    year_info = ""
                    if "學年度" in new_df.columns:
                        years = new_df["學年度"].dropna().unique()
                        year_info = f"（{', '.join(map(str, sorted(years)))}）"

                    st.session_state.uploaded_hashes[fhash] = uf.name
                    st.session_state.upload_log.append(
                        f"✅ {uf.name}：{row_count} 筆 {year_info}"
                    )
                except Exception as e:
                    st.session_state.upload_log.append(
                        f"❌ {uf.name}：{str(e)}"
                    )

    # 示範資料
    st.markdown("---")
    if st.button("📊 載入示範資料", use_container_width=True):
        sample_df = generate_sample_data()
        sample_df = parse_coordinates(sample_df)
        st.session_state.main_data = merge_data(
            st.session_state.main_data, sample_df
        )
        st.session_state.upload_log.append("✅ 示範資料已載入")
        st.rerun()

    # 匯入紀錄
    if st.session_state.upload_log:
        st.markdown("### 📋 匯入紀錄")
        for log in st.session_state.upload_log[-10:]:
            st.caption(log)

    # 資料狀態
    st.markdown("---")
    data = st.session_state.main_data
    if not data.empty:
        st.success(f"📊 目前共 **{len(data):,}** 筆資料")
        if "學年度" in data.columns:
            years = sorted(data["學年度"].dropna().unique())
            st.info(f"📅 涵蓋學年度：{', '.join(map(str, years))}")
        if "報考科系" in data.columns:
            st.info(f"🏫 科系數：{data['報考科系'].nunique()}")
    else:
        st.warning("尚未匯入資料")

    # 清除按鈕
    st.markdown("---")
    if st.button("🗑️ 清除所有資料", use_container_width=True, type="secondary"):
        st.session_state.main_data = pd.DataFrame()
        st.session_state.uploaded_hashes = {}
        st.session_state.upload_log = []
        st.rerun()

# ============================================================
# 主畫面
# ============================================================
st.title("🎓 中華醫事科技大學 招生數據分析系統")

data = st.session_state.main_data

if data.empty:
    st.markdown("""
    <div style='text-align:center; padding:60px 20px;'>
        <h2>👈 請先從左側匯入資料</h2>
        <p style='color:gray; font-size:18px;'>
            支援 Excel 格式，需包含欄位：學年度、報考科系、姓名、畢業學校、經緯度、身分證字號
        </p>
        <p style='color:gray;'>或點擊「載入示範資料」快速體驗</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ============================================================
# 分頁
# ============================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🗺️ 地圖視覺化",
    "📊 科系招生分析",
    "🏫 來源學校分析",
    "📈 跨年度比較",
    "🔍 資料檢視與匯出"
])

# ============================================================
# TAB 1: 地圖視覺化
# ============================================================
with tab1:
    st.header("🗺️ 報考生分布地圖")

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        if "學年度" in data.columns:
            years = ["全部"] + sorted(data["學年度"].dropna().unique().tolist())
            sel_year_map = st.selectbox("選擇學年度", years, key="map_year")
        else:
            sel_year_map = "全部"
    with col_f2:
        if "報考科系" in data.columns:
            depts = ["全部"] + sorted(data["報考科系"].dropna().unique().tolist())
            sel_dept_map = st.selectbox("選擇科系", depts, key="map_dept")
        else:
            sel_dept_map = "全部"

    map_data = data.copy()
    if sel_year_map != "全部" and "學年度" in map_data.columns:
        map_data = map_data[map_data["學年度"] == sel_year_map]
    if sel_dept_map != "全部" and "報考科系" in map_data.columns:
        map_data = map_data[map_data["報考科系"] == sel_dept_map]

    if "緯度" in map_data.columns and "經度" in map_data.columns:
        valid_coords = map_data.dropna(subset=["緯度", "經度"])

        if len(valid_coords) > 0:
            col_m1, col_m2 = st.columns([3, 1])

            with col_m1:
                # 中華醫事科大位置
                hwu_lat, hwu_lon = 22.9908, 120.2133
                m = folium.Map(
                    location=[23.5, 120.5],
                    zoom_start=8,
                    tiles="CartoDB positron"
                )

                # 學校標記
                folium.Marker(
                    location=[hwu_lat, hwu_lon],
                    popup="中華醫事科技大學",
                    icon=folium.Icon(color="red", icon="star", prefix="fa"),
                ).add_to(m)

                # 考生標記（使用 MarkerCluster 的概念，但用 CircleMarker）
                for _, row in valid_coords.iterrows():
                    popup_text = ""
                    if "畢業學校" in row.index:
                        popup_text += f"學校：{row['畢業學校']}<br>"
                    if "報考科系" in row.index:
                        popup_text += f"科系：{row['報考科系']}<br>"
                    if "學年度" in row.index:
                        popup_text += f"學年：{row['學年度']}"

                    folium.CircleMarker(
                        location=[row["緯度"], row["經度"]],
                        radius=5,
                        color="#E8792F",
                        fill=True,
                        fill_color="#E8792F",
                        fill_opacity=0.6,
                        popup=folium.Popup(popup_text, max_width=200),
                    ).add_to(m)

                st_folium(m, width=None, height=550, use_container_width=True)

            with col_m2:
                st.metric("篩選結果", f"{len(map_data):,} 人")
                st.metric("有座標資料", f"{len(valid_coords):,} 人")
                st.metric("座標涵蓋率", f"{len(valid_coords)/len(map_data)*100:.1f}%")

                if "畢業學校" in valid_coords.columns:
                    st.markdown("**📍 主要來源地區**")
                    school_counts = valid_coords["畢業學校"].value_counts().head(8)
                    for school, cnt in school_counts.items():
                        st.caption(f"• {school}：{cnt} 人")
        else:
            st.warning("⚠️ 篩選後無有效座標資料")
    else:
        st.warning("⚠️ 無法解析經緯度資料，請確認格式（如：25.033, 121.565）")

# ============================================================
# TAB 2: 科系招生分析
# ============================================================
with tab2:
    st.header("📊 科系招生分析")

    if "報考科系" not in data.columns:
        st.warning("資料中未包含「報考科系」欄位")
    else:
        # 篩選
        if "學年度" in data.columns:
            years_t2 = ["全部"] + sorted(data["學年度"].dropna().unique().tolist())
            sel_year_t2 = st.selectbox("選擇學年度", years_t2, key="t2_year")
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
                dept_counts,
                x="報考人數",
                y="科系",
                orientation="h",
                title="各科系報考人數",
                color="報考人數",
                color_continuous_scale=["#FDD7B4", "#E8792F", "#B84500"],
            )
            fig_bar.update_layout(
                yaxis={"categoryorder": "total ascending"},
                height=450,
                showlegend=False,
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        with col_c2:
            fig_pie = px.pie(
                dept_counts,
                values="報考人數",
                names="科系",
                title="科系佔比分布",
                color_discrete_sequence=px.colors.sequential.Oranges_r,
            )
            fig_pie.update_layout(height=450)
            st.plotly_chart(fig_pie, use_container_width=True)

        # 跨年度科系比較（如有多學年）
        if "學年度" in data.columns and data["學年度"].nunique() > 1 and sel_year_t2 == "全部":
            st.subheader("📈 各科系歷年報考趨勢")
            yearly_dept = data.groupby(["學年度", "報考科系"]).size().reset_index(name="人數")
            fig_line = px.line(
                yearly_dept,
                x="學年度",
                y="人數",
                color="報考科系",
                markers=True,
                title="各科系歷年報考人數趨勢",
            )
            fig_line.update_layout(height=450)
            st.plotly_chart(fig_line, use_container_width=True)

        # KPI 卡片
        st.subheader("📋 科系報考統計摘要")
        kpi_cols = st.columns(4)
        with kpi_cols[0]:
            st.metric("總報考人數", f"{len(t2_data):,}")
        with kpi_cols[1]:
            st.metric("科系數", f"{t2_data['報考科系'].nunique()}")
        with kpi_cols[2]:
            top_dept = dept_counts.iloc[0]["科系"] if len(dept_counts) > 0 else "-"
            st.metric("最熱門科系", top_dept)
        with kpi_cols[3]:
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
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            if "學年度" in data.columns:
                years_t3 = ["全部"] + sorted(data["學年度"].dropna().unique().tolist())
                sel_year_t3 = st.selectbox("選擇學年度", years_t3, key="t3_year")
            else:
                sel_year_t3 = "全部"
        with col_f2:
            top_n = st.slider("顯示前 N 所學校", 5, 30, 15, key="t3_topn")

        t3_data = data.copy()
        if sel_year_t3 != "全部" and "學年度" in t3_data.columns:
            t3_data = t3_data[t3_data["學年度"] == sel_year_t3]

        school_counts = t3_data["畢業學校"].value_counts().head(top_n).reset_index()
        school_counts.columns = ["畢業學校", "報考人數"]

        # 長條圖
        fig_school = px.bar(
            school_counts,
            x="報考人數",
            y="畢業學校",
            orientation="h",
            title=f"報考人數 Top {top_n} 來源學校",
            color="報考人數",
            color_continuous_scale=["#FDD7B4", "#E8792F", "#B84500"],
        )
        fig_school.update_layout(
            yaxis={"categoryorder": "total ascending"},
            height=max(400, top_n * 30),
            showlegend=False,
        )
        st.plotly_chart(fig_school, use_container_width=True)

        # 來源學校 × 科系 交叉分析
        if "報考科系" in t3_data.columns:
            st.subheader("🔀 來源學校 × 科系 交叉分析")
            top_schools = school_counts["畢業學校"].tolist()
            cross_data = t3_data[t3_data["畢業學校"].isin(top_schools)]
            cross_table = pd.crosstab(cross_data["畢業學校"], cross_data["報考科系"])

            fig_heat = px.imshow(
                cross_table,
                title="來源學校 vs 報考科系（人數）",
                color_continuous_scale="Oranges",
                aspect="auto",
            )
            fig_heat.update_layout(height=max(400, top_n * 28))
            st.plotly_chart(fig_heat, use_container_width=True)

        # 學校管理建議
        st.subheader("⭐ 來源學校管理建議")
        all_school_counts = t3_data["畢業學校"].value_counts()
        total = len(t3_data)

        recs = []
        cumulative = 0
        for school, cnt in all_school_counts.items():
            cumulative += cnt
            ratio = cnt / total * 100
            cum_ratio = cumulative / total * 100
            if cum_ratio <= 50:
                level = "⭐⭐⭐ 重點經營"
            elif cum_ratio <= 80:
                level = "⭐⭐ 持續關注"
            else:
                level = "⭐ 一般維護"
            recs.append({
                "學校": school,
                "報考人數": cnt,
                "佔比(%)": round(ratio, 1),
                "累積佔比(%)": round(cum_ratio, 1),
                "建議等級": level
            })

        rec_df = pd.DataFrame(recs)
        st.dataframe(
            rec_df.head(top_n),
            use_container_width=True,
            hide_index=True,
        )

# ============================================================
# TAB 4: 跨年度比較
# ============================================================
with tab4:
    st.header("📈 跨年度比較分析")

    if "學年度" not in data.columns or data["學年度"].nunique() < 2:
        st.info("💡 需要至少兩個學年度的資料才能進行跨年比較。目前僅有一個學年度。")
    else:
        years_sorted = sorted(data["學年度"].dropna().unique())

        # 總報考人數趨勢
        st.subheader("📊 總報考人數趨勢")
        yearly_total = data.groupby("學年度").size().reset_index(name="報考人數")

        col_t1, col_t2 = st.columns([2, 1])
        with col_t1:
            fig_trend = px.bar(
                yearly_total,
                x="學年度",
                y="報考人數",
                title="歷年總報考人數",
                text="報考人數",
                color_discrete_sequence=["#E8792F"],
            )
            fig_trend.update_traces(textposition="outside")
            fig_trend.update_layout(height=400)
            st.plotly_chart(fig_trend, use_container_width=True)

        with col_t2:
            st.markdown("**年度增減分析**")
            for i in range(1, len(yearly_total)):
                prev = yearly_total.iloc[i - 1]
                curr = yearly_total.iloc[i]
                diff = curr["報考人數"] - prev["報考人數"]
                pct = diff / prev["報考人數"] * 100 if prev["報考人數"] > 0 else 0
                emoji = "📈" if diff > 0 else "📉" if diff < 0 else "➡️"
                st.metric(
                    f"{int(curr['學年度'])} 學年度",
                    f"{int(curr['報考人數']):,} 人",
                    f"{diff:+d} 人（{pct:+.1f}%）"
                )

        # 科系增減
        if "報考科系" in data.columns and len(years_sorted) >= 2:
            st.subheader("📊 科系增減比較")
            col_y1, col_y2 = st.columns(2)
            with col_y1:
                year_a = st.selectbox("基準年", years_sorted[:-1], key="cmp_ya")
            with col_y2:
                later_years = [y for y in years_sorted if y > year_a]
                year_b = st.selectbox("比較年", later_years, key="cmp_yb")

            da = data[data["學年度"] == year_a]["報考科系"].value_counts()
            db = data[data["學年度"] == year_b]["報考科系"].value_counts()

            all_depts = sorted(set(da.index) | set(db.index))
            comparison = []
            for dept in all_depts:
                a_val = da.get(dept, 0)
                b_val = db.get(dept, 0)
                diff = b_val - a_val
                pct = (diff / a_val * 100) if a_val > 0 else (100 if b_val > 0 else 0)
                comparison.append({
                    "科系": dept,
                    f"{int(year_a)}學年": a_val,
                    f"{int(year_b)}學年": b_val,
                    "增減": diff,
                    "增減率(%)": round(pct, 1),
                })

            cmp_df = pd.DataFrame(comparison).sort_values("增減", ascending=False)

            fig_cmp = go.Figure()
            fig_cmp.add_trace(go.Bar(
                name=f"{int(year_a)}學年",
                x=cmp_df["科系"],
                y=cmp_df[f"{int(year_a)}學年"],
                marker_color="#FDD7B4"
            ))
            fig_cmp.add_trace(go.Bar(
                name=f"{int(year_b)}學年",
                x=cmp_df["科系"],
                y=cmp_df[f"{int(year_b)}學年"],
                marker_color="#E8792F"
            ))
            fig_cmp.update_layout(
                barmode="group",
                title=f"科系人數比較：{int(year_a)} vs {int(year_b)}",
                height=450,
            )
            st.plotly_chart(fig_cmp, use_container_width=True)

            st.dataframe(cmp_df, use_container_width=True, hide_index=True)

        # 來源學校變化
        if "畢業學校" in data.columns and len(years_sorted) >= 2:
            st.subheader("🏫 來源學校年度變化")

            school_year = data.groupby(["學年度", "畢業學校"]).size().reset_index(name="人數")
            top_schools_overall = data["畢業學校"].value_counts().head(10).index.tolist()
            school_year_top = school_year[school_year["畢業學校"].isin(top_schools_overall)]

            fig_school_trend = px.line(
                school_year_top,
                x="學年度",
                y="人數",
                color="畢業學校",
                markers=True,
                title="Top 10 來源學校歷年趨勢",
            )
            fig_school_trend.update_layout(height=450)
            st.plotly_chart(fig_school_trend, use_container_width=True)

# ============================================================
# TAB 5: 資料檢視與匯出
# ============================================================
with tab5:
    st.header("🔍 資料檢視與匯出")

    # 篩選
    col_e1, col_e2, col_e3 = st.columns(3)
    export_data = data.copy()

    with col_e1:
        if "學年度" in data.columns:
            yr_options = ["全部"] + sorted(data["學年度"].dropna().unique().tolist())
            sel_yr_exp = st.selectbox("學年度", yr_options, key="exp_yr")
            if sel_yr_exp != "全部":
                export_data = export_data[export_data["學年度"] == sel_yr_exp]

    with col_e2:
        if "報考科系" in data.columns:
            dept_options = ["全部"] + sorted(data["報考科系"].dropna().unique().tolist())
            sel_dept_exp = st.selectbox("科系", dept_options, key="exp_dept")
            if sel_dept_exp != "全部":
                export_data = export_data[export_data["報考科系"] == sel_dept_exp]

    with col_e3:
        if "畢業學校" in data.columns:
            search_school = st.text_input("搜尋畢業學校", key="exp_school")
            if search_school:
                export_data = export_data[
                    export_data["畢業學校"].str.contains(search_school, na=False)
                ]

    st.caption(f"篩選結果：{len(export_data):,} 筆")

    # 顯示欄位（隱藏身分證字號）
    display_cols = [c for c in export_data.columns if c != "身分證字號"]
    st.dataframe(
        export_data[display_cols],
        use_container_width=True,
        hide_index=True,
        height=400,
    )

    # 匯出
    st.subheader("📥 匯出資料")
    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        csv_buffer = io.BytesIO()
        export_data.to_csv(csv_buffer, index=False, encoding="utf-8-sig")
        st.download_button(
            label="⬇️ 下載 CSV",
            data=csv_buffer.getvalue(),
            file_name="招生分析資料.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with col_dl2:
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            export_data.to_excel(writer, index=False, sheet_name="招生資料")
        st.download_button(
            label="⬇️ 下載 Excel",
            data=excel_buffer.getvalue(),
            file_name="招生分析資料.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # 資料品質報告
    st.subheader("📋 資料品質報告")
    quality = []
    for col in data.columns:
        non_null = data[col].notna().sum()
        total = len(data)
        quality.append({
            "欄位": col,
            "非空筆數": non_null,
            "總筆數": total,
            "完整率": f"{non_null/total*100:.1f}%",
            "空值數": total - non_null,
        })
    st.dataframe(pd.DataFrame(quality), use_container_width=True, hide_index=True)
