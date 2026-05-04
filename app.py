# -*- coding: utf-8 -*-
"""
中華醫事科技大學 招生數據分析系統 v6.0
- 多年度標籤式介面
- 經緯度從一階資料讀取，所有階段共用
- 最終入學管道自動讀取
- 跨年度比較分析（Module 7）
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re, io, copy

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
st.markdown('<div class="sub-header">Enrollment Analytics v6.0 ｜ 多年度標籤式分析 ｜ 經緯度從一階讀取 ｜ 跨年度比較</div>', unsafe_allow_html=True)
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

# ============================================================
# 常數
# ============================================================
KNOWN_CHANNELS = [
    "聯合免試","甄選入學","技優甄審","運動績優","身障甄試",
    "單獨招生","進修部","產學攜手","運動單獨招生","四技申請","繁星計畫",
]
CHANNEL_KEYWORDS = {
    "聯合免試":["免試","聯合免試","聯免"],
    "甄選入學":["甄選","甄選入學"],
    "技優甄審":["技優","技優甄審"],
    "運動績優":["運動績優","運績"],
    "身障甄試":["身障","身障甄試"],
    "單獨招生":["單獨招生","單招","單獨"],
    "進修部":["進修部","進修","夜間"],
    "產學攜手":["產攜","產學攜手","產學"],
    "四技申請":["四技申請","申請入學"],
    "繁星計畫":["繁星"],
}
FINAL_CH_CANDIDATES = [
    "入學方式","入學管道","錄取管道","招生管道","管道","入學途徑","錄取方式","報名管道"
]
HWU = {"lat":22.9340,"lon":120.2756}

# ============================================================
# 工具函式
# ============================================================
def detect_school_col(df):
    for kw in ["畢業學校","來源學校","原就讀學校","高中職校名","學校名稱","校名","畢業高中職","畢業學校名稱","就讀學校","原學校"]:
        for c in df.columns:
            if kw in str(c): return c
    return None

def detect_dept_col(df):
    for kw in ["科系","系所","報名科系","錄取科系","志願科系","就讀科系","科系名稱","系科","學系","錄取系所","就讀系所","註冊科系"]:
        for c in df.columns:
            if kw in str(c): return c
    return None

def detect_final_ch_col(df):
    for kw in FINAL_CH_CANDIDATES:
        for c in df.columns:
            if kw in str(c): return c
    return None

def detect_channel_from_filename(fn):
    if not fn: return None
    for ch,kws in CHANNEL_KEYWORDS.items():
        for kw in kws:
            if kw in fn: return ch
    return None

def detect_channel_from_cols(df):
    for cn in ["管道","入學管道","招生管道","報名管道"]:
        if cn in df.columns:
            top=df[cn].dropna().value_counts()
            if not top.empty:
                v=str(top.index[0])
                for ch,kws in CHANNEL_KEYWORDS.items():
                    for kw in kws:
                        if kw in v: return ch
                return v
    return None

def norm_school(name):
    if not isinstance(name,str): return str(name)
    name=name.strip().replace("臺","台").replace("（","(").replace("）",")")
    name=re.sub(r"\s+","",name)
    for sfx in ["附設進修學校","進修學校","進修部"]: name=name.replace(sfx,"")
    return name

def detect_lat_lon_cols(df):
    lat_col=lon_col=None
    for c in df.columns:
        s=str(c).strip().lower()
        if lat_col is None and any(k in s for k in ["緯度","lat","latitude"]): lat_col=c
        if lon_col is None and any(k in s for k in ["經度","lon","lng","longitude"]): lon_col=c
    return lat_col,lon_col

def build_geo_from_p1(p1):
    sc=detect_school_col(p1)
    lat_c,lon_c=detect_lat_lon_cols(p1)
    if sc is None or lat_c is None or lon_c is None:
        return None
    g=p1[[sc,lat_c,lon_c]].copy()
    g.columns=["學校_raw","lat","lon"]
    g["lat"]=pd.to_numeric(g["lat"],errors="coerce")
    g["lon"]=pd.to_numeric(g["lon"],errors="coerce")
    g=g.dropna(subset=["lat","lon"])
    g["_std"]=g["學校_raw"].apply(norm_school)
    geo=g.groupby("_std").agg(lat=("lat","mean"),lon=("lon","mean")).reset_index()
    return geo

def enrich_geo(df,geo):
    sc=detect_school_col(df)
    if sc is None or geo is None or geo.empty: return df
    df=df.copy()
    df["_std"]=df[sc].apply(norm_school)
    for c in ["緯度","經度","lat","lon","latitude","longitude"]:
        if c in df.columns and c!=sc: df.drop(columns=[c],inplace=True,errors="ignore")
    df=df.merge(geo,on="_std",how="left").drop(columns=["_std"],errors="ignore")
    return df

def eff_stars(r):
    if r>=70: return "⭐⭐⭐"
    elif r>=40: return "⭐⭐"
    else: return "⭐"

# ============================================================
# 統計構建
# ============================================================
def build_school_stats(p1,p2=None,p3=None):
    sc1=detect_school_col(p1)
    if sc1 is None: return None
    s=p1[sc1].value_counts().reset_index(); s.columns=["學校","一階人數"]
    if p2 is not None:
        sc2=detect_school_col(p2)
        if sc2:
            t2=p2[sc2].value_counts().reset_index(); t2.columns=["學校","二階人數"]
            s=s.merge(t2,on="學校",how="left")
    if "二階人數" not in s.columns: s["二階人數"]=np.nan
    if p3 is not None:
        sc3=detect_school_col(p3)
        if sc3:
            t3=p3[sc3].value_counts().reset_index(); t3.columns=["學校","最終入學"]
            s=s.merge(t3,on="學校",how="left")
    if "最終入學" not in s.columns: s["最終入學"]=np.nan
    s["二階人數"]=s["二階人數"].fillna(0).astype(int)
    s["最終入學"]=s["最終入學"].fillna(0).astype(int)
    s["一→二階(%)"]=( s["二階人數"]/s["一階人數"]*100).round(1)
    s["一→最終(%)"]=( s["最終入學"]/s["一階人數"]*100).round(1)
    s["流失人數"]=s["一階人數"]-s["最終入學"]
    s["效率評等"]=s["一→最終(%)"].apply(eff_stars)
    return s

def build_dept_stats(p1,p2=None,p3=None):
    dc1=detect_dept_col(p1)
    if dc1 is None: return None
    s=p1[dc1].value_counts().reset_index(); s.columns=["科系","一階人數"]
    if p2 is not None:
        dc2=detect_dept_col(p2)
        if dc2:
            t2=p2[dc2].value_counts().reset_index(); t2.columns=["科系","二階人數"]
            s=s.merge(t2,on="科系",how="left")
    if "二階人數" not in s.columns: s["二階人數"]=np.nan
    if p3 is not None:
        dc3=detect_dept_col(p3)
        if dc3:
            t3=p3[dc3].value_counts().reset_index(); t3.columns=["科系","最終入學"]
            s=s.merge(t3,on="科系",how="left")
    if "最終入學" not in s.columns: s["最終入學"]=np.nan
    s["二階人數"]=s["二階人數"].fillna(0).astype(int)
    s["最終入學"]=s["最終入學"].fillna(0).astype(int)
    s["一→二階(%)"]=( s["二階人數"]/s["一階人數"]*100).round(1)
    s["一→最終(%)"]=( s["最終入學"]/s["一階人數"]*100).round(1)
    s["效率評等"]=s["一→最終(%)"].apply(eff_stars)
    return s

# ============================================================
# 視覺化
# ============================================================
def fig_funnel(labels,values,title="招生漏斗"):
    colors=["#2196F3","#FF9800","#4CAF50","#E91E63"]
    fig=go.Figure(go.Funnel(y=labels,x=values,textinfo="value+percent initial",
        marker=dict(color=colors[:len(labels)]),
        connector=dict(line=dict(color="royalblue",width=2))))
    fig.update_layout(title=title,height=420,font=dict(size=14))
    return fig

def fig_bar_h(df,y,x,title,color="#667eea"):
    fig=px.bar(df.sort_values(x,ascending=True),x=x,y=y,orientation="h",text=x,title=title)
    fig.update_traces(marker_color=color,texttemplate="%{text:.1f}%",textposition="outside")
    fig.update_layout(height=max(380,len(df)*28),xaxis_title="轉換率 (%)",yaxis_title="")
    return fig

def fig_grouped_bar(df,y,vals,title):
    fig=go.Figure()
    colors=["#2196F3","#FF9800","#4CAF50"]
    for i,v in enumerate(vals):
        if v in df.columns:
            fig.add_trace(go.Bar(name=v,y=df[y],x=df[v],orientation="h",
                marker_color=colors[i%3],text=df[v],textposition="outside"))
    fig.update_layout(barmode="group",title=title,height=max(400,len(df)*35),yaxis=dict(autorange="reversed"))
    return fig

def fig_map(df,size_col,title,color_col=None):
    if "lat" not in df.columns or "lon" not in df.columns: return None
    m=df.dropna(subset=["lat","lon"]).copy()
    if m.empty: return None
    m[size_col]=pd.to_numeric(m[size_col],errors="coerce").fillna(1)
    sc=detect_school_col(m)
    fig=px.scatter_mapbox(m,lat="lat",lon="lon",size=size_col,
        color=color_col if color_col and color_col in m.columns else None,
        hover_name=sc if sc and sc in m.columns else None,
        hover_data={size_col:True,"lat":":.4f","lon":":.4f"},
        title=title,size_max=30,zoom=7,
        center={"lat":HWU["lat"],"lon":HWU["lon"]},mapbox_style="carto-positron")
    fig.add_trace(go.Scattermapbox(lat=[HWU["lat"]],lon=[HWU["lon"]],mode="markers+text",
        marker=dict(size=18,color="red",symbol="star"),
        text=["中華醫事科技大學"],textposition="top center",name="本校",showlegend=True))
    fig.update_layout(height=600,margin=dict(l=0,r=0,t=40,b=0))
    return fig

def fig_heatmap(df,x,y,v,title):
    pv=df.pivot_table(index=y,columns=x,values=v,aggfunc="sum").fillna(0)
    fig=px.imshow(pv,text_auto=True,aspect="auto",color_continuous_scale="YlOrRd",title=title)
    fig.update_layout(height=max(400,len(pv)*25))
    return fig

# ============================================================
# Session State
# ============================================================
if "years" not in st.session_state:
    st.session_state["years"] = {}          # {year_label: {p1_name,p2_name,p3_name,files:{name:df},channel_col,selected_channels,...}}
if "all_files" not in st.session_state:
    st.session_state["all_files"] = {}      # {filename: df}
if "analysis_ready" not in st.session_state:
    st.session_state["analysis_ready"] = False
if "analysis_version" not in st.session_state:
    st.session_state["analysis_version"] = 0

# ============================================================
# Sidebar
# ============================================================
with st.sidebar:
    st.header("📂 資料管理")

    # ── 上傳全部檔案 ──
    uploaded=st.file_uploader("上傳所有招生資料 (Excel/CSV)",type=["xlsx","xls","csv"],
        accept_multiple_files=True,help="可一次上傳多個年度的檔案")
    if uploaded:
        for uf in uploaded:
            if uf.name not in st.session_state["all_files"]:
                try:
                    df=pd.read_csv(uf) if uf.name.endswith(".csv") else pd.read_excel(uf)
                    st.session_state["all_files"][uf.name]=df
                    st.success(f"✅ {uf.name}（{len(df)} 筆）")
                except Exception as e:
                    st.error(f"❌ {uf.name}: {e}")

    if st.session_state["all_files"]:
        st.caption(f"📁 已載入 {len(st.session_state['all_files'])} 個檔案")
        if st.button("🗑️ 清除全部檔案"):
            st.session_state["all_files"]={}
            st.session_state["years"]={}
            st.session_state["analysis_ready"]=False
            st.rerun()

    st.markdown("---")

    # ── 年度管理 ──
    st.header("📅 年度管理")
    st.markdown('<div style="font-size:.8rem;color:#888;margin-bottom:8px;">'
                '每個年度各自指定一階/二階/最終入學資料<br>'
                '經緯度從一階資料的緯度/經度欄位讀取</div>',unsafe_allow_html=True)

    new_year=st.text_input("新增年度標籤：",placeholder="例如：113學年",key="new_year_input")
    if st.button("➕ 新增年度") and new_year:
        new_year=new_year.strip()
        if new_year and new_year not in st.session_state["years"]:
            st.session_state["years"][new_year]={
                "p1":None,"p2":None,"p3":None,
                "channel_col":None,"selected_channels":None,
            }
            st.success(f"✅ 已新增「{new_year}」")
            st.rerun()
        elif new_year in st.session_state["years"]:
            st.warning("⚠️ 此年度已存在")

    file_opts=["-- 未選擇 --"]+list(st.session_state["all_files"].keys())

    for yr in list(st.session_state["years"].keys()):
        ydata=st.session_state["years"][yr]
        with st.expander(f"📅 {yr}",expanded=True):
            ydata["p1"]=st.selectbox(f"🔵 一階（含經緯度）",file_opts,key=f"p1_{yr}",
                index=file_opts.index(ydata["p1"]) if ydata["p1"] in file_opts else 0)
            ydata["p2"]=st.selectbox(f"🟠 二階",file_opts,key=f"p2_{yr}",
                index=file_opts.index(ydata["p2"]) if ydata["p2"] in file_opts else 0)
            ydata["p3"]=st.selectbox(f"🟢 最終入學",file_opts,key=f"p3_{yr}",
                index=file_opts.index(ydata["p3"]) if ydata["p3"] in file_opts else 0)

            # 偵測一階經緯度
            if ydata["p1"] and ydata["p1"]!="-- 未選擇 --":
                p1df=st.session_state["all_files"][ydata["p1"]]
                lat_c,lon_c=detect_lat_lon_cols(p1df)
                if lat_c and lon_c:
                    n_valid=p1df[[lat_c,lon_c]].dropna().shape[0]
                    st.caption(f"📍 經緯度：{lat_c}/{lon_c}（{n_valid}筆有效）")
                else:
                    st.caption("⚠️ 一階未偵測到經緯度欄位")

            # 最終入學管道
            if ydata["p3"] and ydata["p3"]!="-- 未選擇 --":
                p3df=st.session_state["all_files"][ydata["p3"]]
                ch_col=detect_final_ch_col(p3df)
                if ch_col:
                    st.caption(f"📌 入學方式欄位：「{ch_col}」")
                    vals=p3df[ch_col].fillna("(空白)").astype(str).str.strip().replace("","(空白)")
                    ch_dist=vals.value_counts()
                    all_chs=ch_dist.index.tolist()
                    # 顯示分布
                    for cn,cnt in ch_dist.head(8).items():
                        st.markdown(f'<span class="channel-tag">{cn}</span> {cnt}人',unsafe_allow_html=True)
                    if len(ch_dist)>8:
                        st.caption(f"... 共 {len(ch_dist)} 種管道")
                    sel_chs=st.multiselect("納入分析的管道：",all_chs,default=all_chs,key=f"chs_{yr}")
                    ydata["channel_col"]=ch_col
                    ydata["selected_channels"]=sel_chs
                else:
                    st.caption("⚠️ 未偵測到入學方式欄位")
                    manual=st.selectbox("手動選擇：",["-- 無 --"]+list(p3df.columns),key=f"mch_{yr}")
                    if manual!="-- 無 --":
                        ydata["channel_col"]=manual
                        vals=p3df[manual].fillna("(空白)").value_counts()
                        all_chs=vals.index.tolist()
                        sel_chs=st.multiselect("納入管道：",all_chs,default=all_chs,key=f"mchs_{yr}")
                        ydata["selected_channels"]=sel_chs

            if st.button(f"🗑️ 刪除 {yr}",key=f"del_{yr}"):
                del st.session_state["years"][yr]
                st.rerun()

    # ── 更新分析 ──
    st.markdown("---")
    if st.button("🔄 更新分析",type="primary",use_container_width=True):
        st.session_state["analysis_ready"]=True
        st.session_state["analysis_version"]+=1
        st.success(f"✅ 分析已更新！版本 #{st.session_state['analysis_version']}")

    if st.session_state["analysis_ready"]:
        n_years=len([y for y in st.session_state["years"] if st.session_state["years"][y].get("p1") not in [None,"-- 未選擇 --"]])
        st.markdown(f'<div class="success-box">✅ 分析就緒　{n_years} 個年度<br>版本 #{st.session_state["analysis_version"]}</div>',unsafe_allow_html=True)


# ============================================================
# 取得某年度三階段資料
# ============================================================
def get_year_dfs(yr):
    ydata=st.session_state["years"].get(yr,{})
    p1=p2=p3=None
    geo=None
    s1=ydata.get("p1")
    s2=ydata.get("p2")
    s3=ydata.get("p3")
    if s1 and s1!="-- 未選擇 --" and s1 in st.session_state["all_files"]:
        p1=st.session_state["all_files"][s1].copy()
        geo=build_geo_from_p1(p1)
    if s2 and s2!="-- 未選擇 --" and s2 in st.session_state["all_files"]:
        p2=st.session_state["all_files"][s2].copy()
        if geo is not None: p2=enrich_geo(p2,geo)
    if s3 and s3!="-- 未選擇 --" and s3 in st.session_state["all_files"]:
        p3=st.session_state["all_files"][s3].copy()
        ch_col=ydata.get("channel_col")
        sel_chs=ydata.get("selected_channels")
        if ch_col and ch_col in p3.columns and sel_chs:
            p3[ch_col]=p3[ch_col].fillna("(空白)").astype(str).str.strip()
            p3.loc[p3[ch_col]=="",ch_col]="(空白)"
            p3=p3[p3[ch_col].isin(sel_chs)]
        if geo is not None: p3=enrich_geo(p3,geo)
    return p1,p2,p3,geo,ydata.get("channel_col")

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
    <p>📍 經緯度將從一階資料的「緯度/經度」欄位讀取</p>
    </div>""",unsafe_allow_html=True)
    st.stop()

valid_years=[yr for yr in st.session_state["years"]
             if st.session_state["years"][yr].get("p1") not in [None,"-- 未選擇 --"]]
if not valid_years:
    st.warning("⚠️ 沒有任何年度設定了一階資料。"); st.stop()


# ============================================================
# 年度單獨分析模組（函式化）
# ============================================================
def render_year_analysis(yr):
    """渲染單一年度的完整分析"""
    p1,p2,p3,geo,ch_col=get_year_dfs(yr)
    if p1 is None:
        st.warning(f"⚠️ {yr}：一階資料未指定或無法讀取。"); return

    # 模組選擇
    mod_opts=["📊 總覽儀表板","🔄 招生漏斗","📈 入學管道","🗺️ 地理分布",
              "🏫 科系熱力圖","🎯 來源學校","⚠️ 流失預警"]
    mod=st.radio("選擇分析模組：",mod_opts,horizontal=True,key=f"mod_{yr}")
    st.markdown('<hr class="section-divider">',unsafe_allow_html=True)

    n1=len(p1)
    n2=len(p2) if p2 is not None else None
    n3=len(p3) if p3 is not None else None

    # ── 總覽 ──
    if "總覽" in mod:
        st.subheader(f"📊 {yr} — 總覽儀表板")
        cols=st.columns(5)
        with cols[0]: st.markdown(f'<div class="metric-card"><h3>一階報名</h3><h1>{n1:,}</h1></div>',unsafe_allow_html=True)
        with cols[1]:
            v=f"{n2:,}" if n2 else "—"
            st.markdown(f'<div class="metric-card metric-orange"><h3>二階報到</h3><h1>{v}</h1></div>',unsafe_allow_html=True)
        with cols[2]:
            v=f"{n3:,}" if n3 else "—"
            st.markdown(f'<div class="metric-card metric-green"><h3>最終入學</h3><h1>{v}</h1></div>',unsafe_allow_html=True)
        with cols[3]:
            r=f"{n2/n1*100:.1f}%" if n2 and n1 else "—"
            st.markdown(f'<div class="metric-card metric-blue"><h3>一→二階</h3><h1>{r}</h1></div>',unsafe_allow_html=True)
        with cols[4]:
            r=f"{n3/n1*100:.1f}%" if n3 and n1 else "—"
            st.markdown(f'<div class="metric-card metric-gold"><h3>一→最終</h3><h1>{r}</h1></div>',unsafe_allow_html=True)

        # 經緯度狀態
        if geo is not None:
            st.caption(f"📍 經緯度資料庫（一階）：{len(geo)} 所學校")
        else:
            st.caption("⚠️ 一階無經緯度欄位，地圖功能不可用")

        # 管道分布
        if p3 is not None and ch_col and ch_col in p3.columns:
            st.markdown("---")
            st.subheader("🟢 最終入學管道分布")
            cd=p3[ch_col].value_counts().reset_index(); cd.columns=["入學管道","人數"]
            cd["佔比(%)"]=( cd["人數"]/cd["人數"].sum()*100).round(1)
            c1,c2=st.columns(2)
            with c1:
                fig=px.pie(cd,names="入學管道",values="人數",title="管道佔比",hole=.35)
                st.plotly_chart(fig,use_container_width=True)
            with c2:
                fig=px.bar(cd.sort_values("人數",ascending=True),x="人數",y="入學管道",
                    orientation="h",text="人數",title="各管道人數",color="佔比(%)",color_continuous_scale="Viridis")
                fig.update_layout(height=max(400,len(cd)*28),yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig,use_container_width=True)

        # 漏斗
        fl,fv=["一階報名"],[n1]
        if n2: fl.append("二階報到"); fv.append(n2)
        if n3: fl.append("最終入學"); fv.append(n3)
        if len(fv)>1:
            st.plotly_chart(fig_funnel(fl,fv,f"{yr} 招生漏斗"),use_container_width=True)

        # 科系 / 學校 TOP
        c1,c2=st.columns(2)
        dc=detect_dept_col(p1); sc=detect_school_col(p1)
        with c1:
            if dc:
                dd=p1[dc].value_counts().reset_index(); dd.columns=["科系","人數"]
                fig=px.pie(dd,names="科系",values="人數",title="一階科系分布",hole=.4)
                st.plotly_chart(fig,use_container_width=True)
        with c2:
            if sc:
                sd=p1[sc].value_counts().head(10).reset_index(); sd.columns=["學校","人數"]
                fig=px.bar(sd,x="人數",y="學校",orientation="h",title="來源學校 TOP 10",text="人數")
                fig.update_traces(marker_color="#667eea")
                fig.update_layout(yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig,use_container_width=True)

        ds=build_dept_stats(p1,p2,p3)
        if ds is not None:
            st.markdown("---"); st.subheader("各科系三階段概覽")
            st.dataframe(ds.sort_values("一階人數",ascending=False),use_container_width=True,hide_index=True)

    # ── 招生漏斗 ──
    elif "漏斗" in mod:
        st.subheader(f"🔄 {yr} — 招生漏斗分析")
        ds=build_dept_stats(p1,p2,p3)
        if ds is not None:
            st.dataframe(ds.sort_values("一→最終(%)",ascending=False),use_container_width=True,hide_index=True)
            rc=[c for c in ["一→二階(%)","一→最終(%)"] if c in ds.columns]
            if rc:
                st.plotly_chart(fig_grouped_bar(ds.sort_values(rc[0],ascending=True),"科系",rc,"各科系轉換率"),use_container_width=True)

            st.subheader("單科系漏斗")
            sel=st.selectbox("選擇科系：",ds["科系"].tolist(),key=f"fun_dept_{yr}")
            row=ds[ds["科系"]==sel].iloc[0]
            fl,fv=["一階報名"],[int(row["一階人數"])]
            if row["二階人數"]>0 or p2: fl.append("二階報到"); fv.append(int(row["二階人數"]))
            if row["最終入學"]>0 or p3: fl.append("最終入學"); fv.append(int(row["最終入學"]))
            st.plotly_chart(fig_funnel(fl,fv,f"{sel} 漏斗"),use_container_width=True)

        ss=build_school_stats(p1,p2,p3)
        if ss is not None:
            st.markdown("---"); st.subheader("各來源學校漏斗")
            mn=st.slider("一階≥",1,50,5,key=f"fun_mn_{yr}")
            sf=ss[ss["一階人數"]>=mn].sort_values("一→最終(%)",ascending=False)
            st.dataframe(sf,use_container_width=True,hide_index=True)
            st.plotly_chart(fig_bar_h(sf.head(20),"學校","一→最終(%)",
                f"來源學校轉換率 TOP 20（一階≥{mn}）",color="#4CAF50"),use_container_width=True)

    # ── 入學管道 ──
    elif "管道" in mod:
        st.subheader(f"📈 {yr} — 入學管道分析")
        if p3 is None or not ch_col or ch_col not in (p3.columns if p3 is not None else []):
            st.warning("⚠️ 需要最終入學資料及入學方式欄位。"); return
        cd=p3[ch_col].value_counts().reset_index(); cd.columns=["入學管道","人數"]
        cd["佔比(%)"]=( cd["人數"]/cd["人數"].sum()*100).round(1)
        cd["累積(%)"]=cd["佔比(%)"].cumsum().round(1)
        c1,c2=st.columns(2)
        with c1:
            fig=px.pie(cd,names="入學管道",values="人數",title="管道佔比",hole=.35)
            st.plotly_chart(fig,use_container_width=True)
        with c2:
            fig=px.bar(cd.sort_values("人數",ascending=True),x="人數",y="入學管道",orientation="h",
                text="人數",title="人數排行",color="佔比(%)",color_continuous_scale="Viridis")
            fig.update_layout(height=max(400,len(cd)*30),yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig,use_container_width=True)
        st.dataframe(cd,use_container_width=True,hide_index=True)

        dc3=detect_dept_col(p3)
        if dc3:
            st.markdown("---"); st.subheader("管道 × 科系")
            cross=p3.groupby([ch_col,dc3]).size().reset_index(name="人數")
            fig=px.bar(cross,x=ch_col,y="人數",color=dc3,barmode="stack",text="人數",title="管道×科系堆疊圖")
            fig.update_layout(height=600,xaxis_tickangle=-45)
            st.plotly_chart(fig,use_container_width=True)
            st.plotly_chart(fig_heatmap(cross,dc3,ch_col,"人數","管道×科系熱力圖"),use_container_width=True)

        sc3=detect_school_col(p3)
        if sc3:
            st.markdown("---"); st.subheader("管道 × 來源學校多元性")
            dv=p3.groupby(ch_col)[sc3].nunique().reset_index(); dv.columns=["入學管道","來源學校數"]
            fig=px.bar(dv.sort_values("來源學校數",ascending=True),x="來源學校數",y="入學管道",orientation="h",
                text="來源學校數",color="來源學校數",color_continuous_scale="Blues",title="學校多元性")
            fig.update_layout(height=max(400,len(dv)*28))
            st.plotly_chart(fig,use_container_width=True)

    # ── 地理分布 ──
    elif "地理" in mod:
        st.subheader(f"🗺️ {yr} — 地理分布")
        if geo is None:
            st.warning("⚠️ 一階資料無經緯度欄位，無法繪製地圖。"); return
        st.caption(f"📍 經緯度資料庫（一階）：{len(geo)} 所學校")
        sc=detect_school_col(p1)
        if not sc: st.warning("⚠️ 未偵測到學校欄位。"); return

        def do_map(src_df,count_label,title_text,phase):
            sc_=detect_school_col(src_df)
            if sc_ is None: return
            agg=src_df.groupby(sc_).size().reset_index(name=count_label)
            agg["_std"]=agg[sc_].apply(norm_school)
            agg=agg.merge(geo,on="_std",how="left").drop(columns=["_std"])
            fig=fig_map(agg,count_label,title_text)
            if fig:
                st.plotly_chart(fig,use_container_width=True)
                ok=agg["lat"].notna().sum()
                st.caption(f"匹配：{ok}/{len(agg)} 校（{ok/len(agg)*100:.1f}%）")
                miss=agg[agg["lat"].isna()]
                if not miss.empty:
                    with st.expander(f"⚠️ {phase} — {len(miss)} 校未匹配"):
                        st.dataframe(miss[[sc_,count_label]].sort_values(count_label,ascending=False),hide_index=True)
            else:
                st.info(f"{phase} 無匹配結果。")

        st.subheader("一階報名地圖")
        do_map(p1,"報名人數",f"{yr} 一階報名來源","一階")
        if p2 is not None:
            st.markdown("---"); st.subheader("二階報到地圖")
            do_map(p2,"報到人數",f"{yr} 二階報到來源","二階")
        if p3 is not None:
            st.markdown("---"); st.subheader("最終入學地圖")
            do_map(p3,"入學人數",f"{yr} 最終入學來源","最終入學")
            if ch_col and ch_col in p3.columns:
                st.markdown("---"); st.subheader("指定管道地圖")
                chl=p3[ch_col].value_counts().index.tolist()
                sel_ch=st.selectbox("選管道：",chl,key=f"mapch_{yr}")
                if sel_ch:
                    sub=p3[p3[ch_col]==sel_ch]
                    do_map(sub,"入學人數",f"{yr}「{sel_ch}」地圖",f"「{sel_ch}」")

    # ── 科系熱力圖 ──
    elif "熱力圖" in mod:
        st.subheader(f"🏫 {yr} — 科系×學校 熱力圖")
        dc=detect_dept_col(p1); sc=detect_school_col(p1)
        if not dc or not sc: st.warning("⚠️ 未偵測到科系或學校欄位。"); return
        mn=st.slider("學校報名≥",1,30,3,key=f"hm_{yr}")
        valid=p1[sc].value_counts(); valid=valid[valid>=mn].index
        filt=p1[p1[sc].isin(valid)]
        cr=filt.groupby([sc,dc]).size().reset_index(name="人數")
        st.plotly_chart(fig_heatmap(cr,dc,sc,"人數",f"一階報名（≥{mn}人）"),use_container_width=True)
        if p3 is not None:
            dc3=detect_dept_col(p3); sc3=detect_school_col(p3)
            if dc3 and sc3:
                cr3=p3.groupby([sc3,dc3]).size().reset_index(name="入學人數")
                cr3=cr3[cr3[sc3].isin(valid)]
                st.plotly_chart(fig_heatmap(cr3,dc3,sc3,"入學人數","最終入學"),use_container_width=True)

    # ── 來源學校 ──
    elif "來源學校" in mod:
        st.subheader(f"🎯 {yr} — 來源學校追蹤")
        ss=build_school_stats(p1,p2,p3)
        if ss is None: st.warning("⚠️ 未偵測到學校欄位。"); return
        def tier(n):
            if n>=30: return "Tier1(≥30)"
            elif n>=10: return "Tier2(10-29)"
            else: return "Tier3(<10)"
        ss["分級"]=ss["一階人數"].apply(tier)
        sel_t=st.multiselect("篩選分級：",["Tier1(≥30)","Tier2(10-29)","Tier3(<10)"],
            default=["Tier1(≥30)","Tier2(10-29)"],key=f"tier_{yr}")
        disp=ss[ss["分級"].isin(sel_t)].sort_values("一階人數",ascending=False)
        st.dataframe(disp,use_container_width=True,hide_index=True)

        st.subheader("個別學校")
        sel=st.selectbox("選擇學校：",ss.sort_values("一階人數",ascending=False)["學校"],key=f"sch_{yr}")
        if sel:
            r=ss[ss["學校"]==sel].iloc[0]
            c1,c2,c3,c4=st.columns(4)
            c1.metric("一階",f'{int(r["一階人數"])}')
            c2.metric("二階",f'{int(r["二階人數"])}')
            c3.metric("最終",f'{int(r["最終入學"])}')
            c4.metric("轉換率",f'{r["一→最終(%)"]}%')
            fl,fv=["一階"],[int(r["一階人數"])]
            if r["二階人數"]>0 or p2: fl.append("二階"); fv.append(int(r["二階人數"]))
            if r["最終入學"]>0 or p3: fl.append("最終"); fv.append(int(r["最終入學"]))
            if len(fv)>1: st.plotly_chart(fig_funnel(fl,fv,f"{sel} 漏斗"),use_container_width=True)

            if p3 is not None and ch_col and ch_col in p3.columns:
                sc3=detect_school_col(p3)
                if sc3:
                    sub=p3[p3[sc3]==sel]
                    if not sub.empty:
                        dd=sub[ch_col].value_counts().reset_index(); dd.columns=["管道","人數"]
                        fig=px.pie(dd,names="管道",values="人數",title=f"{sel} 入學管道",hole=.35)
                        st.plotly_chart(fig,use_container_width=True)

    # ── 流失預警 ──
    elif "流失" in mod:
        st.subheader(f"⚠️ {yr} — 流失預警分析")
        if p2 is None and p3 is None:
            st.warning("⚠️ 需要至少二階或最終入學資料。"); return
        ss=build_school_stats(p1,p2,p3)
        if ss is None: st.warning("⚠️ 未偵測到學校欄位。"); return
        has_final=p3 is not None
        rc="一→最終(%)" if has_final else "一→二階(%)"
        ll="最終入學" if has_final else "二階人數"
        sl="一→最終" if has_final else "一→二階"
        ss["流失人數"]=ss["一階人數"]-ss[ll]

        mn=st.slider("一階≥",1,50,10,key=f"loss_mn_{yr}")
        pool=ss[ss["一階人數"]>=mn]
        avg=pool[rc].mean()
        warn=pool[pool[rc]<avg].sort_values("流失人數",ascending=False)
        if warn.empty: st.success("✅ 沒有預警學校！")
        else:
            st.markdown(f'<div class="warning-box">⚠️ {len(warn)}所學校低於平均 {avg:.1f}%</div>',unsafe_allow_html=True)
            st.dataframe(warn,use_container_width=True,hide_index=True)

        st.subheader("IPA 矩陣")
        ana=ss[ss["一階人數"]>=mn].copy()
        if not ana.empty:
            med_x=ana["一階人數"].median(); med_y=ana[rc].median()
            fig=px.scatter(ana,x="一階人數",y=rc,size="流失人數",hover_name="學校",
                hover_data={"一階人數":True,ll:True,rc:True},
                title=f"IPA（{sl}）",size_max=40,color=rc,color_continuous_scale="RdYlGn")
            fig.add_hline(y=med_y,line_dash="dash",line_color="red",annotation_text=f"中位數 {med_y:.1f}%")
            fig.add_vline(x=med_x,line_dash="dash",line_color="blue",annotation_text=f"中位數 {med_x:.0f}")
            fig.update_layout(height=600)
            st.plotly_chart(fig,use_container_width=True)

        if p3 is not None and ch_col and ch_col in p3.columns:
            st.markdown("---"); st.subheader("Pareto 分析")
            cs=p3[ch_col].value_counts().reset_index(); cs.columns=["入學管道","人數"]
            cs["佔比(%)"]=( cs["人數"]/cs["人數"].sum()*100).round(1)
            cs["累積(%)"]=cs["佔比(%)"].cumsum().round(1)
            fig=go.Figure()
            fig.add_trace(go.Bar(x=cs["入學管道"],y=cs["人數"],name="人數",
                marker_color="#4CAF50",text=cs["人數"],textposition="outside"))
            fig.add_trace(go.Scatter(x=cs["入學管道"],y=cs["累積(%)"],name="累積%",
                yaxis="y2",line=dict(color="#FF5722",width=3),marker=dict(size=8)))
            fig.update_layout(title="Pareto圖",yaxis=dict(title="人數"),
                yaxis2=dict(title="累積%",overlaying="y",side="right",range=[0,105]),
                height=500,xaxis_tickangle=-45)
            st.plotly_chart(fig,use_container_width=True)

        st.markdown("---"); st.subheader("科系流失")
        ds=build_dept_stats(p1,p2,p3)
        if ds is not None:
            drc="一→最終(%)" if has_final else "一→二階(%)"
            dll="最終入學" if has_final else "二階人數"
            ds["流失人數"]=ds["一階人數"]-ds[dll]
            fig=px.scatter(ds,x="一階人數",y=drc,size="流失人數",hover_name="科系",text="科系",
                title=f"科系 IPA（{sl}）",size_max=50,color=drc,color_continuous_scale="RdYlGn")
            fig.update_traces(textposition="top center")
            fig.update_layout(height=500)
            st.plotly_chart(fig,use_container_width=True)
            st.dataframe(ds.sort_values("流失人數",ascending=False),use_container_width=True,hide_index=True)


# ============================================================
# 跨年度比較模組
# ============================================================
def render_cross_year():
    st.header("📊 跨年度比較分析")

    if len(valid_years)<2:
        st.warning("⚠️ 需要至少 2 個年度才能進行跨年度比較。"); return

    # 收集各年度摘要
    summaries=[]
    for yr in valid_years:
        p1,p2,p3,geo,ch_col=get_year_dfs(yr)
        if p1 is None: continue
        n1=len(p1); n2=len(p2) if p2 is not None else 0; n3=len(p3) if p3 is not None else 0
        sc=detect_school_col(p1); dc=detect_dept_col(p1)
        n_sch=p1[sc].nunique() if sc else 0
        n_dept=p1[dc].nunique() if dc else 0
        summaries.append({
            "年度":yr,"一階人數":n1,"二階人數":n2,"最終入學":n3,
            "一→二階(%)":round(n2/n1*100,1) if n1 and n2 else 0,
            "一→最終(%)":round(n3/n1*100,1) if n1 and n3 else 0,
            "來源學校數":n_sch,"科系數":n_dept
        })
    if not summaries:
        st.warning("⚠️ 無有效年度資料。"); return
    sdf=pd.DataFrame(summaries)

    # 7-1 年度總覽表
    st.subheader("7-1. 年度總覽")
    st.dataframe(sdf,use_container_width=True,hide_index=True)

    # 7-2 趨勢圖
    st.subheader("7-2. 招生量趨勢")
    fig=go.Figure()
    fig.add_trace(go.Bar(x=sdf["年度"],y=sdf["一階人數"],name="一階",marker_color="#2196F3"))
    if sdf["二階人數"].sum()>0:
        fig.add_trace(go.Bar(x=sdf["年度"],y=sdf["二階人數"],name="二階",marker_color="#FF9800"))
    if sdf["最終入學"].sum()>0:
        fig.add_trace(go.Bar(x=sdf["年度"],y=sdf["最終入學"],name="最終入學",marker_color="#4CAF50"))
    fig.update_layout(barmode="group",title="各年度招生量",height=450)
    st.plotly_chart(fig,use_container_width=True)

    # 7-3 轉換率趨勢
    st.subheader("7-3. 轉換率趨勢")
    fig=go.Figure()
    if sdf["一→二階(%)"].sum()>0:
        fig.add_trace(go.Scatter(x=sdf["年度"],y=sdf["一→二階(%)"],name="一→二階",
            mode="lines+markers+text",text=sdf["一→二階(%)"],textposition="top center",
            line=dict(width=3,color="#FF9800"),marker=dict(size=12)))
    if sdf["一→最終(%)"].sum()>0:
        fig.add_trace(go.Scatter(x=sdf["年度"],y=sdf["一→最終(%)"],name="一→最終",
            mode="lines+markers+text",text=sdf["一→最終(%)"],textposition="top center",
            line=dict(width=3,color="#4CAF50"),marker=dict(size=12)))
    fig.update_layout(title="轉換率趨勢",yaxis_title="轉換率(%)",height=400)
    st.plotly_chart(fig,use_container_width=True)

    # 7-4 科系跨年度比較
    st.markdown("---")
    st.subheader("7-4. 科系跨年度比較")
    all_depts=set()
    dept_data={}
    for yr in valid_years:
        p1,p2,p3,_,_=get_year_dfs(yr)
        if p1 is None: continue
        ds=build_dept_stats(p1,p2,p3)
        if ds is not None:
            dept_data[yr]=ds
            all_depts.update(ds["科系"].tolist())
    if dept_data and all_depts:
        sel_dept=st.selectbox("選擇科系：",sorted(all_depts),key="cross_dept")
        rows=[]
        for yr,ds in dept_data.items():
            r=ds[ds["科系"]==sel_dept]
            if not r.empty:
                r=r.iloc[0]
                rows.append({"年度":yr,"一階":int(r["一階人數"]),"二階":int(r["二階人數"]),
                    "最終":int(r["最終入學"]),"一→最終(%)":r["一→最終(%)"]})
        if rows:
            rdf=pd.DataFrame(rows)
            st.dataframe(rdf,use_container_width=True,hide_index=True)
            fig=go.Figure()
            fig.add_trace(go.Bar(x=rdf["年度"],y=rdf["一階"],name="一階",marker_color="#2196F3"))
            fig.add_trace(go.Bar(x=rdf["年度"],y=rdf["最終"],name="最終",marker_color="#4CAF50"))
            fig.add_trace(go.Scatter(x=rdf["年度"],y=rdf["一→最終(%)"],name="轉換率",yaxis="y2",
                mode="lines+markers",line=dict(color="#E91E63",width=3),marker=dict(size=10)))
            fig.update_layout(barmode="group",title=f"「{sel_dept}」跨年度趨勢",
                yaxis=dict(title="人數"),yaxis2=dict(title="轉換率(%)",overlaying="y",side="right"),height=450)
            st.plotly_chart(fig,use_container_width=True)

    # 7-5 來源學校跨年度比較
    st.markdown("---")
    st.subheader("7-5. 來源學校跨年度比較")
    all_schs=set()
    sch_data={}
    for yr in valid_years:
        p1,p2,p3,_,_=get_year_dfs(yr)
        if p1 is None: continue
        ss=build_school_stats(p1,p2,p3)
        if ss is not None:
            sch_data[yr]=ss
            all_schs.update(ss["學校"].tolist())
    if sch_data and all_schs:
        sel_sch=st.selectbox("選擇學校：",sorted(all_schs),key="cross_sch")
        rows=[]
        for yr,ss in sch_data.items():
            r=ss[ss["學校"]==sel_sch]
            if not r.empty:
                r=r.iloc[0]
                rows.append({"年度":yr,"一階":int(r["一階人數"]),"二階":int(r["二階人數"]),
                    "最終":int(r["最終入學"]),"一→最終(%)":r["一→最終(%)"],"流失":int(r["流失人數"])})
        if rows:
            rdf=pd.DataFrame(rows)
            st.dataframe(rdf,use_container_width=True,hide_index=True)
            fig=go.Figure()
            fig.add_trace(go.Bar(x=rdf["年度"],y=rdf["一階"],name="一階",marker_color="#2196F3"))
            fig.add_trace(go.Bar(x=rdf["年度"],y=rdf["最終"],name="最終",marker_color="#4CAF50"))
            fig.add_trace(go.Scatter(x=rdf["年度"],y=rdf["一→最終(%)"],name="轉換率",yaxis="y2",
                mode="lines+markers",line=dict(color="#E91E63",width=3),marker=dict(size=10)))
            fig.update_layout(barmode="group",title=f"「{sel_sch}」跨年度趨勢",
                yaxis=dict(title="人數"),yaxis2=dict(title="轉換率(%)",overlaying="y",side="right"),height=450)
            st.plotly_chart(fig,use_container_width=True)

    # 7-6 入學管道跨年度比較
    st.markdown("---")
    st.subheader("7-6. 入學管道跨年度比較")
    ch_year_data={}
    all_channels=set()
    for yr in valid_years:
        p1,p2,p3,_,ch_col=get_year_dfs(yr)
        if p3 is None or ch_col is None: continue
        if ch_col not in p3.columns: continue
        cd=p3[ch_col].value_counts().reset_index(); cd.columns=["管道","人數"]
        cd["年度"]=yr
        ch_year_data[yr]=cd
        all_channels.update(cd["管道"].tolist())

    if ch_year_data and all_channels:
        combined=pd.concat(ch_year_data.values(),ignore_index=True)

        # 全管道對比
        pv=combined.pivot_table(index="管道",columns="年度",values="人數",aggfunc="sum").fillna(0)
        fig=px.imshow(pv,text_auto=True,aspect="auto",color_continuous_scale="YlGnBu",
            title="入學管道×年度 人數矩陣")
        fig.update_layout(height=max(400,len(pv)*30))
        st.plotly_chart(fig,use_container_width=True)

        # 單管道趨勢
        sel_ch=st.selectbox("選擇管道查看趨勢：",sorted(all_channels),key="cross_ch")
        if sel_ch:
            sub=combined[combined["管道"]==sel_ch].sort_values("年度")
            if not sub.empty:
                fig=go.Figure()
                fig.add_trace(go.Bar(x=sub["年度"],y=sub["人數"],name="入學人數",
                    marker_color="#4CAF50",text=sub["人數"],textposition="outside"))
                fig.add_trace(go.Scatter(x=sub["年度"],y=sub["人數"],mode="lines+markers",
                    name="趨勢線",line=dict(color="#FF5722",width=3,dash="dash")))
                fig.update_layout(title=f"「{sel_ch}」跨年度趨勢",height=400)
                st.plotly_chart(fig,use_container_width=True)

    # 7-7 年度增減分析
    st.markdown("---")
    st.subheader("7-7. 年度增減分析")
    if len(sdf)>=2:
        latest=sdf.iloc[0]; prev=sdf.iloc[1]
        chg1=latest["一階人數"]-prev["一階人數"]
        chg3=latest["最終入學"]-prev["最終入學"]
        chg_r=latest["一→最終(%)"]-prev["一→最終(%)"]

        cols=st.columns(3)
        with cols[0]:
            icon="📈" if chg1>=0 else "📉"
            st.metric(f"一階人數變化（{latest['年度']} vs {prev['年度']}）",
                f"{latest['一階人數']:,}",f"{chg1:+,}")
        with cols[1]:
            st.metric("最終入學變化",f"{latest['最終入學']:,}",f"{chg3:+,}")
        with cols[2]:
            st.metric("轉換率變化",f"{latest['一→最終(%)']}%",f"{chg_r:+.1f}%")

        # 科系增減
        if len(dept_data)>=2:
            yrs=list(dept_data.keys())
            d1=dept_data[yrs[0]].set_index("科系")
            d2=dept_data[yrs[1]].set_index("科系")
            common=list(set(d1.index)&set(d2.index))
            if common:
                chg_rows=[]
                for dept in common:
                    v1=d1.loc[dept,"一階人數"]; v2=d2.loc[dept,"一階人數"]
                    f1=d1.loc[dept,"最終入學"]; f2=d2.loc[dept,"最終入學"]
                    chg_rows.append({"科系":dept,
                        f"{yrs[0]}一階":int(v1),f"{yrs[1]}一階":int(v2),"一階增減":int(v1-v2),
                        f"{yrs[0]}最終":int(f1),f"{yrs[1]}最終":int(f2),"最終增減":int(f1-f2)})
                chg_df=pd.DataFrame(chg_rows).sort_values("最終增減",ascending=False)
                st.dataframe(chg_df,use_container_width=True,hide_index=True)

                fig=px.bar(chg_df,x="科系",y="最終增減",text="最終增減",
                    title=f"科系最終入學增減（{yrs[0]} vs {yrs[1]}）",
                    color="最終增減",color_continuous_scale="RdYlGn")
                fig.update_layout(height=450)
                st.plotly_chart(fig,use_container_width=True)


# ============================================================
# 主畫面：標籤式介面
# ============================================================
tab_names=valid_years+["📊 跨年度比較"]
tabs=st.tabs(tab_names)

for i,yr in enumerate(valid_years):
    with tabs[i]:
        st.markdown(f'<span class="year-tag">📅 {yr}</span>',unsafe_allow_html=True)
        render_year_analysis(yr)

with tabs[-1]:
    render_cross_year()


# ============================================================
# Footer
# ============================================================
st.markdown('<hr class="section-divider">',unsafe_allow_html=True)
st.markdown(f"""
<div style="text-align:center;color:#aaa;font-size:.85rem;padding:10px;">
    🎓 中華醫事科技大學 招生數據分析系統 v6.0<br>
    多年度標籤式分析 ｜ 經緯度從一階讀取 ｜ 入學管道自動偵測<br>
    跨年度比較（科系/學校/管道/增減分析）<br>
    分析版本 #{st.session_state.get("analysis_version",0)}　年度數 {len(valid_years)}<br>
    Built with Streamlit + Plotly ｜ © 2024 HWU Admissions Office
</div>
""",unsafe_allow_html=True)
