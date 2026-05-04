import streamlit as st
import pandas as pd
import json
import os
import time
import re
import requests
from datetime import datetime
from math import radians, cos, sin, asin, sqrt
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
from io import BytesIO

# ============================================================
# 系統設定
# ============================================================
st.set_page_config(
    page_title="學校座標查詢系統 v3.1",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 資料庫檔案路徑
DB_FILE = "school_coordinates_db.json"
BACKUP_DIR = "backups"
LOG_FILE = "search_log.json"

# API 設定
API_DELAY = 0.15
MAX_WORKERS = 4

# ============================================================
# 工具函數
# ============================================================
def load_database():
    """載入本地座標資料庫"""
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_database(db):
    """儲存座標資料庫"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    if os.path.exists(DB_FILE):
        backup_name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        try:
            import shutil
            shutil.copy2(DB_FILE, os.path.join(BACKUP_DIR, backup_name))
        except:
            pass
    with open(DB_FILE, 'w', encoding='utf-8') as f:
        json.dump(db, f, ensure_ascii=False, indent=2)

def load_search_log():
    """載入搜尋記錄"""
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []
    return []

def save_search_log(log):
    """儲存搜尋記錄"""
    with open(LOG_FILE, 'w', encoding='utf-8') as f:
        json.dump(log[-1000:], f, ensure_ascii=False, indent=2)

def normalize_school_name(name):
    """標準化學校名稱"""
    if not isinstance(name, str):
        return ""
    name = name.strip()
    name = re.sub(r'\s+', '', name)
    # 移除常見前綴的變體
    replacements = {
        '台': '臺',
        '臺北縣': '新北市',
        '桃園縣': '桃園市',
    }
    for old, new in replacements.items():
        if old in name and old != '台':
            name = name.replace(old, new)
    # 台/臺 統一
    name_normalized = name.replace('台', '臺')
    return name_normalized

def generate_search_variants(name):
    """產生搜尋變體"""
    variants = [name]
    
    # 台/臺互換
    if '臺' in name:
        variants.append(name.replace('臺', '台'))
    if '台' in name:
        variants.append(name.replace('台', '臺'))
    
    # 移除縣市前綴
    prefixes = [
        '臺北市', '台北市', '新北市', '桃園市', '臺中市', '台中市',
        '臺南市', '台南市', '高雄市', '基隆市', '新竹市', '新竹縣',
        '苗栗縣', '彰化縣', '南投縣', '雲林縣', '嘉義市', '嘉義縣',
        '屏東縣', '宜蘭縣', '花蓮縣', '臺東縣', '台東縣', '澎湖縣',
        '金門縣', '連江縣'
    ]
    for prefix in prefixes:
        if name.startswith(prefix):
            short = name[len(prefix):]
            if short and len(short) >= 2:
                variants.append(short)
            break
    
    # 補上縣市前綴
    has_prefix = any(name.startswith(p) for p in prefixes)
    if not has_prefix and len(name) >= 2:
        common_cities = ['臺北市', '新北市', '桃園市', '臺中市', '臺南市', '高雄市']
        for city in common_cities:
            variants.append(city + name)
    
    # 學校類型變體
    type_map = {
        '國小': ['國民小學', '小學'],
        '國民小學': ['國小', '小學'],
        '國中': ['國民中學', '中學'],
        '國民中學': ['國中'],
        '高中': ['高級中學', '高級中等學校'],
        '高級中學': ['高中'],
        '高工': ['高級工業職業學校', '高級工商職業學校'],
        '高商': ['高級商業職業學校'],
    }
    for short_form, long_forms in type_map.items():
        if short_form in name:
            for lf in long_forms:
                variants.append(name.replace(short_form, lf))
    
    return list(dict.fromkeys(variants))

def haversine(lon1, lat1, lon2, lat2):
    """計算兩點間距離（公里）"""
    lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a))
    return 6371 * c

# ============================================================
# 地理編碼引擎
# ============================================================
class GeocodingEngine:
    """多引擎地理編碼"""
    
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'SchoolGeocoder/3.1 (Educational Research Project)'
        })
        self.lock = threading.Lock()
        self.search_log = load_search_log()
    
    def _log_search(self, name, result, engine, elapsed):
        """記錄搜尋"""
        entry = {
            'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'name': name,
            'success': result is not None,
            'engine': engine,
            'elapsed': round(elapsed, 3)
        }
        if result:
            entry['lat'] = result[0]
            entry['lon'] = result[1]
        with self.lock:
            self.search_log.append(entry)
    
    def search_nominatim(self, query):
        """Nominatim 搜尋"""
        try:
            url = "https://nominatim.openstreetmap.org/search"
            params = {
                'q': query,
                'format': 'json',
                'limit': 3,
                'countrycodes': 'tw',
                'addressdetails': 1
            }
            resp = self.session.get(url, params=params, timeout=10)
            if resp.status_code == 200:
                data = resp.json()
                if data:
                    # 優先找學校類型
                    for item in data:
                        item_type = item.get('type', '')
                        item_class = item.get('class', '')
                        if item_class == 'amenity' and item_type == 'school':
                            return (float(item['lat']), float(item['lon']))
                    # 否則回傳第一個
                    return (float(data[0]['lat']), float(data[0]['lon']))
        except:
            pass
        return None
    
    def search_photon(self, query):
        """Photon 搜尋"""
        try:
            url = "https://photon.komoot.io/api/"
            params = {
                'q': query + ' Taiwan',
                'limit': 3,
                'lang': 'zh'
            }
            resp = self.session.get(url, params=params, timeout=10)
            if resp.status_code == 200:
                data = resp.json()
                features = data.get('features', [])
                if features:
                    # 優先學校
                    for f in features:
                        props = f.get('properties', {})
                        if props.get('osm_value') == 'school':
                            coords = f['geometry']['coordinates']
                            return (coords[1], coords[0])
                    coords = features[0]['geometry']['coordinates']
                    return (coords[1], coords[0])
        except:
            pass
        return None
    
    def search_osm_structured(self, query):
        """OSM 結構化搜尋"""
        try:
            url = "https://nominatim.openstreetmap.org/search"
            params = {
                'q': query,
                'format': 'json',
                'limit': 5,
                'countrycodes': 'tw'
            }
            resp = self.session.get(url, params=params, timeout=10)
            if resp.status_code == 200:
                data = resp.json()
                for item in data:
                    lat = float(item['lat'])
                    lon = float(item['lon'])
                    # 確認在台灣範圍
                    if 21.5 < lat < 25.5 and 119.5 < lon < 122.5:
                        return (lat, lon)
        except:
            pass
        return None
    
    def geocode_school(self, school_name, db):
        """多引擎搜尋學校座標"""
        start_time = time.time()
        normalized = normalize_school_name(school_name)
        
        # 1. 先查資料庫
        if normalized in db:
            entry = db[normalized]
            elapsed = time.time() - start_time
            self._log_search(school_name, (entry['lat'], entry['lon']), 'database', elapsed)
            return {
                'name': school_name,
                'normalized': normalized,
                'lat': entry['lat'],
                'lon': entry['lon'],
                'source': 'database',
                'success': True
            }
        
        # 2. 檢查變體是否在資料庫
        variants = generate_search_variants(normalized)
        for v in variants:
            v_norm = normalize_school_name(v)
            if v_norm in db:
                entry = db[v_norm]
                elapsed = time.time() - start_time
                self._log_search(school_name, (entry['lat'], entry['lon']), 'database_variant', elapsed)
                return {
                    'name': school_name,
                    'normalized': normalized,
                    'lat': entry['lat'],
                    'lon': entry['lon'],
                    'source': 'database (variant)',
                    'success': True
                }
        
        # 3. 線上搜尋
        engines = [
            ('Nominatim', self.search_nominatim),
            ('Photon', self.search_photon),
            ('OSM', self.search_osm_structured),
        ]
        
        for engine_name, engine_func in engines:
            for variant in variants[:4]:
                try:
                    result = engine_func(variant)
                    if result:
                        lat, lon = result
                        # 驗證在台灣範圍
                        if 21.5 < lat < 25.5 and 119.5 < lon < 122.5:
                            # 存入資料庫
                            db[normalized] = {
                                'lat': lat,
                                'lon': lon,
                                'source': engine_name,
                                'query': variant,
                                'updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            }
                            elapsed = time.time() - start_time
                            self._log_search(school_name, (lat, lon), engine_name, elapsed)
                            return {
                                'name': school_name,
                                'normalized': normalized,
                                'lat': lat,
                                'lon': lon,
                                'source': engine_name,
                                'success': True
                            }
                    time.sleep(API_DELAY)
                except:
                    time.sleep(API_DELAY)
        
        elapsed = time.time() - start_time
        self._log_search(school_name, None, 'all_failed', elapsed)
        return {
            'name': school_name,
            'normalized': normalized,
            'lat': None,
            'lon': None,
            'source': 'not_found',
            'success': False
        }

# ============================================================
# Excel 工具函數
# ============================================================
def detect_columns(df):
    """自動偵測可能的欄位"""
    result = {'id': None, 'name': None, 'school': None}
    
    # 學號/座號偵測
    id_keywords = ['學號', '座號', '編號', 'id', 'ID', 'number', '序號']
    for col in df.columns:
        col_str = str(col).strip()
        for kw in id_keywords:
            if kw in col_str:
                result['id'] = col
                break
        if result['id']:
            break
    
    # 姓名偵測
    name_keywords = ['姓名', '名字', '學生', 'name', 'Name', '名稱']
    for col in df.columns:
        col_str = str(col).strip()
        for kw in name_keywords:
            if kw in col_str and '學校' not in col_str:
                result['name'] = col
                break
        if result['name']:
            break
    
    # 學校偵測
    school_keywords = ['學校', '校名', '畢業', '國小', '國中', '就讀', 'school', 'School']
    for col in df.columns:
        col_str = str(col).strip()
        for kw in school_keywords:
            if kw in col_str:
                result['school'] = col
                break
        if result['school']:
            break
    
    return result

def create_sample_excel():
    """建立範例 Excel"""
    data = {
        '學號': ['S001', 'S002', 'S003', 'S004', 'S005'],
        '姓名': ['王小明', '李小華', '張小美', '陳小強', '林小玲'],
        '畢業國小': ['臺北市大安國小', '新北市板橋國小', '桃園市中壢國小', '臺中市西區大同國小', '高雄市前鎮國小']
    }
    return pd.DataFrame(data)

def df_to_excel_bytes(df):
    """DataFrame 轉 Excel bytes"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='資料')
    output.seek(0)
    return output.getvalue()

# ============================================================
# 主介面
# ============================================================
def main():
    # 側邊欄
    with st.sidebar:
        st.title("🏫 學校座標查詢系統")
        st.caption("v3.1 Excel 版 · 本地加速")
        st.divider()
        
        # 資料庫狀態
        db = load_database()
        st.metric("📦 資料庫學校數", f"{len(db)} 所")
        
        if db:
            sources = {}
            for v in db.values():
                src = v.get('source', 'unknown')
                sources[src] = sources.get(src, 0) + 1
            with st.expander("資料來源分布"):
                for src, cnt in sorted(sources.items(), key=lambda x: -x[1]):
                    st.write(f"- {src}: {cnt}")
        
        st.divider()
        st.markdown("### 🗄️ 資料庫管理")
        
        # 匯出資料庫
        if db:
            db_json = json.dumps(db, ensure_ascii=False, indent=2)
            st.download_button(
                "📥 匯出資料庫 (JSON)",
                data=db_json,
                file_name=f"school_db_{datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json"
            )
        
        # 匯入資料庫
        uploaded_db = st.file_uploader("📤 匯入資料庫", type=['json'], key='db_import')
        if uploaded_db:
            try:
                imported = json.loads(uploaded_db.read().decode('utf-8'))
                if isinstance(imported, dict):
                    db.update(imported)
                    save_database(db)
                    st.success(f"✅ 匯入成功！合併後共 {len(db)} 筆")
                    st.rerun()
            except Exception as e:
                st.error(f"匯入失敗：{e}")
    
    # 主頁面標籤
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📤 上傳 Excel",
        "🗺️ 地圖",
        "📊 統計分析",
        "✏️ 手動編輯",
        "🔧 進階工具"
    ])
    
    # ========================================
    # Tab 1: 上傳 Excel
    # ========================================
    with tab1:
        st.header("📤 上傳 Excel 檔案")
        
        # 範例下載
        col_sample1, col_sample2 = st.columns([1, 3])
        with col_sample1:
            sample_df = create_sample_excel()
            st.download_button(
                "📋 下載範例 Excel",
                data=df_to_excel_bytes(sample_df),
                file_name="範例_學生資料.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_sample2:
            st.info("💡 Excel 檔案需包含「學校名稱」欄位，系統會自動偵測。")
        
        # 上傳
        uploaded_file = st.file_uploader(
            "上傳學生資料 Excel 檔案",
            type=['xlsx', 'xls'],
            help="支援 .xlsx 和 .xls 格式"
        )
        
        if uploaded_file:
            try:
                # 讀取 Excel
                if uploaded_file.name.endswith('.xls'):
                    df_raw = pd.read_excel(uploaded_file, engine='xlrd')
                else:
                    df_raw = pd.read_excel(uploaded_file, engine='openpyxl')
                
                st.success(f"✅ 成功讀取：{len(df_raw)} 筆資料，{len(df_raw.columns)} 個欄位")
                
                # 預覽
                with st.expander("📋 資料預覽（前 10 筆）", expanded=True):
                    st.dataframe(df_raw.head(10), use_container_width=True)
                
                # 欄位對應
                st.subheader("🔗 欄位對應")
                detected = detect_columns(df_raw)
                
                col_map1, col_map2, col_map3 = st.columns(3)
                columns_list = ['（不使用）'] + list(df_raw.columns)
                
                with col_map1:
                    id_default = columns_list.index(detected['id']) if detected['id'] in columns_list else 0
                    id_col = st.selectbox(
                        "學號/座號 欄位",
                        columns_list,
                        index=id_default,
                        help="選擇學號或座號欄位（可不選）"
                    )
                
                with col_map2:
                    name_default = columns_list.index(detected['name']) if detected['name'] in columns_list else 0
                    name_col = st.selectbox(
                        "姓名 欄位",
                        columns_list,
                        index=name_default,
                        help="選擇學生姓名欄位（可不選）"
                    )
                
                with col_map3:
                    school_default = columns_list.index(detected['school']) if detected['school'] in columns_list else 0
                    school_col = st.selectbox(
                        "⭐ 學校名稱 欄位（必選）",
                        columns_list,
                        index=school_default,
                        help="選擇包含學校名稱的欄位"
                    )
                
                if school_col == '（不使用）':
                    st.warning("⚠️ 請選擇「學校名稱」欄位後才能開始查詢")
                else:
                    # 整理資料
                    result_data = []
                    for idx, row in df_raw.iterrows():
                        entry = {'原始索引': idx}
                        if id_col != '（不使用）':
                            entry['學號'] = row[id_col]
                        if name_col != '（不使用）':
                            entry['姓名'] = row[name_col]
                        entry['學校名稱'] = str(row[school_col]).strip() if pd.notna(row[school_col]) else ''
                        result_data.append(entry)
                    
                    df_work = pd.DataFrame(result_data)
                    df_work = df_work[df_work['學校名稱'].str.len() > 0]
                    
                    st.write(f"📊 有效資料：**{len(df_work)}** 筆")
                    
                    # 查詢按鈕
                    col_btn1, col_btn2 = st.columns([1, 3])
                    with col_btn1:
                        start_search = st.button(
                            "🚀 開始查詢座標",
                            type="primary",
                            use_container_width=True
                        )
                    with col_btn2:
                        workers = st.slider("並行數", 1, 6, MAX_WORKERS, help="越高越快，但過高可能被限速")
                    
                    if start_search:
                        db = load_database()
                        engine = GeocodingEngine()
                        
                        # 取得不重複學校名
                        unique_schools = df_work['學校名稱'].unique().tolist()
                        st.write(f"🏫 不重複學校：**{len(unique_schools)}** 所")
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        results_container = st.empty()
                        
                        school_results = {}
                        success_count = 0
                        fail_count = 0
                        start_time = time.time()
                        
                        # 先處理資料庫已有的
                        to_search = []
                        for school in unique_schools:
                            normalized = normalize_school_name(school)
                            found = False
                            
                            if normalized in db:
                                school_results[school] = {
                                    'lat': db[normalized]['lat'],
                                    'lon': db[normalized]['lon'],
                                    'source': 'database'
                                }
                                success_count += 1
                                found = True
                            
                            if not found:
                                variants = generate_search_variants(normalized)
                                for v in variants:
                                    v_norm = normalize_school_name(v)
                                    if v_norm in db:
                                        school_results[school] = {
                                            'lat': db[v_norm]['lat'],
                                            'lon': db[v_norm]['lon'],
                                            'source': 'database (variant)'
                                        }
                                        success_count += 1
                                        found = True
                                        break
                            
                            if not found:
                                to_search.append(school)
                        
                        status_text.write(f"📦 資料庫命中：{success_count} 所 | 需線上查詢：{len(to_search)} 所")
                        
                        # 線上搜尋
                        if to_search:
                            def search_one(school_name):
                                return engine.geocode_school(school_name, db)
                            
                            completed = 0
                            total = len(to_search)
                            
                            with ThreadPoolExecutor(max_workers=workers) as executor:
                                futures = {executor.submit(search_one, s): s for s in to_search}
                                for future in as_completed(futures):
                                    result = future.result()
                                    completed += 1
                                    
                                    if result['success']:
                                        school_results[result['name']] = {
                                            'lat': result['lat'],
                                            'lon': result['lon'],
                                            'source': result['source']
                                        }
                                        success_count += 1
                                    else:
                                        fail_count += 1
                                    
                                    progress = (success_count + fail_count) / len(unique_schools)
                                    progress_bar.progress(min(progress, 1.0))
                                    elapsed = time.time() - start_time
                                    status_text.write(
                                        f"⏱️ {elapsed:.1f}s | "
                                        f"✅ {success_count} | ❌ {fail_count} | "
                                        f"🔄 查詢中 {completed}/{total}"
                                    )
                        
                        # 儲存資料庫
                        save_database(db)
                        save_search_log(engine.search_log)
                        
                        progress_bar.progress(1.0)
                        total_time = time.time() - start_time
                        status_text.write(
                            f"🎉 完成！耗時 {total_time:.1f} 秒 | "
                            f"✅ 成功 {success_count} | ❌ 失敗 {fail_count}"
                        )
                        
                        # 合併結果
                        lat_list = []
                        lon_list = []
                        source_list = []
                        for _, row in df_work.iterrows():
                            school = row['學校名稱']
                            if school in school_results:
                                lat_list.append(school_results[school]['lat'])
                                lon_list.append(school_results[school]['lon'])
                                source_list.append(school_results[school]['source'])
                            else:
                                lat_list.append(None)
                                lon_list.append(None)
                                source_list.append('not_found')
                        
                        df_work['緯度'] = lat_list
                        df_work['經度'] = lon_list
                        df_work['來源'] = source_list
                        
                        # 存到 session
                        st.session_state['result_df'] = df_work
                        st.session_state['school_results'] = school_results
                        
                        # 顯示結果
                        st.subheader("📋 查詢結果")
                        
                        # 成功/失敗分開顯示
                        df_success = df_work[df_work['緯度'].notna()]
                        df_fail = df_work[df_work['緯度'].isna()]
                        
                        tab_s, tab_f = st.tabs([
                            f"✅ 成功 ({len(df_success)})",
                            f"❌ 未找到 ({len(df_fail)})"
                        ])
                        
                        with tab_s:
                            if len(df_success) > 0:
                                st.dataframe(df_success, use_container_width=True)
                        
                        with tab_f:
                            if len(df_fail) > 0:
                                st.dataframe(df_fail, use_container_width=True)
                                st.info("💡 未找到的學校可以在「手動編輯」分頁手動輸入座標")
                        
                        # 下載結果
                        st.subheader("📥 下載結果")
                        col_dl1, col_dl2 = st.columns(2)
                        
                        with col_dl1:
                            st.download_button(
                                "📥 下載完整結果 (Excel)",
                                data=df_to_excel_bytes(df_work),
                                file_name=f"座標查詢結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        
                        with col_dl2:
                            if len(df_fail) > 0:
                                st.download_button(
                                    "📥 下載未找到清單 (Excel)",
                                    data=df_to_excel_bytes(df_fail),
                                    file_name=f"未找到學校_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
            
            except Exception as e:
                st.error(f"❌ 讀取檔案失敗：{e}")
                st.info("請確認檔案是有效的 Excel 格式（.xlsx 或 .xls）")
    
    # ========================================
    # Tab 2: 地圖
    # ========================================
    with tab2:
        st.header("🗺️ 學校分布地圖")
        
        if 'result_df' not in st.session_state:
            st.info("📤 請先在「上傳 Excel」分頁查詢座標")
        else:
            df_map = st.session_state['result_df']
            df_valid = df_map[df_map['緯度'].notna() & df_map['經度'].notna()].copy()
            
            if len(df_valid) == 0:
                st.warning("沒有可顯示的座標資料")
            else:
                try:
                    import folium
                    from folium.plugins import MarkerCluster, HeatMap
                    from streamlit_folium import st_folium
                    
                    map_type = st.radio(
                        "地圖模式",
                        ["標記地圖", "聚合地圖", "熱力圖"],
                        horizontal=True
                    )
                    
                    center_lat = df_valid['緯度'].mean()
                    center_lon = df_valid['經度'].mean()
                    
                    m = folium.Map(location=[center_lat, center_lon], zoom_start=8)
                    
                    if map_type == "標記地圖":
                        for _, row in df_valid.iterrows():
                            popup_text = f"🏫 {row['學校名稱']}"
                            if '姓名' in row and pd.notna(row.get('姓名')):
                                popup_text += f"<br>👤 {row['姓名']}"
                            
                            folium.CircleMarker(
                                location=[row['緯度'], row['經度']],
                                radius=6,
                                popup=folium.Popup(popup_text, max_width=200),
                                tooltip=row['學校名稱'],
                                color='#3388ff',
                                fill=True,
                                fillColor='#3388ff',
                                fillOpacity=0.7
                            ).add_to(m)
                    
                    elif map_type == "聚合地圖":
                        cluster = MarkerCluster().add_to(m)
                        for _, row in df_valid.iterrows():
                            popup_text = f"🏫 {row['學校名稱']}"
                            if '姓名' in row and pd.notna(row.get('姓名')):
                                popup_text += f"<br>👤 {row['姓名']}"
                            
                            folium.Marker(
                                location=[row['緯度'], row['經度']],
                                popup=folium.Popup(popup_text, max_width=200),
                                tooltip=row['學校名稱']
                            ).add_to(cluster)
                    
                    else:  # 熱力圖
                        heat_data = df_valid[['緯度', '經度']].values.tolist()
                        HeatMap(heat_data, radius=15).add_to(m)
                    
                    st_folium(m, width=None, height=600)
                    st.caption(f"📍 顯示 {len(df_valid)} 個地點")
                
                except ImportError:
                    st.error("需要安裝 folium 和 streamlit-folium：")
                    st.code("pip install folium streamlit-folium")
    
    # ========================================
    # Tab 3: 統計分析
    # ========================================
    with tab3:
        st.header("📊 統計分析")
        
        if 'result_df' not in st.session_state:
            st.info("📤 請先在「上傳 Excel」分頁查詢座標")
        else:
            df_stats = st.session_state['result_df']
            df_valid = df_stats[df_stats['緯度'].notna()].copy()
            
            if len(df_valid) == 0:
                st.warning("沒有可分析的資料")
            else:
                # 基本統計
                col_s1, col_s2, col_s3, col_s4 = st.columns(4)
                with col_s1:
                    st.metric("總筆數", len(df_stats))
                with col_s2:
                    st.metric("成功查詢", len(df_valid))
                with col_s3:
                    st.metric("不重複學校", df_valid['學校名稱'].nunique())
                with col_s4:
                    rate = len(df_valid) / len(df_stats) * 100 if len(df_stats) > 0 else 0
                    st.metric("成功率", f"{rate:.1f}%")
                
                st.divider()
                
                # 學校統計
                st.subheader("🏫 各學校學生數")
                school_counts = df_valid['學校名稱'].value_counts()
                
                # 長條圖
                top_n = min(20, len(school_counts))
                st.bar_chart(school_counts.head(top_n))
                
                # 表格
                with st.expander(f"完整列表（{len(school_counts)} 所學校）"):
                    df_school_stats = pd.DataFrame({
                        '學校': school_counts.index,
                        '學生數': school_counts.values
                    })
                    st.dataframe(df_school_stats, use_container_width=True)
                
                st.divider()
                
                # 區域分析
                st.subheader("📍 區域分布")
                
                def get_region(name):
                    regions = {
                        '臺北': '臺北市', '台北': '臺北市',
                        '新北': '新北市',
                        '桃園': '桃園市',
                        '臺中': '臺中市', '台中': '臺中市',
                        '臺南': '臺南市', '台南': '臺南市',
                        '高雄': '高雄市',
                        '基隆': '基隆市',
                        '新竹': '新竹',
                        '苗栗': '苗栗縣',
                        '彰化': '彰化縣',
                        '南投': '南投縣',
                        '雲林': '雲林縣',
                        '嘉義': '嘉義',
                        '屏東': '屏東縣',
                        '宜蘭': '宜蘭縣',
                        '花蓮': '花蓮縣',
                        '臺東': '臺東縣', '台東': '臺東縣',
                        '澎湖': '澎湖縣',
                        '金門': '金門縣',
                        '連江': '連江縣',
                    }
                    for key, region in regions.items():
                        if key in str(name):
                            return region
                    return '其他'
                
                df_valid['區域'] = df_valid['學校名稱'].apply(get_region)
                region_counts = df_valid['區域'].value_counts()
                
                st.bar_chart(region_counts)
                
                # 距離分析
                st.divider()
                st.subheader("📏 距離分析")
                
                ref_lat = st.number_input("參考點緯度", value=24.9936, format="%.4f", help="預設為學校所在地")
                ref_lon = st.number_input("參考點經度", value=121.3010, format="%.4f")
                
                if st.button("計算距離"):
                    distances = []
                    for _, row in df_valid.iterrows():
                        d = haversine(ref_lon, ref_lat, row['經度'], row['緯度'])
                        distances.append(round(d, 2))
                    df_valid['距離(km)'] = distances
                    
                    st.write(f"平均距離：**{sum(distances)/len(distances):.2f}** km")
                    st.write(f"最近：**{min(distances):.2f}** km")
                    st.write(f"最遠：**{max(distances):.2f}** km")
                    
                    # 距離分佈
                    bins = [0, 5, 10, 20, 50, 100, float('inf')]
                    labels = ['0-5km', '5-10km', '10-20km', '20-50km', '50-100km', '100km+']
                    df_valid['距離區間'] = pd.cut(distances, bins=bins, labels=labels)
                    dist_counts = df_valid['距離區間'].value_counts().sort_index()
                    st.bar_chart(dist_counts)
    
    # ========================================
    # Tab 4: 手動編輯
    # ========================================
    with tab4:
        st.header("✏️ 手動編輯座標")
        
        db = load_database()
        
        # 新增
        st.subheader("➕ 新增/修改學校座標")
        col_e1, col_e2, col_e3 = st.columns([2, 1, 1])
        with col_e1:
            edit_name = st.text_input("學校名稱", placeholder="例：臺北市大安國小")
        with col_e2:
            edit_lat = st.number_input("緯度", value=25.0330, format="%.6f", key='edit_lat')
        with col_e3:
            edit_lon = st.number_input("經度", value=121.5654, format="%.6f", key='edit_lon')
        
        if st.button("💾 儲存", type="primary"):
            if edit_name:
                normalized = normalize_school_name(edit_name)
                db[normalized] = {
                    'lat': edit_lat,
                    'lon': edit_lon,
                    'source': 'manual',
                    'updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                save_database(db)
                st.success(f"✅ 已儲存：{edit_name} ({edit_lat}, {edit_lon})")
                st.rerun()
        
        st.divider()
        
        # 查詢失敗的學校
        if 'result_df' in st.session_state:
            df_fail = st.session_state['result_df'][st.session_state['result_df']['緯度'].isna()]
            if len(df_fail) > 0:
                st.subheader(f"❌ 查詢失敗的學校（{len(df_fail)} 筆）")
                
                fail_schools = df_fail['學校名稱'].unique()
                for school in fail_schools:
                    with st.expander(f"🏫 {school}"):
                        col_f1, col_f2 = st.columns(2)
                        with col_f1:
                            fix_lat = st.number_input(
                                "緯度", value=25.0, format="%.6f",
                                key=f"fix_lat_{school}"
                            )
                        with col_f2:
                            fix_lon = st.number_input(
                                "經度", value=121.5, format="%.6f",
                                key=f"fix_lon_{school}"
                            )
                        
                        google_url = f"https://www.google.com/maps/search/{school}"
                        st.markdown(f"🔍 [Google Maps 搜尋]({google_url})")
                        
                        if st.button(f"儲存 {school}", key=f"save_{school}"):
                            normalized = normalize_school_name(school)
                            db[normalized] = {
                                'lat': fix_lat,
                                'lon': fix_lon,
                                'source': 'manual_fix',
                                'updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            }
                            save_database(db)
                            st.success(f"✅ 已儲存 {school}")
        
        st.divider()
        
        # 瀏覽資料庫
        st.subheader("📖 瀏覽資料庫")
        if db:
            search_term = st.text_input("🔍 搜尋學校", placeholder="輸入關鍵字篩選...")
            
            filtered_db = {k: v for k, v in db.items()
                         if not search_term or search_term in k}
            
            if filtered_db:
                db_display = []
                for name, info in sorted(filtered_db.items()):
                    db_display.append({
                        '學校': name,
                        '緯度': info['lat'],
                        '經度': info['lon'],
                        '來源': info.get('source', ''),
                        '更新時間': info.get('updated', '')
                    })
                
                st.dataframe(pd.DataFrame(db_display), use_container_width=True)
                st.caption(f"顯示 {len(filtered_db)} / {len(db)} 筆")
            
            # 刪除功能
            with st.expander("🗑️ 刪除學校"):
                del_name = st.selectbox(
                    "選擇要刪除的學校",
                    sorted(db.keys())
                )
                if st.button("刪除", type="secondary"):
                    if del_name in db:
                        del db[del_name]
                        save_database(db)
                        st.success(f"已刪除：{del_name}")
                        st.rerun()
    
    # ========================================
    # Tab 5: 進階工具
    # ========================================
    with tab5:
        st.header("🔧 進階工具")
        
        # 批次查詢
        st.subheader("📝 批次貼上查詢")
        batch_text = st.text_area(
            "貼上學校名稱（每行一個）",
            height=150,
            placeholder="臺北市大安國小\n新北市板橋國小\n桃園市中壢國小"
        )
        
        if st.button("🔍 批次查詢") and batch_text.strip():
            schools = [s.strip() for s in batch_text.strip().split('\n') if s.strip()]
            
            db = load_database()
            engine = GeocodingEngine()
            
            progress = st.progress(0)
            results = []
            
            for i, school in enumerate(schools):
                result = engine.geocode_school(school, db)
                results.append(result)
                progress.progress((i + 1) / len(schools))
            
            save_database(db)
            save_search_log(engine.search_log)
            
            # 顯示結果
            result_data = []
            for r in results:
                result_data.append({
                    '學校': r['name'],
                    '緯度': r['lat'],
                    '經度': r['lon'],
                    '來源': r['source'],
                    '狀態': '✅' if r['success'] else '❌'
                })
            
            df_batch = pd.DataFrame(result_data)
            st.dataframe(df_batch, use_container_width=True)
            
            # 下載
            st.download_button(
                "📥 下載結果 (Excel)",
                data=df_to_excel_bytes(df_batch),
                file_name=f"批次查詢結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        st.divider()
        
        # 搜尋記錄
        st.subheader("📜 搜尋記錄")
        search_log = load_search_log()
        if search_log:
            st.write(f"共 {len(search_log)} 筆記錄")
            
            df_log = pd.DataFrame(search_log[-50:][::-1])
            st.dataframe(df_log, use_container_width=True)
            
            # 統計
            success_count = sum(1 for l in search_log if l.get('success'))
            fail_count = len(search_log) - success_count
            st.write(f"成功率：{success_count}/{len(search_log)} ({success_count/len(search_log)*100:.1f}%)")
        else:
            st.info("尚無搜尋記錄")
        
        st.divider()
        
        # 資料庫健康檢查
        st.subheader("🏥 資料庫健康檢查")
        if st.button("執行檢查"):
            db = load_database()
            issues = []
            
            for name, info in db.items():
                lat = info.get('lat')
                lon = info.get('lon')
                
                if lat is None or lon is None:
                    issues.append(f"❌ {name}: 缺少座標")
                elif not (21.5 < lat < 25.5):
                    issues.append(f"⚠️ {name}: 緯度 {lat} 超出台灣範圍")
                elif not (119.5 < lon < 122.5):
                    issues.append(f"⚠️ {name}: 經度 {lon} 超出台灣範圍")
            
            if issues:
                st.warning(f"發現 {len(issues)} 個問題：")
                for issue in issues:
                    st.write(issue)
            else:
                st.success(f"✅ 資料庫健康！共 {len(db)} 筆資料均正常")

if __name__ == '__main__':
    main()
