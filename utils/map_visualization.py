import folium
import pandas as pd
import numpy as np


class MapVisualization:
    """地圖視覺化模組"""

    TAIWAN_COORDS = {
        "台北市": (25.0330, 121.5654),
        "新北市": (25.0120, 121.4650),
        "桃園市": (24.9936, 121.3010),
        "台中市": (24.1477, 120.6736),
        "台南市": (22.9998, 120.2269),
        "高雄市": (22.6273, 120.3014),
        "基隆市": (25.1276, 121.7392),
        "新竹市": (24.8138, 120.9675),
        "新竹縣": (24.8387, 121.0178),
        "苗栗縣": (24.5602, 120.8214),
        "彰化縣": (24.0518, 120.5161),
        "南投縣": (23.9610, 120.9718),
        "雲林縣": (23.7092, 120.4313),
        "嘉義市": (23.4801, 120.4491),
        "嘉義縣": (23.4518, 120.2551),
        "屏東縣": (22.5519, 120.5487),
        "宜蘭縣": (24.7570, 121.7533),
        "花蓮縣": (23.9872, 121.6016),
        "台東縣": (22.7583, 121.1444),
        "澎湖縣": (23.5711, 119.5793),
        "金門縣": (24.4493, 118.3767),
        "連江縣": (26.1505, 119.9499),
    }

    HWU_LOCATION = (23.0048, 120.2210)  # 中華醫事科技大學

    @staticmethod
    def create_distribution_map(city_data: pd.DataFrame) -> folium.Map:
        """
        建立學生分布地圖
        city_data 需包含：縣市、申請人數、入學人數
        """
        m = folium.Map(
            location=[23.5, 120.9],
            zoom_start=8,
            tiles='CartoDB positron'
        )

        # 標記學校位置
        folium.Marker(
            MapVisualization.HWU_LOCATION,
            popup="中華醫事科技大學",
            tooltip="中華醫事科技大學",
            icon=folium.Icon(color='red', icon='university', prefix='fa')
        ).add_to(m)

        if len(city_data) == 0:
            return m

        max_count = city_data['申請人數'].max()

        for _, row in city_data.iterrows():
            city = row['縣市']
            if city not in MapVisualization.TAIWAN_COORDS:
                continue

            lat, lon = MapVisualization.TAIWAN_COORDS[city]
            count = row['申請人數']
            enrolled = row.get('入學人數', 0)
            rate = (enrolled / count * 100) if count > 0 else 0

            radius = max(8, min(40, count / max_count * 40))

            if rate >= 40:
                color = '#28a745'
            elif rate >= 25:
                color = '#ffc107'
            else:
                color = '#dc3545'

            folium.CircleMarker(
                location=[lat, lon],
                radius=radius,
                popup=f"{city}<br>申請：{count}<br>入學：{enrolled}<br>轉換率：{rate:.1f}%",
                tooltip=f"{city}：{count}人",
                color=color,
                fill=True,
                fillColor=color,
                fillOpacity=0.6
            ).add_to(m)

        return m
