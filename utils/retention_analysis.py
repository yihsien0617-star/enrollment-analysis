import pandas as pd
import plotly.graph_objects as go
import plotly.express as px


class RetentionAnalysis:
    """在學穩定度分析模組"""

    @staticmethod
    def calculate_retention_by_channel(df: pd.DataFrame) -> pd.DataFrame:
        """按入學管道計算穩定度"""
        stats = df.groupby('入學管道').agg(
            總人數=('學生編號', 'count'),
            在學=('目前狀態', lambda x: (x == '在學').sum()),
            畢業=('目前狀態', lambda x: (x == '畢業').sum()),
            休學=('目前狀態', lambda x: (x == '休學').sum()),
            退學=('目前狀態', lambda x: (x == '退學').sum()),
        ).reset_index()

        stats['穩定率(%)'] = ((stats['在學'] + stats['畢業']) / stats['總人數'] * 100).round(1)
        stats['休學率(%)'] = (stats['休學'] / stats['總人數'] * 100).round(1)
        stats['退學率(%)'] = (stats['退學'] / stats['總人數'] * 100).round(1)

        return stats.sort_values('穩定率(%)', ascending=False)

    @staticmethod
    def calculate_retention_by_dept(df: pd.DataFrame) -> pd.DataFrame:
        """按科系計算穩定度"""
        stats = df.groupby('入學科系').agg(
            總人數=('學生編號', 'count'),
            休退學=('目前狀態', lambda x: x.isin(['休學', '退學']).sum()),
        ).reset_index()

        stats['休退學率(%)'] = (stats['休退學'] / stats['總人數'] * 100).round(1)
        return stats.sort_values('休退學率(%)')
