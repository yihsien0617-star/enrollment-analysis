import pandas as pd
import plotly.graph_objects as go


class FunnelAnalysis:
    """招生漏斗分析模組"""

    STAGE_ORDER = [
        "第一階段報名",
        "通過第一階段",
        "完成二階面試",
        "錄取",
        "已報到"
    ]

    @staticmethod
    def calculate_funnel(df: pd.DataFrame) -> pd.DataFrame:
        """計算漏斗各階段人數"""
        stage_map = {s: i for i, s in enumerate(FunnelAnalysis.STAGE_ORDER)}
        df_calc = df.copy()
        df_calc['stage_idx'] = df_calc['階段狀態'].map(stage_map)

        results = []
        total = len(df_calc)
        for i, stage in enumerate(FunnelAnalysis.STAGE_ORDER):
            count = len(df_calc[df_calc['stage_idx'] >= i])
            results.append({
                '階段': stage,
                '人數': count,
                '佔比(%)': round(count / total * 100, 1) if total > 0 else 0
            })

        return pd.DataFrame(results)

    @staticmethod
    def create_funnel_chart(funnel_df: pd.DataFrame, title: str = "") -> go.Figure:
        """建立漏斗圖"""
        fig = go.Figure(go.Funnel(
            y=funnel_df['階段'],
            x=funnel_df['人數'],
            textinfo="value+percent initial",
            marker=dict(
                color=['#667eea', '#7c8cf5', '#a3b1ff', '#48bb78', '#28a745']
            )
        ))
        fig.update_layout(title=title, height=450)
        return fig
