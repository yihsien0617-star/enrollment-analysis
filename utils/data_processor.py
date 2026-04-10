import pandas as pd
import numpy as np


class DataProcessor:
    """資料處理與清洗模組"""

    # 台灣縣市名稱標準化對照表
    CITY_MAPPING = {
        "臺北市": "台北市", "臺中市": "台中市", "臺南市": "台南市",
        "臺東縣": "台東縣", "臺北縣": "新北市",
        "桃園縣": "桃園市",  # 2014年升格
    }

    @staticmethod
    def clean_applicant_data(df: pd.DataFrame) -> pd.DataFrame:
        """清理申請入學資料"""
        df = df.copy()

        # 欄位名稱清理（去除空白）
        df.columns = df.columns.str.strip()

        # 標準化縣市名稱
        if '畢業學校縣市' in df.columns:
            df['畢業學校縣市'] = df['畢業學校縣市'].str.strip()
            df['畢業學校縣市'] = df['畢業學校縣市'].replace(DataProcessor.CITY_MAPPING)

        if '住家縣市' in df.columns:
            df['住家縣市'] = df['住家縣市'].str.strip()
            df['住家縣市'] = df['住家縣市'].replace(DataProcessor.CITY_MAPPING)

        # 確保學年度為整數
        if '學年度' in df.columns:
            df['學年度'] = pd.to_numeric(df['學年度'], errors='coerce').astype('Int64')

        # 標準化科系名稱
        if '報考科系' in df.columns:
            df['報考科系'] = df['報考科系'].str.strip()

        # 標準化最終入學欄位
        if '最終入學' in df.columns:
            df['最終入學'] = df['最終入學'].str.strip()
            df['最終入學'] = df['最終入學'].replace({
                'Y': '是', 'y': '是', '1': '是', 'True': '是', 'true': '是',
                'N': '否', 'n': '否', '0': '否', 'False': '否', 'false': '否'
            })

        return df

    @staticmethod
    def clean_retention_data(df: pd.DataFrame) -> pd.DataFrame:
        """清理在學穩定度資料"""
        df = df.copy()
        df.columns = df.columns.str.strip()

        if '學年度' in df.columns:
            df['學年度'] = pd.to_numeric(df['學年度'], errors='coerce').astype('Int64')

        if '目前狀態' in df.columns:
            df['目前狀態'] = df['目前狀態'].str.strip()

        if '入學管道' in df.columns:
            df['入學管道'] = df['入學管道'].str.strip()

        return df

    @staticmethod
    def validate_applicant_data(df: pd.DataFrame) -> dict:
        """驗證申請入學資料欄位"""
        required_cols = ['學年度', '報考科系', '畢業學校', '畢業學校縣市', '階段狀態', '最終入學']
        missing = [col for col in required_cols if col not in df.columns]
        return {
            'valid': len(missing) == 0,
            'missing_columns': missing,
            'row_count': len(df),
            'columns': df.columns.tolist()
        }

    @staticmethod
    def validate_retention_data(df: pd.DataFrame) -> dict:
        """驗證在學穩定度資料欄位"""
        required_cols = ['學年度', '入學管道', '入學科系', '目前狀態']
        missing = [col for col in required_cols if col not in df.columns]
        return {
            'valid': len(missing) == 0,
            'missing_columns': missing,
            'row_count': len(df),
            'columns': df.columns.tolist()
        }

    @staticmethod
    def extract_city_from_address(address: str) -> str:
        """從地址提取縣市名稱"""
        if pd.isna(address):
            return "未知"

        cities = [
            "台北市", "新北市", "桃園市", "台中市", "台南市", "高雄市",
            "基隆市", "新竹市", "嘉義市", "新竹縣", "苗栗縣", "彰化縣",
            "南投縣", "雲林縣", "嘉義縣", "屏東縣", "宜蘭縣", "花蓮縣",
            "台東縣", "澎湖縣", "金門縣", "連江縣",
            "臺北市", "臺中市", "臺南市", "臺東縣",
        ]

        for city in cities:
            if city in str(address):
                return DataProcessor.CITY_MAPPING.get(city, city)

        return "未知"
