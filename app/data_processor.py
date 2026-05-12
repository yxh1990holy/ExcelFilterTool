import pandas as pd
from typing import List, Dict, Any

class ExcelProcessor():
    """Excel数据处理器"""

    @staticmethod
    def get_sheet_names(file_path: str) -> List[str]:
        """获取Excel文件中所有工作表名称"""
        try:
            xl = pd.ExcelFile(file_path)
            return xl.sheet_names
        except Exception as e:
            raise Exception(f"读取工作表失败：{str(e)}")
        
    @staticmethod
    def load_sheet(file_path: str, sheet_name: str) -> pd.DataFrame:
        """加载指定工作表的数据"""
        try:
            df = pd.read_excel(file_path, sheet_name= sheet_name)
            return df
        except Exception as e:
            raise Exception(f"加载工作表{sheet_name}失败：{str(e)}")
        
    @staticmethod
    def load_all_sheets(file_path: str) -> Dict[str, pd.DataFrame]:
        """加载Excel文件中的所有工作表"""
        try:
            # 一次性读取所有工作表
            all_sheets = pd.read_excel(file_path, sheet_name=None)
            return all_sheets
        except Exception as e:
            raise Exception(f"加载所有工作表失败：{str(e)}")

    @staticmethod
    def filter_data(df: pd.DataFrame, column: str, operator: str, value: Any) -> pd.DataFrame:
        """
        根据条件过滤数据
        operator支持：==, !=, >, >=, <, <=, contains(字符串包含), in(列表内)
        """
        if operator == "==":
            return df[df[column] == value]
        elif operator == "!=":
            return df[df[column] != value]
        elif operator == ">":
            return df[df[column] > value]
        elif operator == ">=":
            return df[df[column] >= value]
        elif operator == "<":
            return df[df[column] < value]
        elif operator == "<=":
            return df[df[column] <= value]
        elif operator == "contains":
            return df[df[column].astype(str).str.contains(str(value), na=False)]
        elif operator == "in":
            if isinstance(value, list):
                return df[df[column].isin(value)]
            return df
        else:
            return df
        
    @staticmethod
    def filter_data_multi(df: pd.DataFrame, filters: List[Dict[str, Any]]) -> pd.DataFrame:
        """
        多条件筛选数据（所有条件用 AND 连接）
        
        Args:
            df: 原始数据
            filters: 筛选条件列表，每个条件包含 column, operator, value
        
        Returns:
            筛选后的DataFrame
        """
        if not filters:
            return df
        
        result_df = df.copy()
        
        for filter_item in filters:
            column = filter_item.get("column")
            operator = filter_item.get("operator")
            value = filter_item.get("value")
            if not column or not value:
                continue
            result_df = ExcelProcessor.single_filter(result_df, column, operator, value)
            if result_df.empty:
                break
        return result_df
    
    @staticmethod
    def single_filter(df: pd.DataFrame, column: str, operator: str, value: Any) -> pd.DataFrame:
        """单条件筛选"""
        try:
            if operator == "==":
                # 尝试转换为数字进行比较
                try:
                    num_value = float(value) if "." in str(value) else int(value)
                    return df[df[column] == num_value]
                except ValueError:
                    return df[df[column].astype(str) == str(value)]
            elif operator == "!=":
                try:
                    num_value = float(value) if "." in str(value) else int(value)
                    return df[df[column] != num_value]
                except ValueError:
                    return df[df[column].astype(str) != str(value)]
            elif operator == ">":
                return df[df[column] > float(value)]
            elif operator == ">=":
                return df[df[column] >= float(value)]
            elif operator == "<":
                return df[df[column] < float(value)]
            elif operator == "<=":
                return df[df[column] <= float(value)]
            elif operator == "contains":
                return df[df[column].astype(str).str.contains(str(value), na=False, case=False)]
            else:
                return df
        except Exception as e:
            print(f"筛选错误: {e}")
            return df
