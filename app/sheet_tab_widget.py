from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QTableView
from PySide6.QtCore import Qt
from app.data_processor import ExcelProcessor
from app.table_model import PandasModel
import pandas as pd

class SheetTabWidget(QWidget):
    """单个工作表的标签页组件"""

    def __init__(self, sheet_name: str, dataframe: pd.DataFrame, parent=None):
        super().__init__(parent)
        self.sheet_name = sheet_name
        self.original_df = dataframe.copy()     # 原始数据
        self.current_df = dataframe.copy()      # 当前显示的数据(筛选后)

        # 添加 headers 属性（从 dataframe 的列名获取）
        self.headers = dataframe.columns.tolist()  # 添加这一行

        self.setup_ui()
        self.update_preview()

    def setup_ui(self):
        """设置标签页的UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 信息标签
        self.info_label = QLabel(f"总行数: {len(self.original_df)} | 当前行数: {len(self.current_df)}")
        self.info_label.setStyleSheet("padding: 5px; background-color: #f5f5f5; border-radius: 3px;")
        layout.addWidget(self.info_label)

        # 表格视图
        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSortingEnabled(True)
        layout.addWidget(self.table_view)

    def update_preview(self):
        """刷新表格预览"""
        model = PandasModel(self.current_df)
        self.table_view.setModel(model)
        self.table_view.resizeColumnsToContents()

        # 更新标签信息标签
        original_rows = len(self.original_df)
        current_rows = len(self.current_df)
        self.info_label.setText(f"总行数: {original_rows} | 当前行数: {current_rows}")

        if current_rows < original_rows:
            self.info_label.setStyleSheet("padding: 5px; background-color: #fff3cd; border-radius: 3px;")
        else:
            self.info_label.setStyleSheet("padding: 5px; background-color: #f5f5f5; border-radius: 3px;")

    def apply_filter(self, column: str, operator: str, value):
        """应用筛选到当前标签页"""
        try:
            filtered_df = ExcelProcessor.filter_data(self.original_df, column, operator, value)
            self.current_df = filtered_df
            self.update_preview()
            return True
        except Exception as e:
            return False
    
    def reset_filter(self):
        """重置筛选"""
        self.current_df = self.original_df.copy()
        self.update_preview()
    
    def export_data(self, output_path: str):
        """导出当前标签页的数据"""
        self.current_df.to_excel(output_path, sheet_name=self.sheet_name, index=False)
