from PySide6.QtCore import QThread, Signal
from app.data_processor import ExcelProcessor
import pandas as pd
import os

class DataProcessWorker(QThread):
    """后台数据处理工作线程"""
    finished = Signal(object)   # 传递处理结果
    error = Signal(str)         # 传递错误信息
    progress = Signal(str)      # 传递进度消息

    def __init__(self, task_type: str, *args, **kwargs):
        super().__init__()
        self.task_type = task_type
        self.args = args
        self.kwargs = kwargs

    def run(self):
        """在线程中执行任务"""
        try:
            if self.task_type == "load_all_sheets":
                file_path = self.args[0]
                self.progress.emit(f"正在加载文件 {os.path.basename(file_path)}...")
                
                # 获取所有工作表名称
                sheet_names = ExcelProcessor.get_sheet_names(file_path)
                self.progress.emit(f"发现 {len(sheet_names)} 个工作表")
                
                # 加载所有工作表
                sheets_data = {}
                for i, sheet_name in enumerate(sheet_names):
                    self.progress.emit(f"正在加载工作表 {i+1}/{len(sheet_names)}: {sheet_name}")
                    df = ExcelProcessor.load_sheet(file_path, sheet_name)
                    sheets_data[sheet_name] = df
                self.finished.emit(sheets_data)
            elif self.task_type == "export_all":
                sheet_tabs = self.args[0]  # 字典 {sheet_name: SheetTabWidget}
                output_path = self.args[1]
                
                self.progress.emit(f"正在导出所有工作表到 {output_path}...")
                
                # 使用 pandas ExcelWriter 写入多个工作表
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    for sheet_name, tab in sheet_tabs.items():
                        self.progress.emit(f"正在写入工作表: {sheet_name}")
                        tab.current_df.to_excel(writer, sheet_name=sheet_name, index=False)
                self.finished.emit(output_path)
            elif self.task_type == "load_sheet":
                file_path = self.args[0]
                sheet_name = self.args[1]
                df = ExcelProcessor.load_sheet(file_path, sheet_name)
                self.finished.emit(df)
            elif self.task_type == "filter":
                # 筛选数据
                df = self.args[0]
                column = self.args[1]
                operator = self.args[2]
                value = self.args[3]
                result = ExcelProcessor.filter_data(df, column, operator, value)
                self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))