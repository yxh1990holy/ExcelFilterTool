from PySide6.QtWidgets import QFileDialog
from PySide6.QtCore import QObject, Signal
import os


class FileHandler(QObject):
    """文件处理类"""
    
    file_selected = Signal(str)  # 文件选择完成信号
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.file_path = None
    
    def select_file(self):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self.parent,
            "请选择文件",
            "",
            "Excel文件(*.xlsx *.xls);;所有文件(*)"
        )
        if file_path:
            self.file_path = file_path
            self.parent.file_label.setText(os.path.basename(file_path))
            self.file_selected.emit(file_path)
            return True
        return False
    
    def get_file_path(self):
        """获取当前文件路径"""
        return self.file_path