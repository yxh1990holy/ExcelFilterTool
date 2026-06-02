from PySide6.QtWidgets import QMessageBox, QApplication
from PySide6.QtCore import QObject
from app.data_processor import ExcelProcessor


class FilterHandler(QObject):
    """筛选处理类"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
    
    def get_all_filters(self):
        """获取所有有效筛选条件"""
        filters = []
        for filter in self.parent.filters:
            if filter.is_valid():
                filters.append(filter.get_filter())
        return filters
    
    def apply_filters(self):
        """应用所有筛选条件"""
        # 使用 tab_handler 获取当前标签页
        current_tab = self.parent.tab_handler.get_current_tab()
        if not current_tab:
            QMessageBox.warning(self.parent, "提示", "请先加载Excel文件")
            return
        
        filters = self.get_all_filters()
        if not filters:
            QMessageBox.warning(self.parent, "提示", "请至少设置一个有效的筛选条件")
            return
        
        self.parent.status_label.setText("正在应用筛选条件...")
        QApplication.processEvents()
        
        try:
            filtered_df = current_tab.original_df.copy()
            for f in filters:
                filtered_df = ExcelProcessor.filter_data(
                    filtered_df, f["column"], f["operator"],
                    self.parent._parse_filter_value(f["value"])
                )
            
            current_tab.current_df = filtered_df
            current_tab.update_preview()
            
            self.parent.status_label.setText(f"筛选完成，共应用 {len(filters)} 个条件")
            
        except Exception as e:
            QMessageBox.critical(self.parent, "错误", f"筛选失败: {str(e)}")
    
    def reset_filters(self):
        """重置所有筛选"""
        current_tab = self.parent.tab_handler.get_current_tab()
        if not current_tab:
            return
        
        for filter in self.parent.filters:
            filter.clear()
        
        current_tab.reset_filter()
        self.parent.status_label.setText("已重置所有筛选条件")
    
    def clear_all_filters(self):
        """清空所有筛选条件"""
        reply = QMessageBox.question(
            self.parent, "确认清空",
            "确定要清空所有筛选条件吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        
        for filter in self.parent.filters:
            filter.clear()
        
        self.parent.status_label.setText("已清空所有筛选条件")