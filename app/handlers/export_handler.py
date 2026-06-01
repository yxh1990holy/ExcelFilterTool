from PySide6.QtWidgets import QFileDialog, QMessageBox, QApplication
from PySide6.QtCore import QObject, QStandardPaths
import pandas as pd


class ExportHandler(QObject):
    """导出处理类"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        # 跨平台获取桌面路径
        self.desk_path = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DesktopLocation)
    
    def export_current(self):
        """导出当前工作表"""
        current_tab = self.parent.tab_handler.get_current_tab()
        if not current_tab:
            return
        
        current_df = current_tab.current_df
        if current_df.empty:
            QMessageBox.warning(self.parent, "提示", "当前工作表没有数据可导出")
            return
        
        output_path, _ = QFileDialog.getSaveFileName(
            self.parent,
            f"保存工作表 '{self.parent.current_sheet_name}'",
            f"{self.desk_path}/{self.parent.current_sheet_name}_处理后.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if output_path:
            try:
                current_df.to_excel(output_path, sheet_name=self.parent.current_sheet_name, index=False)
                QMessageBox.information(self.parent, "成功", f"已导出到: {output_path}")
                self.parent.status_label.setText(f"已导出: {output_path}")
            except Exception as e:
                QMessageBox.critical(self.parent, "错误", f"导出失败: {str(e)}")
    
    def export_all(self):
        """导出所有工作表"""
        if not self.parent.sheet_tabs:
            QMessageBox.warning(self.parent, "提示", "没有可导出的工作表")
            return
        
        # 先处理未加载的工作表
        unloaded_sheets = [name for name, tab in self.parent.sheet_tabs.items() if tab is None]
        
        if unloaded_sheets:
            reply = QMessageBox.question(
                self.parent, "确认导出",
                f"还有 {len(unloaded_sheets)} 个工作表未加载，导出前需要加载，是否继续？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
            )
            if reply != QMessageBox.Yes:
                return
            
            self.parent.status_label.setText(f"正在加载 {len(unloaded_sheets)} 个工作表...")
            QApplication.processEvents()
            
            for sheet_name in unloaded_sheets:
                if sheet_name in self.parent.sheets_data_cache:
                    df = self.parent.sheets_data_cache[sheet_name]
                    if isinstance(df, pd.DataFrame):
                        for i in range(self.parent.tab_widget.count()):
                            if self.parent.tab_widget.tabText(i) == sheet_name:
                                from app.sheet_tab_widget import SheetTabWidget
                                new_tab = SheetTabWidget(sheet_name, df, self.parent)
                                self.parent.tab_widget.blockSignals(True)
                                self.parent.tab_widget.removeTab(i)
                                self.parent.tab_widget.insertTab(i, new_tab, sheet_name)
                                self.parent.tab_widget.blockSignals(False)
                                self.parent.sheet_tabs[sheet_name] = new_tab
                                self.parent.sheet_loaded[sheet_name] = True
                                break
                    QApplication.processEvents()
        
        # 收集数据
        sheets_data = {}
        for sheet_name, tab in self.parent.sheet_tabs.items():
            if tab is not None and hasattr(tab, 'current_df') and isinstance(tab.current_df, pd.DataFrame):
                if not tab.current_df.empty:
                    sheets_data[sheet_name] = tab.current_df
        
        if not sheets_data:
            QMessageBox.warning(self.parent, "提示", "没有可导出的数据")
            return
        
        output_path, _ = QFileDialog.getSaveFileName(
            self.parent, "保存所有工作表",
            f"{self.desk_path}/所有工作表处理后.xlsx", "Excel文件 (*.xlsx)"
        )
        
        if output_path:
            try:
                self.parent.progress_bar.setVisible(True)
                self.parent.progress_bar.setRange(0, 0)
                self.parent.setEnabled(False)
                
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    for sheet_name, df in sheets_data.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                self.parent.setEnabled(True)
                self.parent.progress_bar.setVisible(False)
                QMessageBox.information(self.parent, "成功", f"已导出到: {output_path}")
                
            except Exception as e:
                self.parent.setEnabled(True)
                self.parent.progress_bar.setVisible(False)
                QMessageBox.critical(self.parent, "错误", f"导出失败: {str(e)}")