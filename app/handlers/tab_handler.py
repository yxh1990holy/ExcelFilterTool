from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QMessageBox, QApplication
from PySide6.QtCore import QObject, Qt
from app.sheet_tab_widget import SheetTabWidget


class TabHandler(QObject):
    """标签页管理类"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
    
    def load_all_tabs_lazy(self, sheets_data):
        """懒加载所有标签页"""
        # 临时断开信号
        self.parent.tab_widget.blockSignals(True)
        
        # 清空现有标签页
        self.parent.tab_widget.clear()
        self.parent.sheet_tabs = {}
        self.parent.sheet_loaded = {}
        
        sheet_names = list(sheets_data.keys())
        
        # 第一个标签页立即加载
        first_sheet = sheet_names[0]
        first_df = sheets_data[first_sheet]
        first_tab = SheetTabWidget(first_sheet, first_df, self.parent)
        self.parent.tab_widget.addTab(first_tab, first_sheet)
        self.parent.sheet_tabs[first_sheet] = first_tab
        self.parent.sheet_loaded[first_sheet] = True
        
        # 其他标签页创建占位符
        for sheet_name in sheet_names[1:]:
            placeholder = self._create_placeholder_tab(sheet_name)
            self.parent.tab_widget.addTab(placeholder, sheet_name)
            self.parent.sheet_tabs[sheet_name] = None
            self.parent.sheet_loaded[sheet_name] = False
        
        self.parent.tab_widget.blockSignals(False)
        self.parent.tab_widget.setCurrentIndex(0)
        
        # 更新UI
        self.parent.current_sheet_name = first_sheet
        self.parent.update_all_filter_columns()
        
        return len(sheets_data)
    
    def _create_placeholder_tab(self, sheet_name):
        """创建占位标签页"""
        placeholder = QWidget()
        layout = QVBoxLayout(placeholder)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignCenter)
        
        loading_label = QLabel(f"📊 工作表 '{sheet_name}'\n\n点击此标签页加载数据...")
        loading_label.setObjectName("loading_label")
        loading_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(loading_label)
        
        return placeholder
    
    def on_tab_changed(self, index):
        """标签页切换处理"""
        if getattr(self.parent, '_loading_tab', False) or index < 0:
            return
        
        sheet_name = self.parent.tab_widget.tabText(index)
        
        # 检查是否需要加载
        if not self.parent.sheet_loaded.get(sheet_name, False):
            # 立即显示加载状态
            self.parent.status_label.setText(f"正在加载工作表 '{sheet_name}'...")
            QApplication.processEvents()

            # 加载数据
            self._load_tab_data(sheet_name, index)
        
        # 更新当前工作表信息
        current_widget = self.parent.tab_widget.widget(index)
        if isinstance(current_widget, SheetTabWidget):
            self.parent.current_sheet_name = current_widget.sheet_name
            self.parent.status_label.setText(f"当前工作表: {self.parent.current_sheet_name}")
            # 更新筛选条件的列名
            if hasattr(self.parent, 'update_all_filter_columns'):
                self.parent.update_all_filter_columns()

    def _load_tab_data(self, sheet_name, index):
        """加载标签页数据"""
        self.parent._loading_tab = True
        
        try:
            df = self.parent.sheets_data_cache.get(sheet_name)
            if df is not None:
                self.parent.tab_widget.blockSignals(True)
                
                new_tab = SheetTabWidget(sheet_name, df, self.parent)
                self.parent.tab_widget.removeTab(index)
                self.parent.tab_widget.insertTab(index, new_tab, sheet_name)
                
                self.parent.sheet_tabs[sheet_name] = new_tab
                self.parent.sheet_loaded[sheet_name] = True
                
                self.parent.tab_widget.blockSignals(False)
                self.parent.tab_widget.setCurrentIndex(index)
                self.parent.status_label.setText(f"已加载工作表 '{sheet_name}'")
            else:
                self.parent.status_label.setText(f"加载工作表 '{sheet_name}' 失败")
        except Exception as e:
            self.parent.status_label.setText(f"加载失败: {str(e)}")
            QMessageBox.critical(self.parent, "错误", f"加载工作表失败: {str(e)}")      
        finally:
            self.parent._loading_tab = False
    
    def get_current_tab(self):
        """获取当前活动的标签页"""
        idx = self.parent.tab_widget.currentIndex()
        if idx < 0:
            return None
        
        widget = self.parent.tab_widget.widget(idx)
        # 检查是否是真正的标签页
        if isinstance(widget, SheetTabWidget):
            return widget
        
        # 占位标签页提示
        sheet_name = self.parent.tab_widget.tabText(idx)
        from PySide6.QtWidgets import QMessageBox
        QMessageBox.information(
            self.parent, "提示",
            f"工作表 '{sheet_name}' 尚未加载\n\n请先点击该标签页加载数据。"
        )
        return None
    
    def get_current_columns(self):
        """获取当前工作表的列名"""
        if self.parent.current_sheet_name and self.parent.current_sheet_name in self.parent.sheet_tabs:
            tab = self.parent.sheet_tabs.get(self.parent.current_sheet_name)
            if tab and hasattr(tab, 'headers'):
                return tab.headers
        return []
 