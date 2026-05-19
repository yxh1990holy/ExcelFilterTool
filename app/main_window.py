from PySide6.QtWidgets import (
    QMainWindow, QWidget, QTabWidget, QApplication,
    QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QPushButton, QGroupBox,
    QProgressBar, QMessageBox
)
from PySide6.QtCore import Qt

from app.worker import DataProcessWorker
from app.handlers import FileHandler, TabHandler, FilterHandler, ExportHandler


class ExcelFilterWindow(QMainWindow):
    """Excel筛选处理工具主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel数据处理工具")
        self.setMinimumSize(1000, 700)
        self.center()
        
        # 数据状态
        self.file_path = None
        self.sheet_tabs = {}
        self.current_sheet_name = None
        self.sheets_data_cache = None
        self.sheet_loaded = {}
        self._loading_tab = False
        self.worker = None
        
        # 筛选条件相关
        self.filter_rows = []
        self.row1_filters = []
        self.row2_filters = []
        self.row1_layout = None
        self.row2_layout = None
        
        # 初始化处理器
        self.file_handler = FileHandler(self)
        self.tab_handler = TabHandler(self)
        self.filter_handler = FilterHandler(self)
        self.export_handler = ExportHandler(self)
        
        # 连接文件处理信号
        self.file_handler.file_selected.connect(self.on_file_selected)
        
        self.setup_ui()
        self.connect_signals()
        self.apply_style()

    def center(self):
        """将窗口居中显示"""
        screen = QApplication.primaryScreen().availableGeometry()
        window = self.frameGeometry()
        x = (screen.width() - window.width()) // 2
        y = (screen.height() - window.height()) // 2
        self.move(x, y)

    def setup_ui(self):
        """构建界面布局"""
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # 文件选择区域
        file_layout = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("padding: 8px; background: #f0f0f0;border-radius: 4px;")
        self.select_btn = QPushButton("📂 选择Excel文件")
        self.select_btn.setFixedWidth(140)
        file_layout.addWidget(QLabel("当前文件："))
        file_layout.addWidget(self.file_label, 1)
        file_layout.addWidget(self.select_btn)
        main_layout.addLayout(file_layout)

        # 筛选条件区域
        filter_group = QGroupBox("筛选条件（所有条件同时生效，AND 逻辑）")
        filter_layout = QFormLayout(filter_group)
        filter_layout.setContentsMargins(10, 15, 10, 10)
        filter_layout.setSpacing(8)

        # 创建2行容器
        self.row1_widget = QWidget()
        self.row1_layout = QHBoxLayout(self.row1_widget)
        self.row1_layout.setContentsMargins(0, 0, 0, 0)
        self.row1_layout.setSpacing(15)
        self.row1_layout.setAlignment(Qt.AlignLeft)

        self.row2_widget = QWidget()
        self.row2_layout = QHBoxLayout(self.row2_widget)
        self.row2_layout.setContentsMargins(0, 0, 0, 0)
        self.row2_layout.setSpacing(15)
        self.row2_layout.setAlignment(Qt.AlignLeft)

        filter_layout.addRow(self.row1_widget)
        filter_layout.addRow(self.row2_widget)

        # 按钮行
        btn_layout = QHBoxLayout()
        self.add_filter_btn = QPushButton("➕ 添加筛选条件")
        self.add_filter_btn.setEnabled(False)
        self.apply_filter_btn = QPushButton("🔍 应用筛选")
        self.apply_filter_btn.setEnabled(False)
        self.reset_filter_btn = QPushButton("🔄 重置筛选")
        self.reset_filter_btn.setEnabled(False)
        self.clear_all_btn = QPushButton("🗑️ 清空所有条件")
        self.clear_all_btn.setEnabled(False)
        self.export_current_btn = QPushButton("💾 导出结果")
        self.export_current_btn.setEnabled(False)
        self.export_all_btn = QPushButton("📦 导出所有工作表")
        self.export_all_btn.setEnabled(False)
        
        btn_layout.setContentsMargins(0, 10, 0, 5)
        btn_layout.setSpacing(10)
        btn_layout.addWidget(self.add_filter_btn)
        btn_layout.addWidget(self.apply_filter_btn)
        btn_layout.addWidget(self.reset_filter_btn)
        btn_layout.addWidget(self.clear_all_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.export_current_btn)
        btn_layout.addWidget(self.export_all_btn)
        filter_layout.addRow(btn_layout)

        main_layout.addWidget(filter_group)

        # 状态栏
        self.status_label = QLabel("就绪")
        self.status_label.setAlignment(Qt.AlignRight)
        self.status_label.setStyleSheet("padding: 5px; color: #666;")
        main_layout.addWidget(self.status_label)

        # 标签页区域
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(False)
        self.tab_widget.setMovable(True)
        self.tab_widget.currentChanged.connect(self.tab_handler.on_tab_changed)
        main_layout.addWidget(self.tab_widget, 1)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # 初始化筛选条件
        self.init_filter_rows()

    def init_filter_rows(self):
        """初始化筛选条件(默认1个)"""
        columns = self.tab_handler.get_current_columns()
        from app.filter_row_widget import FilterRowWidget
        row1_filter = FilterRowWidget(columns, show_delete=False)
        row1_filter.set_index(1)
        self.filter_rows.append(row1_filter)
        self.row1_layout.addWidget(row1_filter)
        self.row1_filters = [row1_filter]
        self.row2_filters = []

    def connect_signals(self):
        """连接信号槽"""
        self.select_btn.clicked.connect(self.file_handler.select_file)
        self.add_filter_btn.clicked.connect(self.add_filter_row)
        self.apply_filter_btn.clicked.connect(self.filter_handler.apply_filters)
        self.reset_filter_btn.clicked.connect(self.filter_handler.reset_filters)
        self.clear_all_btn.clicked.connect(self.filter_handler.clear_all_filters)
        self.export_current_btn.clicked.connect(self.export_handler.export_current)
        self.export_all_btn.clicked.connect(self.export_handler.export_all)

    def on_file_selected(self, file_path):
        """文件选择完成回调"""
        self.file_path = file_path
        self.load_all_sheets()

    def load_all_sheets(self):
        """加载Excel文件中的所有工作表"""
        if not self.file_path:
            return
        
        self.cleanup_worker()
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.setEnabled(False)
        self.status_label.setText("正在读取工作表信息...")

        self.worker = DataProcessWorker("load_all_sheets", self.file_path)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_all_sheets_loaded)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_all_sheets_loaded(self, sheets_data):
        """所有工作表加载完成"""
        self.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        if not sheets_data:
            QMessageBox.critical(self, "错误", "没有加载到任何工作表数据")
            return
        
        self.sheets_data_cache = sheets_data
        self._loading_tab = False
        
        # 使用 TabHandler 加载标签页
        count = self.tab_handler.load_all_tabs_lazy(sheets_data)
        
        # 启用按钮
        for btn in [self.add_filter_btn, self.apply_filter_btn, 
                    self.reset_filter_btn, self.clear_all_btn,
                    self.export_current_btn, self.export_all_btn]:
            btn.setEnabled(True)
        
        self.status_label.setText(f"已加载 {count} 个工作表（当前: {self.current_sheet_name}）")
        QMessageBox.information(self, "加载完成", 
            f"成功加载 {count} 个工作表\n"
            f"总数据行数: {sum(len(df) for df in sheets_data.values())}\n\n"
            f"💡 提示：点击其他标签页时加载对应数据")
    
    def add_filter_row(self):
        """添加新的筛选条件"""
        if len(self.filter_rows) >= 4:
            QMessageBox.warning(self, "提示", "最多支持4个筛选条件")
            return
        
        from app.filter_row_widget import FilterRowWidget
        columns = self.tab_handler.get_current_columns()
        new_filter = FilterRowWidget(columns, show_delete=True)
        new_filter.deleted.connect(self.remove_filter_row)

        if len(self.row1_filters) < 2:
            self.row1_layout.addWidget(new_filter)
            self.row1_filters.append(new_filter)
        else:
            self.row2_layout.addWidget(new_filter)
            self.row2_filters.append(new_filter)
        
        self.filter_rows.append(new_filter)
        self.update_filter_indices()
        self.status_label.setText(f"已添加筛选条件，当前共 {len(self.filter_rows)} 个")

    def remove_filter_row(self, filter_row):
        """删除筛选条件"""
        if len(self.filter_rows) <= 1:
            QMessageBox.warning(self, "提示", "至少保留一个筛选条件")
            return
        
        if filter_row in self.row1_filters:
            self.row1_layout.removeWidget(filter_row)
            self.row1_filters.remove(filter_row)
        elif filter_row in self.row2_filters:
            self.row2_layout.removeWidget(filter_row)
            self.row2_filters.remove(filter_row)
        
        self.filter_rows.remove(filter_row)
        filter_row.deleteLater()
        self.rearrange_filters()
        self.update_filter_indices()
        self.status_label.setText(f"已删除筛选条件，当前共 {len(self.filter_rows)} 个")

    def rearrange_filters(self):
        """重新整理筛选条件布局"""
        all_filters = self.row1_filters + self.row2_filters
        
        for f in self.row1_filters:
            self.row1_layout.removeWidget(f)
        for f in self.row2_filters:
            self.row2_layout.removeWidget(f)
        
        self.row1_filters.clear()
        self.row2_filters.clear()
        
        for i, f in enumerate(all_filters):
            if i < 2:
                self.row1_layout.addWidget(f)
                self.row1_filters.append(f)
            else:
                self.row2_layout.addWidget(f)
                self.row2_filters.append(f)

    def update_filter_indices(self):
        """更新筛选条件序号"""
        all_filters = self.row1_filters + self.row2_filters
        for i, f in enumerate(all_filters):
            f.set_index(i + 1)

    def get_all_filters(self):
        """获取所有有效的筛选条件"""
        filters = []
        for row in self.filter_rows:
            if row.is_valid():
                filters.append(row.get_filter())
        return filters

    def update_all_filter_columns(self):
        """更新所有筛选条件的列下拉框"""
        columns = self.tab_handler.get_current_columns()
        for row in self.filter_rows:
            row.set_columns(columns)
    
    def on_progress(self, message: str):
        """进度更新"""
        self.status_label.setText(message)
        self.progress_bar.setFormat(message[:50])
        QApplication.processEvents()

    def on_error(self, error_msg: str):
        """错误处理"""
        self.setEnabled(True)
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "错误", error_msg)
        self.status_label.setText(f"错误: {error_msg}")

    def _parse_filter_value(self, value: str):
        """解析筛选值"""
        if "," in value:
            items = [item.strip() for item in value.split(",")]
            numeric_items = []
            for item in items:
                try:
                    numeric_items.append(float(item) if "." in item else int(item))
                except ValueError:
                    numeric_items.append(item)
            return numeric_items
        try:
            return float(value) if "." in value else int(value)
        except ValueError:
            return value

    def apply_style(self):
        """应用样式表"""
        self.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ccc;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QPushButton {
                padding: 6px 12px;
                background: #0078d4;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover { background: #106ebe; }
            QPushButton:disabled { background: #ccc; }
            QTableView {
                alternate-background-color: #f5f5f5;
                selection-background-color: #0078d4;
            }
            QTabWidget::pane {
                border: 1px solid #ccc;
                border-radius: 4px;
            }
            QTabBar::tab {
                padding: 8px 16px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: #0078d4;
                color: white;
            }
        """)

    def cleanup_worker(self):
        """清理工作线程"""
        if self.worker and self.worker.isRunning():
            self.worker.quit()
            self.worker.wait(250)
            self.worker.deleteLater()
            self.worker = None

    def closeEvent(self, event):
        """窗口关闭清理"""
        self._loading_tab = False
        self.cleanup_worker()
        event.accept()