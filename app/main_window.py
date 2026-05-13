import os
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QTabWidget, QApplication,
    QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QPushButton, QGroupBox,
    QProgressBar, QFileDialog, QMessageBox
)
from PySide6.QtCore import Qt

from app.worker import DataProcessWorker
from app.data_processor import ExcelProcessor
from app.sheet_tab_widget import SheetTabWidget
from app.filter_row_widget import FilterRowWidget

import pandas as pd


class ExcelFilterWindow(QMainWindow):
    """Excel筛选处理工具主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel数据处理工具")
        self.setMinimumSize(1000, 600)
        self.center()
        
        # 数据状态
        self.file_path = None
        self.sheet_tabs = {}                # 存储每个标签页对象
        self.current_sheet_name = None      # 当前活动的工作表名称
        self.sheets_data_cache = None       # 缓存原始数据
        self.sheet_loaded = {}              # 记录已加载的工作表
        self._loading_tab = False           # 防重入锁

        self.worker = None
        
        # 筛选条件相关
        self.filter_rows = []               # 存储所有筛选条件行
        self.row1_filters = []
        self.row2_filters = []
        self.row1_layout = None
        self.row2_layout = None
        self.add_filter_btn = None
        self.apply_filter_btn = None
        self.reset_filter_btn = None
        self.clear_all_btn = None
        # self.info_label = None
        
        self.setup_ui()
        self.connect_signals()
        self.apply_style()

    def center(self):
        """将窗口居中显示"""
        # 获取当前屏幕的几何信息
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        
        # 获取窗口自身的大小
        window_geometry = self.frameGeometry()
        
        # 计算居中位置
        x = (screen_geometry.width() - window_geometry.width()) // 2
        y = (screen_geometry.height() - window_geometry.height()) // 2
        
        # 移动窗口到计算出的位置
        self.move(x, y)

    def setup_ui(self):
        """构建界面布局"""
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(15, 15, 15, 15)  # 设置外边距
        main_layout.setSpacing(10)

        # 上部分:文件选择区域
        file_layout = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("padding: 8px; background: #f0f0f0;border-radius: 4px;")
        self.select_btn = QPushButton("📂 选择Excel文件")
        self.select_btn.setFixedWidth(140)
        file_layout.addWidget(QLabel("当前文件："))
        file_layout.addWidget(self.file_label, 1)
        file_layout.addWidget(self.select_btn)
        main_layout.addLayout(file_layout)

        # 中部：条件筛选区(固定2行,最多4个条件)
        filter_group = QGroupBox("筛选条件（应用于当前活动工作表）")
        filter_layout = QFormLayout(filter_group)
        filter_layout.setContentsMargins(10, 15, 10, 10)  # 内边距
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

        # 信息标签
        # self.info_label = QLabel("原始数据行数：0 | 筛选后行数：0")
        # self.info_label.setAlignment(Qt.AlignLeft)
        # self.info_label.setStyleSheet("padding: 5px; color: #666;")
        # filter_layout.addWidget(self.info_label)

        main_layout.addWidget(filter_group)

         # ========== 中部：状态栏 ==========
        self.status_label = QLabel("就绪")
        self.status_label.setAlignment(Qt.AlignRight)
        self.status_label.setStyleSheet("padding: 5px; color: #666;")
        main_layout.addWidget(self.status_label)

        # ========== 核心区域：多工作表标签页 ==========
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(False)  # 不显示关闭按钮
        self.tab_widget.setMovable(True)  # 允许移动标签页顺序
        self.tab_widget.currentChanged.connect(self.on_tab_changed)
        main_layout.addWidget(self.tab_widget, 1)  # 拉伸因子为1，占据主要空间

        # 底部：进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # 初始化筛选条件列表 
        self.init_filter_rows()

    def init_filter_rows(self):
        """初始化筛选条件(默认1个)"""
        # 获取列名(如果已加载数据)
        columns = []
        if hasattr(self, 'current_sheet_name') and self.current_sheet_name:
            tab = self.sheet_tabs.get(self.current_sheet_name)
            if tab and hasattr(tab, "headers"):
                columns = tab.headers

        # 创建第一个条件(无删除按钮) 
        row1_filter = FilterRowWidget(columns, show_delete=False)
        row1_filter.set_index(1)
        self.filter_rows.append(row1_filter)
        self.row1_layout.addWidget(row1_filter)

        # 第一行占位(最多2个)
        self.row1_filters = [row1_filter]
        self.row2_filters = []

    def connect_signals(self):
        """连接信号槽"""
        self.select_btn.clicked.connect(self.on_select_file)
        self.add_filter_btn.clicked.connect(self.add_filter_row)
        self.apply_filter_btn.clicked.connect(self.on_apply_filters)
        self.reset_filter_btn.clicked.connect(self.on_reset_filters)
        self.clear_all_btn.clicked.connect(self.on_clear_all_filters)
        self.export_current_btn.clicked.connect(self.on_export_current)
        self.export_all_btn.clicked.connect(self.on_export_all)

    def add_filter_row(self):
        """添加新的筛选条件(最多4个,分行显示)"""
        if len(self.filter_rows) >= 4:
            QMessageBox.warning(self, "提示", "最多支持4个筛选条件")
            return 
        # 获取列表名
        columns = self.get_current_columns()

        # 创建新条件
        new_filter = FilterRowWidget(columns, show_delete=True)
        new_filter.deleted.connect(self.remove_filter_row)

        # 决定放到哪一行
        if len(self.row1_filters) < 2:
            # 第一行还有位置
            self.row1_layout.addWidget(new_filter)
            self.row1_filters.append(new_filter)
        else:
            # 放到第二行
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
        
        # 从布局中移除
        if filter_row in self.row1_filters:
            self.row1_layout.removeWidget(filter_row)
            self.row1_filters.remove(filter_row)
        elif filter_row in self.row2_filters:
            self.row2_layout.removeWidget(filter_row)
            self.row2_filters.remove(filter_row)
        
        # 从列表中移除
        self.filter_rows.remove(filter_row)
        filter_row.deleteLater()
        
        # 重新整理布局（将第二行的条件移到第一行如果第一行有空位）
        self.rearrange_filters()
        self.update_filter_indices()
        self.status_label.setText(f"已删除筛选条件，当前共 {len(self.filter_rows)} 个")

    def rearrange_filters(self):
        """重新整理筛选条件布局（保持第一行优先填满）"""
        all_filters = self.row1_filters + self.row2_filters
        
        # 清空两行
        for f in self.row1_filters:
            self.row1_layout.removeWidget(f)
        for f in self.row2_filters:
            self.row2_layout.removeWidget(f)
        
        self.row1_filters.clear()
        self.row2_filters.clear()
        
        # 重新分配：第一行最多2个
        for i, f in enumerate(all_filters):
            if i < 2:
                self.row1_layout.addWidget(f)
                self.row1_filters.append(f)
            else:
                self.row2_layout.addWidget(f)
                self.row2_filters.append(f)

    def update_filter_indices(self):
        """更新所有筛选条件的序号"""
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

    def get_current_columns(self) -> list:
        """获取当前工作表的列名"""
        # 从当前标签页获取
        if hasattr(self, 'current_sheet_name') and self.current_sheet_name:
            tab = self.sheet_tabs.get(self.current_sheet_name)
            if tab and hasattr(tab, 'headers'):
                return tab.headers
        
        # 从 current_sheet_df 获取
        if hasattr(self, 'current_sheet_df') and self.current_sheet_df is not None:
            return self.current_sheet_df.columns.tolist()
        
        return []

    def update_all_filter_columns(self):
        """更新所有筛选条件的列下拉框"""
        columns = self.get_current_columns()
        for row in self.filter_rows:
            row.set_columns(columns)

    def on_clear_all_filters(self):
        """清空所有筛选条件（保留第一个）"""
        reply = QMessageBox.question(
            self,
            "确认清空",
            "确定要清空所有筛选条件吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 清空所有输入框
        for row in self.filter_rows:
            row.clear()
        
        self.status_label.setText("已清空所有筛选条件")

    def on_select_file(self):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(self, "请选择文件", "", "Excel文件(*.xlsx *.xls);;所有文件(*)")
        if file_path:
            self.file_path = file_path
            self.file_label.setText(os.path.basename(file_path))
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

        # 创建工作线程加载所有工作表
        self.worker = DataProcessWorker("load_all_sheets", self.file_path)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_all_sheets_loaded)
        self.worker.error.connect(self.on_error)
        self.worker.error.connect(self.worker.deleteLater)
        self.worker.start()

    def on_all_sheets_loaded(self, sheets_data):
        """所有工作表加载完成 - 懒加载优化版"""
        self.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        # import time
        # start = time.perf_counter()
        
        if not sheets_data:
            QMessageBox.critical(self, "错误", "没有加载到任何工作表数据")
            return
        
        # 保存原始数据到缓存
        self.sheets_data_cache = sheets_data
        
        # 初始化防重入锁
        self._loading_tab = False
        
        # 临时断开信号，避免触发 on_tab_changed
        self.tab_widget.blockSignals(True)
        
        # 清空现有标签页
        self.tab_widget.clear()
        self.sheet_tabs = {}
        self.sheet_loaded = {}
        
        # 获取所有工作表名称
        sheet_names = list(sheets_data.keys())
        
        # 先创建第一个标签页（当前显示的）- 立即加载数据
        first_sheet = sheet_names[0]
        first_df = sheets_data[first_sheet]
        
        first_tab = SheetTabWidget(first_sheet, first_df, self)
        self.tab_widget.addTab(first_tab, first_sheet)
        self.sheet_tabs[first_sheet] = first_tab
        self.sheet_loaded[first_sheet] = True
        
        # 为其他工作表创建占位标签页
        for sheet_name in sheet_names[1:]:
            placeholder = QWidget()
            layout = QVBoxLayout(placeholder)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setAlignment(Qt.AlignCenter)
            
            # 添加加载提示
            loading_label = QLabel(f"📊 工作表 '{sheet_name}'\n\n点击此标签页加载数据...")
            loading_label.setAlignment(Qt.AlignCenter)
            loading_label.setStyleSheet("font-size: 14px; color: #666; padding: 20px;")
            layout.addWidget(loading_label)
            
            self.tab_widget.addTab(placeholder, sheet_name)
            self.sheet_tabs[sheet_name] = None
            self.sheet_loaded[sheet_name] = False
        
        # 重新连接信号
        self.tab_widget.blockSignals(False)
        
        # 手动设置当前选中的标签页（不触发信号，因为信号已临时断开）
        self.tab_widget.setCurrentIndex(0)
        
        # 手动更新当前工作表的UI
        self.current_sheet_name = first_sheet
        # columns = first_df.columns.tolist()
        self.update_all_filter_columns()
        # self.column_combo.blockSignals(True)
        # self.column_combo.clear()
        # self.column_combo.addItems(columns)
        # self.column_combo.blockSignals(False)
        
        # 启用控件
        # self.column_combo.setEnabled(True)
        # self.operator_combo.setEnabled(True)
        # self.value_input.setEnabled(True)
        self.add_filter_btn.setEnabled(True)
        self.apply_filter_btn.setEnabled(True)
        self.reset_filter_btn.setEnabled(True)
        self.clear_all_btn.setEnabled(True)
        self.export_current_btn.setEnabled(True)
        self.export_all_btn.setEnabled(True)
        
        # 更新状态栏
        self.status_label.setText(f"已加载 {len(sheets_data)} 个工作表（当前: {first_sheet}）")
        
        # end = time.perf_counter()
        # print(f"界面加载耗时: {end - start:.6f} 秒")
        
        QMessageBox.information(self, 
            "加载完成", 
            f"成功加载 {len(sheets_data)} 个工作表\n"
            f"总数据行数: {sum(len(df) for df in sheets_data.values())}\n\n"
            f"💡 提示：点击其他标签页时加载对应数据"
        ) 

    def on_tab_changed(self, index):
        """标签页切换时的处理 - 支持懒加载（带防重入锁）"""
        # column_combo_value = self.column_combo.currentText()  # 切换前保存上一次的选值

        # 防重入锁：避免在处理过程中被再次调用
        if hasattr(self, '_loading_tab') and self._loading_tab:
            # print("正在加载中，跳过重复调用")
            return
        
        if index < 0:
            return
        
        sheet_name = self.tab_widget.tabText(index)
        # print(f"切换到工作表: {sheet_name}, 已加载: {self.sheet_loaded.get(sheet_name, False)}")
        
        # 检查是否已加载
        if not self.sheet_loaded.get(sheet_name, False):
            # 设置加载标志
            self._loading_tab = True
            
            try:
                self.status_label.setText(f"正在加载工作表 '{sheet_name}'...")
                QApplication.processEvents()
                
                # 从缓存中获取数据
                df = self.sheets_data_cache.get(sheet_name)
                
                if df is not None:
                    # import time
                    # start = time.perf_counter()
                    
                    # 临时断开信号，避免连锁反应
                    self.tab_widget.blockSignals(True)
                    
                    current_index = self.tab_widget.currentIndex()
                    
                    # 创建真正的标签页控件
                    new_tab = SheetTabWidget(sheet_name, df, self)
                    
                    # 移除占位标签页
                    self.tab_widget.removeTab(current_index)
                    # 插入新标签页
                    self.tab_widget.insertTab(current_index, new_tab, sheet_name)
                    
                    # 更新字典
                    self.sheet_tabs[sheet_name] = new_tab
                    self.sheet_loaded[sheet_name] = True
                    
                    # 重新连接信号并设置当前索引
                    self.tab_widget.blockSignals(False)
                    self.tab_widget.setCurrentIndex(current_index)
                    
                    # end = time.perf_counter()
                    # print(f"加载工作表 '{sheet_name}' 耗时: {end - start:.6f} 秒")
                    self.status_label.setText(f"已加载工作表 '{sheet_name}'")
                    
                    # 更新当前工作表的列选择框
                    self.current_sheet_name = sheet_name
                    # columns = df.columns.tolist()
                    # self.column_combo.blockSignals(True)
                    # self.column_combo.clear()
                    # self.column_combo.addItems(columns)                 
                    # if column_combo_value in columns:                       # 如果新的tab也有该值，则直接保留上次的选值
                    #     self.column_combo.setCurrentText(column_combo_value)
                    # self.column_combo.blockSignals(False)
                else:
                    self.status_label.setText(f"加载工作表 '{sheet_name}' 失败")
                    return
            finally:
                # 清除加载标志
                self._loading_tab = False
        
        # 更新当前工作表信息（已加载的情况）
        current_widget = self.tab_widget.widget(index)
        
        if isinstance(current_widget, SheetTabWidget):
            self.current_sheet_name = current_widget.sheet_name
            
            # 更新列选择下拉框
            # current_df = current_widget.original_df
            # if current_df is not None:
            #     self.column_combo.blockSignals(True)
            #     columns = current_df.columns.tolist()
            #     self.column_combo.clear()
            #     self.column_combo.addItems(columns)                 
            #     if column_combo_value in columns:                       # 如果新的tab也有该值，则直接保留上次的选值
            #         self.column_combo.setCurrentText(column_combo_value)
            #     self.column_combo.blockSignals(False)
            
            self.status_label.setText(f"当前工作表: {self.current_sheet_name}")
            self.update_all_filter_columns()

    def get_current_tab(self):
        """获取当前活动的标签页（只返回已加载的）"""
        current_index = self.tab_widget.currentIndex()
        if current_index < 0:
            return None
        
        current_widget = self.tab_widget.widget(current_index)
        
        # 检查是否是 SheetTabWidget 类型
        if isinstance(current_widget, SheetTabWidget):
            return current_widget
        else:
            # 占位标签页，提示用户需要先加载
            sheet_name = self.tab_widget.tabText(current_index)
            QMessageBox.information(
                self, 
                "提示", 
                f"工作表 '{sheet_name}' 尚未加载\n\n"
                f"请先点击该标签页加载数据，然后再导出。"
            )
            return None

    # def on_apply_filter(self):
    #     """应用筛选到当前工作表"""
    #     current_tab = self.get_current_tab()
        
    #     if not current_tab:
    #         return  # get_current_tab 已经显示了提示
        
    #     column = self.column_combo.currentText()
    #     operator = self.operator_combo.currentText()
    #     value = self.value_input.text().strip()
        
    #     if not column:
    #         QMessageBox.warning(self, "提示", "请选择筛选列")
    #         return
    #     if not value:
    #         QMessageBox.warning(self, "提示", "请输入筛选值")
    #         return
        
    #     # 解析筛选值
    #     processed_value = self._parse_filter_value(value)
        
    #     self.status_label.setText(f"正在筛选工作表 '{current_tab.sheet_name}'...")
    #     QApplication.processEvents()
        
    #     success = current_tab.apply_filter(column, operator, processed_value)
        
    #     if success:
    #         self.status_label.setText(f"筛选完成，当前显示 {len(current_tab.current_df)} 行")
    #     else:
    #         self.status_label.setText("筛选失败")
    #         QMessageBox.warning(self, "错误", "筛选条件无效")

    # def on_reset_filter(self):
    #     """重置当前工作表的筛选"""
    #     current_tab = self.get_current_tab()
        
    #     if not current_tab:
    #         return
        
    #     current_tab.reset_filter()
    #     self.value_input.clear()
    #     self.status_label.setText(f"已重置工作表 '{current_tab.sheet_name}' 的筛选")

    def on_apply_filters(self):
        """应用所有筛选条件"""
        if not self.get_current_tab():
            QMessageBox.warning(self, "提示", "请先加载Excel文件")
            return
        
        filters = self.get_all_filters()
        
        if not filters:
            QMessageBox.warning(self, "提示", "请至少设置一个有效的筛选条件")
            return
        
        current_tab = self.get_current_tab()
        if not current_tab:
            return
        
        self.status_label.setText("正在应用筛选条件...")
        QApplication.processEvents()
        
        try:
            # 逐个应用筛选（AND逻辑）
            filtered_df = current_tab.original_df.copy()
            for f in filters:
                filtered_df = ExcelProcessor.filter_data(
                    filtered_df, 
                    f["column"], 
                    f["operator"], 
                    self._parse_filter_value(f["value"])
                )
            
            current_tab.current_df = filtered_df
            current_tab.update_preview()
            
            # original_rows = len(current_tab.original_df)
            # filtered_rows = len(filtered_df)
            # self.info_label.setText(f"原始数据行数：{original_rows} | 筛选后行数：{filtered_rows}")
            self.status_label.setText(f"筛选完成，共应用 {len(filters)} 个条件")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"筛选失败: {str(e)}")

    def on_reset_filters(self):
        """重置所有筛选"""
        current_tab = self.get_current_tab()
        if not current_tab:
            return
        
        # 清空所有输入框
        for row in self.filter_rows:
            row.clear()
        
        # 重置数据
        current_tab.reset_filter()
        
        # original_rows = len(current_tab.original_df)
        # self.info_label.setText(f"原始数据行数：{original_rows} | 筛选后行数：{original_rows}")
        self.status_label.setText("已重置所有筛选条件")

    def on_clear_all_filters(self):
        """清空所有筛选条件（保留第一个）"""
        reply = QMessageBox.question(
            self,
            "确认清空",
            "确定要清空所有筛选条件吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 清空所有输入框
        for row in self.filter_rows:
            row.clear()
        
        self.status_label.setText("已清空所有筛选条件")


    def on_export_current(self):
        """导出当前工作表"""
        current_tab = self.get_current_tab()
        
        if not current_tab:
            return
        
        # 获取当前数据
        if hasattr(current_tab, 'current_df') and isinstance(current_tab.current_df, pd.DataFrame):
            current_df = current_tab.current_df
        elif hasattr(current_tab, 'original_df') and isinstance(current_tab.original_df, pd.DataFrame):
            current_df = current_tab.original_df
        else:
            QMessageBox.warning(self, "提示", "当前工作表没有有效数据")
            return
        
        if current_df.empty:
            QMessageBox.warning(self, "提示", "当前工作表没有数据可导出")
            return
        
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            f"保存工作表 '{self.current_sheet_name}'",
            f"{self.current_sheet_name}_处理后.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if output_path:
            try:
                # 直接导出
                current_df.to_excel(output_path, sheet_name=self.current_sheet_name, index=False)
                QMessageBox.information(self, "成功", f"已导出到: {output_path}")
                self.status_label.setText(f"已导出: {output_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")

    def on_export_all(self):
        """导出所有工作表"""
        if not self.sheet_tabs:
            QMessageBox.warning(self, "提示", "没有可导出的工作表")
            return
        
        # 先找出所有未加载的工作表
        unloaded_sheets = [name for name, tab in self.sheet_tabs.items() if tab is None]
        
        if unloaded_sheets:
            # 询问用户是否要加载所有未加载的工作表
            reply = QMessageBox.question(
                self,
                "确认导出",
                f"还有 {len(unloaded_sheets)} 个工作表未加载：\n"
                f"{', '.join(unloaded_sheets[:5])}{'...' if len(unloaded_sheets) > 5 else ''}\n\n"
                f"导出前需要加载所有工作表，这可能需要一些时间。\n"
                f"是否继续？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            
            if reply != QMessageBox.Yes:
                return
            
            # 开始加载所有未加载的工作表
            self.status_label.setText(f"正在加载 {len(unloaded_sheets)} 个工作表...")
            QApplication.processEvents()
            
            # 逐个加载未加载的工作表
            for sheet_name in unloaded_sheets:
                if sheet_name in self.sheets_data_cache:
                    df = self.sheets_data_cache[sheet_name]
                    
                    # 确保 df 是 DataFrame
                    if isinstance(df, pd.DataFrame):
                        # 找到该工作表的标签页索引
                        for i in range(self.tab_widget.count()):
                            if self.tab_widget.tabText(i) == sheet_name:
                                # 创建真正的标签页控件
                                new_tab = SheetTabWidget(sheet_name, df, self)
                                
                                # 替换占位标签页
                                self.tab_widget.blockSignals(True)
                                self.tab_widget.removeTab(i)
                                self.tab_widget.insertTab(i, new_tab, sheet_name)
                                self.tab_widget.blockSignals(False)
                                
                                self.sheet_tabs[sheet_name] = new_tab
                                self.sheet_loaded[sheet_name] = True
                                break
                    else:
                        QMessageBox.warning(self, "警告", f"工作表 '{sheet_name}' 的数据不是 DataFrame，而是 {type(df)}")
                        # print(f"警告：工作表 '{sheet_name}' 的数据不是 DataFrame，而是 {type(df)}")
                    
                    QApplication.processEvents()
            
            self.status_label.setText("所有工作表加载完成，准备导出...")
            QApplication.processEvents()
        
        # 收集所有已加载的工作表数据
        sheets_data = {}
        
        for sheet_name, tab in self.sheet_tabs.items():
            if tab is not None:
                # 确保获取的是 DataFrame
                if hasattr(tab, 'current_df') and isinstance(tab.current_df, pd.DataFrame):
                    sheets_data[sheet_name] = tab.current_df
                elif hasattr(tab, 'original_df') and isinstance(tab.original_df, pd.DataFrame):
                    sheets_data[sheet_name] = tab.original_df
                else:
                    # print(f"警告：工作表 '{sheet_name}' 没有有效的 DataFrame")
                    QMessageBox.warning(self, "警告", f"工作表 '{sheet_name}' 没有有效数据，将被跳过")
        
        if not sheets_data:
            QMessageBox.warning(self, "提示", "没有可导出的数据")
            return
        
        # 确认导出
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存所有工作表",
            "所有工作表处理后.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if output_path:
            try:
                self.cleanup_worker()
                self.progress_bar.setVisible(True)
                self.progress_bar.setRange(0, 0)
                self.setEnabled(False)
                self.status_label.setText("正在导出所有工作表...")
                QApplication.processEvents()
                
                # 直接在当前线程导出（避免类型传递问题）
                success = self._export_all_sheets_sync(sheets_data, output_path)
                
                if success:
                    self.setEnabled(True)
                    self.progress_bar.setVisible(False)
                    QMessageBox.information(self, "成功", f"已导出所有工作表到: {output_path}")
                    self.status_label.setText(f"已导出所有工作表: {output_path}")
                else:
                    raise Exception("导出失败")
                
            except Exception as e:
                self.setEnabled(True)
                self.progress_bar.setVisible(False)
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")

    def _export_all_sheets_sync(self, sheets_data: dict, output_path: str) -> bool:
        """同步导出所有工作表（避免线程类型传递问题）"""
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in sheets_data.items():
                    if isinstance(df, pd.DataFrame) and not df.empty:
                        # 写入工作表
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        # print(f"已写入工作表: {sheet_name}, 行数: {len(df)}")
                    else:
                        QMessageBox.warning(self, "警告", f"工作表 '{sheet_name}' 数据无效，跳过")
                        # print(f"警告：工作表 '{sheet_name}' 数据无效，跳过")
            return True
        except Exception as e:
            print(f"导出失败: {e}")
            return False

    def on_export_all_finished(self, result):
        """所有工作表导出完成"""
        self.setEnabled(True)
        self.progress_bar.setVisible(False)
        if result:
            QMessageBox.information(self, "成功", f"已导出所有工作表到: {result}")
            self.status_label.setText(f"已导出所有工作表: {result}")

    def on_progress(self, message: str):
        """接收工作线程的进度消息"""
        # 更新状态栏
        self.status_label.setText(message)
        # 可选：也显示在进度条上
        self.progress_bar.setFormat(message[:50])  # 限制长度
        # 强制立即更新界面
        QApplication.processEvents()

    def on_error(self, error_msg: str):
        """错误处理"""
        self.setEnabled(True)
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "错误", error_msg)
        self.status_label.setText(f"错误: {error_msg}")

    def _parse_filter_value(self, value: str):
        """智能分析筛选值：尝试转换为数字，如果包含逗号则解析为列表"""
        if "," in value:
            items = [items.strip() for item in value.split(",")]
            # 尝试将每项转换为数字
            numeric_items = []
            for item in items:
                try:
                    numeric_items.append(float(item) if "." in item else int(item))
                except ValueError:
                    numeric_items.append(item)
            return numeric_items
        else:
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
            QPushButton:hover {
                background: #106ebe;
            }
            QPushButton:disabled {
                background: #ccc;
            }
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
        """清理旧的工作线程"""
        if hasattr(self, 'worker') and self.worker is not None:
            if self.worker.isRunning():
                self.worker.quit()
                self.worker.wait(250)  # 等待250毫秒
            self.worker.deleteLater()
            self.worker = None

    def closeEvent(self, event):
        """窗口关闭时的清理"""
        self._loading_tab = False
        self.cleanup_worker()
        event.accept()