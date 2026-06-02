from PySide6.QtWidgets import QWidget, QHBoxLayout, QComboBox, QLineEdit, QPushButton, QLabel, QStyle
from PySide6.QtCore import Signal


class FilterWidget(QWidget):
    """单行筛选条件组件"""
    
    # 当点击删除按钮时发出信号
    deleted = Signal(object)
    
    def __init__(self, columns: list, show_delete: bool = True, parent=None):
        super().__init__(parent)
        self.columns = columns if columns else []   # 处理空列表
        self.show_delete = show_delete
        self.setup_ui()
    
    def setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 2, 0, 2)
        layout.setSpacing(8)

        # 序号标签
        self.index_label = QLabel()
        self.index_label.setFixedWidth(12)
        self.index_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.index_label)
        
        # 列选择下拉框
        self.column_combo = QComboBox()
        if self.columns:
            self.column_combo.addItems(self.columns)
        else:
            self.column_combo.addItem("请先加载数据")
        self.column_combo.setMinimumWidth(130)
        layout.addWidget(self.column_combo)
        
        # 运算符下拉框
        self.operator_combo = QComboBox()           
        self.operator_combo.setEnabled(False)
        self.operator_combo.addItems(["==", "!=", ">", ">=", "<", "<=", "contains"])
        self.operator_combo.setMinimumWidth(80)
        layout.addWidget(self.operator_combo)
        
        # 筛选值输入框
        self.value_input = QLineEdit()
        self.value_input.setEnabled(False)
        self.value_input.setPlaceholderText("请输入筛选值（多个值用逗号分隔）")
        self.value_input.setMinimumWidth(100)
        layout.addWidget(self.value_input)
        
        # 删除按钮
        self.delete_btn = QPushButton(icon=self.style().standardIcon(QStyle.SP_MessageBoxCritical))
        self.delete_btn.setObjectName("delete_filter_btn")
        self.delete_btn.setFixedSize(28, 28)
        self.delete_btn.setVisible(self.show_delete)
        self.delete_btn.clicked.connect(lambda: self.deleted.emit(self))
        layout.addWidget(self.delete_btn)
        
        # 添加伸缩空间
        layout.addStretch()

        if self.columns:
            self.column_combo.setEnabled(True)
            self.operator_combo.setEnabled(True)
            self.value_input.setEnabled(True)
        else:
            self.column_combo.setEnabled(False)
            self.operator_combo.setEnabled(False)
            self.value_input.setEnabled(False)
    
    def set_index(self, idx: int):
        """设置序号"""
        self.index_label.setText(f"{idx}.")
        self.index_label.setStyleSheet("font-weight: bold;")
    
    def get_filter(self) -> dict:
        """获取当前筛选条件"""
        return {
            "column": self.column_combo.currentText(),
            "operator": self.operator_combo.currentText(),
            "value": self.value_input.text().strip()
        }
    
    def set_columns(self, columns: list):
        """更新列选择下拉框"""
        current = self.column_combo.currentText()
        self.column_combo.clear()
        self.column_combo.addItems(columns)
        self.column_combo.setEnabled(True)
        self.operator_combo.setEnabled(True)
        self.value_input.setEnabled(True)
        if current in columns:
            self.column_combo.setCurrentText(current)
        else:
            self.value_input.clear()
    
    def is_valid(self) -> bool:
        """检查筛选条件是否有效"""
        column = self.column_combo.currentText()
        value = self.value_input.text().strip()
        return bool(column and value and column != "请先加载数据")
    
    def clear(self):
        """清空输入并恢复下拉选为默认值"""
        self.column_combo.setCurrentIndex(0)
        self.operator_combo.setCurrentIndex(0)
        self.value_input.clear()