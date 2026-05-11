from PySide6.QtCore import QAbstractTableModel, Qt, QModelIndex
import pandas as pd

class PandasModel(QAbstractTableModel):
    """将pandas DataFrame适配为QAbstractTableModel，用于QTableView显示"""

    def __init__(self, data: pd.DataFrame):
        super().__init__()
        self._data = data

    def rowCount(self, parent = QModelIndex()):
        return 0 if self._data is None else self._data.shape[0]
    
    def columnCount(self, parent = QModelIndex()):
        return 0 if self._data is None else self._data.shape[1]
    
    def data(self, index, role = Qt.DisplayRole):
        if not index.isValid() or self._data is None:
            return None
        
        if role == Qt.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]
            return str(value)
        
        return None
    
    def headerData(self, section, orientation, role = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        
        if orientation == Qt.Horizontal:
            return str(self._data.columns[section])
        else:
            return str(section+1)