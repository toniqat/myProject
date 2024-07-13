import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QVBoxLayout, QWidget, QFileDialog, QMessageBox
from PyQt5.QtCore import QAbstractTableModel, Qt
from fileinput import filename
from PyQt5 import uic

import os

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

class PandasModel(QAbstractTableModel):
    def __init__(self, df=pd.DataFrame(), parent=None):
        QAbstractTableModel.__init__(self, parent)
        self._data = df

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return self._data.columns[section]
            else:
                return self._data.index[section]
        return None

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel to QTableView")
        uic.loadUi(resource_path('untitled.ui'), self)

        self.actionOpen.triggered.connect(self.openFileDialog)
        #self.setCentralWidget(self.centralwidget)
        self.centralwidget.setLayout(self.verticalLayout)

        # 레이아웃 설정
        #layout = QVBoxLayout()
        #layout.addWidget(self.table_view)

        # 중앙 위젯 설정
        #container = QWidget()
        #container.setLayout(layout)
        #self.setCentralWidget(container)

    def openFileDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xls *.xlsx);;All Files (*)", options=options)
        if fileName:
            self.loadExcelFile(fileName)

    def loadExcelFile(self, fileName):
        try:
            df = pd.read_excel(fileName)
            self.model = PandasModel(df)
            self.tableView.setModel(self.model)
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())