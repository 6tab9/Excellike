from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QToolBar, QMainWindow, QComboBox, QTableView, QLabel, QFileDialog,QPushButton
from PySide6.QtGui import QAction
from PySide6.QtCore import QAbstractTableModel, Qt
import pandas as pd
import sys
class Window(QMainWindow):
    def __init__(self):
        def setSpreadSheetData():
            model = TableModel(self.openFileExplorer())
            spreadsheet.setModel(model)
        super().__init__()
        self.setWindowTitle("Excel")
        self.setGeometry(900,500,900,500)
        layout = QVBoxLayout()
        widget = QWidget()



        openButton = QAction("open",self)
        openButton.triggered.connect(setSpreadSheetData)
        saveButton = QAction("save as",self)

        menu = self.menuBar()

        fileMenu = menu.addMenu("File")
        fileMenu.addAction(openButton)
        fileMenu.addAction(saveButton)

        filterButton = menu.addAction("Filter")
        helpButton = menu.addAction("Help")

        spreadsheet = QTableView(self)
        spreadsheet.showGrid = True
        layout.addWidget(spreadsheet)

        widget.setLayout(layout)
        self.setCentralWidget(widget)
    def openFileExplorer(self):
        fileExplorer = QFileDialog.getOpenFileName(self,"Open an Excel File")
        file = pd.read_excel(fileExplorer[0])
        print(file)
        return file
class TableModel(QAbstractTableModel):

    def __init__(self, data):
        super().__init__()
        self._data = data

    def data(self, index, role):
        if role == Qt.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]
            return str(value)

    def rowCount(self, index):
        return self._data.shape[0]

    def columnCount(self, index):
        return self._data.shape[1]

    def headerData(self, section, orientation, role):
        # section is the index of the column/row.
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._data.columns[section])

            if orientation == Qt.Vertical:
                return str(self._data.index[section])

    # def rowCount(self, index):
    #     # The length of the outer list.
    #     return len(self.data)
    #
    # def columnCount(self, index):
    #     # The following takes the first sub-list, and returns
    #     # the length (only works if all rows are an equal length)
    #     return len(self.data[0])

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec())