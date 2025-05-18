from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QToolBar, QMainWindow, QComboBox, QTableView, QLabel, QFileDialog,QPushButton, QMessageBox
from PySide6.QtGui import QAction
from PySide6.QtCore import QAbstractTableModel, Qt
import pandas as pd
import sys
from array import array
class Window(QMainWindow):
    def __init__(self):
        def saveSpreadSheetData():
            keys = list(self.xlsxFile)
            values = list(self.xlsxFile.values())
            with pd.ExcelWriter(self.saveFileExplorer()) as writer:
                for (key,value) in enumerate(values):
                    value.to_excel(writer,index=False,sheet_name=keys[key])
        super().__init__()
        self.url = ""
        self.setWindowTitle("Excel")
        self.model = TableModel(self)
        self.setGeometry(900,500,900,500)
        self.columns = {}
        layout = QVBoxLayout()
        widget = QWidget()
        openButton = QAction("open",self)
        openButton.triggered.connect(self.openFileExplorer)
        saveButton = QAction("save as",self)
        saveButton.triggered.connect(saveSpreadSheetData)
        self.menu = self.menuBar()
        fileMenu = self.menu.addMenu("File")
        self.sheetsMenu = self.menu.addMenu("Sheets")
        fileMenu.addAction(openButton)
        fileMenu.addAction(saveButton)

        filterButton = self.menu.addAction("Filter")
        helpButton = self.menu.addAction("Help")

        self.spreadsheet = QTableView(self)
        self.spreadsheet.showGrid = True
        layout.addWidget(self.spreadsheet)

        widget.setLayout(layout)
        self.setCentralWidget(widget)
    def openFileExplorer(self):
        fileExplorer = QFileDialog.getOpenFileName(self,"Open an Excel File","","","Excel files (*.xlsx)")
        if fileExplorer[0].split(".")[1]=="xlsx":
            self.url = fileExplorer[0]
            self.xlsxFile = pd.read_excel(self.url,sheet_name=None)
            defaultFile = list(self.xlsxFile.values())[0]
            self.setWindowTitle(f"{self.url.split("/")[-1]}")
            self.columns = {}
            for (i,column) in enumerate(list(defaultFile.columns.values)):
                self.columns.update({i:column})
            self.sheetsMenu.clear()
            if len(self.xlsxFile.keys())>1:
                self.setWindowTitle(f"{self.url.split("/")[-1]} : {list(self.xlsxFile.keys())[0]}")
                for (i,sheetName) in enumerate(self.xlsxFile.keys()):
                    sheetAction = self.sheetsMenu.addAction(f"Open {sheetName}")
                    sheetAction.triggered.connect(lambda state, x=i:self.openSheet(x))
            self.model = TableModel(defaultFile)
            self.spreadsheet.setModel(self.model)
        else:
            message = QMessageBox(self,"Error","Nieprawdłowy rodzaj pliku")
            message.setText("Nieprawidłowy rodzaj pliku")
            message.show()
    def openSheet(self,sheetID):
        file = list(self.xlsxFile.values())[sheetID]
        for (i, column) in enumerate(list(file.columns.values)):
            self.columns = {}
            self.columns.update({i: column})
        self.model = TableModel(file)
        self.spreadsheet.setModel(self.model)

    def saveFileExplorer(self):
        fileExplorer = QFileDialog.getSaveFileName(self,"Save Excel file","","","Excel files (*.xlsx)")
        if fileExplorer[0][-5:len(fileExplorer[0])]==".xlsx":
            return fileExplorer[0]
        else:
            return fileExplorer[0] + ".xlsx"
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

    def setData(self, index, value, role):
        if role == Qt.EditRole:
            self._data.iloc[index.row(), index.column()] = value
            return True
        return False

    def flags(self, index):
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec())