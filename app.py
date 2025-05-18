from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QToolBar, QMainWindow, QComboBox, QTableView, QLabel, QFileDialog,QPushButton, QMessageBox
from PySide6.QtGui import QAction
from PySide6.QtCore import QAbstractTableModel, Qt
import pandas as pd
import sys
from array import array
class Window(QMainWindow):
    def __init__(self):
        def setSpreadSheetData():
            self.model = TableModel(self.openFileExplorer())
            spreadsheet.setModel(self.model)
        def saveSpreadSheetData():
            print(self.model.getData())
            with pd.ExcelWriter(self.saveFileExplorer()) as writer:
                self.model.getData().to_excel(writer,index=False)
        super().__init__()
        self.setWindowTitle("Excel")
        self.model = TableModel(self)
        self.setGeometry(900,500,900,500)
        self.columns = {}
        layout = QVBoxLayout()
        widget = QWidget()
        openButton = QAction("open",self)
        openButton.triggered.connect(setSpreadSheetData)
        saveButton = QAction("save as",self)
        saveButton.triggered.connect(saveSpreadSheetData)
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
        fileExplorer = QFileDialog.getOpenFileName(self,"Open an Excel File","","","Excel files (*.xlsx)")
        if fileExplorer[0].split(".")[1]=="xlsx":
            file = pd.read_excel(fileExplorer[0])
            self.setWindowTitle(fileExplorer[0].split("/")[-1])
            for (i,column) in enumerate(list(file.columns.values)):
                self.columns.update({i:column})
            print(self.columns)
            return file
        else:
            message = QMessageBox(self,"Error","Nieprawdłowy rodzaj pliku")
            message.setText("Nieprawidłowy rodzaj pliku")
            message.show()
            return None
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
    def getData(self):
        data = []
        headers = []
        indexes = []
        if len(self._data) > 0:
            for x in range(0,self.rowCount(0)):
                for y in range(0,self.columnCount(0)):
                    data.append(self._data.iloc[x, y])
                    headers.append(self._data.columns[y])
                    indexes.append(self._data.index[x])
            print(data)
            return pd.DataFrame([data], columns=headers)
        return None

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