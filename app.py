from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QToolBar, QMainWindow, QComboBox, QTableView, QLabel, QFileDialog,QPushButton
from PySide6.QtGui import QAction
from PySide6.QtCore import QAbstractTableModel
import pandas as pd
import sys
class Window(QMainWindow):
    def __init__(self):
        def setSpreadSheetData():
            spreadsheet.setModel(self.openFileExplorer())
        super().__init__()
        self.setWindowTitle("Excel")
        self.setGeometry(900,500,900,500)
        layout = QVBoxLayout()
        widget = QWidget()



        openButton = QAction("open",self)
        openButton.triggered.connect(self.openFileExplorer)
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
        # return file
        print(file)
class dataModel(QAbstractTableModel):
    def __init__(self,data):
        self.data = data
    def data

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec())