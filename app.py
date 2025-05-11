from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QToolBar, QMainWindow, QComboBox, QTableView, QLabel, QFileDialog,QPushButton
from PySide6.QtGui import QAction
import sys
class Window(QMainWindow):
    def __init__(self):
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
        fileExplorer = QFileDialog.getOpenFileName(self)
        print(fileExplorer)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec())