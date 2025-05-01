import sys
import gspread
from google.oauth2.service_account import Credentials
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication
from PyQt5.uic import loadUi
from PyQt5.QtGui import QIcon

class SpreadsheetInfo(QMainWindow):
    def __init__(self):
        super(SpreadsheetInfo, self).__init__()
        loadUi("main.ui", self)
        self.confirmButton.clicked.connect(self.validateSpreadsheet)
        self.setWindowIcon(QIcon("fav.jpg"))
        self.setFixedSize(411, 260)

    def validateSpreadsheet(self):
        sheetID = self.SpreadsheetIDInput.text()
        print(sheetID)

def main():
    app = QApplication(sys.argv)
    ui = SpreadsheetInfo()
    ui.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()