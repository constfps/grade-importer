import sys
import gspread
from google.oauth2.service_account import Credentials
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication
from PyQt5.uic import loadUi
from PyQt5.QtGui import QIcon

scopes = ["https://www.googleapis.com/auth/spreadsheets"]
cred = Credentials.from_service_account_file("keys.json", scopes=scopes)
client = gspread.authorize(cred)

class SpreadsheetInfo(QMainWindow):
    def __init__(self):
        super(SpreadsheetInfo, self).__init__()
        loadUi("main.ui", self)
        self.confirmButton.clicked.connect(self.validateSpreadsheet)
        self.setWindowIcon(QIcon("fav.jpg"))
        self.setFixedSize(411, 260)

    def validateSpreadsheet(self):
        sheetID = self.SpreadsheetIDInput.text()
        gradesIndex = int(self.GradesIndexInput.text())
        studentsIDIndex = int(self.StudentIDIndexInput.text())

        try:
            ids_sheet = client.open_by_key(sheetID).get_worksheet(studentsIDIndex);
            grades_sheet = client.open_by_key(sheetID).get_worksheet(gradesIndex)
            print("pog it works")
        except Exception as e:
            print(e)

def main():
    app = QApplication(sys.argv)
    ui = SpreadsheetInfo()
    ui.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()