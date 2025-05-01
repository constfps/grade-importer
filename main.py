import sys
import gspread
import re
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
        self.setFixedSize(411, 270)

    def validateSpreadsheet(self):
        self.sheetID = self.SpreadsheetIDInput.text()
        self.gradesIndex = int(self.GradesIndexInput.text())
        self.studentsIDIndex = int(self.StudentIDIndexInput.text())

        try:
            self.statusBar.showMessage("Validating input...")
            self.ids_sheet = client.open_by_key(self.sheetID).get_worksheet(self.studentsIDIndex)
            self.grades_sheet = client.open_by_key(self.sheetID).get_worksheet(self.gradesIndex)
            self.statusBar.showMessage("Input valid!")
        except gspread.exceptions.SpreadsheetNotFound:
            self.statusBar.showMessage("Spreadsheet ID invalid. Did you share it to the API email?")
        except Exception as e:
            self.statusBar.showMessage(f"Error: {e}")
        else:
            self.pullNames()
    
    def pullNames(self):
        self.studentNames = list(map(lambda cell: cell.value, self.ids_sheet.findall(re.compile("[A-Z]\w+, [A-Z]\w+"))))

def main():
    app = QApplication(sys.argv)
    ui = SpreadsheetInfo()
    ui.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()