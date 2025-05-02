import sys
import gspread
import re
from google.oauth2.service_account import Credentials
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication, QTableWidgetItem
from PyQt5.uic import loadUi
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt

# auth stuff
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
            self.infoTable = [
                # Student Names
                list(map(lambda cell: cell.value, self.ids_sheet.findall(re.compile("[A-Z]\\w+, [A-Z]\\w+")))),

                # Student IDs
                list(map(lambda cell: int(cell.value), self.ids_sheet.findall(re.compile("\\d+"), in_column=2))),

                # Section IDs
                list(map(lambda cell: int(cell.value), self.ids_sheet.findall(re.compile("\\d+"), in_column=3)))
            ]

            # Check if any info is missing
            assert len(self.infoTable[0]) == len(self.infoTable[1]) and len(self.infoTable[1]) == len(self.infoTable[2]) and len(self.infoTable[2]) == len(self.infoTable[0])
        except AssertionError:
            self.statusBar.showMessage("Looks like there's some missing parts of the table")
        except gspread.exceptions.SpreadsheetNotFound:
            self.statusBar.showMessage("Spreadsheet ID invalid. Did you share it to the API email?")
        except Exception as e:
            self.statusBar.showMessage(f"Error: {e}")
        else:
            loadUi("list.ui", self)
            self.table.setRowCount(len(self.infoTable[0]))

            for info in range(0, len(self.infoTable)):
                for data in range(0, len(self.infoTable[info])):
                    if info == 0:
                        item = QTableWidgetItem(str(self.infoTable[info][data]))
                        item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
                        item.setCheckState(Qt.CheckState.Checked)
                        self.table.setItem(data, info, item)
                    else: 
                        item = QTableWidgetItem(str(self.infoTable[info][data]))
                        item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                        self.table.setItem(data, info, item)

def main():
    # boilerplate to do everything
    app = QApplication(sys.argv)
    ui = SpreadsheetInfo()
    ui.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()