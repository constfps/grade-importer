import sys
import gspread
import re
from enum import Enum
from google.oauth2.service_account import Credentials
from PyQt5.QtWidgets import QApplication, QTableWidgetItem, QWidget, QStackedWidget
from PyQt5.uic import loadUi
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt

# auth stuff
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
cred = Credentials.from_service_account_file("keys.json", scopes=scopes)
client = gspread.authorize(cred)
studentSelection = None

class WidgetIndexes(Enum):
    SPREADSHEET_INFO = 0
    STUDENTS_LIST = 1
    UNITS_LIST = 2
    IDLE = 3

class SpreadsheetInfo(QWidget):
    def __init__(self):
        super(SpreadsheetInfo, self).__init__()
        loadUi("main.ui", self)
        self.confirmButton.clicked.connect(self.validateSpreadsheet)

    def validateSpreadsheet(self):
        # Extract input
        self.sheetID = self.SpreadsheetIDInput.text()
        self.gradesIndex = int(self.GradesIndexInput.text())
        self.studentsIDIndex = int(self.StudentIDIndexInput.text())

        try:
            self.ids_sheet = client.open_by_key(self.sheetID).get_worksheet(self.studentsIDIndex)
            self.grades_sheet = client.open_by_key(self.sheetID).get_worksheet(self.gradesIndex)
            self.infoTable = [
                # Student Names
                list(map(lambda cell: cell.value, self.ids_sheet.findall(re.compile("[A-Z]\\w+, [A-Z]\\w+")))),

                # Student IDs
                list(map(lambda cell: int(cell.value), self.ids_sheet.findall(re.compile("^\\d+$"), in_column=2))),

                # Section IDs
                list(map(lambda cell: int(cell.value), self.ids_sheet.findall(re.compile("^\\d+$"), in_column=3)))
            ]

            # Check if any info is missing
            assert len(self.infoTable[0]) == len(self.infoTable[1]) and len(self.infoTable[1]) == len(self.infoTable[2]) and len(self.infoTable[2]) == len(self.infoTable[0])
        except AssertionError:
            print("Missing data")
        except gspread.exceptions.SpreadsheetNotFound:
            print("Spreadsheet not found")
        except Exception as e:
            print(e)
        else:
            # Update UI & go to students list
            self.parentWidget().widget(WidgetIndexes.STUDENTS_LIST.value).updateTable(self.infoTable)
            self.parentWidget().setCurrentIndex(WidgetIndexes.STUDENTS_LIST.value)

class StudentsList(QWidget):
    def __init__(self):
        super(StudentsList, self).__init__()
        loadUi("list.ui", self)

        # Connect buttons to respective methods
        self.DeselectAll.clicked.connect(self.deselectAll)
        self.SelectAll.clicked.connect(self.selectAll)
        self.ConfirmButton.clicked.connect(self.generateSelection)

    # I love RegEx
    def deselectAll(self):
        for item in self.table.findItems("[A-Z]\\w+, [A-Z]\\w+", Qt.MatchFlag.MatchRegularExpression):
            item.setCheckState(Qt.CheckState.Unchecked)

    def selectAll(self):
        for item in self.table.findItems("[A-Z]\\w+, [A-Z]\\w+", Qt.MatchFlag.MatchRegularExpression):
            item.setCheckState(Qt.CheckState.Checked)

    def generateSelection(self):
        global studentSelection
        studentSelection = list(map(lambda student: bool(student.checkState() == Qt.CheckState.Checked), self.table.findItems("[A-Z]\\w+, [A-Z]\\w+", Qt.MatchFlag.MatchRegularExpression)))

    def updateTable(self, infoTable: list):
        self.table.setRowCount(len(infoTable[0]))
        self.table.setColumnWidth(0, 150)

        for info in range(0, len(infoTable)):
            for data in range(0, len(infoTable[info])):
                if info == 0:
                    item = QTableWidgetItem(str(infoTable[info][data]))
                    item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
                    item.setCheckState(Qt.CheckState.Checked)
                    self.table.setItem(data, info, item)
                else: 
                    item = QTableWidgetItem(str(infoTable[info][data]))
                    item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                    self.table.setItem(data, info, item)

def main():
    # boilerplate to do everything
    app = QApplication(sys.argv)

    ui = QStackedWidget()
    spreadsheets = SpreadsheetInfo()
    students = StudentsList()

    ui.addWidget(spreadsheets)
    ui.addWidget(students)
    
    ui.setWindowIcon(QIcon("fav.jpg"))
    ui.setFixedSize(411, 270)
    ui.setWindowTitle("Grade Importer")
    ui.show()
    
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()