import sys
import gspread
import re
from enum import Enum
from google.oauth2.service_account import Credentials
from PyQt5.QtWidgets import QApplication, QTableWidgetItem, QWidget, QStackedWidget, QTreeWidgetItem
from PyQt5.uic import loadUi
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt

# Awtorisasyon
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
cred = Credentials.from_service_account_file("keys.json", scopes=scopes)
client = gspread.authorize(cred)

studentSelection = None

spreadsheet_id = None
ids_index = None
grades_index = None

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
        global spreadsheet_id
        global ids_index
        global grades_index
        
        # Kunin ang mga nilagay
        spreadsheet_id = self.SpreadsheetIDInput.text()
        ids_index = int(self.GradesIndexInput.text())
        grades_index = int(self.StudentIDIndexInput.text())

        try:
            self.ids_sheet = client.open_by_key(spreadsheet_id).get_worksheet(grades_index)
            self.grades_sheet = client.open_by_key(spreadsheet_id).get_worksheet(ids_index)
            self.infoTable = [
                # Mga pangalan ng mga estudyante
                list(map(lambda cell: cell.value, self.ids_sheet.findall(re.compile("[A-Z]\\w+, [A-Z]\\w+")))),

                # Mga ID ng estudyante
                list(map(lambda cell: int(cell.value), self.ids_sheet.findall(re.compile("^\\d+$"), in_column=2))),

                # Mga ID ng mga section
                list(map(lambda cell: int(cell.value), self.ids_sheet.findall(re.compile("^\\d+$"), in_column=3)))
            ]

            # Tignan kung may nawawalang impormasyon
            assert len(self.infoTable[0]) == len(self.infoTable[1]) and len(self.infoTable[1]) == len(self.infoTable[2]) and len(self.infoTable[2]) == len(self.infoTable[0])
        except AssertionError:
            print("Missing data")
        except gspread.exceptions.SpreadsheetNotFound:
            print("Spreadsheet not found")
        except Exception as e:
            print(e)
        else:
            # Baguhin yung impormasyon sa mga susunod na widget at pumunta sa listahan ng mga estudyante
            self.parentWidget().widget(WidgetIndexes.STUDENTS_LIST.value).updateTable(self.infoTable)
            self.parentWidget().widget(WidgetIndexes.UNITS_LIST.value).updateTree(self.grades_sheet)
            self.parentWidget().setCurrentIndex(WidgetIndexes.STUDENTS_LIST.value)

class StudentsList(QWidget):
    def __init__(self):
        super(StudentsList, self).__init__()
        loadUi("list.ui", self)

        # I-konekta ang mga pindotan sa kani-kanilang function
        self.DeselectAll.clicked.connect(self.deselectAll)
        self.SelectAll.clicked.connect(self.selectAll)
        self.ConfirmButton.clicked.connect(self.generateSelection)

    def deselectAll(self):
        for item in self.table.findItems("[A-Z]\\w+, [A-Z]\\w+", Qt.MatchFlag.MatchRegularExpression):
            item.setCheckState(Qt.CheckState.Unchecked)

    def selectAll(self):
        for item in self.table.findItems("[A-Z]\\w+, [A-Z]\\w+", Qt.MatchFlag.MatchRegularExpression):
            item.setCheckState(Qt.CheckState.Checked)

    def generateSelection(self):
        global studentSelection
        studentSelection = list(map(lambda student: bool(student.checkState() == Qt.CheckState.Checked), self.table.findItems("[A-Z]\\w+, [A-Z]\\w+", Qt.MatchFlag.MatchRegularExpression)))
        self.parentWidget().setCurrentIndex(WidgetIndexes.UNITS_LIST.value)

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

class UnitsList(QWidget):
    def __init__(self):
        global spreadsheet_id
        global grades_index
        
        super(UnitsList, self).__init__()
        loadUi("units.ui", self)

    def updateTree(self, grades_sheet: gspread.Worksheet):
        # Kunin lahat ng possibleng units mula sa worksheet
        data = list(map(lambda cell: cell.split(" "), list(map(lambda cell: cell.value, grades_sheet.findall(query=re.compile("^\\d{1,2}\\.\\d{1,2}\\s\\w+$"))))))
        self.tree.itemChanged.connect(self.onItemChanged)

        temp = []
        for i in data:
            if i[0] not in temp:
                temp.append(i[0])
        data = temp

        self.units_dict = {}
        for i in range(1, 11):
            self.units_dict.setdefault(str(i), [])
        for i in data:
            x = i.split(".")
            self.units_dict[str(x[0])].append(x[1])
        
        # ilagay ang kontento ng dict sa QTreeWidget
        for key, value in self.units_dict.items():
            unit = QTreeWidgetItem([f"Unit {key}"])
            unit.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
            unit.setCheckState(0, Qt.CheckState.Checked)

            self.tree.addTopLevelItem(unit)
            for submodule in value:
                child = QTreeWidgetItem([f"{key}.{submodule}"])
                child.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
                child.setCheckState(0, Qt.CheckState.Checked)
                unit.addChild(child)
    
    # gitnang function ng signal
    def onItemChanged(self, item, column):
        if column != 0:
            return
        
        # para walang di nagtatapos na recursion
        self.tree.blockSignals(True)
        try:
            state = item.checkState(0)
            self.updateChildrenCheckstate(item, state)
            self.updateParentCheckstate(item)
        finally:
            self.tree.blockSignals(False)

    def updateChildrenCheckstate(self, item, state):
        for i in range(item.childCount()):
            child = item.child(i)
            child.setCheckState(0, state)
            self.updateChildrenCheckstate(child, state)

    def updateParentCheckstate(self, item):
        parent = item.parent()
        if parent is not None:
            checked = 0
            unchecked = 0
            childCount = parent.childCount()

            for i in range(childCount):
                child = parent.child(i)
                if child.checkState(0) == Qt.CheckState.Checked:
                    checked += 1
                elif child.checkState(0) == Qt.CheckState.Unchecked:
                    unchecked += 1
            
            if checked == childCount:
                parent.setCheckState(0, Qt.CheckState.Checked)
            elif unchecked == childCount:
                parent.setCheckState(0, Qt.CheckState.Unchecked)
            else:
                parent.setCheckState(0, Qt.CheckState.PartiallyChecked)
            
            self.updateParentCheckstate(parent)

def main():
    # pag package ng mga widget sa isang pangunahing bintana
    app = QApplication(sys.argv)

    ui = QStackedWidget()
    ui.addWidget(SpreadsheetInfo())
    ui.addWidget(StudentsList())
    ui.addWidget(UnitsList())
    
    ui.setWindowIcon(QIcon("fav.jpg"))
    ui.setFixedSize(411, 270)
    ui.setWindowTitle("Grade Importer")
    ui.show()
    
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()