import sys
import time
import gspread
import re
import json
import threading
from enum import Enum
from google.oauth2.service_account import Credentials
from PyQt5.QtWidgets import QApplication, QTableWidgetItem, QWidget, QStackedWidget, QTreeWidgetItem
from PyQt5.uic import loadUi
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# Awtorisasyon
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
cred = Credentials.from_service_account_file("keys.json", scopes=scopes)
client = gspread.authorize(cred)

grades_cells = []
studentNames = []
studentSelection = {}
unitSelection = {}

spreadsheet_id = None
ids_index = None
grades_index = None
grades_worksheet = None

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
        global grades_cells
        global grades_worksheet
        
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
                list(map(lambda cell: int(cell.value), self.ids_sheet.findall(re.compile("^\\d+$"), in_column=self.ids_sheet.find(re.compile("[A-Z]\\w+, [A-Z]\\w+")).col+1))),

                # Mga ID ng mga section
                list(map(lambda cell: int(cell.value), self.ids_sheet.findall(re.compile("^\\d+$"), in_column=self.ids_sheet.find(re.compile("[A-Z]\\w+, [A-Z]\\w+")).col+2)))
            ]

            # Tignan kung may nawawalang impormasyon
            assert len(self.infoTable[0]) == len(self.infoTable[1]) and len(self.infoTable[1]) == len(self.infoTable[2]) and len(self.infoTable[2]) == len(self.infoTable[0])
            grades_cells = self.grades_sheet.get_all_cells()
            grades_worksheet = self.grades_sheet
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
        global studentNames
        
        for student in range(self.table.rowCount()):
            if self.table.item(student, 0).checkState() == Qt.CheckState.Checked:
                studentSelection[self.table.item(student, 1).text()] = self.table.item(student, 2).text()
                studentNames.append(self.table.item(student, 0).text())

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

        self.tree.itemChanged.connect(self.onItemChanged)
        self.ConfirmButton.clicked.connect(self.generateSelection)
        self.SelectAll.clicked.connect(self.selectAll)
        self.DeselectAll.clicked.connect(self.deselectAll)

    def selectAll(self):
        for i in range(self.tree.topLevelItemCount()):
            self.tree.topLevelItem(i).setCheckState(0, Qt.CheckState.Checked)

    def deselectAll(self):
        for i in range(self.tree.topLevelItemCount()):
            self.tree.topLevelItem(i).setCheckState(0, Qt.CheckState.Unchecked)
    
    def generateSelection(self):
        global unitSelection
        
        unitSelection = {}
        for i in range(self.tree.topLevelItemCount()):
            if self.tree.topLevelItem(i).checkState(0) != Qt.CheckState.Unchecked:
                unitSelection[str(i + 1)] = []
                for x in range(self.tree.topLevelItem(i).childCount()):
                    unitSelection[str(i + 1)].append(bool(self.tree.topLevelItem(i).child(x).checkState(0) == Qt.CheckState.Checked))
        
        extraction()
        '''
        t1 = threading.Thread(target=extraction)
        t1.start()
        '''

    def updateTree(self, grades_sheet: gspread.Worksheet):
        # Kunin lahat ng possibleng units mula sa worksheet
        data = list(map(lambda cell: cell.split(" "), list(map(lambda cell: cell.value, grades_sheet.findall(query=re.compile("^\\d{1,2}\\.\\d{1,2}\\s\\w+$"))))))

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
        # kung may kagaguhan na yung column hindi 0
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

def extraction():
    global studentSelection
    global unitSelection
    global grades_worksheet
    
    inpage = False
    
    # pagka-login sa CodeHS
    with open("cookies.json", "r") as file:
        cookies = json.load(file)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get("https://codehs.com")
    driver.maximize_window()
    for cookie in cookies:
        driver.add_cookie({
            "name": cookie,
            "value": cookies[cookie]
        })
    driver.refresh()

    for student in studentSelection:
        studentName = studentNames[list(studentSelection.keys()).index(student)]
        
        driver.get(f"https://codehs.com/student/{student}/section/{studentSelection[student]}/assignments")
        for i in range(50):
            try:
                WebDriverWait(driver, 60).until(
                EC.visibility_of_element_located((
                    By.XPATH, "//span[text()='Unit 0: Nitro']"
                )))
            except:
                driver.refresh()
                time.sleep(1)
            else:
                break
        inpage = False

        for target_unit in unitSelection:
            update_queue = []
            if inpage:
                driver.get(f"https://codehs.com/student/{student}/section/{studentSelection[student]}")
                for i in range(50):
                    try:
                        WebDriverWait(driver, 60).until(
                        EC.visibility_of_element_located((
                            By.XPATH, f"//span[text()='Unit {unit}: Nitro']"
                        )))
                    except:
                        driver.refresh()
                        time.sleep(1)
                    else:
                        break
            
            driver.find_element(By.XPATH, f"//span[text()='Unit {target_unit}: Nitro']").click()
            time.sleep(2)
            inpage = False

            for subunit in range(len(unitSelection[target_unit])):
                print(f"Processing {target_unit}.{subunit+1}")
                if unitSelection[target_unit][subunit]:
                    micro_units_xpath = f"//div[@class='lessons-sec module-expand' and @style='display: block;']/div[@class='lazy-wrap']/div[{subunit+1}]/div[@class='lesson-header']/div[@class='right']/div[@class='lesson-items']/a"

                    micro_units = driver.find_elements(By.XPATH, micro_units_xpath)
                    micro_units_href = []
                    micro_units_type = []
                    unit_type = ["example", "exercise", "quiz"]

                    for micro_unit in micro_units:
                        if micro_unit.get_attribute("aria-label").find("Finalized") == -1:
                            continue
                        elif micro_unit.get_attribute("class").split(" ")[0] in unit_type:
                            micro_units_href.append(micro_unit.get_attribute("href"))
                            micro_units_type.append(micro_unit.get_attribute("class").split(" ")[0])
                    
                    units = dict(zip(micro_units_href, micro_units_type))
                    
                    quiz_sum = 0
                    examples_sum = list(units.values()).count('example')
                    exercise_sum = 0
                    
                    for unit in units:
                        if units[unit] == "quiz":
                            driver.get(unit)
                            inpage = True
                            for i in range(50):
                                try:
                                    WebDriverWait(driver, 60).until(EC.visibility_of_element_located((
                                        By.CLASS_NAME, "num-correct"
                                    )))
                                except:
                                    driver.refresh()
                                else:
                                    break
                            quiz_sum += int(driver.find_element(By.CLASS_NAME, "num-correct").text)
                        elif units[unit] == "exercise":
                            driver.get(unit)
                            inpage = True

                            for i in range(50):
                                try:
                                    WebDriverWait(driver, 60).until(EC.visibility_of_element_located((
                                        By.XPATH, "//div[text()='Grade']"
                                    )))
                                except:
                                    driver.refresh()
                                else:
                                    break
                            
                            driver.find_element(By.XPATH, "//div[text()='Grade']").click()
                            
                            for i in range(50):
                                try:
                                    WebDriverWait(driver, 60).until(EC.visibility_of_element_located((
                                        By.CLASS_NAME, "grade-score"
                                    )))
                                except:
                                    driver.refresh()
                                else:
                                    break
                            
                            grade = driver.find_element(By.CLASS_NAME, "grade-score").text

                            if grade.isnumeric:
                                if units[unit] == "example":
                                    examples_sum += int(grade)
                                else:
                                    exercise_sum += int(grade)
                    
                    print(f"Examples: {examples_sum}\nQuiz: {quiz_sum}\nExercises: {exercise_sum}")
                
                    for i in unit_type:
                        try:
                            grade_group = f"{target_unit}.{subunit+1} "

                            match i:
                                case "quiz":
                                    grade_group += "Quiz"
                                case "exercise":
                                    grade_group += "Programs"
                                case "example":
                                    grade_group += "Examples"

                            cell = triangulate(find_cell(studentName, grades_cells), find_cell(grade_group, grades_cells))
                            update_queue.append({
                                'range': cell.address,
                                'values': [[str(quiz_sum)]]
                            })
                        except:
                            continue

                
                if inpage:
                    # punta sa inisyal sa pahina
                    driver.get("https://codehs.com/student/" + student + "/section/" + studentSelection[student] + "/assignments")
                    WebDriverWait(driver, 60).until(
                        EC.presence_of_element_located((
                            By.XPATH, f"//span[text()='Unit {target_unit}: Nitro']"
                        ))
                    )
                    driver.find_element(By.XPATH, f"//span[text()='Unit {target_unit}: Nitro']").click()
                    inpage = False
            
            print(update_queue)
            grades_worksheet.batch_update(update_queue)

def find_cell(query: str, grades_cells: list):
    return grades_cells[list(map(lambda cell: cell.value, grades_cells)).index(query)]

def triangulate(row: gspread.Cell, col: gspread.Cell):
    return gspread.Cell(row.row, col.col)

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