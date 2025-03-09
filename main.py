import gspread
import json
import re
import curses
from curses import wrapper
from curses.textpad import Textbox
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time

scopes = [
    "https://www.googleapis.com/auth/spreadsheets"
]
cred = Credentials.from_service_account_file("keys.json", scopes=scopes)
client = gspread.authorize(cred)

ids_sheet_id = "1DJ2Q8NHtYyrINqI8JiZGvh2F52R7Z06Qmb3_cQpLguQ"
ids_sheet = client.open_by_key(ids_sheet_id).get_worksheet(0)

grades_sheet_id = "1jEcSOD_5aTKX77XKgWjY5RSveHfXF3c6Op58tApObNA"
grades_sheet = client.open_by_key(grades_sheet_id).get_worksheet(1)

student_selection = []
module_selection = []
submodule_selection = {}

def main(stdscr: curses.window):
    global submodule_selection
    global module_selection
    global student_selection
    
    selection_list = ['student', 'module', 'submodule']
    entered_selections = []
    current_selection = 1
    
    for selection in selection_list:
        while True:
            stdscr.clear()
            stdscr.refresh()
            stdscr.addstr(f"Select {selection} mode:\n")

            stdscr.addstr("1. Single\n")
            stdscr.addstr("2. Multiple\n")
            stdscr.addstr("3. All")
            
            match current_selection:
                case 1:
                    stdscr.addstr(1, 0, "1. Single", curses.A_REVERSE)
                case 2:
                    stdscr.addstr(2, 0, "2. Multiple", curses.A_REVERSE)
                case 3:
                    stdscr.addstr(3, 0, "3. All", curses.A_REVERSE)

            stdscr.refresh()

            key = stdscr.getch()

            key_ups = [curses.KEY_UP, 115, 97]
            key_downs = [curses.KEY_DOWN, 119, 100]

            try:
                if key in key_downs:
                    current_selection += 1
                elif key in key_ups:
                    current_selection -= 1
                assert current_selection < 4 and current_selection > 0
            except:
                if key in key_downs:
                    current_selection -= 1
                elif key in key_ups:
                    current_selection += 1

            if key == 10 or key == 32:
                entered_selections.append(current_selection)
                break
        
        stdscr.clear()
        stdscr.refresh()
    
    selection_dict = dict(zip(selection_list, entered_selections))

    for selection_type in selection_dict:
        stdscr.clear()
        match selection_type:
            case "student":
                match selection_dict[selection_type]:
                    case 1:
                        stdscr.addstr("Enter student's last name")
                        win = curses.newwin(1, 30, 1, 0)
                        box = Textbox(win)

                        stdscr.refresh()
                        box.edit()
                        student_selection.append(box.gather().strip().capitalize())
                    case 2:
                        stdscr.addstr("Enter students' last name (separate with space)")
                        win = curses.newwin(1, 30, 1, 0)
                        box = Textbox(win)

                        stdscr.refresh()
                        box.edit()
                        student_selection = box.gather().strip().title().split(" ")
                    case 3:
                        student_selection = "all"
            case "module":
                match selection_dict[selection_type]:
                    case 1:
                        stdscr.addstr("Enter module")
                        win = curses.newwin(1, 30, 1, 0)
                        box = Textbox(win)

                        stdscr.refresh()
                        box.edit()
                        module_selection.append(int(box.gather().strip()))
                    case 2:
                        stdscr.addstr("Enter modules (separate with space)")
                        win = curses.newwin(1, 30, 1, 0)
                        box = Textbox(win)

                        stdscr.refresh()
                        box.edit()
                        module_selection = [int(x) for x in box.gather().strip().split(" ")]
                    case 3:
                        module_selection = range(1, 11)
            case "submodule":
                match selection_dict[selection_type]:
                    case 1:
                        if type(module_selection) == list:
                            for i in module_selection:
                                stdscr.clear()
                                stdscr.addstr(f"Enter submodule for Module {i}")
                                win = curses.newwin(1, 30, 1, 0)
                                box = Textbox(win)

                                stdscr.refresh()
                                box.edit()
                                submodule_selection.setdefault(i, int(box.gather().strip()))
                        else:
                            stdscr.addstr(f"Enter submodule for Module {module_selection}")
                            win = curses.newwin(1, 30, 1, 0)
                            box = Textbox(win)

                            stdscr.refresh()
                            box.edit()
                            submodule_selection.setdefault(module_selection, int(box.gather().strip()))
                    case 2:
                        if type(module_selection) == list:
                            for i in module_selection:
                                stdscr.clear()
                                stdscr.addstr(f"Enter submodules for Module {i} (separate with space)")
                                win = curses.newwin(1, 30, 1, 0)
                                box = Textbox(win)

                                stdscr.refresh()
                                box.edit()
                                submodule_selection.setdefault(i, [int(x) for x in box.gather().strip().split(" ")])
                        else:
                            stdscr.addstr(f"Enter submodules for Module {module_selection} (separate with space)")
                            win = curses.newwin(1, 30, 1, 0)
                            box = Textbox(win)

                            stdscr.refresh()
                            box.edit()
                            submodule_selection.setdefault(module_selection, [int(x) for x in box.gather().strip().split(" ")])
                    case 3:
                        submodule_selection = "all"

wrapper(main)

# get student(s) info
studentnames = []
studentids = []
sectionids = []

for student in student_selection:
    studentrow = ids_sheet.find(re.compile(student)).row
    
    studentnames.append(ids_sheet.cell(studentrow, 1).value)
    studentids.append(ids_sheet.cell(studentrow, 2).value)
    sectionids.append(ids_sheet.cell(studentrow, 3).value)

if module_selection == "all":
    module_selection = range(1, 11)

# load driver and cookies
with open('cookies.json', 'r') as file:
    cookies = json.load(file)

driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))

#preload to fix cookie adverse issues
driver.get("https://codehs.com")

# set driver cookies
for item in cookies:
    driver.add_cookie({
        'name': item,
        'value': cookies[item]
    })

# iterate through every student
for student in range(0, len(student_selection)):
    # iterate through every module
    for target_module in module_selection:
        #go to assignments page and wait and click on main module
        driver.get("https://codehs.com/student/" + studentids[student] + "/section/" + sectionids[student] + "/assignments")
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((
                By.XPATH, f"//span[text()='Unit {target_module}: Nitro']"
            ))
        )
        driver.find_element(By.XPATH, f"//span[text()='Unit {target_module}: Nitro']").click()
        time.sleep(0.5)

        #wait for sub-modules to load
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((
                By.XPATH, 
                f"//div[@class='lazy-wrap']"
            ))
        )

        # set max sub-module count
        target_module_max = len(driver.find_elements(By.XPATH, "//div[@class='lazy-wrap']/div"))

        # set sub-modules selection appropriately if all mode is selected
        if submodule_selection == "all":
            submodules_list = range(1, target_module_max)
        else:
            submodules_list = submodule_selection[target_module]

        # iterate through every sub-module
        for submodule in submodules_list:
            # assign xpath of list of micro modules
            micro_modules_xpath = f"//div[@class='lessons-sec module-expand' and @style='display: block;']/div[@class='lazy-wrap']/div[{submodule}]/div[@class='lesson-header']/div[@class='right']/div[@class='lesson-items']/a"

            #convert DOM elements to dict with href and type
            micro_modules = driver.find_elements(By.XPATH, micro_modules_xpath)
            micro_modules_href = []
            micro_modules_type = []

            for module in micro_modules:
                # skip anything that isnt finalized
                if module.get_attribute("aria-label").find("Finalized") == -1:
                    continue

                micro_modules_href.append(module.get_attribute("href"))
                micro_modules_type.append(module.get_attribute("class").split(" ")[0])

            modules = dict(zip(micro_modules_href, micro_modules_type))

            # set module types that were looking for
            module_type = ["example", "exercise", "quiz"]

            # initialize counters
            quiz_sum = 0
            examples_sum = 0
            exercise_sum = 0

            # iterate through every micro module and add to score counter
            for module in modules:
                if modules[module] in module_type:
                    driver.get(module)
                    
                    if (modules[module] == "quiz"):
                        WebDriverWait(driver, 5).until(EC.presence_of_element_located((
                            By.CLASS_NAME, "num-correct"
                        )))

                        # will always be a number even if unattempted
                        quiz_sum += int(driver.find_element(By.CLASS_NAME, "num-correct").text)
                    else:
                        WebDriverWait(driver, 5).until(EC.presence_of_element_located((
                            By.XPATH, "//div[text()='Grade']"
                        )))
                        driver.find_element(By.XPATH, "//div[text()='Grade']").click()
                        
                        WebDriverWait(driver, 5).until(EC.presence_of_element_located((
                            By.CLASS_NAME, "grade-score"
                        )))
                        # sometimes grade-score loads without text for a bit so wait a bit more
                        time.sleep(0.1)
                        grade = driver.find_element(By.CLASS_NAME, "grade-score").text

                        if grade.isnumeric():
                            if modules[module] == "example":
                                examples_sum += int(grade)
                            else:
                                exercise_sum += int(grade)
            
            # print results
            print(f"{target_module}.{submodule} scores:")
            print(f"Quiz: {quiz_sum}")
            print(f"Examples: {examples_sum}")
            print(f"Programs: {exercise_sum}")

            for type in module_type:
                match type:
                    case "example":
                        try:
                            col = grades_sheet.find(f"{target_module}.{submodule} Examples").col
                        except:
                            continue
                        temp = examples_sum
                    case "exercise":
                        try:
                            col = grades_sheet.find(f"{target_module}.{submodule} Programs").col
                        except:
                            continue
                        temp = exercise_sum
                    case "quiz":
                        col = grades_sheet.find(f"{target_module}.{submodule} Quiz").col
                        temp = quiz_sum
            
                row = grades_sheet.find(studentnames[student]).row
                cell = gspread.cell.Cell(row, col)
                
                grades_sheet.update_acell(cell.address, temp)

            # go back to initial assignments page
            driver.get("https://codehs.com/student/" + studentids[student] + "/section/" + sectionids[student] + "/assignments")
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((
                    By.XPATH, f"//span[text()='Unit {target_module}: Nitro']"
                ))
            )
            driver.find_element(By.XPATH, f"//span[text()='Unit {target_module}: Nitro']").click()

            #wait for sub-modules to load
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((
                    By.XPATH, 
                    f"//div[@class='lazy-wrap']"
                ))
            )

            # to avoid rate limiter
            time.sleep(2)

time.sleep(5)
driver.close()