import gspread
import json
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
ids_sheet = client.open_by_key(ids_sheet_id)

grades_sheet_id = "1jEcSOD_5aTKX77XKgWjY5RSveHfXF3c6Op58tApObNA"
grades_sheet = client.open_by_key(grades_sheet_id)

# get student info
studentname = ids_sheet.get_worksheet(0).acell('A14').value
studentid = ids_sheet.get_worksheet(0).acell('B14').value
sectionid = ids_sheet.get_worksheet(0).acell('C14').value

# load driver and cookies
with open('cookies.json', 'r') as file:
    cookies = json.load(file)

driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))

#to fix cookie adverse issues
driver.get("https://codehs.com")

# set driver cookies
for item in cookies:
    driver.add_cookie({
        'name': item,
        'value': cookies[item]
    })

# set target module
target_module = 2

#go to assignments page and wait and click on main module
driver.get("https://codehs.com/student/" + studentid + "/section/" + sectionid + "/assignments")
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

for i in range(1, target_module_max):
    # assign xpath of list of micro modules
    micro_modules_xpath = f"//div[@class='lessons-sec module-expand' and @style='display: block;']/div[@class='lazy-wrap']/div[{i}]/div[@class='lesson-header']/div[@class='right']/div[@class='lesson-items']/a"
    #convert DOM elements to dict
    micro_modules = driver.find_elements(By.XPATH, micro_modules_xpath)
    micro_modules_href = []
    micro_modules_type = []
    for module in micro_modules:
        micro_modules_href.append(module.get_attribute("href"))
        micro_modules_type.append(module.get_attribute("class").split(" ")[0])
    modules = dict(zip(micro_modules_href, micro_modules_type))

    # set module types that were looking for
    module_type = ["example", "exercise", "quiz"]

    # initiallize counters
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
                grade = driver.find_element(By.CLASS_NAME, "grade-score").text

                # grade will be a dash if not graded yet
                if grade.isnumeric():
                    if modules[module] == "example":
                        examples_sum += int(grade)
                    else:
                        exercise_sum += int(grade)
    
    # print results
    print(f"{target_module}.{i} scores:")
    print(f"Quiz: {quiz_sum}")
    print(f"Examples: {examples_sum}")
    print(f"Programs: {exercise_sum}")

    for type in module_type:
        match type:
            case "example":
                col = grades_sheet.get_worksheet(1).find(f"{target_module}.{i} Examples").col
                temp = examples_sum
            case "exercise":
                col = grades_sheet.get_worksheet(1).find(f"{target_module}.{i} Programs").col
                temp = exercise_sum
            case "quiz":
                col = grades_sheet.get_worksheet(1).find(f"{target_module}.{i} Quiz").col
                temp = quiz_sum
    
        row = grades_sheet.get_worksheet(1).find(studentname).row
        cell = gspread.cell.Cell(row, col)
        
        grades_sheet.get_worksheet(1).update_acell(cell.address, temp)

    # go back to initial assignments page
    driver.get("https://codehs.com/student/" + studentid + "/section/" + sectionid + "/assignments")
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
    time.sleep(2.5)

time.sleep(5)
driver.close()