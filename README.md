# Grade Importer
A CodeHS Paywall Bypass because I'm lazy
## Setting up
* Make sure Google Chrome is installed
* Make a file named ***cookies.json*** and populate it with keys of of your CodeHS sessionid, uid, and username seen under Inspect Element > Application > Cookies
    - Example:
    ```
    {
        "sessionid": "qwertyuiopasdfghjklzxcvbnm1234567890",
        "uid": "123456789",
        "username": "\"b'John Doe'\""
    }
    ```
* Include the ***keys.json*** file provided by me
* Make a seperate sheet with students' CodeHS name, Student ID, and Section ID
## Usage
* Upon running ***run me.bat***, you will be prompted to enter the spreadsheet's ID of where the student IDs are and the index of the sheet
* You will be prompted for a spreadsheet ID and index of sheet again but for where to put the grades
* Select the range for Students, Modules, and Sub-module
* Sit back and watch as glory happens