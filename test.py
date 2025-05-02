import gspread
import re
from google.oauth2.service_account import Credentials

scopes = ["https://www.googleapis.com/auth/spreadsheets"]
cred = Credentials.from_service_account_file("keys.json", scopes=scopes)
client = gspread.authorize(cred)

query = re.compile("^\\d{1,2}\\.\\d{1,2}\\s\\w+$")
grades = client.open_by_key("1jEcSOD_5aTKX77XKgWjY5RSveHfXF3c6Op58tApObNA").get_worksheet(1)
data = list(map(lambda cell: cell.split(" "), list(map(lambda cell: cell.value, grades.findall(query=query)))))

for cell in data:
    cell.pop(1)

temp = []
for i in data:
    if i[0] not in temp:
        temp.append(i[0])
data = temp

asdf = {}
for i in range(1, 11):
    asdf.setdefault(str(i), [])
for i in data:
    x = i.split(".")
    asdf[str(x[0])].append(x[1])

print(asdf)