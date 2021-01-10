# python-DZ
import openpyxl
import random

people=['Аббасов Байрам','Архипова Анна','Богданов Даниил','Гусев Александр','Коконов Александр ','Коваленко Иван','Криволуцкий Михаил','Марков Михаил','Пушкин Александр','Савельев Иван','Стрыкало Валентин','Тимошенко Мария','Федорук Евгения','Хайрулин Тимур']

def list ():
for i in range(len(people)):
value = people[i]
cell = sheet.cell(row = i+2, column = 1)
cell.value = value

wb = openpyxl.Workbook()

wb.create_sheet(title = 'investments', index = 0)

sheet = wb['investments']

sheet.column_dimensions['A'].width = 30

value = "Name"
cell = sheet.cell(row = 1, column = 1)
cell.value = value

value = "Investment"
cell = sheet.cell(row = 1, column = 2)
cell.value = value

value = "September"
cell = sheet.cell(row = 1, column =3 )
cell.value = value

value = "October"
cell = sheet.cell(row = 1, column = 4)
cell.value = value

value = "November"
cell = sheet.cell(row = 1, column = 5)
cell.value = value

value = "Sum"
cell = sheet.cell(row = 1, column = 6)
cell.value = value
list()


for i in range(len(people)):
a=random.randint(0,3)
if (a==0):
value = "no"
cell = sheet.cell(row = i+2, column =2)
cell.value = value
value = "0"
cell = sheet.cell(row = i+2, column =6)
cell.value = value

if (a==1):
value = "yes"
cell = sheet.cell(row = i+2, column =2)
cell.value = value
value = "20"
cell = sheet.cell(row = i+2, column =3)
cell.value = value
cell = sheet.cell(row = i+2, column =6)
cell.value = value
if (a==2):
value = "yes"
cell = sheet.cell(row = i+2, column =2)
cell.value = value
value = "50"
cell = sheet.cell(row = i+2, column =4)
cell.value = value
cell = sheet.cell(row = i+2, column =6)
cell.value = value
if (a==3):
value = "yes"
cell = sheet.cell(row = i+2, column =2)
cell.value = value
value = "80"
cell = sheet.cell(row = i+2, column =5)
cell.value = value
cell = sheet.cell(row = i+2, column =6)
cell.value = value

wb.create_sheet(title = 'Card of admission', index =1)
sheet = wb['Card of admission']
sheet.column_dimensions['A'].width = 30
sheet.column_dimensions['B'].width = 30
sheet.column_dimensions['C'].width = 30

value = "Name"
cell = sheet.cell(row = 1, column = 1)
cell.value = value

list()

value = "Status"
cell = sheet.cell(row = 1, column = 2)
cell.value = value

value = "Camps"
cell = sheet.cell(row = 1, column = 3)
cell.value = value

for i in range(len(people)):
sheet = wb['investments']
cell = sheet.cell(row = i+2, column =2)
if (cell.value == "yes"):
sheet = wb['Card of admission']
value = "accepted"
cell = sheet.cell(row = i+2, column =2)
cell.value = value
sheet = wb['investments']
cell = sheet.cell(row = i+2, column =6)
b=int(cell.value)

if (b>60 ):
sheet = wb['Card of admission']
value = "Artek"
cell = sheet.cell(row = i+2, column =3)
cell.value = value
else:
if (b >40):
sheet = wb['Card of admission']
value = "Ocean"
cell = sheet.cell(row = i+2, column =3)
cell.value = value
else:
if (b >10):
sheet = wb['Card of admission']
value = "Orlenok"
cell = sheet.cell(row = i+2, column =3)
cell.value = value
else:
sheet = wb['Card of admission']
value = "refusing"
cell = sheet.cell(row = i+2, column =2)
cell.value = value
value = "_"
cell = sheet.cell(row = i+2, column =3)
cell.value = value

wb.save('Getting card od admission.xlsx')
