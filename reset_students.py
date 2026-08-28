from openpyxl import Workbook

wb = Workbook()

sheet = wb.active
sheet.title = "Students"

data = [
    ["Name", "Marks", "AGE"],
    ["Nikhitha", 90, 22],
    ["Rahul", 85, 24],
    ["Priya", 95, 27],
    ["Kiran", 78, 21],
    ["Anu", 88, 19]
]

for row in data:
    sheet.append(row)

wb.save("students.xlsx")

print("Students Excel reset successfully")