from openpyxl import load_workbook

wb = load_workbook("students.xlsx")

sheet = wb["Students"]

students = [
    ["Sita", 87, 21],
    ["Arjun", 76, 22],
    ["Meena", 93, 20]
]

for student in students:
    sheet.append(student)

wb.save("students.xlsx")

print("Multiple students added")