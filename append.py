from openpyxl import load_workbook

wb = load_workbook("students.xlsx")

sheet = wb["Students"]

sheet.append(["Ravi", 91, 23])

wb.save("students.xlsx")
print("New student added")