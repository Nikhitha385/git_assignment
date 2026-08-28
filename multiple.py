from openpyxl import load_workbook

wb = load_workbook("students.xlsx")
sheet = wb["Students"]

print(sheet["A4"].value)
print(sheet["B4"].value)
print(sheet["C4"].value)
print(sheet["D4"].value)