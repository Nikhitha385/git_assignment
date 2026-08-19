from openpyxl import load_workbook
wb = load_workbook("students.xlsx")
sheet = wb["Students"]
print(sheet["J28"].value)
print(sheet["K28"].value)
print(sheet["L28"].value)
print(sheet["M28"].value)