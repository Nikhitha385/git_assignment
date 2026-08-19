from openpyxl import load_workbook
wb=load_workbook("students.xlsx")
sheet=wb.active
print(sheet["B6"].value)
print(sheet["C6"].value)
print(sheet["D6"].value)
print(sheet["E6"].value)