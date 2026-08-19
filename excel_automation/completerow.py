from openpyxl import load_workbook
wb = load_workbook("students.xlsx")
sheet = wb["Students"]
for cell in sheet[4][0:4]:
    print(cell.value)