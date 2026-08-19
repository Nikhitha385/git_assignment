from openpyxl import workbook,load_workbook
wb=load_workbook("students.xlsx")
sheet=wb.active
for cell in sheet[6][0:4]:
    print(cell.value)