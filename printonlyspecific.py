from openpyxl import load_workbook
wb=load_workbook("students.xlsx")
sheet=wb["Students"]
for row in range(4,8):
    print(sheet[f"B{row}"].value)