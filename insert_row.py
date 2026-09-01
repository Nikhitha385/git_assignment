from openpyxl import load_workbook

wb = load_workbook("students.xlsx")

sheet = wb["Students"]

sheet.insert_rows(3)

wb.save("students.xlsx")

print("Row inserted successfully")