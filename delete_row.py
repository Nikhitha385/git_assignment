from openpyxl import load_workbook

wb = load_workbook("students.xlsx")

sheet = wb["Students"]

sheet.delete_rows(3)

wb.save("students.xlsx")

print("Row deleted successfully")