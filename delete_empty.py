from openpyxl import load_workbook

wb = load_workbook("students.xlsx")
sheet = wb["Students"]

# Delete empty row 3
sheet.delete_rows(3, 1)

# Delete empty columns B and C
sheet.delete_cols(2, 2)

wb.save("students.xlsx")

print("Empty rows and columns deleted successfully")