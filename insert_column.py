from openpyxl import load_workbook
wb = load_workbook("students.xlsx")
sheet = wb["Students"]
sheet.insert_cols(2)
wb.save("students.xlsx")
print("Column inserted successfully")