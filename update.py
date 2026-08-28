from openpyxl import load_workbook
wb = load_workbook("students.xlsx")
sheet = wb["Students"]
sheet["B2"] = 95
wb.save("students.xlsx")
print("Marks updated")