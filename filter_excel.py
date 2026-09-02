from openpyxl import load_workbook

wb = load_workbook("students.xlsx")

sheet = wb["Students"]

# Add filter to the used range
sheet.auto_filter.ref = sheet.dimensions

wb.save("students.xlsx")

print("Filter added successfully")