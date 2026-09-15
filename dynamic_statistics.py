from openpyxl import load_workbook

wb = load_workbook("students.xlsx")
sheet = wb["Students"]

# Dynamic Statistics

sheet["F2"] = "Total Marks"
sheet["G2"] = "=SUM(B2:B15)"
sheet.column_dimensions["G"].width = 18

sheet["F3"] = "Average Marks"
sheet["G3"] = "=AVERAGE(B2:B15)"

sheet["F4"] = "Highest Marks"
sheet["G4"] = "=MAX(B2:B15)"

sheet["F5"] = "Lowest Marks"
sheet["G5"] = "=MIN(B2:B15)"

wb.save("students.xlsx")

print("Dynamic statistics added successfully")