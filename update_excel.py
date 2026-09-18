from openpyxl import load_workbook

# Open existing Excel file
wb = load_workbook("students_new.xlsx")

# Open the worksheet
sheet = wb.active

# Update Rahul's marks
sheet["B2"] = 95

# Save the updated file
wb.save("students_new.xlsx")

print("Excel file updated successfully")