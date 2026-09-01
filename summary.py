from openpyxl import load_workbook
wb = load_workbook("students.xlsx")
summary = wb.create_sheet("Summary")
summary["A1"] = "Student Summary"
wb.save("students.xlsx")
print("Summary sheet created")