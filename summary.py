from openpyxl import load_workbook

wb = load_workbook("students.xlsx")

students = wb["Students"]

# Use existing Summary sheet, otherwise create it
if "Summary" in wb.sheetnames:
    summary = wb["Summary"]
else:
    summary = wb.create_sheet("Summary")

summary["A1"] = "Student Summary"

# Total students
total_students = students.max_row - 1
summary["A2"] = "Total Students"
summary["B2"] = total_students

# Total marks
total_marks = 0

for row in range(2, students.max_row + 1):
    marks = students.cell(row=row, column=2).value
    total_marks += marks

# Average
average = total_marks / total_students

summary["A3"] = "Average Marks"
summary["B3"] = average

# Highest marks
highest = max(
    students.cell(row=row, column=2).value
    for row in range(2, students.max_row + 1)
)

summary["A4"] = "Highest Marks"
summary["B4"] = highest

# Lowest marks
lowest = min(
    students.cell(row=row, column=2).value
    for row in range(2, students.max_row + 1)
)

summary["A5"] = "Lowest Marks"
summary["B5"] = lowest

wb.save("students.xlsx")

print("Summary updated successfully")