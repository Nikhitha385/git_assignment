from openpyxl import load_workbook
from openpyxl.formatting.rule import DataBarRule

wb = load_workbook("students.xlsx")
sheet = wb["Students"]

rule = DataBarRule(
    start_type="min",
    end_type="max",
    color="638EC6",
    showValue=True
)

sheet.conditional_formatting.add("D2:D16", rule)

wb.save("students.xlsx")

print("Data bars added successfully")