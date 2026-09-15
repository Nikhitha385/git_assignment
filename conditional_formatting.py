from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

wb = load_workbook("students.xlsx")
sheet = wb["Students"]

yellow_fill = PatternFill(
    start_color="FFFF00",
    end_color="FFFF00",
    fill_type="solid"
)

rule = CellIsRule(
    operator="lessThan",
    formula=["80"],
    fill=yellow_fill
)

sheet.conditional_formatting.add("D2:D16", rule)

wb.save("students.xlsx")

print("Below 80 marks highlighted")