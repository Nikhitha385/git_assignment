from openpyxl import load_workbook
from openpyxl.styles import Font

wb = load_workbook("students.xlsx")
sheet = wb["Students"]

for cell in sheet[1]:
    cell.font = Font(
        bold=True,
        size=14,
        color="FFFFFF"
    )

wb.save("students.xlsx")

print("Font formatting completed")