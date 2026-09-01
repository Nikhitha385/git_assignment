from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# Open the Excel file
wb = load_workbook("students.xlsx")

# Select the Students sheet
sheet = wb["Students"]

# Header formatting
for cell in sheet[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(
        fill_type="solid",
        fgColor="4472C4"
    )
    cell.alignment = Alignment(horizontal="center")

# Add borders to the table
thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

for row in sheet.iter_rows(
    min_row=1,
    max_row=sheet.max_row,
    min_col=1,
    max_col=4
):
    for cell in row:
        cell.border = thin_border

# Set column widths
sheet.column_dimensions["A"].width = 15
sheet.column_dimensions["B"].width = 10
sheet.column_dimensions["C"].width = 10
sheet.column_dimensions["D"].width = 12

# Align data
for row in sheet.iter_rows(
    min_row=2,
    max_row=sheet.max_row,
    min_col=1,
    max_col=4
):
    for cell in row:
        cell.alignment = Alignment(horizontal="center")

# Save the formatted workbook
wb.save("students.xlsx")

print("Excel formatting completed")