from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# Open workbook
wb = load_workbook("students.xlsx")

# Select Students sheet
students = wb["Students"]

# -----------------------------
# 1. Header formatting
# -----------------------------

header_fill = PatternFill(
    fill_type="solid",
    fgColor="4472C4"
)

header_font = Font(
    bold=True,
    color="FFFFFF"
)

for cell in students[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")

# -----------------------------
# 2. Borders
# -----------------------------

thin = Side(style="thin")

border = Border(
    left=thin,
    right=thin,
    top=thin,
    bottom=thin
)

for row in students.iter_rows(
    min_row=1,
    max_row=students.max_row,
    min_col=1,
    max_col=4
):
    for cell in row:
        cell.border = border

# -----------------------------
# 3. Alignment
# -----------------------------

for row in students.iter_rows(
    min_row=2,
    max_row=students.max_row,
    min_col=1,
    max_col=4
):
    for cell in row:
        cell.alignment = Alignment(horizontal="center")

# -----------------------------
# 4. Column widths
# -----------------------------

students.column_dimensions["A"].width = 15
students.column_dimensions["B"].width = 10
students.column_dimensions["C"].width = 10
students.column_dimensions["D"].width = 12

# -----------------------------
# 5. Format Summary sheet
# -----------------------------

if "Summary" in wb.sheetnames:

    summary = wb["Summary"]

    summary["A1"].font = Font(
        bold=True,
        color="FFFFFF",
        size=14
    )

    summary["A1"].fill = header_fill

    summary["A1"].alignment = Alignment(
        horizontal="center"
    )

    # Format summary rows
    for row in summary.iter_rows(
        min_row=1,
        max_row=5,
        min_col=1,
        max_col=2
    ):
        for cell in row:
            cell.border = border

    summary.column_dimensions["A"].width = 20
    summary.column_dimensions["B"].width = 15

    # Make labels bold
    for row in range(2, 6):
        summary[f"A{row}"].font = Font(bold=True)

    # Average to 2 decimal places
    summary["B3"].number_format = "0.00"

# -----------------------------
# 6. Save workbook
# -----------------------------

wb.save("students.xlsx")

print("Excel formatting completed")