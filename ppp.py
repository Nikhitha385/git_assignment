import pandas as pd

data = {
    "Name": ["Rahul", "Priya", "Arun"],
    "Marks": [95, 90, 82]
}

df = pd.DataFrame(data)

df.to_excel("student_report.xlsx", sheet_name="Student Data", index=False)

print("Excel file created successfully")
#Read Excel using Pandas and modify it using OpenPyXL
import pandas as pd
from openpyxl import load_workbook

# Read Excel using Pandas
data = pd.read_excel("student_report.xlsx")

print(data)

# Open the same Excel file using OpenPyXL
wb = load_workbook("student_report.xlsx")
sheet = wb.active

# Add a title
sheet["D1"] = "Report"

# Save the file
wb.save("student_report.xlsx")

print("Excel file updated successfully")