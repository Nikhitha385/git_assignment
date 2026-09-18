from openpyxl import Workbook

wb = Workbook()
sheet = wb.active

sheet.append(["Name", "Marks", "Age"])
sheet.append(["Rahul", 95, 21])
sheet.append(["Priya", None, 22])
sheet.append(["Arun", 82, None])
sheet.append(["Sneha", 88, 21])

wb.save("dirty_data.xlsx")

print("dirty_data.xlsx created successfully")