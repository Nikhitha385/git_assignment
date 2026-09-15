from openpyxl import Workbook

wb = Workbook()
sheet = wb.active

sheet.append(["Name", "Marks"])
sheet.append(["Divya", 95])
sheet.append(["Ravi", 81])
sheet.append(["Sita", 87])

wb.save("March.xlsx")

print("March file created")