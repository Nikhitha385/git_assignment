from openpyxl import Workbook

# Old file
wb = Workbook()
sheet = wb.active

sheet.append(["Name", "Marks"])
sheet.append(["Rahul", 85])
sheet.append(["Priya", 90])
sheet.append(["Arun", 78])

wb.save("students_old.xlsx")


# New file
wb = Workbook()
sheet = wb.active

sheet.append(["Name", "Marks"])
sheet.append(["Rahul", 90])
sheet.append(["Priya", 90])
sheet.append(["Arun", 82])

wb.save("students_new.xlsx")

print("Comparison files created successfully")