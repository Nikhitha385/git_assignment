from openpyxl import load_workbook     # openpyxl is a library
wb=load_workbook("students.xlsx")    # opens the excel file
sheet=wb.active                      # gets the active sheet
print(sheet["A2"].value)             # read data from excel file
print(sheet["B2"].value)
print(sheet["C2"].value)
 