from openpyxl import load_workbook     # openpyxl is a library
wb=load_workbook("students.xlsx")    # opens the excel file
sheet=wb["Students"]                     # gets the active sheet
print(sheet["J28"].value)             # read data from excel file


