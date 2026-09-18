import pandas as pd

data = pd.read_excel("students_new.xlsx")

sorted_data = data.sort_values("Marks")

print(sorted_data)
#Sort marks from highest to lowest
import pandas as pd

data = pd.read_excel("students_new.xlsx")

sorted_data = data.sort_values("Marks", ascending=False)

print(sorted_data)
# Sort by Name
import pandas as pd

data = pd.read_excel("students_new.xlsx")

sorted_data = data.sort_values("Name")

print(sorted_data)