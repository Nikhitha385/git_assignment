import pandas as pd

data = pd.read_excel("students_new.xlsx")

print(data)
print(type(data))

import pandas as pd

data = pd.read_excel("students_new.xlsx")

filtered_data = data[data["Marks"] > 85]

print(filtered_data)


import pandas as pd

data = pd.read_excel("students_new.xlsx")

filtered_data = data[(data["Marks"] >= 80) & (data["Marks"] <= 90)]

print(filtered_data)

#using or condition
import pandas as pd

data = pd.read_excel("students_new.xlsx")

filtered_data = data[(data["Marks"] < 80) | (data["Marks"] > 90)]

print(filtered_data)
# filter using Specific data
import pandas as pd

data = pd.read_excel("students_new.xlsx")

filtered_data = data[data["Name"] == "Rahul"]

print(filtered_data)
#Filter using isin()
import pandas as pd

data = pd.read_excel("students_new.xlsx")

filtered_data = data[data["Name"].isin(["Rahul", "Arun"])]

print(filtered_data)