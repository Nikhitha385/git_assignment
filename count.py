import pandas as pd

data = {
    "Name": ["Rahul", "Priya", "Arun", "Sneha"],
    "Department": ["IT", "HR", "IT", "HR"],
    "Marks": [95, 90, 82, 88]
}

df = pd.DataFrame(data)

grouped = df.groupby("Department")["Name"].count()

print(grouped)
#Highest mark
import pandas as pd

data = {
    "Name": ["Rahul", "Priya", "Arun", "Sneha"],
    "Department": ["IT", "HR", "IT", "HR"],
    "Marks": [95, 90, 82, 88]
}

df = pd.DataFrame(data)

grouped = df.groupby("Department")["Marks"].max()

print(grouped)
#Lowest marks
import pandas as pd

data = {
    "Name": ["Rahul", "Priya", "Arun", "Sneha"],
    "Department": ["IT", "HR", "IT", "HR"],
    "Marks": [95, 90, 82, 88]
}

df = pd.DataFrame(data)

grouped = df.groupby("Department")["Marks"].min()

print(grouped)