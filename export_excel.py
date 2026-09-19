import pandas as pd

data = {
    "Name": ["Rahul", "Priya", "Arun"],
    "Marks": [95, 90, 82]
}

df = pd.DataFrame(data)

df.to_excel("student_report.xlsx", index=False)

print("Excel file created successfully")


#Export a filtered DataFrame to Excel
import pandas as pd

data = {
    "Name": ["Rahul", "Priya", "Arun"],
    "Marks": [95, 90, 82]
}

df = pd.DataFrame(data)

filtered_data = df[df["Marks"] >= 90]

filtered_data.to_excel("top_students.xlsx", index=False)

print("Filtered data exported successfully")
#Export multiple columns to Excel
import pandas as pd

data = {
    "Name": ["Rahul", "Priya", "Arun"],
    "Department": ["IT", "HR", "IT"],
    "Marks": [95, 90, 82]
}

df = pd.DataFrame(data)

df.to_excel("student_details.xlsx", index=False)

print("Student details exported successfully")
#Export filtered data with selected columns
import pandas as pd

data = {
    "Name": ["Rahul", "Priya", "Arun"],
    "Department": ["IT", "HR", "IT"],
    "Marks": [95, 90, 82]
}

df = pd.DataFrame(data)

filtered_data = df[df["Marks"] >= 90]

selected_data = filtered_data[["Name", "Marks"]]

selected_data.to_excel("top_students.xlsx", index=False)

print("Selected data exported successfully")