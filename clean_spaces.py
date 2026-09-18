import pandas as pd

data = pd.read_excel("dirty_data.xlsx")

# Add unwanted spaces for practice
data.loc[0, "Name"] = " Rahul "
data.loc[1, "Name"] = " Priya "

# Remove unwanted spaces
data["Name"] = data["Name"].str.strip()

# Save to the same Excel file
data.to_excel("dirty_data.xlsx", index=False)

print("Unwanted spaces removed successfully")   
import pandas as pd

data = pd.read_excel("dirty_data.xlsx")

# Keep only valid marks
data = data[data["Marks"] <= 100]

# Save to the same Excel file
data.to_excel("dirty_data.xlsx", index=False)

print("Incorrect data removed successfully")