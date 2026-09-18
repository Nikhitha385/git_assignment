import pandas as pd

data = pd.read_excel("dirty_data.xlsx")

print(data.isnull())
#fill missing values
import pandas as pd

data = pd.read_excel("dirty_data.xlsx")

data["Marks"] = data["Marks"].fillna(85)
data["Age"] = data["Age"].fillna(21)

print(data)
#Remove Missing Rows
import pandas as pd

data = pd.read_excel("dirty_data.xlsx")

cleaned_data = data.dropna()

print(cleaned_data)
import pandas as pd

data = pd.read_excel("dirty_data.xlsx")

cleaned_data = data.dropna()

print(cleaned_data)
#Remove duplicates
import pandas as pd

data = pd.read_excel("dirty_data.xlsx")

# Add a duplicate row for practice
data.loc[len(data)] = data.iloc[0]

# Remove duplicate rows
data = data.drop_duplicates()

# Save back to the same Excel file
data.to_excel("dirty_data.xlsx", index=False)

print("Duplicates removed and Excel file updated")