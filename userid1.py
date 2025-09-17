import pandas as pd

# Load sheets
sheet1 = pd.read_excel("file1.xlsx", sheet_name="Sheet1")
sheet2 = pd.read_excel("file2.xlsx", sheet_name="Sheet2")

# Bring only the needed column from sheet2
merged = sheet1.merge(
    sheet2[["Acc_No", "Balance"]],  # only keep key + target column
    left_on="Account Number", 
    right_on="Acc_No", 
    how="left"
)

# Drop duplicate key column if needed
merged = merged.drop(columns=["Acc_No"])

# Save result
merged.to_excel("output.xlsx", index=False)
