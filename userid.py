import pandas as pd

# Load both sheets
sheet1 = pd.read_excel("file1.xlsx", sheet_name="Sheet1")
sheet2 = pd.read_excel("file2.xlsx", sheet_name="Sheet2")

# Merge on account number (different column names)
merged = sheet1.merge(sheet2, left_on="Account Number", right_on="Acc_No", how="left")

# Save result
merged.to_excel("merged_output.xlsx", index=False)
