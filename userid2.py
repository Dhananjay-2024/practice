import pandas as pd

# Load both sheets
sheet1 = pd.read_excel("file1.xlsx", sheet_name="ONSE")
sheet2 = pd.read_excel("file2.xlsx", sheet_name="Details")

# Merge (keeps only first match per account)
merged = sheet1.merge(
    sheet2[["Acc_No", "Balance"]],
    left_on="Account Number",
    right_on="Acc_No",
    how="left"
).drop(columns=["Acc_No"])

merged.to_excel("output_first_match.xlsx", index=False)
